from __future__ import annotations

from contextlib import asynccontextmanager
from datetime import datetime, timedelta
import json
from pathlib import Path
from uuid import uuid4

from fastapi import BackgroundTasks, FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from .audit import AuditService
from .config import BotFaturasSettings
from .database import BotFaturasDatabase
from .exporter import ResultExporter
from .parser import SpreadsheetParser
from .queue import QueueManager
from .security import CredentialCipher
from .services import LineProcessingService, LotService
from .storage import StorageManager
from .tiflux_sheet_service import TifluxSheetService


class BotFaturasRuntime:
    def __init__(self) -> None:
        self.settings = BotFaturasSettings.from_env()
        self.database = BotFaturasDatabase(self.settings.database_path, self.settings.timezone_name)
        self.storage = StorageManager(self.settings.storage_root)
        self.parser = SpreadsheetParser()
        self.cipher = CredentialCipher.from_settings(self.settings)
        self.audit = AuditService(self.database)
        self.exporter = ResultExporter(self.storage)
        self.processing_service = LineProcessingService(
            settings=self.settings,
            database=self.database,
            storage=self.storage,
            cipher=self.cipher,
            audit=self.audit,
            exporter=self.exporter,
        )
        self.queue = QueueManager(
            settings=self.settings,
            database=self.database,
            processing_service=self.processing_service,
        )
        self.lot_service = LotService(
            settings=self.settings,
            database=self.database,
            storage=self.storage,
            parser=self.parser,
            cipher=self.cipher,
            audit=self.audit,
            processing_service=self.processing_service,
            queue_manager=self.queue,
        )
        self.tiflux_sheet_service = TifluxSheetService(
            browser_timeout_ms=self.settings.browser_timeout_ms,
            headless=self.settings.headless,
            output_dir=Path("output/tiflux"),
            session_file=Path("storage/tiflux/session.json"),
        )
        self.tiflux_jobs: dict[str, dict[str, object]] = {}
        self.tiflux_jobs_root = self.storage.root / "_tiflux_jobs"
        self.tiflux_jobs_root.mkdir(parents=True, exist_ok=True)
        self._load_persisted_tiflux_jobs()

    def tiflux_job_file(self, job_id: str) -> Path:
        return self.tiflux_jobs_root / f"{job_id}.json"

    def save_tiflux_job(self, payload: dict[str, object]) -> None:
        job_id = str(payload["job_id"])
        self.tiflux_jobs[job_id] = payload
        self.tiflux_job_file(job_id).write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def load_tiflux_job(self, job_id: str) -> dict[str, object] | None:
        cached = self.tiflux_jobs.get(job_id)
        if cached:
            return cached
        target = self.tiflux_job_file(job_id)
        if not target.exists():
            return None
        payload = json.loads(target.read_text(encoding="utf-8"))
        self.tiflux_jobs[job_id] = payload
        return payload

    def _load_persisted_tiflux_jobs(self) -> None:
        for target in sorted(self.tiflux_jobs_root.glob("*.json")):
            try:
                payload = json.loads(target.read_text(encoding="utf-8"))
                job_id = str(payload.get("job_id", ""))
                if job_id:
                    self.tiflux_jobs[job_id] = payload
            except Exception:
                continue


@asynccontextmanager
async def lifespan(app: FastAPI):
    runtime = BotFaturasRuntime()
    app.state.runtime = runtime
    await runtime.queue.start()
    yield
    await runtime.queue.stop()


app = FastAPI(
    title="BOT DE FATURAS API",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=BotFaturasSettings.from_env().cors_allow_origins or ["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def runtime_from_request(request: Request) -> BotFaturasRuntime:
    return request.app.state.runtime


class TifluxGoogleSheetRunRequest(BaseModel):
    spreadsheet_id: str = Field(..., min_length=3)
    worksheet_name: str = Field(default="Sheet1", min_length=1)


class TifluxGoogleSheetJobResponse(BaseModel):
    ok: bool
    job_id: str
    status: str
    spreadsheet_id: str
    worksheet_name: str
    message: str


class TifluxBatchRunRequest(BaseModel):
    tickets: list[str] = Field(default_factory=list)
    updates: dict[str, str] = Field(default_factory=dict)


class TifluxBatchJobResponse(BaseModel):
    ok: bool
    job_id: str
    status: str
    ticket_count: int
    message: str


@app.get("/health")
async def health(request: Request) -> dict[str, object]:
    runtime = runtime_from_request(request)
    return {
        "ok": True,
        "service": "bot-faturas-fastapi",
        "queue_size": runtime.queue.queue.qsize(),
        "workers": runtime.settings.worker_count,
    }


@app.get("/v1/tiflux/google-sheet/template")
async def tiflux_google_sheet_template(request: Request) -> dict[str, object]:
    runtime = runtime_from_request(request)
    return {
        "ok": True,
        "headers": runtime.tiflux_sheet_service.template_headers,
        "worksheet_default_name": "TIFLUX_PREENCHIMENTO",
    }


@app.get("/v1/tiflux/batch/template")
async def tiflux_batch_template(request: Request) -> dict[str, object]:
    runtime = runtime_from_request(request)
    return {
        "ok": True,
        "fields": runtime.tiflux_sheet_service.template_fields,
    }


@app.post("/v1/tiflux/google-sheet/run", response_model=TifluxGoogleSheetJobResponse)
async def tiflux_google_sheet_run(
    request: Request,
    payload: TifluxGoogleSheetRunRequest,
    background_tasks: BackgroundTasks,
) -> TifluxGoogleSheetJobResponse:
    runtime = runtime_from_request(request)
    job_id = f"tiflux-sheet-{uuid4().hex[:12]}"
    queued_job = {
        "ok": True,
        "job_id": job_id,
        "status": "queued",
        "spreadsheet_id": payload.spreadsheet_id,
        "worksheet_name": payload.worksheet_name,
        "message": "Processamento enfileirado.",
        "updated_at": _job_timestamp(),
    }
    runtime.save_tiflux_job(queued_job)
    runtime.tiflux_sheet_service.set_job_banner(
        payload.spreadsheet_id,
        payload.worksheet_name,
        job_id=job_id,
        status_text="Processamento enfileirado.",
    )
    background_tasks.add_task(
        _run_tiflux_google_sheet_job,
        runtime,
        job_id,
        payload.spreadsheet_id,
        payload.worksheet_name,
    )
    return TifluxGoogleSheetJobResponse(
        ok=True,
        job_id=job_id,
        status="queued",
        spreadsheet_id=payload.spreadsheet_id,
        worksheet_name=payload.worksheet_name,
        message="Processamento da planilha iniciado.",
    )


@app.get("/v1/tiflux/google-sheet/jobs/{job_id}")
async def tiflux_google_sheet_job_status(request: Request, job_id: str) -> dict[str, object]:
    runtime = runtime_from_request(request)
    job = runtime.load_tiflux_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job do TiFlux nao encontrado.")
    job = _fail_stale_tiflux_job(runtime, job)
    return job


@app.post("/v1/tiflux/batch/run", response_model=TifluxBatchJobResponse)
async def tiflux_batch_run(
    request: Request,
    payload: TifluxBatchRunRequest,
    background_tasks: BackgroundTasks,
) -> TifluxBatchJobResponse:
    runtime = runtime_from_request(request)
    ticket_count = len([ticket for ticket in payload.tickets if str(ticket).strip()])
    if ticket_count == 0:
        raise HTTPException(status_code=400, detail="Informe ao menos um ticket para atualizar no TiFlux.")
    job_id = f"tiflux-batch-{uuid4().hex[:12]}"
    queued_job = {
        "ok": True,
        "job_id": job_id,
        "status": "queued",
        "ticket_count": ticket_count,
        "message": "Lote do TiFlux enfileirado.",
        "updated_at": _job_timestamp(),
        "updates": payload.updates,
        "tickets": payload.tickets,
    }
    runtime.save_tiflux_job(queued_job)
    background_tasks.add_task(
        _run_tiflux_batch_job,
        runtime,
        job_id,
        payload.tickets,
        payload.updates,
    )
    return TifluxBatchJobResponse(
        ok=True,
        job_id=job_id,
        status="queued",
        ticket_count=ticket_count,
        message="Processamento do lote TiFlux iniciado.",
    )


@app.post("/v1/lots/upload")
async def create_lot(request: Request, file: UploadFile = File(...)) -> dict[str, object]:
    runtime = runtime_from_request(request)
    try:
        content = await file.read()
        payload = await runtime.lot_service.create_lot_from_upload(file_name=file.filename or "upload.xlsx", content=content)
        return _enrich_lot_payload(payload)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.get("/v1/lots")
async def list_lots(request: Request, limit: int = Query(default=30, ge=1, le=200)) -> dict[str, object]:
    runtime = runtime_from_request(request)
    lots = [_enrich_lot(runtime, item) for item in runtime.lot_service.list_history(limit=limit)]
    return {"ok": True, "lots": lots}


@app.get("/v1/lots/{lot_id}")
async def get_lot(request: Request, lot_id: str) -> dict[str, object]:
    runtime = runtime_from_request(request)
    try:
        payload = runtime.lot_service.get_lot_payload(lot_id)
        return _enrich_lot_payload(payload)
    except ValueError as error:
        raise HTTPException(status_code=404, detail=str(error)) from error


@app.get("/v1/lines/{line_id}")
async def get_line(request: Request, line_id: str) -> dict[str, object]:
    runtime = runtime_from_request(request)
    try:
        detail = runtime.lot_service.get_line_detail(line_id)
        line = detail["line"]
        detail["line"] = _enrich_line(runtime, line)
        detail["log_preview"] = _read_log_preview(line.get("log_execucao", ""))
        return {"ok": True, **detail}
    except ValueError as error:
        raise HTTPException(status_code=404, detail=str(error)) from error


@app.post("/v1/lots/{lot_id}/reprocess")
async def reprocess_lot(request: Request, lot_id: str) -> dict[str, object]:
    runtime = runtime_from_request(request)
    try:
        payload = await runtime.lot_service.reprocess_lot(lot_id)
        return _enrich_lot_payload(payload)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/v1/lines/{line_id}/reprocess")
async def reprocess_line(request: Request, line_id: str) -> dict[str, object]:
    runtime = runtime_from_request(request)
    try:
        payload = await runtime.lot_service.reprocess_line(line_id)
        line = payload["line"]
        payload["line"] = _enrich_line(runtime, line)
        payload["log_preview"] = _read_log_preview(line.get("log_execucao", ""))
        return {"ok": True, **payload}
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.get("/v1/lots/{lot_id}/download/report")
async def download_lot_report(request: Request, lot_id: str):
    runtime = runtime_from_request(request)
    lot = runtime.database.get_lot(lot_id)
    if not lot or not lot.get("report_path"):
        raise HTTPException(status_code=404, detail="Planilha de retorno nao encontrada para este lote.")
    return FileResponse(path=lot["report_path"], filename=Path(lot["report_path"]).name)


@app.get("/v1/lots/{lot_id}/download/archive")
async def download_lot_archive(request: Request, lot_id: str):
    runtime = runtime_from_request(request)
    lot = runtime.database.get_lot(lot_id)
    if not lot or not lot.get("archive_path"):
        raise HTTPException(status_code=404, detail="Arquivo ZIP do lote nao encontrado.")
    return FileResponse(path=lot["archive_path"], filename=Path(lot["archive_path"]).name)


@app.get("/v1/lots/{lot_id}/download/log")
async def download_lot_log(request: Request, lot_id: str):
    runtime = runtime_from_request(request)
    lot = runtime.database.get_lot(lot_id)
    if not lot or not lot.get("consolidated_log_path"):
        raise HTTPException(status_code=404, detail="Log consolidado do lote nao encontrado.")
    return FileResponse(path=lot["consolidated_log_path"], filename=Path(lot["consolidated_log_path"]).name)


@app.get("/v1/lines/{line_id}/file/{kind}")
async def download_line_file(request: Request, line_id: str, kind: str):
    runtime = runtime_from_request(request)
    line = runtime.database.get_line(line_id)
    if not line:
        raise HTTPException(status_code=404, detail="Linha nao encontrada.")

    file_path = ""
    if kind == "pdf":
        file_path = str(line.get("arquivo_pdf_path") or "")
    elif kind == "ae":
        file_path = str(line.get("arquivo_ae_path") or "")
    elif kind == "screenshot":
        file_path = str(line.get("screenshot_erro_path") or "")
    elif kind == "html":
        file_path = str(line.get("html_erro_path") or "")
    elif kind == "log":
        file_path = str(line.get("log_execucao") or "")
    else:
        raise HTTPException(status_code=400, detail="Tipo de arquivo invalido.")

    if not file_path:
        raise HTTPException(status_code=404, detail="Arquivo solicitado nao esta disponivel para esta linha.")
    return FileResponse(path=file_path, filename=Path(file_path).name)


def _enrich_lot_payload(payload: dict[str, object]) -> dict[str, object]:
    lot = payload["lot"]
    lines = payload["lines"]
    runtime = None
    return {"ok": True, "lot": _enrich_lot(runtime, lot), "lines": [_enrich_line(runtime, line) for line in lines]}


def _enrich_lot(_runtime: BotFaturasRuntime | None, lot: dict[str, object]) -> dict[str, object]:
    enriched = dict(lot)
    lot_id = str(lot.get("id") or "")
    enriched["download_urls"] = {
        "report": f"/v1/lots/{lot_id}/download/report" if lot.get("report_path") else "",
        "archive": f"/v1/lots/{lot_id}/download/archive" if lot.get("archive_path") else "",
        "log": f"/v1/lots/{lot_id}/download/log" if lot.get("consolidated_log_path") else "",
    }
    return enriched


def _enrich_line(_runtime: BotFaturasRuntime | None, line: dict[str, object]) -> dict[str, object]:
    enriched = dict(line)
    line_id = str(line.get("id") or "")
    enriched["download_urls"] = {
        "pdf": f"/v1/lines/{line_id}/file/pdf" if line.get("arquivo_pdf_path") else "",
        "ae": f"/v1/lines/{line_id}/file/ae" if line.get("arquivo_ae_path") else "",
        "screenshot": f"/v1/lines/{line_id}/file/screenshot" if line.get("screenshot_erro_path") else "",
        "html": f"/v1/lines/{line_id}/file/html" if line.get("html_erro_path") else "",
        "log": f"/v1/lines/{line_id}/file/log" if line.get("log_execucao") else "",
    }
    return enriched


def _read_log_preview(file_path: str, limit: int = 20) -> list[str]:
    if not file_path:
        return []
    path = Path(file_path)
    if not path.exists():
        return []
    return path.read_text(encoding="utf-8").splitlines()[-limit:]


def _run_tiflux_google_sheet_job(
    runtime: BotFaturasRuntime,
    job_id: str,
    spreadsheet_id: str,
    worksheet_name: str,
) -> None:
    running_job = {
        "ok": True,
        "job_id": job_id,
        "status": "running",
        "spreadsheet_id": spreadsheet_id,
        "worksheet_name": worksheet_name,
        "message": "Processamento em andamento.",
        "updated_at": _job_timestamp(),
    }
    runtime.save_tiflux_job(running_job)
    runtime.tiflux_sheet_service.set_job_banner(
        spreadsheet_id,
        worksheet_name,
        job_id=job_id,
        status_text="Processamento em andamento.",
    )
    try:
        result = runtime.tiflux_sheet_service.process_google_sheet(
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet_name,
        )
        completed_job = {
            "ok": True,
            "job_id": job_id,
            "status": "completed",
            "spreadsheet_id": spreadsheet_id,
            "worksheet_name": worksheet_name,
            "updated_at": _job_timestamp(),
            **result,
        }
        runtime.save_tiflux_job(completed_job)
        runtime.tiflux_sheet_service.set_job_banner(
            spreadsheet_id,
            worksheet_name,
            job_id=job_id,
            status_text="Processamento conclu?do.",
            reset_checkbox=True,
        )
    except Exception as error:  # noqa: BLE001
        failed_job = {
            "ok": False,
            "job_id": job_id,
            "status": "failed",
            "spreadsheet_id": spreadsheet_id,
            "worksheet_name": worksheet_name,
            "message": str(error),
            "updated_at": _job_timestamp(),
        }
        runtime.save_tiflux_job(failed_job)
        runtime.tiflux_sheet_service.set_job_banner(
            spreadsheet_id,
            worksheet_name,
            job_id=job_id,
            status_text=f"Falha: {error}",
            reset_checkbox=True,
        )


def _run_tiflux_batch_job(
    runtime: BotFaturasRuntime,
    job_id: str,
    tickets: list[str],
    updates: dict[str, str],
) -> None:
    ticket_count = len([ticket for ticket in tickets if str(ticket).strip()])
    running_job = {
        "ok": True,
        "job_id": job_id,
        "status": "running",
        "ticket_count": ticket_count,
        "message": "Processamento do lote TiFlux em andamento.",
        "updated_at": _job_timestamp(),
        "updates": updates,
        "tickets": tickets,
    }
    runtime.save_tiflux_job(running_job)

    def save_progress(summary: list[dict[str, object]], total: int) -> None:
        runtime.save_tiflux_job(
            {
                "ok": True,
                "job_id": job_id,
                "status": "running",
                "ticket_count": total,
                "processed": len(summary),
                "updated": sum(1 for item in summary if item.get("status") == "OK"),
                "failed": sum(1 for item in summary if item.get("status") in {"ERRO", "ALERTA"}),
                "message": f"Processando lote TiFlux: {len(summary)} de {total} ticket(s).",
                "updated_at": _job_timestamp(),
                "updates": updates,
                "tickets": tickets,
                "results": summary,
            }
        )

    try:
        result = runtime.tiflux_sheet_service.process_ticket_batch(
            tickets=tickets,
            raw_updates=updates,
            progress_callback=save_progress,
        )
        completed_job = {
            "ok": True,
            "job_id": job_id,
            "status": "completed",
            "ticket_count": len(result.get("tickets", [])),
            "updated_at": _job_timestamp(),
            "message": "Lote TiFlux concluído.",
            **result,
        }
        runtime.save_tiflux_job(completed_job)
    except Exception as error:  # noqa: BLE001
        failed_job = {
            "ok": False,
            "job_id": job_id,
            "status": "failed",
            "ticket_count": len([ticket for ticket in tickets if str(ticket).strip()]),
            "message": str(error),
            "updated_at": _job_timestamp(),
            "updates": updates,
            "tickets": tickets,
        }
        runtime.save_tiflux_job(failed_job)


def _job_timestamp() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _fail_stale_tiflux_job(runtime: BotFaturasRuntime, job: dict[str, object]) -> dict[str, object]:
    if str(job.get("status") or "") not in {"queued", "running"}:
        return job
    updated_at = str(job.get("updated_at") or "")
    try:
        updated = datetime.fromisoformat(updated_at)
    except ValueError:
        return job
    # Jobs created before the latest deploy may have been timestamped in UTC while
    # the app now runs in the configured local timezone, so check both clocks.
    possible_ages = (datetime.now() - updated, datetime.utcnow() - updated)
    if max(possible_ages) < timedelta(minutes=12):
        return job

    failed_job = {
        **job,
        "ok": False,
        "status": "failed",
        "message": (
            "Job ficou sem retorno por mais de 12 minutos. Provavel autenticacao do TiFlux "
            "pendente/codigo de verificacao ou navegador travado antes do primeiro ticket."
        ),
        "updated_at": _job_timestamp(),
    }
    runtime.save_tiflux_job(failed_job)
    return failed_job

from __future__ import annotations

import asyncio

from playwright.async_api import async_playwright

from .constants import LineStatus


class QueueManager:
    def __init__(self, *, settings, database, processing_service) -> None:
        self.settings = settings
        self.database = database
        self.processing_service = processing_service
        self.queue: asyncio.Queue[str] = asyncio.Queue()
        self._workers: list[asyncio.Task] = []
        self._playwright = None
        self._browser = None
        self._started = False

    async def start(self) -> None:
        if self._started:
            return
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(headless=self.settings.headless)
        self._workers = [
            asyncio.create_task(self._worker_loop(index), name=f"bot-faturas-worker-{index}")
            for index in range(self.settings.worker_count)
        ]
        self._started = True
        pending = self.database.list_lines_by_statuses([str(LineStatus.NA_FILA), str(LineStatus.EM_PROCESSAMENTO)])
        for line in pending:
            await self.enqueue(str(line.get("id")))

    async def stop(self) -> None:
        for worker in self._workers:
            worker.cancel()
        for worker in self._workers:
            try:
                await worker
            except asyncio.CancelledError:
                pass
        self._workers.clear()
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        self._browser = None
        self._playwright = None
        self._started = False

    async def enqueue(self, line_id: str) -> None:
        await self.queue.put(line_id)

    async def enqueue_many(self, line_ids: list[str]) -> None:
        for line_id in line_ids:
            await self.enqueue(line_id)

    async def _worker_loop(self, index: int) -> None:
        while True:
            line_id = await self.queue.get()
            try:
                await self.processing_service.process_line(self._browser, line_id)
            finally:
                self.queue.task_done()

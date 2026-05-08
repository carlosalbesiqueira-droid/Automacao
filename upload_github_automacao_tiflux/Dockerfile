FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    NODE_ENV=production \
    NPM_CONFIG_UPDATE_NOTIFIER=false \
    PORT=3210 \
    BOT_FATURAS_API_HOST=127.0.0.1 \
    BOT_FATURAS_API_PORT=8321 \
    BOT_FATURAS_API_BASE=http://127.0.0.1:8321

RUN apt-get update \
    && apt-get install -y --no-install-recommends curl ca-certificates gnupg bash \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && python -m pip install --upgrade pip \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY requirements-automation.txt requirements-automation-lock.txt ./
RUN pip install --no-cache-dir -r requirements-automation.txt \
    && python -m playwright install --with-deps chromium

COPY . .

RUN chmod +x scripts/start_railway.sh

EXPOSE 3210

CMD ["bash", "./scripts/start_railway.sh"]

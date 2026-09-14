FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

# Dependencias del sistema para Chromium/Playwright (unión de lo que pedían
# los Dockerfiles individuales de Bolívar/Previsora/Mundial)
RUN apt-get update && apt-get install -y \
    libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 libcups2 libdrm2 \
    libxkbcommon0 libxcomposite1 libxdamage1 libxfixes3 libxrandr2 libgbm1 \
    libpango-1.0-0 libcairo2 libasound2 libatspi2.0-0 libxshmfence1 \
    libx11-6 libxcb1 libxext6 libdbus-1-3 libglib2.0-0 libwayland-client0 \
    fonts-liberation fonts-unifont libfontconfig1 libfreetype6 \
    wget curl gnupg ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Chromium de Playwright (dependencias de sistema ya instaladas arriba)
RUN playwright install chromium

COPY . .

RUN mkdir -p downloads/estado downloads/sura downloads/bolivar downloads/previsora downloads/mundial

ENV PORT=8080
EXPOSE 8080

# Un solo worker (cada bot guarda su estado en memoria del proceso) con
# varios hilos para servir el polling concurrente de las 5 páginas sin
# bloquearse entre sí. 24 hilos porque las automatizaciones son de espera
# (I/O: navegación y descargas), no de cómputo puro — el plan real tiene
# 8 vCPU / 8 GB de margen de sobra para sostenerlos.
CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:${PORT:-8080} --workers 1 --threads 24 --timeout 600 --keep-alive 5"]

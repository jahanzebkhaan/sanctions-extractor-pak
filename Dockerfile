# ── Base image ────────────────────────────────────────────────────────────────
FROM python:3.11-slim

# ── System deps for Playwright / Chromium ─────────────────────────────────────
RUN apt-get update && apt-get install -y \
    wget curl gnupg ca-certificates \
    libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libxkbcommon0 libxcomposite1 \
    libxdamage1 libxfixes3 libxrandr2 libgbm1 libasound2 \
    libpango-1.0-0 libcairo2 libatspi2.0-0 libwayland-client0 \
    fonts-liberation libappindicator3-1 xdg-utils \
    && rm -rf /var/lib/apt/lists/*

# ── App directory ─────────────────────────────────────────────────────────────
WORKDIR /app

# ── Python deps ───────────────────────────────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── Install Playwright Chromium ───────────────────────────────────────────────
RUN playwright install chromium --with-deps

# ── Copy app files ────────────────────────────────────────────────────────────
COPY app.py .
COPY sanctions_extractor.py .

# ── Create cache directory ────────────────────────────────────────────────────
RUN mkdir -p /app/sanctions_cache

# ── Run with gunicorn ─────────────────────────────────────────────────────────
EXPOSE 8080
ENV PORT=8080

CMD ["gunicorn", "app:app", \
     "--bind", "0.0.0.0:8080", \
     "--workers", "1", \
     "--threads", "4", \
     "--timeout", "300", \
     "--keep-alive", "5"]

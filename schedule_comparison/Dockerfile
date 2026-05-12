FROM python:3.11-slim

WORKDIR /app

# Install dependencies first (layer cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Cache dir inside container (override via -v for persistence)
RUN mkdir -p /tmp/sce_cache

EXPOSE 5000

ENV PYTHONUNBUFFERED=1 \
    SCE_CACHE_DIR=/tmp/sce_cache \
    SCE_UPLOAD_MAX_MB=50 \
    PORT=5000

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "--access-logfile", "-", "app:app"]

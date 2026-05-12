FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /tmp/sce_cache

EXPOSE 7860

ENV PYTHONUNBUFFERED=1 \
    SCE_CACHE_DIR=/tmp/sce_cache \
    SCE_UPLOAD_MAX_MB=50 \
    PORT=7860

CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "2", "--timeout", "120", "--access-logfile", "-", "app:app"]

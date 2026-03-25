FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Preview and PDF generation depend on LibreOffice's soffice binary.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-impress \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY mss_ai_ppt_sample_assets/backend/requirements.txt /tmp/requirements.txt
RUN pip install -r /tmp/requirements.txt

COPY mss_ai_ppt_sample_assets /app/mss_ai_ppt_sample_assets

EXPOSE 8000

CMD ["python", "-m", "uvicorn", "mss_ai_ppt_sample_assets.backend.app:app", "--host", "0.0.0.0", "--port", "8000"]

FROM python:3.11-slim

# Install LibreOffice (needed for DOCX → PDF on Linux)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App files
COPY cloud_app.py .
COPY "DN_INVOICE _FORMAT.docx" .

# LibreOffice profile dir (must be writable at runtime)
RUN mkdir -p /tmp/lo_profile && chmod 777 /tmp/lo_profile

EXPOSE 8000

CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "2", "--timeout", "300", "cloud_app:app"]

FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    HOME=/tmp \
    PORT=8000

WORKDIR /app

# Прил.7 / расстановка: headless LibreOffice (soffice) на Linux
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-writer \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY pyproject.toml requirements.txt README.md ./
COPY sk_reporter ./sk_reporter
COPY webapp ./webapp
COPY data ./data

RUN pip install --no-cache-dir -e .

WORKDIR /app/webapp
EXPOSE 8000

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]

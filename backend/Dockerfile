# Backend API image. Build context is the REPO ROOT so it can copy backend/ and
# the shared gunicorn config:  docker build -f backend/Dockerfile .
FROM python:3.11-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# LibreOffice (headless) + fonts so page_locator can render DOCX -> PDF in-process
# and read REAL page numbers — no separate Gotenberg service needed. If you prefer
# the lighter Gotenberg-service route, you can drop this block and set GOTENBERG_URL.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-calc \
        fonts-dejavu fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY backend/requirements.txt ./backend/requirements.txt
RUN pip install -r backend/requirements.txt

COPY backend/ ./backend/
COPY gunicorn.conf.py ./gunicorn.conf.py

EXPOSE 10000

# gunicorn.conf.py sets chdir=./backend and bind 0.0.0.0:10000.
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]

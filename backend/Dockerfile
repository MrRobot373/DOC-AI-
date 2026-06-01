# Backend API image. Build context is the REPO ROOT so it can copy backend/ and
# the shared gunicorn config:  docker build -f backend/Dockerfile .
FROM python:3.11-slim

WORKDIR /app

# System deps kept minimal; PDF page-rendering is done by the Gotenberg service,
# not a local LibreOffice, so no heavy packages are needed here.
ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

COPY backend/requirements.txt ./backend/requirements.txt
RUN pip install -r backend/requirements.txt

COPY backend/ ./backend/
COPY gunicorn.conf.py ./gunicorn.conf.py

EXPOSE 10000

# gunicorn.conf.py sets chdir=./backend and bind 0.0.0.0:10000.
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]

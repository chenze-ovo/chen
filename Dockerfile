FROM python:3.12-slim

WORKDIR /app

COPY backend/requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt

COPY backend/app /app/app
COPY engines /app/engines
COPY frontend/dist /app/frontend/dist

RUN mkdir -p /app/data

ENV ERP_REPO_ROOT=/app
ENV ERP_DATA_DIR=/app/data
ENV ERP_DIST_DIR=/app/frontend/dist
ENV ERP_UI_MODE=web
ENV PYTHONUNBUFFERED=1
ENV PORT=7860

EXPOSE 7860

CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-7860}"]

FROM python:3.11-slim

WORKDIR /app
COPY . /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOST=0.0.0.0 \
    PORT=8000

EXPOSE 8000

CMD ["python", "web_server.py"]

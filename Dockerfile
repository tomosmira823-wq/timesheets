FROM python:3.12-slim
# tesseract: optional OCR for scanned receipts/bills (the app works without it)
RUN apt-get update && apt-get install -y --no-install-recommends tesseract-ocr tzdata && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY web/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY timesheet_app.py /app/timesheet_app.py
COPY web/server.py /app/web/server.py
COPY timesheets.png /app/timesheets.png
ENV DATA_ROOT=/data SECURE_COOKIES=1 PORT=8000 TZ=Europe/London
VOLUME ["/data"]
EXPOSE 8000
CMD ["sh","-c","uvicorn web.server:app --host 0.0.0.0 --port ${PORT}"]

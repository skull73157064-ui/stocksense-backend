# StockSense 抽圖後端 Dockerfile
# 基於 Python 3.11 slim,裝 LibreOffice 用來轉 .xls -> .xlsx
FROM python:3.11-slim

# 裝 LibreOffice 與必要字型(支援中文)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-calc \
    fonts-noto-cjk \
    fonts-wqy-zenhei \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 先複製 requirements 再 pip install,讓 Docker 能快取這層
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

# Render 會提供 PORT 環境變數,預設 10000
ENV PORT=10000
EXPOSE 10000

CMD uvicorn main:app --host 0.0.0.0 --port ${PORT}

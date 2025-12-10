# 使用 Python 3.11 作為基礎映像
FROM python:3.11-slim

# 設定工作目錄
WORKDIR /app

# 複製依賴檔案
COPY requirements.txt .

# 安裝依賴
RUN pip install --no-cache-dir -r requirements.txt gunicorn

# 複製程式碼
COPY . .

# 暴露端口
EXPOSE 5000

# 啟動應用（使用 gunicorn 作為生產環境伺服器）
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]


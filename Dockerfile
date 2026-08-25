FROM python:3.11-slim

WORKDIR /app

# Cài dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ source code
COPY . .

# Hugging Face Spaces dùng port 7860
EXPOSE 7860

# Set CLOUD_MODE tự động khi chạy trên HF
ENV HF_SPACE=1

# Render và các Cloud Host tự cấp cổng qua biến $PORT
CMD gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --timeout 120 server:app

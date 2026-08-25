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

# Chạy bằng gunicorn (production WSGI server)
CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "1", "--timeout", "120", "--preload", "server:app"]

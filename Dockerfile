FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 80
CMD ["gunicorn", "--bind", "0.0.0.0:80", "--timeout", "600", "--workers", "2", "--threads", "4", "app:app"]
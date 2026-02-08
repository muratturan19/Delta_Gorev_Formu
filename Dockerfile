FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create data directories
RUN mkdir -p /app/data /app/uploads

# Environment defaults (override via docker-compose or env)
ENV DEV_MODE=0
ENV DATA_FOLDER=/app/data
ENV UPLOAD_FOLDER=/app/uploads
ENV PORT=5002

EXPOSE 5002

CMD ["gunicorn", "--bind", "0.0.0.0:5002", "--workers", "2", "--timeout", "120", "web_app:app"]

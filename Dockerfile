FROM python:3.11-slim

WORKDIR /app

# Copy only backend files
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/ .

# Run the application (Railway provides PORT env var)
CMD gunicorn bon_a_envoye:app --host 0.0.0.0 --port $PORT 
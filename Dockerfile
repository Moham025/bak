FROM python:3.11-slim

WORKDIR /app

# Copy only backend files
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/ .

# Set default port
ENV PORT=8080

# Run the application with debug logging
CMD ["sh", "-c", "echo 'Starting server on port $PORT' && gunicorn bon_a_envoye:app --bind 0.0.0.0:$PORT --log-level info --access-logfile -"] 
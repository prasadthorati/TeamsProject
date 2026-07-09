# --- Teams Message Sender: Cloud Run image ---
FROM python:3.12-slim

# Prevents .pyc files and forces stdout/stderr to be unbuffered so
# print()/traceback.print_exc() show up immediately in Cloud Run Logs.
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Install dependencies first so Docker can cache this layer between builds.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code.
COPY teamspython.py .

# Run as a non-root user (Cloud Run best practice). Uploaded files go to
# /tmp, which stays writable for any user, so this does not break uploads.
RUN useradd --create-home appuser
USER appuser

# Cloud Run injects PORT (defaults to 8080) and expects the container to
# listen on 0.0.0.0:$PORT.
ENV PORT=8080
EXPOSE 8080

# workers=1 keeps the in-memory token/Excel state consistent (see note
# below); threads=8 lets it still handle several requests concurrently.
# timeout=0 disables gunicorn's worker timeout so a long "Send Actual
# Messages" run (many chats + delay-per-message) is not killed mid-way.
CMD ["sh", "-c", "exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 teamspython:app"]

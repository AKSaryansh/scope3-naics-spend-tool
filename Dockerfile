# Streamlit deployment container (works for Cloud Run / any container platform)
FROM python:3.12-slim

WORKDIR /app

# System deps (optional but useful for troubleshooting/healthcheck)
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
  && rm -rf /var/lib/apt/lists/*

# Install Python deps
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copy app
COPY . /app

# Streamlit default port is 8501, but Cloud Run expects $PORT (usually 8080)
ENV PORT=8080
EXPOSE 8080

HEALTHCHECK CMD curl --fail http://localhost:${PORT}/_stcore/health || exit 1

CMD ["sh", "-c", "streamlit run app.py --server.address=0.0.0.0 --server.port=${PORT}"]

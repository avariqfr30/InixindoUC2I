import os

from waitress import serve

from app import app


host = os.getenv("HOST", "0.0.0.0")
port = int(os.getenv("PORT", "8000"))
threads = int(os.getenv("PILOT_THREADS", "8"))
connection_limit = int(os.getenv("PILOT_CONNECTION_LIMIT", "100"))
channel_timeout = int(os.getenv("PILOT_CHANNEL_TIMEOUT", "240"))

serve(
    app,
    host=host,
    port=port,
    threads=threads,
    connection_limit=connection_limit,
    channel_timeout=channel_timeout,
)

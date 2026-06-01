"""
RQ worker process for DOC-AI.

Runs review jobs enqueued by the web app onto the Redis 'docai' queue. Crash-safe:
if the worker dies mid-review, the job is requeued; review state lives in the
shared JSON store / Supabase, so progress polling keeps working.

Start with:  python worker.py     (or:  python -m rq worker docai)
Requires REDIS_URL (default redis://localhost:6379).
"""

import os

# Load .env so the worker shares the web app's config (Supabase, Gotenberg, etc.)
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))
except ImportError:
    pass

from redis import Redis
from rq import Queue, Worker

REDIS_URL = os.environ.get("REDIS_URL", "redis://localhost:6379")


def main():
    conn = Redis.from_url(REDIS_URL)
    queue = Queue("docai", connection=conn)
    print(f"[worker] connected to {REDIS_URL}, listening on 'docai'")
    Worker([queue], connection=conn).work(with_scheduler=True)


if __name__ == "__main__":
    main()

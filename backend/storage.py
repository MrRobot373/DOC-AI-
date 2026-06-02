"""
Shared object storage for uploads and generated reports.

When Supabase is configured, files are stored in a Supabase Storage bucket so the
web service and a separate worker (different containers / disks) can both reach
them — the prerequisite for the durable Redis+RQ queue and for auto-fix surviving
restarts. Falls back to the local disk when Supabase Storage isn't available, so
single-instance / local-dev keeps working unchanged.

Public API:
    init(supabase_client)
    put_upload(local_path, key) -> key
    put_report(local_path, key) -> key
    fetch_to(key, dest_path, bucket="uploads"|"reports") -> dest_path | None
    ensure_local(key, dest_dir, bucket) -> local_path | None
"""

from __future__ import annotations

import os

UPLOADS_BUCKET = "docai-uploads"
REPORTS_BUCKET = "docai-reports"

_supabase = None


def init(supabase_client):
    """Wire in the Supabase client (called once from app startup). Creates buckets."""
    global _supabase
    _supabase = supabase_client
    if not _supabase:
        return
    for bucket in (UPLOADS_BUCKET, REPORTS_BUCKET):
        try:
            _supabase.storage.create_bucket(bucket, options={"public": False})
        except Exception:
            pass  # already exists (or insufficient perms) — fine


def enabled():
    return _supabase is not None


def _upload(bucket, key, local_path, content_type):
    if not _supabase or not os.path.exists(local_path):
        return None
    try:
        with open(local_path, "rb") as fh:
            data = fh.read()
        # upsert so re-runs / regenerated reports overwrite cleanly
        _supabase.storage.from_(bucket).upload(
            key, data, {"content-type": content_type, "upsert": "true"}
        )
        return key
    except Exception as e:
        print(f"[storage] upload {bucket}/{key} failed: {e}")
        return None


def put_upload(local_path, key):
    return _upload(UPLOADS_BUCKET, key,
                   local_path, "application/octet-stream") or key


def put_report(local_path, key):
    ct = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          if key.lower().endswith(".xlsx") else "application/octet-stream")
    return _upload(REPORTS_BUCKET, key, local_path, ct) or key


def fetch_to(key, dest_path, bucket):
    """Download an object to dest_path. Returns dest_path on success, else None."""
    if not _supabase:
        return None
    try:
        data = _supabase.storage.from_(bucket).download(key)
        if not data:
            return None
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        with open(dest_path, "wb") as fh:
            fh.write(data)
        return dest_path
    except Exception as e:
        print(f"[storage] download {bucket}/{key} failed: {e}")
        return None


def ensure_local(key, dest_dir, bucket):
    """
    Return a local path for `key`: use the on-disk copy if present, otherwise
    pull it from Supabase Storage into dest_dir. None if unavailable anywhere.
    """
    local = os.path.join(dest_dir, os.path.basename(key))
    if os.path.exists(local):
        return local
    return fetch_to(key, local, bucket)

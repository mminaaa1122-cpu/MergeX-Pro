"""
storage_utils.py
----------------
Small persistence layer shared by ops_log.py and settings.py.
Uses the same Supabase bucket that batch_cleaner.py already uses,
with a local-file fallback (local files vanish when the app restarts on
Streamlit Cloud, so Supabase is what makes the data permanent).
"""
from __future__ import annotations

import os

import streamlit as st

BUCKET_NAME = "batch-cleaner"


_last_error = ""


def _client():
    global _last_error
    try:
        from supabase import create_client

        return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])
    except Exception as exc:
        # Keep only the exception type and a short message (never the key).
        _last_error = f"{type(exc).__name__}: {str(exc)[:120]}"
        return None


def read_bytes(filename: str) -> bytes | None:
    """Read from Supabase first, then from the local file."""
    client = _client()
    if client is not None:
        try:
            data = client.storage.from_(BUCKET_NAME).download(filename)
            if data:
                return data
        except Exception:
            pass
    if os.path.exists(filename):
        try:
            with open(filename, "rb") as f:
                return f.read()
        except Exception:
            pass
    return None


def write_bytes(filename: str, content: bytes) -> bool:
    """
    Write to Supabase (permanent) and to a local copy.
    Returns True only if the Supabase write succeeded.
    """
    remote_ok = False
    client = _client()
    if client is not None:
        try:
            try:
                client.storage.from_(BUCKET_NAME).upload(
                    filename, content, {"upsert": "true"}
                )
            except Exception:
                try:
                    client.storage.from_(BUCKET_NAME).remove([filename])
                except Exception:
                    pass
                client.storage.from_(BUCKET_NAME).upload(filename, content)
            remote_ok = True
        except Exception as exc:
            print(f"Supabase write failed for {filename}: {exc}")
    try:
        with open(filename, "wb") as f:
            f.write(content)
    except Exception:
        pass
    return remote_ok


@st.cache_data(ttl=60, show_spinner=False)
def storage_status() -> dict:
    """Is permanent storage reachable? Cached for 60s to avoid a network call per rerun."""
    client = _client()
    if client is None:
        return {"ok": False, "message": f"تعذر إنشاء اتصال Supabase — {_last_error}"}
    try:
        client.storage.from_(BUCKET_NAME).list()
        return {"ok": True, "message": "متصل"}
    except Exception as exc:
        return {"ok": False, "message": f"تعذر الوصول للـ bucket ({type(exc).__name__})"}

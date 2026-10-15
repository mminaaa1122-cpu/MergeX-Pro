"""
ops_log.py
----------
One persistent log for every real operation in MergeX Pro
(Merge, Batch Cleaner, Member Checker, and login/logout events).
Reports and Settings read from here, so nothing is demo data and nothing
disappears on restart (stored in Supabase, same bucket as the master DB).

Never raises: a logging failure must never break the user's actual work.
"""
from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd
import pytz
import streamlit as st

from storage_utils import read_bytes, write_bytes

LOG_NAME = "operations_log.csv"

NUMERIC_COLUMNS = [
    "total", "members", "old_members", "not_found", "needs_review",
    "adib", "xclusives", "empty_product", "name_mismatch", "duplicates", "clean",
]
COLUMNS = [
    "timestamp", "username", "module", "action", "status", "file_name",
    *NUMERIC_COLUMNS, "ip", "note",
]
COLUMN_LABELS = {
    "timestamp": "الوقت", "username": "المستخدم", "module": "القسم",
    "action": "العملية", "status": "الحالة", "file_name": "الملف",
    "total": "الإجمالي", "members": "Members", "old_members": "Old Members",
    "not_found": "Not Found", "needs_review": "Needs Review", "adib": "ADIB",
    "xclusives": "Xclusives", "empty_product": "Product فاضي",
    "name_mismatch": "اختلاف اسم", "duplicates": "مكرر", "clean": "Clean",
    "ip": "IP", "note": "ملاحظات",
}


def _now() -> str:
    return datetime.now(pytz.timezone("Africa/Cairo")).strftime("%Y-%m-%d %H:%M:%S")


def _load_raw() -> pd.DataFrame:
    data = read_bytes(LOG_NAME)
    if not data:
        return pd.DataFrame(columns=COLUMNS)
    try:
        df = pd.read_csv(BytesIO(data), dtype=str, keep_default_na=False)
    except Exception:
        return pd.DataFrame(columns=COLUMNS)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNS]


def log_operation(
    module: str,
    action: str,
    status: str = "نجاح",
    username: str | None = None,
    file_name: str = "",
    metrics: dict | None = None,
    note: str = "",
    ip: str = "",
) -> None:
    try:
        row = {c: "" for c in COLUMNS}
        row.update(
            timestamp=_now(),
            username=username or st.session_state.get("username", "-"),
            module=module,
            action=action,
            status=status,
            file_name=file_name or "",
            note=(note or "")[:300],
            ip=ip or "",
        )
        for key, value in (metrics or {}).items():
            if key in NUMERIC_COLUMNS:
                row[key] = int(value)
        df = pd.concat([_load_raw(), pd.DataFrame([row])], ignore_index=True)
        write_bytes(LOG_NAME, df.to_csv(index=False).encode("utf-8-sig"))
    except Exception as exc:
        print(f"ops_log error: {exc}")


def load_operations() -> pd.DataFrame:
    """All logged operations, newest first, with a parsed `ts` column and int metrics."""
    df = _load_raw()
    df["ts"] = pd.to_datetime(df["timestamp"], errors="coerce")
    for col in NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    return df.sort_values("ts", ascending=False).reset_index(drop=True)

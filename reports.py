"""
reports.py — Executive Reports Center for MergeX Pro
"""
from __future__ import annotations

from datetime import date, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st

from ops_log import load_operations, COLUMN_LABELS


def _n(df: pd.DataFrame, col: str) -> int:
    if df is None or df.empty or col not in df.columns:
        return 0
    return int(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())


def _excel(df: pd.DataFrame, sheet: str = "Report") -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=str(sheet)[:31])
    return buf.getvalue()


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Report") -> bytes:
    return _excel(df, sheet_name)


def _kpi(title: str, value, sub: str = "", color: str = "#8b5cf6"):
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(145deg, rgba(15,23,42,.95), rgba(30,41,59,.9));
            border: 1px solid rgba(148,163,184,.18);
            border-top: 3px solid {color};
            border-radius: 16px;
            padding: 18px 16px;
            min-height: 110px;
            box-shadow: 0 8px 24px rgba(0,0,0,.25);
        ">
            <div style="color:#94a3b8;font-size:.78rem;font-weight:700;letter-spacing:.04em;text-transform:uppercase;margin-bottom:8px;">
                {title}
            </div>
            <div style="color:#f8fafc;font-size:1.85rem;font-weight:800;line-height:1.1;margin-bottom:6px;">
                {value}
            </div>
            <div style="color:#64748b;font-size:.8rem;">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_reports():
    # ---------- Header ----------
    st.markdown(
        """
        <div style="
            padding:28px 24px;border-radius:22px;margin-bottom:20px;
            background:linear-gradient(120deg, rgba(15,23,42,.98), rgba(49,46,129,.35), rgba(15,23,42,.98));
            border:1px solid rgba(99,102,241,.35);
            box-shadow:0 0 40px rgba(99,102,241,.12);
            text-align:center;
        ">
            <div style="font-size:.75rem;letter-spacing:.2em;color:#a5b4fc;font-weight:700;margin-bottom:8px;">
                MERGEX PRO · BUSINESS INTELLIGENCE
            </div>
            <h1 style="margin:0 0 8px 0;font-size:2rem;color:#f1f5f9;">📋 مركز التقارير التنفيذي</h1>
            <p style="margin:0;color:#94a3b8;font-size:.95rem;">
                نظرة شاملة على الأداء التشغيلي • حسب الفترة والمستخدم والقسم
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---------- Load ----------
    try:
        df = load_operations()
    except Exception as exc:
        st.error("تعذر تحميل سجل العمليات.")
        st.caption(str(exc)[:240])
        return

    if df is None or df.empty:
        st.info("لا توجد عمليات مسجّلة بعد. ستظهر التقارير فور تنفيذ أول عملية.")
        return

    if "ts" not in df.columns:
        df["ts"] = pd.to_datetime(df.get("timestamp"), errors="coerce")
    else:
        df["ts"] = pd.to_datetime(df["ts"], errors="coerce")
    df = df[df["ts"].notna()].copy()
    if df.empty:
        st.warning("تعذر قراءة التواريخ من سجل العمليات.")
        return

    df["date_only"] = df["ts"].dt.date
    min_day, max_day = min(df["date_only"]), max(df["date_only"])
    today = min(max(date.today(), min_day), max_day)

    # ---------- Filters ----------
    st.markdown(
        """
        <div style="
            padding:4px 0 0 0;margin-bottom:6px;
            border-bottom:1px solid rgba(148,163,184,.15);
        ">
            <span style="color:#cbd5e1;font-weight:700;font-size:.95rem;">تحديد نطاق التحليل</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    f1, f2, f3 = st.columns(3)
    with f1:
        period = st.selectbox(
        "الفترة الزمنية",
        ["هذا الشهر", "اليوم", "آخر 7 أيام", "هذه السنة", "كل الفترات", "تحديد فترة"],
        key="rpt_period",
    )
    with f2:
        users = sorted(
        df.get("username", pd.Series(dtype=str))
        .fillna("")
        .astype(str)
        .replace("", "غير معروف")
        .unique()
        .tolist()
    )
    selected_user = st.selectbox("المستخدم", ["كل المستخدمين", *users], key="rpt_user")
    with f3:
        modules = sorted(
        df.get("module", pd.Series(dtype=str))
        .fillna("")
        .astype(str)
        .replace("", "غير محدد")
        .unique()
        .tolist()
    )
        selected_module = st.selectbox("القسم", ["كل الأقسام", *modules], key="rpt_module")

    calendar_min = min_day - timedelta(days=60)
    calendar_max = max(max_day, date.today())    

    for k in ("rpt_start", "rpt_end"):
        if k in st.session_state:
            v = st.session_state.get(k)
            if not isinstance(v, date) or v < calendar_min or v > calendar_max:
                del st.session_state[k]

    if period == "تحديد فترة":
        d1, d2 = st.columns(2)
        with d1:
            start_day = st.date_input(
            "من تاريخ",
            value=min_day,
            min_value=calendar_min,
            max_value=calendar_max,
            key="rpt_start",
        )
        with d2:
            end_day = st.date_input(
            "إلى تاريخ",
            value=max_day,
            min_value=calendar_min,
            max_value=calendar_max,
            key="rpt_end",
        )
    elif period == "اليوم":
        start_day = end_day = today
    elif period == "آخر 7 أيام":
        start_day = max(min_day, today - timedelta(days=6))
        end_day = max_day
    elif period == "هذه السنة":
        start_day = max(min_day, date(today.year, 1, 1))
        end_day = max_day
    elif period == "كل الفترات":
         start_day, end_day = min_day, max_day
    else:
        start_day = max(min_day, date(today.year, today.month, 1))
        end_day = max_day

    if start_day > end_day:
        st.error("تاريخ البداية يجب أن يسبق تاريخ النهاية أو يساويه.")
        return

    filtered = df[(df["date_only"] >= start_day) & (df["date_only"] <= end_day)].copy()

    # غير الأدمن يشوف عملياته فقط
    if st.session_state.get("role", "user") != "admin" and "username" in filtered.columns:
        filtered = filtered[
            filtered["username"].fillna("").astype(str).str.lower()
            == str(st.session_state.get("username", "")).lower()
        ]

    if selected_user != "كل المستخدمين" and "username" in filtered.columns:
        filtered = filtered[
            filtered["username"].fillna("").astype(str).replace("", "غير معروف") == selected_user
        ]
    if selected_module != "كل الأقسام" and "module" in filtered.columns:
        filtered = filtered[
            filtered["module"].fillna("").astype(str).replace("", "غير محدد") == selected_module
        ]

    st.caption(
        f"نطاق التحليل: {start_day:%d/%m/%Y} — {end_day:%d/%m/%Y}  ·  "
        f"{selected_user}  ·  {selected_module}  ·  {len(filtered):,} عملية"
    )
    st.markdown("---")

    # ---------- Executive KPIs ----------
    total_ops = len(filtered)
    total_rows = _n(filtered, "total")
    success = 0
    if not filtered.empty and "status" in filtered.columns:
        success = int(
            filtered["status"].astype(str)
            .str.contains("نجاح|success", case=False, na=False).sum()
        )
    success_rate = f"{(success / total_ops * 100):.0f}%" if total_ops else "—"

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        _kpi("إجمالي العمليات", f"{total_ops:,}", "Operations executed", "#6366f1")
    with k2:
        _kpi("معدل النجاح", success_rate, f"{success:,} عملية ناجحة", "#22c55e")
    with k3:
        _kpi("حجم السجلات", f"{total_rows:,}", "Rows processed", "#06b6d4")
    with k4:
        _kpi("يحتاج مراجعة", f"{_n(filtered, 'needs_review'):,}", "Needs Review", "#f59e0b")

    st.markdown("")

    # ---------- Module performance strip ----------
    st.markdown("#### أداء الأقسام")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        _kpi("Members", f"{_n(filtered, 'members'):,}", "عضويات سارية", "#4ade80")
    with m2:
        _kpi("Not Found", f"{_n(filtered, 'not_found'):,}", "غير موجودين", "#f87171")
    with m3:
        _kpi("ADIB", f"{_n(filtered, 'adib'):,}", "عملاء ADIB", "#34d399")
    with m4:
        _kpi("Xclusives", f"{_n(filtered, 'xclusives'):,}", "عملاء Xclusives", "#38bdf8")

    st.markdown("")
    m5, m6, m7, m8 = st.columns(4)
    with m5:
        _kpi("Old Members", f"{_n(filtered, 'old_members'):,}", "عضويات منتهية", "#fb923c")
    with m6:
        _kpi("Duplicates", f"{_n(filtered, 'duplicates'):,}", "سجلات مكررة", "#eab308")
    with m7:
        _kpi("Clean", f"{_n(filtered, 'clean'):,}", "سجلات نظيفة", "#a3e635")
    with m8:
        _kpi("Name Mismatch", f"{_n(filtered, 'name_mismatch'):,}", "اختلاف اسم", "#c084fc")

    st.markdown("---")

    # ---------- Tabs ----------
    tab_ops, tab_trend = st.tabs(
       ["📑 سجل العمليات التفصيلي", "📈 توزيع الأداء"]
    )

    with tab_ops:
        if filtered.empty:
            st.info("لا توجد عمليات في النطاق المحدد.")
        else:
            show = filtered.sort_values("ts", ascending=False).copy()
            show["الوقت"] = show["ts"].dt.strftime("%Y-%m-%d %H:%M")
            drop_cols = [c for c in ["ts", "date_only", "timestamp"] if c in show.columns]
            show = show.drop(columns=drop_cols)
            rename = {
    "username": "المستخدم",
    "module": "القسم",
    "action": "العملية",
    "status": "الحالة",
    "file_name": "الملف",
    "note": "ملاحظات",
    "total": "الإجمالي",
    "members": "Members",
    "old_members": "Old Members",
    "not_found": "Not Found",
    "needs_review": "Needs Review",
    "adib": "ADIB",
    "xclusives": "Xclusives",
    "empty_product": "Product فاضي",
    "name_mismatch": "اختلاف اسم",
    "duplicates": "مكرر",
    "clean": "Clean",
    "ip": "IP",
}
            show = show.rename(columns=rename)
            show = show.loc[:, ~show.columns.duplicated()]
            cols = [
                c for c in [
                    "الوقت", "المستخدم", "القسم", "العملية", "الحالة", "الملف",
                    "الإجمالي", "Members", "Old Members", "Not Found",
                    "ADIB", "Xclusives", "مكرر", "Clean", "ملاحظات",
                ] if c in show.columns
            ]
            st.dataframe(show[cols], use_container_width=True, hide_index=True, height=420)

            st.download_button(
                "⬇️ تحميل السجل التفصيلي Excel",
                data=_excel(show[cols], "Operations"),
                file_name=f"MergeX_Ops_{start_day}_{end_day}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with tab_trend:
        if filtered.empty:
            st.info("لا بيانات للعرض.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**العمليات حسب القسم**")
                if "module" in filtered.columns:
                    st.bar_chart(filtered.groupby("module").size().rename("عمليات"))
            with c2:
                st.markdown("**العمليات حسب المستخدم**")
                if "username" in filtered.columns:
                    st.bar_chart(
                        filtered.groupby("username").size().rename("عمليات").sort_values(ascending=False)
                    )

            st.markdown("**ملخص تشغيلي حسب القسم**")
            rows = []
            for mod, g in filtered.groupby("module"):
                rows.append({
                    "القسم": mod,
                    "العمليات": len(g),
                    "السجلات": _n(g, "total"),
                    "Members": _n(g, "members"),
                    "Not Found": _n(g, "not_found"),
                    "ADIB": _n(g, "adib"),
                    "Xclusives": _n(g, "xclusives"),
                    "مكرر": _n(g, "duplicates"),
                    "Clean": _n(g, "clean"),
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    
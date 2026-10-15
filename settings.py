"""
settings.py - ⚙️ Settings module for MergeX Pro (admin only).

Tabs: templates | connections | users & permissions | login/events | general.
Everything is stored in app_settings.json (Supabase + local fallback), so it
survives restarts. Passwords set from the UI are stored as PBKDF2 hashes.
Users defined in the code (USERS in test1.py) keep working as a fallback.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import secrets
import time

import pandas as pd
import streamlit as st

from ops_log import COLUMN_LABELS, load_operations
from reports import to_excel_bytes
from storage_utils import read_bytes, storage_status, write_bytes

SETTINGS_FILE = "app_settings.json"
ADMIN_ONLY_MODULES = {"⚙️ Settings"}
NEW_USER = "➕ مستخدم جديد"
DEFAULT_GENERAL = {"name_max_len": 23, "district_cutoff": 0.4}


# ────────────────────────── storage ──────────────────────────
@st.cache_data(ttl=30, show_spinner=False)
def _read_settings() -> dict:
    raw = read_bytes(SETTINGS_FILE)
    if not raw:
        return {}
    try:
        data = json.loads(raw.decode("utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def load_settings() -> dict:
    s = _read_settings()
    s.setdefault("templates", {})
    s.setdefault("users", {})
    s.setdefault("general", {})
    return s


def save_settings(s: dict) -> bool:
    ok = write_bytes(SETTINGS_FILE, json.dumps(s, ensure_ascii=False, indent=2).encode("utf-8"))
    _read_settings.clear()
    return ok


def _flash(msg: str, kind: str = "success") -> None:
    st.session_state["_settings_flash"] = (kind, msg)


def _show_flash() -> None:
    item = st.session_state.pop("_settings_flash", None)
    if item:
        getattr(st, item[0])(item[1])


def _persist(settings: dict, ok_msg: str) -> None:
    if save_settings(settings):
        _flash(ok_msg)
    else:
        _flash("⚠️ اتحفظ محليًا بس — Supabase مش متاح، فممكن يضيع عند إعادة التشغيل.", "warning")
    st.rerun()


# ────────────────────── functions used by test1.py ──────────────────────
def get_general(default_districts: list[str] | None = None) -> dict:
    stored = load_settings()["general"]
    g = {**DEFAULT_GENERAL, **stored}
    districts = stored.get("districts") or default_districts or []
    g["districts"] = [str(d).strip().upper() for d in districts if str(d).strip()]
    return g


def apply_template_overrides(template_key: str, defaults: dict) -> dict:
    override = load_settings()["templates"].get(template_key)
    return dict(override) if override else dict(defaults)


def _hash(password: str, salt: bytes | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return salt.hex(), digest.hex()


def _check(password: str, salt_hex: str, hash_hex: str) -> bool:
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), 200_000
    )
    return hmac.compare_digest(digest.hex(), hash_hex)


def verify_user(username: str, password: str, fallback_users: dict) -> dict | None:
    """Stored (hashed) credentials win; the USERS dict in code is the fallback."""
    u = (username or "").strip().lower()
    stored = load_settings()["users"].get(u)
    fb = fallback_users.get(u)
    if not stored and not fb:
        return None
    if stored and stored.get("active") is False:
        return None

    if stored and stored.get("hash"):
        ok = _check(password, stored["salt"], stored["hash"])
    elif fb:
        ok = hmac.compare_digest(password.encode("utf-8"), str(fb["password"]).encode("utf-8"))
    else:
        ok = False
    if not ok:
        return None

    return {
        "role": (stored or {}).get("role") or (fb or {}).get("role", "user"),
        "name": (stored or {}).get("name") or (fb or {}).get("name", u),
    }


def get_allowed_modules(username: str, role: str, all_modules: list[str]) -> list[str]:
    if role == "admin":
        return list(all_modules)
    mods = load_settings()["users"].get(username, {}).get("modules")
    allowed = [
        m for m in all_modules
        if m not in ADMIN_ONLY_MODULES and (mods is None or m in mods)
    ]
    return allowed or [all_modules[0]]


# ────────────────────────── UI ──────────────────────────
def render_settings(templates: dict, fallback_users: dict, all_modules: list[str],
                    default_districts: list[str] | None = None) -> None:
    st.markdown("<h1>⚙️ Settings</h1>", unsafe_allow_html=True)

    if st.session_state.get("role") != "admin":
        st.error("⛔ الإعدادات متاحة للأدمن فقط.")
        return

    _show_flash()
    tabs = st.tabs(["📄 القوالب", "🔌 الاتصال", "👥 المستخدمون", "🛡️ الدخول والأحداث", "🧩 إعدادات عامة"])
    with tabs[0]:
        _tab_templates(templates)
    with tabs[1]:
        _tab_connections()
    with tabs[2]:
        _tab_users(fallback_users, all_modules)
    with tabs[3]:
        _tab_events()
    with tabs[4]:
        _tab_general(default_districts)


def _tab_templates(templates: dict) -> None:
    settings = load_settings()
    st.caption("القيم دي بتتطبق على كل الصفوف عند الدمج. عمود Call Date بيتحط تلقائيًا بتاريخ النهارده.")
    for key, info in templates.items():
        exists = os.path.exists(info["file"])
        label = f"{key} — {info['file']} {'✅' if exists else '❌ الملف غير موجود'}"
        with st.expander(label):
            override = settings["templates"].get(key)
            current = override if override else info["defaults"]
            df = pd.DataFrame({"Column": list(current.keys()), "Value": [str(v) for v in current.values()]})
            edited = st.data_editor(
                df, num_rows="dynamic", use_container_width=True,
                hide_index=True, key=f"tpl_editor_{key}",
            )
            c1, c2 = st.columns(2)
            if c1.button("💾 حفظ", key=f"tpl_save_{key}", use_container_width=True):
                new: dict[str, str] = {}
                for _, r in edited.iterrows():
                    col = "" if pd.isna(r["Column"]) else str(r["Column"]).strip()
                    if col:
                        new[col] = "" if pd.isna(r["Value"]) else str(r["Value"])
                if not new:
                    st.error("لازم يفضل عمود واحد على الأقل.")
                else:
                    settings["templates"][key] = new
                    _persist(settings, f"✅ اتحفظت قيم {key}")
            if c2.button("↩️ رجوع للافتراضي", key=f"tpl_reset_{key}", use_container_width=True):
                settings["templates"].pop(key, None)
                _persist(settings, f"✅ رجعت قيم {key} للافتراضي")


def _test_sql(cfg: dict) -> tuple[bool, str]:
    try:
        import pyodbc
    except ImportError:
        return False, "مكتبة pyodbc غير مثبتة."
    encrypt = str(cfg["encrypt"]).lower() in {"1", "true", "yes", "on"}
    trust = str(cfg["trust_server_certificate"]).lower() in {"1", "true", "yes", "on"}
    conn_str = (
        f"DRIVER={{{cfg['driver']}}};SERVER={cfg['server']};DATABASE={cfg['database']};"
        f"UID={cfg['username']};PWD={cfg['password']};"
        f"Encrypt={'yes' if encrypt else 'no'};TrustServerCertificate={'yes' if trust else 'no'};"
        "Connection Timeout=10;"
    )
    conn = None
    started = time.perf_counter()
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        conn.cursor().execute("SELECT 1").fetchone()
        return True, f"تم الاتصال بنجاح ({(time.perf_counter() - started) * 1000:.0f} ms)"
    except Exception as exc:
        return False, f"{type(exc).__name__}: {str(exc)[:160]}"
    finally:
        if conn is not None:
            conn.close()


def _tab_connections() -> None:
    st.markdown("#### 🗄️ قاعدة بيانات الشركة (SQL Server)")
    try:
        from member_checker import get_sql_config

        cfg = get_sql_config()
    except Exception as exc:
        cfg = None
        st.error(str(exc))

    if cfg:
        user = str(cfg["username"])
        st.dataframe(
            pd.DataFrame({
                "البند": ["Server", "Database", "Driver", "Username", "Password", "Encrypt"],
                "القيمة": [cfg["server"], cfg["database"], cfg["driver"],
                           user[:2] + "***", "••••••••", str(cfg["encrypt"])],
            }),
            hide_index=True, use_container_width=True,
        )
        if st.button("🧪 اختبار الاتصال", key="test_sql"):
            with st.spinner("جاري الاختبار..."):
                ok, msg = _test_sql(cfg)
            (st.success if ok else st.error)(("✅ " if ok else "❌ ") + msg)

    st.markdown("#### ☁️ التخزين الدائم (Supabase)")
    status = storage_status()
    if status["ok"]:
        st.success(f"✅ {status['message']}")
    else:
        st.error(f"❌ {status['message']}")
    if st.button("🔄 إعادة فحص التخزين", key="recheck_storage"):
        storage_status.clear()
        st.rerun()


def _tab_users(fallback: dict, all_modules: list[str]) -> None:
    settings = load_settings()
    stored = settings["users"]
    names = list(dict.fromkeys(list(fallback) + list(stored)))
    assignable = [m for m in all_modules if m not in ADMIN_ONLY_MODULES]
    me = st.session_state.get("username", "")

    rows = []
    for u in names:
        fb, s = fallback.get(u, {}), stored.get(u, {})
        role = s.get("role") or fb.get("role", "user")
        mods = s.get("modules")
        rows.append({
            "اليوزر": u,
            "الاسم": s.get("name") or fb.get("name", u),
            "الدور": role,
            "مفعّل": s.get("active", True),
            "الأقسام": "الكل" if role == "admin" else ("الكل (ماعدا الإعدادات)" if mods is None else "، ".join(mods)),
            "الباسورد": "مخزّن ومشفّر" if s.get("hash") else "من الكود",
        })
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    st.markdown("#### ✏️ إضافة / تعديل مستخدم")
    choice = st.selectbox("اختر مستخدم", [NEW_USER] + names, key="user_choice")
    is_new = choice == NEW_USER
    k = "uf0" if is_new else f"uf{names.index(choice) + 1}"
    s, fb = stored.get(choice, {}), fallback.get(choice, {})
    cur_role = s.get("role") or fb.get("role", "user")
    cur_mods = s.get("modules")
    cur_mods = assignable if cur_mods is None else [m for m in cur_mods if m in assignable]

    with st.form(f"user_form_{k}"):
        username = st.text_input("اسم المستخدم (إنجليزي صغير)", value="" if is_new else choice,
                                 disabled=not is_new, key=f"{k}_u")
        name = st.text_input("الاسم الظاهر", value=s.get("name") or fb.get("name", ""), key=f"{k}_n")
        role = st.selectbox("الدور", ["user", "admin"],
                            index=1 if cur_role == "admin" else 0, key=f"{k}_r")
        active = st.checkbox("مفعّل", value=s.get("active", True), key=f"{k}_a")
        modules = st.multiselect("الأقسام المسموحة (للمستخدم العادي)", assignable,
                                 default=cur_mods, key=f"{k}_m")
        pwd = st.text_input("كلمة مرور جديدة (٨ حروف على الأقل)", type="password",
                            help="سيبها فاضية لو مش عايز تغيّر الباسورد الحالي.", key=f"{k}_p")
        submitted = st.form_submit_button("💾 حفظ المستخدم", use_container_width=True)

    if submitted:
        uname = (username if is_new else choice).strip().lower()
        errors = []
        if not re.fullmatch(r"[a-z0-9_.\-]{3,30}", uname):
            errors.append("اسم المستخدم: ٣–٣٠ حرف إنجليزي صغير/أرقام/_ . -")
        if is_new and uname in names:
            errors.append("اسم المستخدم موجود بالفعل.")
        if is_new and not pwd:
            errors.append("لازم تحط كلمة مرور للمستخدم الجديد.")
        if pwd and len(pwd) < 8:
            errors.append("كلمة المرور لازم تكون ٨ حروف على الأقل.")
        if uname == me and (role != "admin" or not active):
            errors.append("مينفعش تلغي أو تنزّل صلاحية حسابك الحالي.")
        if errors:
            for e in errors:
                st.error(e)
        else:
            entry = dict(stored.get(uname, {}))
            entry.update(name=name.strip() or uname, role=role, active=active, modules=modules)
            if pwd:
                entry["salt"], entry["hash"] = _hash(pwd)
            settings["users"][uname] = entry
            _persist(settings, f"✅ اتحفظ المستخدم {uname}")

    if not is_new and choice in stored:
        st.markdown("---")
        in_code = choice in fallback
        confirm = st.checkbox("أؤكد", key=f"{k}_del_ok")
        label = "↩️ رجوع لإعدادات الكود (يمسح التعديلات المخزّنة)" if in_code else "🗑️ حذف المستخدم"
        if st.button(label, disabled=not confirm, key=f"{k}_del"):
            if choice == me:
                st.error("مينفعش تحذف حسابك الحالي.")
            else:
                settings["users"].pop(choice, None)
                _persist(settings, f"✅ تم: {choice}")


def _tab_events() -> None:
    df = load_operations()
    auth = df[(df["module"] == "Auth") & df["ts"].notna()].copy()
    if auth.empty:
        st.info("لا توجد أحداث دخول مسجلة بعد.")
        return

    ok = int((auth["status"] == "نجاح").sum())
    bad = int((auth["status"] == "فشل").sum())
    since = pd.Timestamp.now() - pd.Timedelta(hours=24)
    bad_24 = int(((auth["status"] == "فشل") & (auth["ts"] >= since)).sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("عمليات ناجحة", ok)
    c2.metric("محاولات فاشلة", bad)
    c3.metric("فاشلة آخر ٢٤ ساعة", bad_24)

    f1, f2 = st.columns(2)
    user_sel = f1.selectbox("المستخدم", ["الكل"] + sorted(auth["username"].unique().tolist()), key="ev_user")
    status_sel = f2.selectbox("الحالة", ["الكل", "نجاح", "فشل"], key="ev_status")
    if user_sel != "الكل":
        auth = auth[auth["username"] == user_sel]
    if status_sel != "الكل":
        auth = auth[auth["status"] == status_sel]

    show = auth[["timestamp", "username", "action", "status", "ip", "note"]].rename(columns=COLUMN_LABELS)
    st.dataframe(show.head(300), use_container_width=True, height=420, hide_index=True)
    st.download_button(
        "📥 تحميل سجل الدخول (Excel)",
        data=to_excel_bytes({"Events": show}),
        file_name="MergeX_Login_Events.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="ev_download",
    )


def _tab_general(default_districts: list[str] | None = None) -> None:
    g = get_general(default_districts)
    st.caption("إعدادات بتأثر فعليًا على ملف الدمج (الـ District وطول الاسم).")
    with st.form("general_form"):
        districts_text = st.text_area(
            "الـ Districts المسموحة (واحد في كل سطر)",
            value="\n".join(g["districts"]), height=220,
        )
        max_len = st.number_input(
            "تلوين اسم الكارت لو طوله وصل لـ", min_value=10, max_value=60,
            value=int(g["name_max_len"]), step=1,
        )
        cutoff = st.slider(
            "حساسية التصحيح التلقائي للـ District (أعلى = أكثر تشددًا)",
            0.1, 0.9, float(g["district_cutoff"]), 0.05,
        )
        submitted = st.form_submit_button("💾 حفظ", use_container_width=True)
    if submitted:
        districts = list(dict.fromkeys(
            d.strip().upper() for d in districts_text.splitlines() if d.strip()
        ))
        if not districts:
            st.error("لازم تحط District واحد على الأقل.")
            return
        settings = load_settings()
        settings["general"] = {
            "districts": districts, "name_max_len": int(max_len), "district_cutoff": float(cutoff),
        }
        _persist(settings, "✅ اتحفظت الإعدادات العامة")

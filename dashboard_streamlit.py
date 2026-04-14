import os
import re
import sqlite3
import hashlib
from io import BytesIO
from datetime import datetime
from typing import Optional

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st


# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="ALI Leads Pro v5",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_PATH = os.environ.get("ALI_LEADS_DB_PATH", "ali_leads.db")

TIPOS_EMPRESA = ["ALIMATICO", "MAGNA"]
TIPOS_CANAL = ["Mostrador", "Whatsapp", "Siniestro", "Taller Magna", "Taller Particular", "Sin clasificar"]
TIPOS_COMPRADO = ["SI", "NO", "EN PROCESO"]
TIPOS_MOTIVO = ["", "Precio", "Sin stock", "Demora", "No respondio", "Compro en otro lado", "No le gusto", "Otros"]
TIPOS_COMPANIA = ["", "BSE", "SURA", "Porto Seguro", "SBI", "HDI", "Berkley", "Sancor", "MAPFRE"]
MARCAS_MAGNA = ["MAZDA", "KIA"]

COLUMNAS_BASE = [
    "EMPRESA",
    "FECHA",
    "CANAL",
    "COMPAÃ‘IA",
    "NÂ° SINIESTRO",
    "CHASIS",
    "NOMBRE CLIENTE",
    "TELEFONO",
    "MARCA",
    "MODELO",
    "CODIGO",
    "REPUESTOS SOLICITADO",
    "VALOR",
    "COMPRADO",
    "MOTIVO",
    "COMENTARIOS",
]


# =========================================================
# ESTILO
# =========================================================
def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 2rem;
        }
        .main-title-card {
            background: linear-gradient(135deg, rgba(255,255,255,0.97) 0%, rgba(248,250,252,0.97) 100%);
            border: 1px solid rgba(15,23,42,0.08);
            border-radius: 24px;
            padding: 1.2rem 1.4rem;
            margin-bottom: 1rem;
            box-shadow: 0 12px 30px rgba(15,23,42,0.10);
        }
        .main-title {
            font-size: 3rem;
            font-weight: 900;
            color: #0f172a;
            margin: 0;
            line-height: 1.05;
            letter-spacing: -0.03em;
            text-shadow: 0 2px 10px rgba(255,255,255,0.45);
        }
        .main-title-accent {
            color: #1d4ed8;
        }
        .subtitle {
            font-size: 1.02rem;
            color: #475569;
            margin-top: 0.45rem;
            margin-bottom: 0;
            font-weight: 500;
        }
        .mode-badge {
            display: inline-block;
            margin-top: 0.85rem;
            padding: 0.42rem 0.8rem;
            border-radius: 999px;
            background: #eff6ff;
            color: #1d4ed8;
            font-size: 0.9rem;
            font-weight: 800;
            border: 1px solid rgba(37,99,235,0.18);
        }
        .hero-card {
            background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);
            color: white;
            border-radius: 22px;
            padding: 1.2rem 1.4rem;
            margin-bottom: 1rem;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.18);
        }
        .hero-title {
            font-size: 1.25rem;
            font-weight: 800;
            margin-bottom: 0.2rem;
        }
        .hero-subtitle {
            font-size: 0.96rem;
            opacity: 0.92;
        }
        .mini-card {
            background: rgba(255,255,255,0.99);
            border: 1px solid rgba(15,23,42,0.06);
            border-radius: 18px;
            padding: 0.95rem 1rem;
            box-shadow: 0 4px 12px rgba(15,23,42,0.05);
            margin-bottom: 0.8rem;
            min-height: 124px;
        }
        .kpi-title {
            font-size: 0.9rem;
            color: #64748b;
            margin-bottom: 0.25rem;
        }
        .kpi-value {
            font-size: 1.55rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.12;
        }
        .status-good {
            background: #ecfdf5;
            color: #166534;
            border-left: 6px solid #22c55e;
            padding: 0.85rem 1rem;
            border-radius: 14px;
            font-weight: 700;
            margin-bottom: 1rem;
        }
        .status-mid {
            background: #fffbeb;
            color: #92400e;
            border-left: 6px solid #f59e0b;
            padding: 0.85rem 1rem;
            border-radius: 14px;
            font-weight: 700;
            margin-bottom: 1rem;
        }
        .status-bad {
            background: #fef2f2;
            color: #991b1b;
            border-left: 6px solid #ef4444;
            padding: 0.85rem 1rem;
            border-radius: 14px;
            font-weight: 700;
            margin-bottom: 1rem;
        }
        .alert-card {
            background: rgba(255,255,255,0.99);
            border: 1px solid rgba(15,23,42,0.06);
            border-radius: 18px;
            padding: 1rem 1rem;
            box-shadow: 0 4px 12px rgba(15,23,42,0.05);
            min-height: 122px;
        }
        .alert-title {
            font-size: 0.9rem;
            color: #64748b;
            margin-bottom: 0.35rem;
            font-weight: 700;
        }
        .alert-value {
            font-size: 1.1rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.2;
        }
        .login-card {
            max-width: 460px;
            margin: 3rem auto;
            background: white;
            padding: 1.5rem;
            border-radius: 22px;
            box-shadow: 0 15px 35px rgba(15,23,42,0.12);
            border: 1px solid rgba(15,23,42,0.08);
        }
        [data-testid="metric-container"] {
            background: rgba(255,255,255,0.99);
            border: 1px solid rgba(15,23,42,0.06);
            padding: 1rem !important;
            border-radius: 18px !important;
            box-shadow: 0 4px 12px rgba(15,23,42,0.05);
        }
        .small-note {
            color: #64748b;
            font-size: 0.9rem;
        }
        .summary-hero-card {
            background: linear-gradient(135deg, #07162c 0%, #123f83 62%, #0a2b58 100%);
            color: #f8fafc;
            border-radius: 26px;
            padding: 1.2rem 1.35rem;
            margin-bottom: 1rem;
            box-shadow: 0 16px 34px rgba(7, 22, 44, 0.24);
            border: 1px solid rgba(147, 197, 253, 0.12);
        }
        .summary-hero-title {
            font-size: 1.45rem;
            font-weight: 900;
            letter-spacing: -0.02em;
            margin-bottom: 0.25rem;
        }
        .summary-hero-subtitle {
            color: rgba(226, 232, 240, 0.92);
            font-size: 0.95rem;
            margin-bottom: 0.8rem;
        }
        .summary-chip-wrap {
            display: flex;
            flex-wrap: wrap;
            gap: 0.45rem;
        }
        .summary-chip {
            display: inline-flex;
            align-items: center;
            padding: 0.36rem 0.72rem;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.10);
            border: 1px solid rgba(191, 219, 254, 0.22);
            color: #e0f2fe;
            font-size: 0.84rem;
            font-weight: 700;
        }
        .summary-kpi-card {
            background: linear-gradient(180deg, #0a2244 0%, #123a73 100%);
            border-radius: 22px;
            padding: 1rem 1rem 1.1rem 1rem;
            min-height: 148px;
            box-shadow: 0 14px 30px rgba(12, 35, 67, 0.18);
            border: 1px solid rgba(147, 197, 253, 0.12);
            margin-bottom: 0.85rem;
        }
        .summary-kpi-title {
            color: #bfdbfe;
            font-size: 0.88rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 0.45rem;
        }
        .summary-kpi-value {
            color: #ffffff;
            font-size: 2rem;
            font-weight: 900;
            line-height: 1.05;
            margin-bottom: 0.35rem;
        }
        .summary-kpi-note {
            color: #cbd5e1;
            font-size: 0.88rem;
            line-height: 1.3;
        }
        .summary-section-label {
            color: #0f172a;
            font-size: 1.05rem;
            font-weight: 900;
            margin: 0.15rem 0 0.6rem 0;
            letter-spacing: -0.01em;
        }
        .summary-chart-note {
            color: #64748b;
            font-size: 0.84rem;
            margin-top: -0.15rem;
            margin-bottom: 0.55rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_css()


# =========================================================
# DB
# =========================================================
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT,
            role TEXT DEFAULT 'user',
            company_scope TEXT DEFAULT 'TODAS',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa TEXT,
            fecha TEXT,
            canal TEXT,
            compania TEXT,
            numero_siniestro TEXT,
            chasis TEXT,
            nombre_cliente TEXT,
            telefono TEXT,
            marca TEXT,
            modelo TEXT,
            codigo TEXT,
            repuesto TEXT,
            valor REAL,
            comprado TEXT,
            motivo TEXT,
            comentarios TEXT,
            created_by TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS app_settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    cur.execute("PRAGMA table_info(users)")
    cols = [r[1] for r in cur.fetchall()]
    if "company_scope" not in cols:
        cur.execute("ALTER TABLE users ADD COLUMN company_scope TEXT DEFAULT 'TODAS'")

    users = [
        ("admin", hash_password("admin123"), "Administrador", "admin", "TODAS"),
        ("demo", hash_password("demo123"), "Usuario Demo", "user", "TODAS"),
        ("magna", hash_password("Magna2026!"), "Usuario Magna", "user", "MAGNA"),
        ("alimatico", hash_password("Alimatico2026!"), "Usuario Alimatico", "user", "ALIMATICO"),
    ]

    for u in users:
        try:
            cur.execute(
                "INSERT INTO users (username, password_hash, full_name, role, company_scope) VALUES (?, ?, ?, ?, ?)",
                u,
            )
        except sqlite3.IntegrityError:
            pass

    conn.commit()
    conn.close()


def authenticate_user(username: str, password: str) -> Optional[dict]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT username, full_name, role, company_scope FROM users WHERE username = ? AND password_hash = ?",
        (username, hash_password(password)),
    )
    row = cur.fetchone()
    conn.close()

    if row:
        return {
            "username": row[0],
            "full_name": row[1],
            "role": row[2],
            "company_scope": row[3] or "TODAS",
        }
    return None


def create_user(username: str, password: str, full_name: str, role: str, company_scope: str) -> tuple[bool, str]:
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO users (username, password_hash, full_name, role, company_scope) VALUES (?, ?, ?, ?, ?)",
            (username.strip(), hash_password(password), full_name.strip(), role, company_scope),
        )
        conn.commit()
        conn.close()
        return True, "Usuario creado."
    except sqlite3.IntegrityError:
        return False, "Ese usuario ya existe."


def get_users_df() -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT id, username, full_name, role, company_scope, created_at FROM users ORDER BY id DESC",
        conn,
    )
    conn.close()
    return df


def get_app_setting(setting_key: str, default: Optional[str] = None) -> Optional[str]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT setting_value FROM app_settings WHERE setting_key = ?", (setting_key,))
    row = cur.fetchone()
    conn.close()
    if row:
        return row[0]
    return default


def get_app_setting_float(setting_key: str, default: float = 0.0) -> float:
    raw_value = get_app_setting(setting_key)
    if raw_value is None:
        return default
    try:
        return float(raw_value)
    except Exception:
        return default


def set_app_setting(setting_key: str, setting_value: object):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO app_settings (setting_key, setting_value, updated_at)
        VALUES (?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(setting_key) DO UPDATE SET
            setting_value = excluded.setting_value,
            updated_at = CURRENT_TIMESTAMP
        """,
        (setting_key, str(setting_value)),
    )
    conn.commit()
    conn.close()


def insert_lead(record: dict):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO leads (
            empresa, fecha, canal, compania, numero_siniestro, chasis, nombre_cliente,
            telefono, marca, modelo, codigo, repuesto, valor, comprado, motivo,
            comentarios, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record["EMPRESA"],
            record["FECHA"],
            record["CANAL"],
            record["COMPAÃ‘IA"],
            record["NÂ° SINIESTRO"],
            record["CHASIS"],
            record["NOMBRE CLIENTE"],
            normalize_phone(record["TELEFONO"]),
            record["MARCA"],
            record["MODELO"],
            record["CODIGO"],
            record["REPUESTOS SOLICITADO"],
            float(record["VALOR"]),
            record["COMPRADO"],
            record["MOTIVO"],
            record["COMENTARIOS"],
            record["CREATED_BY"],
        ),
    )
    conn.commit()
    conn.close()


def replace_company_leads(company: str, df: pd.DataFrame, created_by: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM leads WHERE UPPER(COALESCE(empresa,'')) = ?", (company.upper(),))
    for _, r in df.iterrows():
        cur.execute(
            """
            INSERT INTO leads (
                empresa, fecha, canal, compania, numero_siniestro, chasis, nombre_cliente,
                telefono, marca, modelo, codigo, repuesto, valor, comprado, motivo,
                comentarios, created_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                r.get("EMPRESA", ""),
                r.get("FECHA", ""),
                r.get("CANAL", ""),
                r.get("COMPAÃ‘IA", ""),
                r.get("NÂ° SINIESTRO", ""),
                r.get("CHASIS", ""),
                r.get("NOMBRE CLIENTE", ""),
                normalize_phone(r.get("TELEFONO", "")),
                r.get("MARCA", ""),
                r.get("MODELO", ""),
                r.get("CODIGO", ""),
                r.get("REPUESTOS SOLICITADO", ""),
                float(r.get("VALOR", 0) or 0),
                r.get("COMPRADO", "NO"),
                r.get("MOTIVO", ""),
                r.get("COMENTARIOS", ""),
                created_by,
            ),
        )
    conn.commit()
    conn.close()


def append_company_leads(df: pd.DataFrame, created_by: str):
    conn = get_conn()
    cur = conn.cursor()
    for _, r in df.iterrows():
        cur.execute(
            """
            INSERT INTO leads (
                empresa, fecha, canal, compania, numero_siniestro, chasis, nombre_cliente,
                telefono, marca, modelo, codigo, repuesto, valor, comprado, motivo,
                comentarios, created_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                r.get("EMPRESA", ""),
                r.get("FECHA", ""),
                r.get("CANAL", ""),
                r.get("COMPAÃ‘IA", ""),
                r.get("NÂ° SINIESTRO", ""),
                r.get("CHASIS", ""),
                r.get("NOMBRE CLIENTE", ""),
                normalize_phone(r.get("TELEFONO", "")),
                r.get("MARCA", ""),
                r.get("MODELO", ""),
                r.get("CODIGO", ""),
                r.get("REPUESTOS SOLICITADO", ""),
                float(r.get("VALOR", 0) or 0),
                r.get("COMPRADO", "NO"),
                r.get("MOTIVO", ""),
                r.get("COMENTARIOS", ""),
                created_by,
            ),
        )
    conn.commit()
    conn.close()


def delete_lead_by_id(lead_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM leads WHERE id = ?", (lead_id,))
    conn.commit()
    conn.close()


def load_db_data() -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM leads ORDER BY id DESC", conn)
    conn.close()
    return df


init_db()


# =========================================================
# HELPERS
# =========================================================
def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_yes_no(value: object) -> str:
    v = normalize_text(value).upper()
    if v in {"SI", "SÃ", "YES", "Y", "1", "TRUE", "OK"}:
        return "SI"
    if v in {"NO", "N", "0", "FALSE"}:
        return "NO"
    if v in {"EN PROCESO", "PENDIENTE", "PEND", "PROCESO", "EN CURSO", "ABIERTO", "OPEN", ""}:
        return "EN PROCESO"
    return "EN PROCESO"


def normalize_canal(value: object) -> str:
    v = normalize_text(value).upper()
    if not v:
        return "Sin clasificar"
    if "MOSTRADOR" in v:
        return "Mostrador"
    if "WHATSAPP" in v or "WSP" in v or "WPP" in v or v == "WP":
        return "Whatsapp"
    if "SINIESTRO" in v or "SEGURO" in v:
        return "Siniestro"
    if "TALLER MAGNA" in v or v == "MAGNA":
        return "Taller Magna"
    if "TALLER" in v:
        return "Taller Particular"
    return normalize_text(value).title()


def normalize_motivo(value: object) -> str:
    v = normalize_text(value).upper()
    if not v:
        return ""
    if "PRECIO" in v:
        return "Precio"
    if "STOCK" in v:
        return "Sin stock"
    if "DEMORA" in v:
        return "Demora"
    if "NO RESP" in v:
        return "No respondio"
    if "OTRO LADO" in v or "COMPRO EN OTRO" in v or "COMPRÃ“ EN OTRO" in v:
        return "Compro en otro lado"
    if "NO LE GUSTO" in v or "NO LE GUSTÃ“" in v:
        return "No le gusto"
    return "Otros"


def normalize_compania(value: object) -> str:
    v = normalize_text(value).upper()
    if not v:
        return ""
    mapping = {
        "BSE": "BSE",
        "SURA": "SURA",
        "PORTO": "Porto Seguro",
        "PORTO SEGURO": "Porto Seguro",
        "SBI": "SBI",
        "HDI": "HDI",
        "BERKLEY": "Berkley",
        "SANCOR": "Sancor",
        "MAPFRE": "MAPFRE",
    }
    return mapping.get(v, normalize_text(value))


def normalize_phone(value: object) -> str:
    s = normalize_text(value)
    if not s:
        return ""

    s = re.sub(r"\.0+$", "", s)
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""

    if digits.startswith("598") and len(digits) > 8:
        digits = digits[3:]

    if digits.startswith("2") or digits.startswith("0"):
        return digits
    return f"0{digits}"


def parse_valor(value: object) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().upper().replace("USD", "").replace("U$S", "").replace("$", "")
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except Exception:
        return 0.0


def infer_business_mode_from_filename(uploaded_filename: str) -> str:
    name = str(uploaded_filename).strip().lower()
    if "ventas_magna" in name:
        return "MAGNA"
    if "ventas_alimatico" in name:
        return "ALIMATICO"
    return "GENERAL"


def infer_client_segment(nombre_cliente: object, canal: object) -> str:
    nombre = normalize_text(nombre_cliente).upper()
    canal_n = normalize_text(canal)
    if canal_n == "Taller Magna":
        return "Taller Magna"
    if canal_n == "Taller Particular":
        return "Mercado Talleres"
    if "TALLER MAGNA" in nombre:
        return "Taller Magna"
    if "TALLER" in nombre:
        return "Mercado Talleres"
    return "Clientes"


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    rename_map = {
        "PRECIO": "VALOR",
        "COMPRADO ": "COMPRADO",
        "TELÃ‰FONO": "TELEFONO",
        "NRO SINIESTRO": "NÂ° SINIESTRO",
        "NRO. SINIESTRO": "NÂ° SINIESTRO",
        "NUMERO SINIESTRO": "NÂ° SINIESTRO",
        "NÂº SINIESTRO": "NÂ° SINIESTRO",
    }
    for old, new in rename_map.items():
        if old in df.columns and new not in df.columns:
            df = df.rename(columns={old: new})
    for col in COLUMNAS_BASE:
        if col not in df.columns:
            df[col] = ""
    return df


def detect_header_row_from_excel(excel_file, sheet_name="BASE DE DATOS", max_rows=15):
    preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=max_rows)
    for i in range(len(preview)):
        vals = [normalize_text(x).upper() for x in preview.iloc[i].tolist()]
        joined = " | ".join(vals)
        if ("FECHA" in vals or "FECHA" in joined) and ("CANAL" in vals or "CANAL" in joined):
            return i
    return 0


def read_excel_smart(uploaded_file, preferred_sheet="BASE DE DATOS"):
    xls = pd.ExcelFile(uploaded_file)
    sheet = preferred_sheet if preferred_sheet in xls.sheet_names else xls.sheet_names[0]
    header_row = detect_header_row_from_excel(uploaded_file, sheet_name=sheet)
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row)


def clean_input_dataframe(df: pd.DataFrame, forced_company: Optional[str] = None) -> pd.DataFrame:
    df = standardize_columns(df)
    df = df[COLUMNAS_BASE].copy()

    if forced_company:
        df["EMPRESA"] = forced_company

    carry_source_cols = ["EMPRESA", "FECHA", "CANAL", "COMPAÃ‘IA", "NÂ° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "TELEFONO", "MARCA", "MODELO"]
    for col in carry_source_cols:
        df[col] = df[col].replace("", pd.NA)

    # Solo arrastramos contexto en filas de continuaciÃ³n reales.
    # Si aparece un cliente nuevo o un siniestro nuevo con datos propios,
    # evitamos heredar compaÃ±Ã­a/cliente de la fila anterior.
    continuation_markers = ["CANAL", "COMPAÃ‘IA", "NÂ° SINIESTRO", "NOMBRE CLIENTE", "TELEFONO"]
    continuation_rows = df[continuation_markers].isna().all(axis=1)

    for col in ["EMPRESA", "FECHA"]:
        df[col] = df[col].ffill()

    inherited_cols = ["CANAL", "COMPAÃ‘IA", "NÂ° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "TELEFONO", "MARCA", "MODELO"]
    for col in inherited_cols:
        carried = df[col].ffill()
        df.loc[continuation_rows, col] = df.loc[continuation_rows, col].combine_first(carried.loc[continuation_rows])

    df["EMPRESA"] = df["EMPRESA"].fillna("").astype(str).str.upper().str.strip()
    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df["FECHA"] = df["FECHA"].dt.strftime("%Y-%m-%d").fillna("")
    df["CANAL"] = df["CANAL"].apply(normalize_canal)
    df["COMPAÃ‘IA"] = df["COMPAÃ‘IA"].apply(normalize_compania)
    df["NÂ° SINIESTRO"] = df["NÂ° SINIESTRO"].fillna("").astype(str).str.strip()
    df["CHASIS"] = df["CHASIS"].fillna("").astype(str).str.strip()
    df["NOMBRE CLIENTE"] = df["NOMBRE CLIENTE"].fillna("").astype(str).str.strip()
    df["TELEFONO"] = df["TELEFONO"].apply(normalize_phone)
    df["MARCA"] = df["MARCA"].fillna("").astype(str).str.strip().str.upper()
    df["MODELO"] = df["MODELO"].fillna("").astype(str).str.strip()
    df["CODIGO"] = df["CODIGO"].fillna("").astype(str).str.strip()
    df["REPUESTOS SOLICITADO"] = df["REPUESTOS SOLICITADO"].fillna("").astype(str).str.strip()
    df["VALOR"] = df["VALOR"].apply(parse_valor)
    df["COMPRADO"] = df["COMPRADO"].apply(normalize_yes_no)
    df["MOTIVO"] = df["MOTIVO"].apply(normalize_motivo)
    df["COMENTARIOS"] = df["COMENTARIOS"].fillna("").astype(str).str.strip()

    df.loc[df["COMPRADO"] == "SI", "MOTIVO"] = ""
    df.loc[df["CANAL"] != "Siniestro", "COMPAÃ‘IA"] = ""

    df = df[
        ~(
            df["CHASIS"].eq("")
            & df["NOMBRE CLIENTE"].eq("")
            & df["REPUESTOS SOLICITADO"].eq("")
            & df["CODIGO"].eq("")
            & df["VALOR"].eq(0)
        )
    ].copy()

    return df.reset_index(drop=True)


def load_analytics_data() -> pd.DataFrame:
    df = load_db_data()
    if df.empty:
        return pd.DataFrame(columns=[
            "ID", "EMPRESA", "FECHA", "CANAL", "COMPAÃ‘IA", "NÂ° SINIESTRO", "CHASIS", "NOMBRE CLIENTE",
            "TELEFONO", "MARCA_ORIG", "MARCA_CAT", "MODELO", "CODIGO", "REPUESTOS SOLICITADO",
            "VALOR", "COMPRADO", "MOTIVO", "COMENTARIOS", "CLIENTE_SEGMENTO", "CREATED_BY"
        ])

    out = pd.DataFrame({
        "ID": df["id"],
        "EMPRESA": df["empresa"].fillna("").astype(str).str.upper(),
        "FECHA": pd.to_datetime(df["fecha"], errors="coerce"),
        "CANAL": df["canal"].fillna(""),
        "COMPAÃ‘IA": df["compania"].fillna(""),
        "NÂ° SINIESTRO": df["numero_siniestro"].fillna(""),
        "CHASIS": df["chasis"].fillna(""),
        "NOMBRE CLIENTE": df["nombre_cliente"].fillna(""),
        "TELEFONO": df["telefono"].apply(normalize_phone),
        "MARCA_ORIG": df["marca"].fillna("").astype(str).str.upper(),
        "MODELO": df["modelo"].fillna(""),
        "CODIGO": df["codigo"].fillna(""),
        "REPUESTOS SOLICITADO": df["repuesto"].fillna(""),
        "VALOR": pd.to_numeric(df["valor"], errors="coerce").fillna(0.0),
        "COMPRADO": df["comprado"].fillna("NO"),
        "MOTIVO": df["motivo"].fillna(""),
        "COMENTARIOS": df["comentarios"].fillna(""),
        "CREATED_BY": df["created_by"].fillna(""),
    })
    out["MARCA_CAT"] = out["MARCA_ORIG"].replace("", "SIN MARCA")
    out["CLIENTE_SEGMENTO"] = out.apply(lambda r: infer_client_segment(r["NOMBRE CLIENTE"], r["CANAL"]), axis=1)
    return out


def build_conversion_table(df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if df.empty or group_col not in df.columns:
        return pd.DataFrame()
    out = (
        df.groupby(group_col)["COMPRADO"]
        .value_counts()
        .unstack(fill_value=0)
        .reindex(columns=["SI", "NO"], fill_value=0)
    )
    out["TOTAL"] = out["SI"] + out["NO"]
    out["TASA (%)"] = ((out["SI"] / out["TOTAL"].replace(0, pd.NA)) * 100).fillna(0).round(1)
    return out.sort_values("TOTAL", ascending=False)


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "FECHA" not in df.columns:
        return pd.DataFrame()
    temp = df[df["FECHA"].notna()].copy()
    if temp.empty:
        return pd.DataFrame()
    temp["MES"] = temp["FECHA"].dt.to_period("M").astype(str)
    out = (
        temp.groupby("MES", as_index=False)
        .agg(
            LEADS=("COMPRADO", "size"),
            VALOR_TOTAL=("VALOR", "sum"),
            GANADAS=("COMPRADO", lambda s: (s.astype(str).str.upper() == "SI").sum()),
        )
    )
    out["CONVERSION_%"] = ((out["GANADAS"] / out["LEADS"]) * 100).round(1)
    return out.sort_values("MES")


def daily_summary(df: pd.DataFrame, limit: int = 21) -> pd.DataFrame:
    if df.empty or "FECHA" not in df.columns:
        return pd.DataFrame()
    temp = df[df["FECHA"].notna()].copy()
    if temp.empty:
        return pd.DataFrame()
    temp["FECHA_DIA"] = temp["FECHA"].dt.normalize()
    out = (
        temp.groupby("FECHA_DIA", as_index=False)
        .agg(
            LEADS=("COMPRADO", "size"),
            VALOR_TOTAL=("VALOR", "sum"),
            GANADAS=("COMPRADO", lambda s: (s.astype(str).str.upper() == "SI").sum()),
        )
        .sort_values("FECHA_DIA")
        .tail(limit)
    )
    out["DIA"] = out["FECHA_DIA"].dt.strftime("%d/%m")
    return out


def top_label(series: pd.Series, default_text="Sin datos") -> str:
    if series.empty:
        return default_text
    return str(series.idxmax())


def build_count_value_summary(df: pd.DataFrame, group_col: str, top_n: Optional[int] = None) -> pd.DataFrame:
    if df.empty or group_col not in df.columns:
        return pd.DataFrame()
    temp = df.copy()
    temp[group_col] = temp[group_col].fillna("").astype(str).str.strip()
    temp = temp[temp[group_col] != ""].copy()
    if temp.empty:
        return pd.DataFrame()
    out = (
        temp.groupby(group_col, as_index=False)
        .agg(CASOS=("COMPRADO", "size"), VALOR=("VALOR", "sum"))
        .sort_values(["CASOS", "VALOR"], ascending=[False, False])
    )
    if top_n:
        out = out.head(top_n)
    return out


def build_client_ranking(good_df: pd.DataFrame) -> pd.DataFrame:
    if good_df.empty:
        return pd.DataFrame()
    temp = good_df[good_df["NOMBRE CLIENTE"].astype(str).str.strip() != ""].copy()
    if temp.empty:
        return pd.DataFrame()
    out = (
        temp.groupby(["NOMBRE CLIENTE", "CLIENTE_SEGMENTO"], as_index=False)
        .agg(
            COMPRAS=("COMPRADO", "size"),
            VALOR_COMPRADO=("VALOR", "sum"),
            TICKET_PROM=("VALOR", "mean"),
        )
        .sort_values(["VALOR_COMPRADO", "COMPRAS"], ascending=[False, False])
    )
    out["TICKET_PROM"] = out["TICKET_PROM"].round(2)
    return out


def build_siniestro_ranking(solo_siniestros: pd.DataFrame) -> pd.DataFrame:
    if solo_siniestros.empty:
        return pd.DataFrame()
    temp = solo_siniestros.copy()
    temp["NÂ° SINIESTRO"] = temp["NÂ° SINIESTRO"].fillna("").astype(str).str.strip()
    temp = temp[temp["NÂ° SINIESTRO"] != ""].copy()
    if temp.empty:
        return pd.DataFrame()
    out = (
        temp.groupby("NÂ° SINIESTRO", as_index=False)
        .agg(
            REPUESTOS=("REPUESTOS SOLICITADO", "count"),
            VALOR_TOTAL=("VALOR", "sum"),
            CLIENTE=("NOMBRE CLIENTE", "first"),
            CHASIS=("CHASIS", "first"),
            COMPANIA=("COMPAÃ‘IA", "first"),
            CLIENTE_SEGMENTO=("CLIENTE_SEGMENTO", "first"),
        )
        .sort_values("VALOR_TOTAL", ascending=False)
    )
    return out


def summarize_brand_mix(series: pd.Series) -> str:
    temp = series.fillna("").astype(str).str.strip().str.upper()
    counts = temp[temp != ""].value_counts()
    if counts.empty:
        return "Sin marca"
    return " | ".join(f"{brand}: {count}" for brand, count in counts.items())


def build_taller_siniestro_ranking(seguros_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["COMPAÃ‘IA", "NOMBRE CLIENTE", "LEADS", "GANADOS", "PERDIDOS", "EN_PROCESO", "VALOR_GANADO", "MARCAS", "ETIQUETA"]
    if seguros_df.empty:
        return pd.DataFrame(columns=columns)
    temp = seguros_df.copy()
    temp["COMPAÃ‘IA"] = temp["COMPAÃ‘IA"].fillna("").astype(str).str.strip()
    temp["NOMBRE CLIENTE"] = temp["NOMBRE CLIENTE"].fillna("").astype(str).str.strip().replace("", "Cliente no informado")
    temp["MARCA_ORIG"] = temp["MARCA_ORIG"].fillna("").astype(str).str.strip().str.upper()
    temp = temp[temp["COMPAÃ‘IA"] != ""].copy()
    if temp.empty:
        return pd.DataFrame(columns=columns)
    temp["GANADOS"] = (temp["COMPRADO"].astype(str).str.upper() == "SI").astype(int)
    temp["PERDIDOS"] = (temp["COMPRADO"].astype(str).str.upper() == "NO").astype(int)
    temp["EN_PROCESO"] = (temp["COMPRADO"].astype(str).str.upper() == "EN PROCESO").astype(int)
    temp["VALOR_GANADO"] = temp["VALOR"].where(temp["GANADOS"] == 1, 0.0)
    out = (
        temp.groupby(["COMPAÃ‘IA", "NOMBRE CLIENTE"], as_index=False)
        .agg(
            LEADS=("COMPRADO", "size"),
            GANADOS=("GANADOS", "sum"),
            PERDIDOS=("PERDIDOS", "sum"),
            EN_PROCESO=("EN_PROCESO", "sum"),
            VALOR_GANADO=("VALOR_GANADO", "sum"),
            MARCAS=("MARCA_ORIG", summarize_brand_mix),
        )
        .sort_values(["GANADOS", "PERDIDOS", "VALOR_GANADO", "LEADS"], ascending=[False, False, False, False])
    )
    out["VALOR_GANADO"] = out["VALOR_GANADO"].round(2)
    out["ETIQUETA"] = out["COMPAÃ‘IA"] + " - " + out["NOMBRE CLIENTE"]
    return out


def build_insurer_ticket_summary(seguros_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["COMPAÃ‘IA", "LEADS", "CLIENTES_UNICOS", "GANADOS", "PERDIDOS", "EN_PROCESO", "VALOR_GANADO", "MARCAS"]
    if seguros_df.empty:
        return pd.DataFrame(columns=columns)
    temp = seguros_df.copy()
    temp["COMPAÃ‘IA"] = temp["COMPAÃ‘IA"].fillna("").astype(str).str.strip()
    temp["NOMBRE CLIENTE"] = temp["NOMBRE CLIENTE"].fillna("").astype(str).str.strip().replace("", "Cliente no informado")
    temp["MARCA_ORIG"] = temp["MARCA_ORIG"].fillna("").astype(str).str.strip().str.upper()
    temp = temp[temp["COMPAÃ‘IA"] != ""].copy()
    if temp.empty:
        return pd.DataFrame(columns=columns)
    temp["GANADOS"] = (temp["COMPRADO"].astype(str).str.upper() == "SI").astype(int)
    temp["PERDIDOS"] = (temp["COMPRADO"].astype(str).str.upper() == "NO").astype(int)
    temp["EN_PROCESO"] = (temp["COMPRADO"].astype(str).str.upper() == "EN PROCESO").astype(int)
    temp["VALOR_GANADO"] = temp["VALOR"].where(temp["GANADOS"] == 1, 0.0)
    out = (
        temp.groupby("COMPAÃ‘IA", as_index=False)
        .agg(
            LEADS=("COMPRADO", "size"),
            CLIENTES_UNICOS=("NOMBRE CLIENTE", "nunique"),
            GANADOS=("GANADOS", "sum"),
            PERDIDOS=("PERDIDOS", "sum"),
            EN_PROCESO=("EN_PROCESO", "sum"),
            VALOR_GANADO=("VALOR_GANADO", "sum"),
            MARCAS=("MARCA_ORIG", summarize_brand_mix),
        )
        .sort_values(["GANADOS", "PERDIDOS", "VALOR_GANADO", "LEADS"], ascending=[False, False, False, False])
    )
    out["VALOR_GANADO"] = out["VALOR_GANADO"].round(2)
    return out.reindex(columns=columns)


def build_insurance_brand_summary(seguros_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["COMPAÃ‘IA", "MARCA_ORIG", "LEADS", "GANADOS", "PERDIDOS", "EN_PROCESO", "VALOR_GANADO", "ETIQUETA"]
    if seguros_df.empty:
        return pd.DataFrame(columns=columns)
    temp = seguros_df.copy()
    temp["COMPAÃ‘IA"] = temp["COMPAÃ‘IA"].fillna("").astype(str).str.strip()
    temp["MARCA_ORIG"] = temp["MARCA_ORIG"].fillna("").astype(str).str.strip().str.upper()
    temp = temp[(temp["COMPAÃ‘IA"] != "") & (temp["MARCA_ORIG"] != "")].copy()
    if temp.empty:
        return pd.DataFrame(columns=columns)
    temp["GANADOS"] = (temp["COMPRADO"].astype(str).str.upper() == "SI").astype(int)
    temp["PERDIDOS"] = (temp["COMPRADO"].astype(str).str.upper() == "NO").astype(int)
    temp["EN_PROCESO"] = (temp["COMPRADO"].astype(str).str.upper() == "EN PROCESO").astype(int)
    temp["VALOR_GANADO"] = temp["VALOR"].where(temp["GANADOS"] == 1, 0.0)
    out = (
        temp.groupby(["COMPAÃ‘IA", "MARCA_ORIG"], as_index=False)
        .agg(
            LEADS=("COMPRADO", "size"),
            GANADOS=("GANADOS", "sum"),
            PERDIDOS=("PERDIDOS", "sum"),
            EN_PROCESO=("EN_PROCESO", "sum"),
            VALOR_GANADO=("VALOR_GANADO", "sum"),
        )
        .sort_values(["GANADOS", "PERDIDOS", "VALOR_GANADO", "LEADS"], ascending=[False, False, False, False])
    )
    out["VALOR_GANADO"] = out["VALOR_GANADO"].round(2)
    out["ETIQUETA"] = out["COMPAÃ‘IA"] + " - " + out["MARCA_ORIG"]
    return out


def build_insurance_brand_totals(seguros_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["MARCA_ORIG", "LEADS", "GANADOS", "PERDIDOS", "EN_PROCESO", "VALOR_GANADO", "CONVERSION_%"]
    base = pd.DataFrame({"MARCA_ORIG": ["MAZDA", "KIA"]})
    if seguros_df.empty:
        empty = base.copy()
        for col in columns[1:]:
            empty[col] = 0.0 if col in {"VALOR_GANADO", "CONVERSION_%"} else 0
        return empty[columns]
    temp = seguros_df.copy()
    temp["MARCA_ORIG"] = temp["MARCA_ORIG"].fillna("").astype(str).str.strip().str.upper()
    temp = temp[temp["MARCA_ORIG"].isin(["MAZDA", "KIA"])].copy()
    if temp.empty:
        empty = base.copy()
        for col in columns[1:]:
            empty[col] = 0.0 if col in {"VALOR_GANADO", "CONVERSION_%"} else 0
        return empty[columns]
    temp["GANADOS"] = (temp["COMPRADO"].astype(str).str.upper() == "SI").astype(int)
    temp["PERDIDOS"] = (temp["COMPRADO"].astype(str).str.upper() == "NO").astype(int)
    temp["EN_PROCESO"] = (temp["COMPRADO"].astype(str).str.upper() == "EN PROCESO").astype(int)
    temp["VALOR_GANADO"] = temp["VALOR"].where(temp["GANADOS"] == 1, 0.0)
    out = (
        temp.groupby("MARCA_ORIG", as_index=False)
        .agg(
            LEADS=("COMPRADO", "size"),
            GANADOS=("GANADOS", "sum"),
            PERDIDOS=("PERDIDOS", "sum"),
            EN_PROCESO=("EN_PROCESO", "sum"),
            VALOR_GANADO=("VALOR_GANADO", "sum"),
        )
    )
    out = base.merge(out, on="MARCA_ORIG", how="left").fillna(0)
    for col in ["LEADS", "GANADOS", "PERDIDOS", "EN_PROCESO"]:
        out[col] = out[col].astype(int)
    out["VALOR_GANADO"] = out["VALOR_GANADO"].astype(float).round(2)
    out["CONVERSION_%"] = np.where(
        (out["GANADOS"] + out["PERDIDOS"]) > 0,
        ((out["GANADOS"] / (out["GANADOS"] + out["PERDIDOS"])) * 100).round(1),
        0.0,
    )
    return out[columns]


def classify_invoice_status(series: pd.Series) -> str:
    statuses = {value for value in series.fillna("").astype(str).str.strip().str.upper() if value}
    if not statuses:
        return "EN PROCESO"
    if statuses == {"SI"}:
        return "GANADA"
    if statuses == {"NO"}:
        return "PERDIDA"
    if statuses == {"EN PROCESO"}:
        return "EN PROCESO"
    return "MIXTA"


def classify_invoice_brand(series: pd.Series) -> str:
    brands = sorted({value for value in series.fillna("").astype(str).str.strip().str.upper() if value})
    if not brands:
        return "SIN MARCA"
    if len(brands) == 1:
        return brands[0]
    return "MIXTA"


def first_existing_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def build_insurance_invoice_base(seguros_df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "FACTURA_ID",
        "COMPAÃ‘IA",
        "NOMBRE CLIENTE",
        "MARCA_FACTURA",
        "ESTADO_FACTURA",
        "REPUESTOS",
        "REP_GANADOS",
        "REP_PERDIDOS",
        "REP_EN_PROCESO",
        "VALOR_TOTAL",
        "VALOR_GANADO",
        "MARCAS",
        "ETIQUETA",
    ]
    if seguros_df.empty:
        return pd.DataFrame(columns=columns)
    temp = seguros_df.copy()
    compania_col = first_existing_column(temp, ["COMPAÃ‘IA", "COMPAÑIA"])
    siniestro_col = first_existing_column(temp, ["NÂ° SINIESTRO", "N° SINIESTRO"])
    if compania_col is None:
        return pd.DataFrame(columns=columns)
    temp[compania_col] = temp[compania_col].fillna("").astype(str).str.strip()
    temp["NOMBRE CLIENTE"] = temp["NOMBRE CLIENTE"].fillna("").astype(str).str.strip().replace("", "Cliente no informado")
    temp["MARCA_ORIG"] = temp["MARCA_ORIG"].fillna("").astype(str).str.strip().str.upper()
    if siniestro_col is not None:
        temp[siniestro_col] = temp[siniestro_col].fillna("").astype(str).str.strip()
    temp = temp[temp[compania_col] != ""].copy()
    if temp.empty:
        return pd.DataFrame(columns=columns)
    temp["FACTURA_ID"] = temp[siniestro_col] if siniestro_col is not None else ""
    temp.loc[temp["FACTURA_ID"] == "", "FACTURA_ID"] = "FILA-" + temp.index.astype(str)
    temp["REP_GANADOS"] = (temp["COMPRADO"].astype(str).str.upper() == "SI").astype(int)
    temp["REP_PERDIDOS"] = (temp["COMPRADO"].astype(str).str.upper() == "NO").astype(int)
    temp["REP_EN_PROCESO"] = (temp["COMPRADO"].astype(str).str.upper() == "EN PROCESO").astype(int)
    temp["VALOR_GANADO"] = temp["VALOR"].where(temp["REP_GANADOS"] == 1, 0.0)
    out = (
        temp.groupby("FACTURA_ID", as_index=False)
        .agg(
            COMPANIA=(compania_col, "first"),
            NOMBRE_CLIENTE=("NOMBRE CLIENTE", "first"),
            MARCA_FACTURA=("MARCA_ORIG", classify_invoice_brand),
            ESTADO_FACTURA=("COMPRADO", classify_invoice_status),
            REPUESTOS=("COMPRADO", "size"),
            REP_GANADOS=("REP_GANADOS", "sum"),
            REP_PERDIDOS=("REP_PERDIDOS", "sum"),
            REP_EN_PROCESO=("REP_EN_PROCESO", "sum"),
            VALOR_TOTAL=("VALOR", "sum"),
            VALOR_GANADO=("VALOR_GANADO", "sum"),
            MARCAS=("MARCA_ORIG", summarize_brand_mix),
        )
        .sort_values(["VALOR_TOTAL", "REPUESTOS"], ascending=[False, False])
    )
    out = out.rename(columns={"NOMBRE_CLIENTE": "NOMBRE CLIENTE", "COMPANIA": "COMPAÃ‘IA"})
    out["VALOR_TOTAL"] = out["VALOR_TOTAL"].round(2)
    out["VALOR_GANADO"] = out["VALOR_GANADO"].round(2)
    out["ETIQUETA"] = out["COMPAÃ‘IA"] + " - " + out["NOMBRE CLIENTE"]
    return out[columns]


def build_insurance_brand_dual_summary(seguros_df: pd.DataFrame) -> pd.DataFrame:
    invoice_base = build_insurance_invoice_base(seguros_df)
    base_brands = ["MAZDA", "KIA"]
    if not invoice_base.empty and "MIXTA" in invoice_base["MARCA_FACTURA"].values:
        base_brands.append("MIXTA")
    columns = [
        "MARCA",
        "FACTURAS",
        "FACT_GANADAS",
        "FACT_PERDIDAS",
        "FACT_EN_PROCESO",
        "FACT_MIXTAS",
        "REPUESTOS",
        "REP_GANADOS",
        "REP_PERDIDOS",
        "REP_EN_PROCESO",
        "VALOR_GANADO",
        "TICKET_FACTURA",
    ]
    base = pd.DataFrame({"MARCA": base_brands})
    if seguros_df.empty:
        empty = base.copy()
        for col in columns[1:]:
            empty[col] = 0.0 if col in {"VALOR_GANADO", "TICKET_FACTURA"} else 0
        return empty[columns]
    temp = seguros_df.copy()
    temp["MARCA_ORIG"] = temp["MARCA_ORIG"].fillna("").astype(str).str.strip().str.upper()
    temp["REP_GANADOS"] = (temp["COMPRADO"].astype(str).str.upper() == "SI").astype(int)
    temp["REP_PERDIDOS"] = (temp["COMPRADO"].astype(str).str.upper() == "NO").astype(int)
    temp["REP_EN_PROCESO"] = (temp["COMPRADO"].astype(str).str.upper() == "EN PROCESO").astype(int)
    temp["VALOR_GANADO"] = temp["VALOR"].where(temp["REP_GANADOS"] == 1, 0.0)
    line_out = (
        temp[temp["MARCA_ORIG"].isin(base_brands)]
        .groupby("MARCA_ORIG", as_index=False)
        .agg(
            REPUESTOS=("COMPRADO", "size"),
            REP_GANADOS=("REP_GANADOS", "sum"),
            REP_PERDIDOS=("REP_PERDIDOS", "sum"),
            REP_EN_PROCESO=("REP_EN_PROCESO", "sum"),
            VALOR_GANADO=("VALOR_GANADO", "sum"),
        )
    )
    if invoice_base.empty:
        invoice_out = pd.DataFrame(columns=["MARCA_FACTURA", "FACTURAS", "FACT_GANADAS", "FACT_PERDIDAS", "FACT_EN_PROCESO", "FACT_MIXTAS"])
    else:
        invoice_out = (
            invoice_base.groupby("MARCA_FACTURA", as_index=False)
            .agg(
                FACTURAS=("FACTURA_ID", "size"),
                FACT_GANADAS=("ESTADO_FACTURA", lambda s: (s == "GANADA").sum()),
                FACT_PERDIDAS=("ESTADO_FACTURA", lambda s: (s == "PERDIDA").sum()),
                FACT_EN_PROCESO=("ESTADO_FACTURA", lambda s: (s == "EN PROCESO").sum()),
                FACT_MIXTAS=("ESTADO_FACTURA", lambda s: (s == "MIXTA").sum()),
            )
        )
    out = base.merge(invoice_out, left_on="MARCA", right_on="MARCA_FACTURA", how="left").drop(columns=["MARCA_FACTURA"], errors="ignore")
    out = out.merge(line_out, left_on="MARCA", right_on="MARCA_ORIG", how="left").drop(columns=["MARCA_ORIG"], errors="ignore")
    out = out.fillna(0)
    for col in ["FACTURAS", "FACT_GANADAS", "FACT_PERDIDAS", "FACT_EN_PROCESO", "FACT_MIXTAS", "REPUESTOS", "REP_GANADOS", "REP_PERDIDOS", "REP_EN_PROCESO"]:
        out[col] = out[col].astype(int)
    out["VALOR_GANADO"] = out["VALOR_GANADO"].astype(float).round(2)
    out["TICKET_FACTURA"] = np.where(
        out["FACT_GANADAS"] > 0,
        (out["VALOR_GANADO"] / out["FACT_GANADAS"]).round(2),
        0.0,
    )
    return out[columns]


def build_taller_market_share(good_df: pd.DataFrame) -> dict:
    valor_taller_magna = float(good_df.loc[good_df["CLIENTE_SEGMENTO"] == "Taller Magna", "VALOR"].sum())
    valor_resto_talleres = float(good_df.loc[good_df["CLIENTE_SEGMENTO"] == "Mercado Talleres", "VALOR"].sum())
    compras_taller_magna = int((good_df["CLIENTE_SEGMENTO"] == "Taller Magna").sum())
    compras_resto_talleres = int((good_df["CLIENTE_SEGMENTO"] == "Mercado Talleres").sum())
    total = valor_taller_magna + valor_resto_talleres
    share_taller_magna = round((valor_taller_magna / total) * 100, 1) if total > 0 else 0.0
    share_resto_talleres = round((valor_resto_talleres / total) * 100, 1) if total > 0 else 0.0
    return {
        "valor_taller_magna": valor_taller_magna,
        "valor_resto_talleres": valor_resto_talleres,
        "share_taller_magna": share_taller_magna,
        "share_resto_talleres": share_resto_talleres,
        "compras_taller_magna": compras_taller_magna,
        "compras_resto_talleres": compras_resto_talleres,
    }


def horizontal_bar(df: pd.DataFrame, category_col: str, value_col: str):
    base = alt.Chart(df).encode(
        y=alt.Y(f"{category_col}:N", sort="-x", title=None),
        x=alt.X(f"{value_col}:Q", title=None),
        tooltip=[category_col, value_col],
    )
    bars = base.mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6)
    text = base.mark_text(align="left", baseline="middle", dx=6).encode(
        text=alt.Text(f"{value_col}:Q", format=",.0f")
    )
    return (bars + text).properties(height=max(220, len(df) * 28))


def dashboard_chart(chart, height: int = 240):
    return (
        chart.properties(height=height, background="#0a2244")
        .configure_view(strokeWidth=0)
        .configure_axis(
            labelColor="#dbeafe",
            titleColor="#bfdbfe",
            domain=False,
            gridColor="rgba(191,219,254,0.14)",
            tickColor="rgba(191,219,254,0.14)",
        )
        .configure_legend(
            titleColor="#bfdbfe",
            labelColor="#dbeafe",
            symbolStrokeColor="#dbeafe",
        )
        .configure_title(color="#f8fafc", fontSize=16)
    )


def donut_chart(df: pd.DataFrame, category_col: str, value_col: str, colors: list[str]):
    if df.empty:
        return None
    chart = alt.Chart(df).mark_arc(innerRadius=58, outerRadius=92).encode(
        theta=alt.Theta(f"{value_col}:Q"),
        color=alt.Color(
            f"{category_col}:N",
            scale=alt.Scale(range=colors),
            legend=alt.Legend(title=None, orient="bottom"),
        ),
        tooltip=[category_col, value_col],
        order=alt.Order(f"{value_col}:Q", sort="descending"),
    )
    return dashboard_chart(chart, height=240)


def conversion_status_html(rate: float) -> str:
    if rate >= 70:
        return f'<div class="status-good">Semaforo comercial: Excelente - {rate:.1f}%</div>'
    if rate >= 40:
        return f'<div class="status-mid">Semaforo comercial: Medio - {rate:.1f}%</div>'
    return f'<div class="status-bad">Semaforo comercial: Bajo - {rate:.1f}%</div>'


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Datos") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


# =========================================================
# LOGIN
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user" not in st.session_state:
    st.session_state.user = None


def login_screen():
    st.markdown(
        """
        <div class="login-card">
            <h1 style="margin-top:0; color:#0f172a;">ALI Leads <span style="color:#1d4ed8;">Pro v5</span></h1>
            <p style="color:#475569;">Ingresa para administrar la base y ver el dashboard comercial.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("Usuario")
        password = st.text_input("Contrasena", type="password")
        submitted = st.form_submit_button("Ingresar", use_container_width=True)
        if submitted:
            user = authenticate_user(username, password)
            if user:
                st.session_state.authenticated = True
                st.session_state.user = user
                st.rerun()
            else:
                st.error("Usuario o contrasena incorrectos.")


if not st.session_state.authenticated:
    login_screen()
    st.stop()

user = st.session_state.user


# =========================================================
# SIDEBAR ADMIN / CARGA
# =========================================================
st.sidebar.success(f"Sesion: {user['full_name']} ({user['username']})")
st.sidebar.caption(f"Rol: {user['role']} | Alcance: {user['company_scope']}")
with st.sidebar.expander("Diagnostico", expanded=False):
    st.caption(f"Script activo: {os.path.abspath(__file__)}")
    st.caption(f"Base activa: {os.path.abspath(DB_PATH)}")

if st.sidebar.button("Cerrar sesion", use_container_width=True):
    st.session_state.authenticated = False
    st.session_state.user = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("Base de datos")

allowed_companies = TIPOS_EMPRESA if user["company_scope"] == "TODAS" else [user["company_scope"]]

load_mode = st.sidebar.radio(
    "Modo de carga",
    ["Usar base de datos", "Importar Excel a base de datos"],
    index=0,
    key="load_mode_sidebar",
)

if load_mode == "Importar Excel a base de datos":
    uploaded_file = st.sidebar.file_uploader("Subir Excel/CSV", type=["xlsx", "xls", "csv"], key="import_file_sidebar")
    import_strategy = st.sidebar.radio(
        "Importacion",
        ["Reemplazar empresa", "Agregar registros"],
        index=0,
        key="import_strategy_sidebar",
    )
    forced_company = st.sidebar.selectbox(
        "Empresa destino",
        allowed_companies,
        index=0,
        key="forced_company_sidebar",
    )

    if uploaded_file is not None:
        try:
            raw_df = read_excel_smart(uploaded_file) if uploaded_file.name.lower().endswith((".xlsx", ".xls")) else pd.read_csv(uploaded_file)
            detected_mode = infer_business_mode_from_filename(uploaded_file.name)
            if detected_mode in allowed_companies:
                forced_company = detected_mode
            clean_df = clean_input_dataframe(raw_df, forced_company=forced_company)
            st.sidebar.info(f"Filas validas detectadas: {len(clean_df)}")

            with st.sidebar.expander("Vista previa importacion", expanded=False):
                st.dataframe(clean_df.head(10), use_container_width=True, hide_index=True)

            if st.sidebar.button("Guardar Excel en base", use_container_width=True):
                if import_strategy == "Reemplazar empresa":
                    replace_company_leads(forced_company, clean_df, user["username"])
                else:
                    append_company_leads(clean_df, user["username"])
                st.sidebar.success("Datos guardados en la base.")
        except Exception as e:
            st.sidebar.error(f"No se pudo importar el archivo: {e}")

st.sidebar.markdown("---")
st.sidebar.subheader("Carga manual")
with st.sidebar.expander("Agregar lead manual", expanded=False):
    with st.form("manual_insert_form", clear_on_submit=True):
        empresa = st.selectbox("Empresa", allowed_companies)
        fecha = st.date_input("Fecha", value=datetime.today())
        canal = st.selectbox("Canal", TIPOS_CANAL)
        compania = st.selectbox("Compania", TIPOS_COMPANIA)
        numero_siniestro = st.text_input("Nro. Siniestro")
        chasis = st.text_input("Chasis")
        nombre_cliente = st.text_input("Nombre cliente")
        telefono = st.text_input("Telefono")
        marca = st.text_input("Marca")
        modelo = st.text_input("Modelo")
        codigo = st.text_input("Codigo")
        repuesto = st.text_input("Repuesto solicitado")
        valor = st.number_input("Valor", min_value=0.0, step=1.0, value=0.0)
        comprado = st.selectbox("Comprado", TIPOS_COMPRADO)
        motivo = st.selectbox("Motivo", TIPOS_MOTIVO)
        comentarios = st.text_area("Comentarios")
        submitted_manual = st.form_submit_button("Guardar lead")
        if submitted_manual:
            record = {
                "EMPRESA": empresa,
                "FECHA": fecha.strftime("%Y-%m-%d"),
                "CANAL": canal,
                "COMPAÃ‘IA": compania if canal == "Siniestro" else "",
                "NÂ° SINIESTRO": numero_siniestro,
                "CHASIS": chasis,
                "NOMBRE CLIENTE": nombre_cliente,
                "TELEFONO": telefono,
                "MARCA": marca.upper().strip(),
                "MODELO": modelo,
                "CODIGO": codigo,
                "REPUESTOS SOLICITADO": repuesto,
                "VALOR": valor,
                "COMPRADO": comprado,
                "MOTIVO": "" if comprado == "SI" else motivo,
                "COMENTARIOS": comentarios,
                "CREATED_BY": user["username"],
            }
            insert_lead(record)
            st.sidebar.success("Lead guardado en base.")
            st.rerun()

if user["role"] == "admin":
    st.sidebar.markdown("---")
    st.sidebar.subheader("Administracion")

    with st.sidebar.expander("Crear usuario", expanded=False):
        with st.form("create_user_form", clear_on_submit=True):
            new_username = st.text_input("Nuevo usuario")
            new_password = st.text_input("Contrasena", type="password")
            full_name = st.text_input("Nombre completo")
            role = st.selectbox("Rol", ["user", "admin"])
            company_scope = st.selectbox("Alcance empresa", ["TODAS"] + TIPOS_EMPRESA)
            create_submit = st.form_submit_button("Crear usuario")
            if create_submit:
                if not new_username or not new_password:
                    st.sidebar.error("Usuario y contrasena son obligatorios.")
                else:
                    ok, msg = create_user(new_username, new_password, full_name, role, company_scope)
                    if ok:
                        st.sidebar.success(msg)
                    else:
                        st.sidebar.error(msg)


# =========================================================
# DATA DESDE DB
# =========================================================
data = load_analytics_data()
if data.empty:
    st.warning("La base de datos esta vacia. Carga un Excel o agrega leads manualmente.")
    st.stop()

if user["company_scope"] != "TODAS":
    data = data[data["EMPRESA"] == user["company_scope"]].copy()


# =========================================================
# FILTRO EMPRESA
# =========================================================
st.sidebar.markdown("---")
if user["company_scope"] == "TODAS":
    empresa_choices = ["TODAS"] + allowed_companies
    empresa_filter = st.sidebar.selectbox("Empresa a analizar", empresa_choices, index=0)
else:
    empresa_filter = user["company_scope"]
    st.sidebar.markdown(f"**Empresa activa:** {empresa_filter}")

filtered_base = data.copy()
if empresa_filter != "TODAS":
    filtered_base = filtered_base[filtered_base["EMPRESA"] == empresa_filter].copy()

business_mode = empresa_filter if empresa_filter in TIPOS_EMPRESA else "GENERAL"
dashboard_subtitle = {
    "MAGNA": "Dashboard comercial enfocado en Mazda y Kia.",
    "ALIMATICO": "Dashboard comercial enfocado en todas las marcas cargadas, incluyendo Mazda y Kia.",
    "GENERAL": "Dashboard comercial general multiempresa.",
}.get(business_mode, "Dashboard comercial general multiempresa.")

if business_mode == "MAGNA":
    filtered_base = filtered_base[filtered_base["MARCA_ORIG"].isin(MARCAS_MAGNA)].copy()

if filtered_base.empty:
    st.warning("No hay datos disponibles para la empresa y filtros elegidos.")
    st.stop()


# =========================================================
# HEADER
# =========================================================
mode_label = {
    "MAGNA": "Modo Magna",
    "ALIMATICO": "Modo Alimatico",
    "GENERAL": "Modo General",
}.get(business_mode, "Modo General")

st.markdown(
    f"""
    <div class="main-title-card">
        <div class="main-title">ALI <span class="main-title-accent">Leads Pro</span> v5</div>
        <div class="subtitle">{dashboard_subtitle}</div>
        <div class="mode-badge">{mode_label}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero-card">
        <div class="hero-title">Panel ejecutivo comercial</div>
        <div class="hero-subtitle">
            Seguimiento de conversion, valor ganado, perdidas, comparacion mensual, alertas, clientes, talleres y siniestros.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)



# =========================================================
# FILTROS ANALITICOS
# =========================================================
st.sidebar.markdown("---")
st.sidebar.subheader("Filtros analiticos")

buscar_siniestro = st.sidebar.text_input("Buscar Nro. Siniestro")
buscar_cliente = st.sidebar.text_input("Buscar cliente")
buscar_codigo = st.sidebar.text_input("Buscar codigo / repuesto")

marca_options = sorted(filtered_base["MARCA_CAT"].dropna().astype(str).unique().tolist())
canal_options = sorted(filtered_base["CANAL"].dropna().astype(str).unique().tolist())
estado_options = sorted(filtered_base["COMPRADO"].dropna().astype(str).unique().tolist())
compania_options = sorted([x for x in filtered_base["COMPAÃ‘IA"].dropna().astype(str).unique().tolist() if x])

marca_filter = st.sidebar.multiselect("Marca", marca_options, default=marca_options)
canal_filter = st.sidebar.multiselect("Canal", canal_options, default=canal_options)
estado_filter = st.sidebar.multiselect("Estado", estado_options, default=estado_options)

if "Siniestro" in canal_filter or not canal_filter:
    compania_filter = st.sidebar.multiselect("Compania", compania_options, default=compania_options)
else:
    compania_filter = []

filtered = filtered_base.copy()

if filtered["FECHA"].notna().any():
    dmin = filtered["FECHA"].min().date()
    dmax = filtered["FECHA"].max().date()
    fecha_rango = st.sidebar.date_input("Rango de fecha", value=(dmin, dmax))
    if isinstance(fecha_rango, tuple) and len(fecha_rango) == 2:
        filtered = filtered[filtered["FECHA"].dt.date.between(fecha_rango[0], fecha_rango[1])]

if marca_filter:
    filtered = filtered[filtered["MARCA_CAT"].isin(marca_filter)]
if canal_filter:
    filtered = filtered[filtered["CANAL"].isin(canal_filter)]
if estado_filter:
    filtered = filtered[filtered["COMPRADO"].isin(estado_filter)]
if compania_filter:
    filtered = filtered[
        (filtered["CANAL"] != "Siniestro") |
        ((filtered["CANAL"] == "Siniestro") & (filtered["COMPAÃ‘IA"].isin(compania_filter)))
    ]

if buscar_siniestro:
    filtered = filtered[
        filtered["NÂ° SINIESTRO"].astype(str).str.contains(buscar_siniestro.strip(), case=False, na=False)
    ]

if buscar_cliente:
    filtered = filtered[
        filtered["NOMBRE CLIENTE"].astype(str).str.contains(buscar_cliente.strip(), case=False, na=False)
    ]

if buscar_codigo:
    filtered = filtered[
        filtered["CODIGO"].astype(str).str.contains(buscar_codigo.strip(), case=False, na=False)
        | filtered["REPUESTOS SOLICITADO"].astype(str).str.contains(buscar_codigo.strip(), case=False, na=False)
    ]

if filtered.empty:
    st.warning("No hay datos con esos filtros.")
    st.stop()


# =========================================================
# METRICAS
# =========================================================
good = filtered[filtered["COMPRADO"] == "SI"].copy()
bad = filtered[filtered["COMPRADO"] == "NO"].copy()
pending = filtered[filtered["COMPRADO"] == "EN PROCESO"].copy()
seguros_df = filtered[
    (filtered["CANAL"] == "Siniestro")
    & (filtered["COMPAÃ‘IA"].astype(str).str.strip() != "")
].copy()
good_seguros = seguros_df[seguros_df["COMPRADO"] == "SI"].copy()

ranking_talleres_siniestro = build_taller_siniestro_ranking(seguros_df)
insurer_ticket_summary = build_insurer_ticket_summary(seguros_df)
insurance_brand_summary = build_insurance_brand_summary(seguros_df)
insurance_brand_totals = build_insurance_brand_totals(seguros_df)
insurance_invoice_base = build_insurance_invoice_base(seguros_df)
insurance_brand_dual_summary = build_insurance_brand_dual_summary(seguros_df)
client_ranking = build_client_ranking(good)
market_share = build_taller_market_share(good)

mensual = monthly_summary(filtered)
if len(mensual) >= 2:
    actual = mensual.iloc[-1]
    anterior = mensual.iloc[-2]
    delta_leads = int(actual["LEADS"] - anterior["LEADS"])
    delta_valor = float(actual["VALOR_TOTAL"] - anterior["VALOR_TOTAL"])
    delta_conv = float(actual["CONVERSION_%"] - anterior["CONVERSION_%"])
else:
    actual = pd.Series({"MES": "N/D", "LEADS": 0, "VALOR_TOTAL": 0.0, "CONVERSION_%": 0.0})
    delta_leads = 0
    delta_valor = 0.0
    delta_conv = 0.0

tot = len(filtered)
won = len(good)
lost = len(bad)
in_process = len(pending)
conv_rate = round((won / (won + lost)) * 100, 1) if (won + lost) else 0.0

total_valor = float(filtered["VALOR"].sum())
valor_ganado = float(good["VALOR"].sum())
valor_perdido = float(bad["VALOR"].sum())
valor_en_proceso = float(pending["VALOR"].sum())
ticket_prom = round(valor_ganado / won, 2) if won else 0.0
promedio_perdido = round(valor_perdido / lost, 2) if lost else 0.0
promedio_en_proceso = round(valor_en_proceso / in_process, 2) if in_process else 0.0

top_cliente = top_label(
    client_ranking.set_index("NOMBRE CLIENTE")["VALOR_COMPRADO"] if not client_ranking.empty else pd.Series(dtype=float),
    "Sin datos",
)
top_cliente_siniestros = top_label(
    ranking_talleres_siniestro.set_index("ETIQUETA")["GANADOS"] if not ranking_talleres_siniestro.empty else pd.Series(dtype=float),
    "Sin compras en seguros",
)
compania_top = top_label(
    insurer_ticket_summary.set_index("COMPAÃ‘IA")["GANADOS"] if not insurer_ticket_summary.empty else pd.Series(dtype=float),
    "Sin datos",
)
motivo_top = top_label(bad["MOTIVO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
canal_top = top_label(good["CANAL"].value_counts(), "Sin ventas")
marca_top = top_label(good["MARCA_ORIG"].value_counts(), "Sin ventas")
aseguradora_ticket_top = top_label(
    insurance_brand_summary.set_index("ETIQUETA")["GANADOS"] if not insurance_brand_summary.empty else pd.Series(dtype=float),
    "Sin datos",
)
repuesto_top = top_label(good["REPUESTOS SOLICITADO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
producto_perdido_top = top_label(bad["REPUESTOS SOLICITADO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
cliente_mas_perdido = top_label(
    bad.loc[bad["NOMBRE CLIENTE"].astype(str).str.strip() != ""].groupby("NOMBRE CLIENTE")["VALOR"].sum()
    if not bad.empty else pd.Series(dtype=float),
    "Sin datos",
)

aseguradoras_activas = int(seguros_df["COMPAÃ‘IA"].replace("", pd.NA).dropna().nunique()) if not seguros_df.empty else 0
clientes_seguro_unicos = int(seguros_df["NOMBRE CLIENTE"].replace("", pd.NA).dropna().nunique()) if not seguros_df.empty else 0
aseguradora_cliente_unicos = int(len(ranking_talleres_siniestro))
facturas_seguro = int(len(insurance_invoice_base))
facturas_ganadas_seguro = int((insurance_invoice_base["ESTADO_FACTURA"] == "GANADA").sum()) if not insurance_invoice_base.empty else 0
facturas_perdidas_seguro = int((insurance_invoice_base["ESTADO_FACTURA"] == "PERDIDA").sum()) if not insurance_invoice_base.empty else 0
facturas_pendientes_seguro = int((insurance_invoice_base["ESTADO_FACTURA"] == "EN PROCESO").sum()) if not insurance_invoice_base.empty else 0
facturas_mixtas_seguro = int((insurance_invoice_base["ESTADO_FACTURA"] == "MIXTA").sum()) if not insurance_invoice_base.empty else 0
repuestos_seguro = int(len(seguros_df))
repuestos_ganados_seguro = int((seguros_df["COMPRADO"] == "SI").sum()) if not seguros_df.empty else 0
repuestos_perdidos_seguro = int((seguros_df["COMPRADO"] == "NO").sum()) if not seguros_df.empty else 0
repuestos_pendientes_seguro = int((seguros_df["COMPRADO"] == "EN PROCESO").sum()) if not seguros_df.empty else 0
ticket_factura_seguro = round(float(good_seguros["VALOR"].sum()) / facturas_ganadas_seguro, 2) if facturas_ganadas_seguro else 0.0
ranking_talleres_siniestro_display = ranking_talleres_siniestro.rename(
    columns={"LEADS": "REPUESTOS", "GANADOS": "REP_GANADOS", "PERDIDOS": "REP_PERDIDOS", "EN_PROCESO": "REP_EN_PROCESO"}
)
insurer_ticket_summary_display = insurer_ticket_summary.rename(
    columns={"LEADS": "REPUESTOS", "GANADOS": "REP_GANADOS", "PERDIDOS": "REP_PERDIDOS", "EN_PROCESO": "REP_EN_PROCESO", "CLIENTES_UNICOS": "CLIENTES"}
)
insurance_brand_summary_display = insurance_brand_summary.rename(
    columns={"LEADS": "REPUESTOS", "GANADOS": "REP_GANADOS", "PERDIDOS": "REP_PERDIDOS", "EN_PROCESO": "REP_EN_PROCESO"}
)

meta_scope = empresa_filter if empresa_filter else user["company_scope"]
meta_setting_key = f"meta_mensual::{str(meta_scope).upper()}"
meta_widget_key = f"meta_mensual_input::{str(meta_scope).upper()}"
meta_mensual_guardada = get_app_setting_float(meta_setting_key, default=10000.0)
if meta_widget_key not in st.session_state:
    st.session_state[meta_widget_key] = meta_mensual_guardada
meta_mensual = st.sidebar.number_input("Objetivo mensual ($)", min_value=0.0, step=100.0, key=meta_widget_key)
if abs(float(meta_mensual) - float(meta_mensual_guardada)) > 1e-9:
    set_app_setting(meta_setting_key, f"{float(meta_mensual):.2f}")

cumplimiento_meta = round((valor_ganado / meta_mensual) * 100, 1) if meta_mensual > 0 else 0.0

ranking_repuestos_vendidos = (
    good[good["REPUESTOS SOLICITADO"].astype(str).str.strip() != ""]
    .groupby("REPUESTOS SOLICITADO", as_index=False)
    .agg(CANTIDAD=("COMPRADO", "size"), VALOR_VENDIDO=("VALOR", "sum"))
    .sort_values(["CANTIDAD", "VALOR_VENDIDO"], ascending=[False, False])
)

ranking_repuestos_perdidos = (
    bad[bad["REPUESTOS SOLICITADO"].astype(str).str.strip() != ""]
    .groupby("REPUESTOS SOLICITADO", as_index=False)
    .agg(CANTIDAD=("COMPRADO", "size"), VALOR_PERDIDO=("VALOR", "sum"))
    .sort_values(["VALOR_PERDIDO", "CANTIDAD"], ascending=[False, False])
)

resumen_diario = daily_summary(filtered)
resumen_canal = build_count_value_summary(filtered, "CANAL", top_n=6)
resumen_estado = build_count_value_summary(filtered, "COMPRADO")
resumen_marca_resumen = build_count_value_summary(good, "MARCA_ORIG", top_n=4)
conversion_marca_resumen = build_conversion_table(filtered, "MARCA_CAT").reset_index()
top_productos_resumen = ranking_repuestos_vendidos.head(8).copy()

perdidas_cliente = (
    bad[bad["NOMBRE CLIENTE"].astype(str).str.strip() != ""]
    .groupby("NOMBRE CLIENTE", as_index=False)
    .agg(
        MONTO_PERDIDO=("VALOR", "sum"),
        CASOS=("COMPRADO", "size"),
        MOTIVO_PRINCIPAL=("MOTIVO", lambda s: s.mode().iloc[0] if not s.mode().empty else ""),
    )
    .sort_values(["MONTO_PERDIDO", "CASOS"], ascending=[False, False])
)

perdidas_motivo = (
    bad[bad["MOTIVO"].astype(str).str.strip() != ""]
    .groupby("MOTIVO", as_index=False)
    .agg(MONTO_PERDIDO=("VALOR", "sum"), CASOS=("COMPRADO", "size"))
    .sort_values(["MONTO_PERDIDO", "CASOS"], ascending=[False, False])
)

detalle_no_compra = bad[
    [
        "FECHA",
        "EMPRESA",
        "CANAL",
        "COMPAÃ‘IA",
        "NÂ° SINIESTRO",
        "NOMBRE CLIENTE",
        "TELEFONO",
        "MARCA_ORIG",
        "MODELO",
        "CODIGO",
        "REPUESTOS SOLICITADO",
        "VALOR",
        "MOTIVO",
        "COMENTARIOS",
    ]
].copy()


def render_panel_ejecutivo():
    panel_chips = [
        f"Empresa: {empresa_filter if empresa_filter else user['company_scope']}",
        f"Meta: ${meta_mensual:,.0f}",
        f"Canal foco: {canal_top}",
        f"Marca foco: {marca_top}",
        f"Compania foco: {compania_top}",
    ]
    chips_html = "".join(f'<span class="summary-chip">{item}</span>' for item in panel_chips)
    st.markdown(
        f"""
        <div class="summary-hero-card">
            <div class="summary-hero-title">Panel Ejecutivo</div>
            <div class="summary-hero-subtitle">Lectura gerencial rapida para conversion, valor, mix comercial y oportunidades.</div>
            <div class="summary-chip-wrap">{chips_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    panel_kpis = [
        ("Leads analizados", f"{tot}", f"Ganadas: {won} | Perdidas: {lost}"),
        ("Valor ganado", f"${valor_ganado:,.0f}", f"Ticket: ${ticket_prom:,.0f}"),
        ("Conversion", f"{conv_rate:.1f}%", f"Meta lograda: {cumplimiento_meta:.1f}%"),
        ("Facturas seguros", f"{facturas_seguro}", f"Ganadas: {facturas_ganadas_seguro} | Mixtas: {facturas_mixtas_seguro}"),
    ]
    pc1, pc2, pc3, pc4 = st.columns(4)
    for col, (title, value, note) in zip([pc1, pc2, pc3, pc4], panel_kpis):
        with col:
            st.markdown(
                f"""
                <div class="summary-kpi-card">
                    <div class="summary-kpi-title">{title}</div>
                    <div class="summary-kpi-value">{value}</div>
                    <div class="summary-kpi-note">{note}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    row1a, row1b = st.columns([1.55, 1.05])
    with row1a:
        st.markdown('<div class="summary-section-label">Evolucion mensual</div>', unsafe_allow_html=True)
        if not mensual.empty:
            monthly_chart = alt.Chart(mensual).mark_bar(color="#60a5fa", cornerRadiusTopLeft=7, cornerRadiusTopRight=7).encode(
                x=alt.X("MES:N", title=None),
                y=alt.Y("VALOR_TOTAL:Q", title=None),
                tooltip=["MES", "VALOR_TOTAL", "LEADS", "CONVERSION_%"],
            )
            st.altair_chart(dashboard_chart(monthly_chart, 250), use_container_width=True)
        else:
            st.info("No hay datos mensuales.")
    with row1b:
        st.markdown('<div class="summary-section-label">Top productos</div>', unsafe_allow_html=True)
        if not top_productos_resumen.empty:
            prod_chart = alt.Chart(top_productos_resumen).mark_bar(color="#38bdf8", cornerRadiusTopRight=6, cornerRadiusBottomRight=6).encode(
                y=alt.Y("REPUESTOS SOLICITADO:N", sort="-x", title=None),
                x=alt.X("CANTIDAD:Q", title=None),
                tooltip=["REPUESTOS SOLICITADO", "CANTIDAD", "VALOR_VENDIDO"],
            )
            st.altair_chart(dashboard_chart(prod_chart, max(250, len(top_productos_resumen) * 28)), use_container_width=True)
        else:
            st.info("No hay productos vendidos.")

    share_taller_df = pd.DataFrame(
        [
            {"SEGMENTO": "Taller Magna", "VALOR": market_share["valor_taller_magna"], "COMPRAS": market_share["compras_taller_magna"]},
            {"SEGMENTO": "Resto talleres", "VALOR": market_share["valor_resto_talleres"], "COMPRAS": market_share["compras_resto_talleres"]},
        ]
    )

    row2a, row2b, row2c = st.columns([1.45, 0.9, 0.9])
    with row2a:
        st.markdown('<div class="summary-section-label">Evolucion diaria</div>', unsafe_allow_html=True)
        if not resumen_diario.empty:
            orden_dias = resumen_diario["DIA"].tolist()
            area = alt.Chart(resumen_diario).mark_area(color="#1d4ed8", opacity=0.25).encode(
                x=alt.X("DIA:N", sort=orden_dias, title=None),
                y=alt.Y("VALOR_TOTAL:Q", title=None),
                tooltip=["DIA", "VALOR_TOTAL", "LEADS", "GANADAS"],
            )
            line = alt.Chart(resumen_diario).mark_line(color="#93c5fd", strokeWidth=3, point=True).encode(
                x=alt.X("DIA:N", sort=orden_dias, title=None),
                y=alt.Y("VALOR_TOTAL:Q", title=None),
                tooltip=["DIA", "VALOR_TOTAL", "LEADS", "GANADAS"],
            )
            st.altair_chart(dashboard_chart(area + line, 240), use_container_width=True)
        else:
            st.info("No hay datos diarios.")
    with row2b:
        st.markdown('<div class="summary-section-label">Mix por canal</div>', unsafe_allow_html=True)
        canal_chart = donut_chart(resumen_canal, "CANAL", "CASOS", ["#60a5fa", "#2563eb", "#22c55e", "#f59e0b", "#f97316", "#a855f7"])
        if canal_chart is not None:
            st.altair_chart(canal_chart, use_container_width=True)
        else:
            st.info("Sin datos.")
    with row2c:
        st.markdown('<div class="summary-section-label">Share Taller Magna</div>', unsafe_allow_html=True)
        share_chart = donut_chart(share_taller_df, "SEGMENTO", "VALOR", ["#22c55e", "#60a5fa"])
        if share_chart is not None and share_taller_df["VALOR"].sum() > 0:
            st.altair_chart(share_chart, use_container_width=True)
            st.caption(
                f"Taller Magna: {market_share['share_taller_magna']:.1f}% del valor ganado | "
                f"Resto: {market_share['share_resto_talleres']:.1f}%"
            )
            st.caption(
                f"Compras ganadas: {market_share['compras_taller_magna']} vs {market_share['compras_resto_talleres']}"
            )
        else:
            st.info("Sin datos.")

    brand_conversion_exec = conversion_marca_resumen.rename(
        columns={"MARCA_CAT": "MARCA", "SI": "GANADAS", "NO": "PERDIDAS", "TOTAL": "CASOS", "TASA (%)": "CONVERSION_%"}
    )
    row3a, row3b, row3c = st.columns([1.1, 1.1, 0.8])
    with row3a:
        st.markdown('<div class="summary-section-label">Valor ganado por marca</div>', unsafe_allow_html=True)
        if not resumen_marca_resumen.empty:
            marca_chart = alt.Chart(resumen_marca_resumen).mark_bar(color="#f59e0b", cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
                x=alt.X("MARCA_ORIG:N", title=None),
                y=alt.Y("VALOR:Q", title=None),
                tooltip=["MARCA_ORIG", "CASOS", "VALOR"],
            )
            st.altair_chart(dashboard_chart(marca_chart, 220), use_container_width=True)
        else:
            st.info("Sin ventas por marca.")
    with row3b:
        st.markdown('<div class="summary-section-label">Conversion por marca</div>', unsafe_allow_html=True)
        if not brand_conversion_exec.empty:
            conv_chart = alt.Chart(brand_conversion_exec).mark_bar(color="#22c55e", cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
                x=alt.X("MARCA:N", title=None),
                y=alt.Y("CONVERSION_%:Q", title=None),
                tooltip=["MARCA", "GANADAS", "PERDIDAS", "CASOS", "CONVERSION_%"],
            )
            st.altair_chart(dashboard_chart(conv_chart, 220), use_container_width=True)
        else:
            st.info("Sin conversion por marca.")
    with row3c:
        st.markdown('<div class="summary-section-label">Focos del periodo</div>', unsafe_allow_html=True)
        focus_cards = [
            ("Cliente top", top_cliente),
            ("Producto top", repuesto_top),
            ("Seguro top", top_cliente_siniestros),
            ("Motivo perdida", motivo_top),
        ]
        for title, value in focus_cards:
            st.markdown(
                f"""
                <div class="mini-card">
                    <div class="kpi-title">{title}</div>
                    <div class="kpi-value">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("---")
    e1, e2 = st.columns(2)
    with e1:
        st.subheader("Top clientes por monto")
        st.dataframe(client_ranking.head(12), use_container_width=True, hide_index=True)
    with e2:
        st.subheader("Mazda vs Kia")
        st.dataframe(insurance_brand_dual_summary, use_container_width=True, hide_index=True)


# =========================================================
# VISTA
# =========================================================
st.markdown(conversion_status_html(conv_rate), unsafe_allow_html=True)

meta1, meta2 = st.columns([2, 1])
with meta1:
    st.progress(min(cumplimiento_meta / 100, 1.0))
    st.caption(f"Valor ganado: ${valor_ganado:,.2f} | Meta: ${meta_mensual:,.2f} | Cumplimiento: {cumplimiento_meta:.1f}%")
with meta2:
    st.metric("Cumplimiento meta", f"{cumplimiento_meta:.1f}%")

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total Leads", f"{tot}")
col2.metric("Ventas Ganadas", f"{won}", delta=f"{conv_rate:.1f}%")
col3.metric("Leads Perdidos", f"{lost}")
col4.metric("En proceso", f"{in_process}")
col5.metric("Ticket Promedio", f"${ticket_prom:,.2f}")

col6, col7, col8, col9, col10 = st.columns(5)
col6.metric("Valor Total Analizado", f"${total_valor:,.2f}")
col7.metric("Valor Ganado", f"${valor_ganado:,.2f}")
col8.metric("Valor Perdido", f"${valor_perdido:,.2f}")
col9.metric("Valor en proceso", f"${valor_en_proceso:,.2f}")
col10.metric("Promedio perdido", f"${promedio_perdido:,.2f}")

r1c1, r1c2, r1c3, r1c4 = st.columns(4)
for col, title, value in [
    (r1c1, "Canal mas efectivo", canal_top),
    (r1c2, "Marca mas vendida", marca_top),
    (r1c3, "Cliente top comprador", top_cliente),
    (r1c4, "Repuesto mas vendido", repuesto_top),
]:
    with col:
        st.markdown(
            f"""
            <div class="mini-card">
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

r2c1, r2c2, r2c3, r2c4 = st.columns(4)
for col, title, value in [
    (r2c1, "Top seguro + cliente - repuestos", top_cliente_siniestros),
    (r2c2, "Compania top - repuestos", compania_top),
    (r2c3, "Producto mas perdido", producto_perdido_top),
    (r2c4, "Cliente mas perdido", cliente_mas_perdido),
]:
    with col:
        st.markdown(
            f"""
            <div class="mini-card">
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

cmp1, cmp2, cmp3 = st.columns(3)
cmp1.metric(f"Leads - {actual['MES']}", f"{int(actual['LEADS'])}", delta=f"{delta_leads:+}")
cmp2.metric(f"Valor - {actual['MES']}", f"${float(actual['VALOR_TOTAL']):,.2f}", delta=f"${delta_valor:,.2f}")
cmp3.metric(f"Conversion - {actual['MES']}", f"{float(actual['CONVERSION_%']):.1f}%", delta=f"{delta_conv:+.1f}%")

alert1, alert2, alert3, alert4, alert5 = st.columns(5)
alerts = [
    ("Motivo de perdida dominante", motivo_top),
    ("Canal en foco", canal_top),
    ("Marca con mejor conversion", marca_top),
    ("Compania mas fuerte - repuestos", compania_top),
    ("Top aseguradora + marca - repuestos", aseguradora_ticket_top),
]
for col, (title, value) in zip([alert1, alert2, alert3, alert4, alert5], alerts):
    with col:
        st.markdown(
            f"""
            <div class="alert-card">
                <div class="alert-title">{title}</div>
                <div class="alert-value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

if buscar_siniestro:
    st.info(f"Mostrando resultados para Nro. Siniestro: {buscar_siniestro}")

tab_names = ["Resumen", "Panel Ejecutivo", "Seguros", "Clientes", "Repuestos", "Perdidas", "Detalle", "Exportar"]
if user["role"] == "admin":
    tab_names.insert(6, "Admin")
tabs = st.tabs(tab_names)
tab_map = dict(zip(tab_names, tabs))

with tab_map["Resumen"]:
    def compact_filter_label(selected, options, fallback):
        if not selected or len(selected) == len(options):
            return fallback
        label = ", ".join(selected[:3])
        if len(selected) > 3:
            label += " +"
        return label

    rango_label = "Sin fecha"
    if filtered["FECHA"].notna().any():
        rango_label = f"{filtered['FECHA'].min().date()} a {filtered['FECHA'].max().date()}"

    summary_chips = [
        f"Empresa: {empresa_filter if empresa_filter else user['company_scope']}",
        f"Marca: {compact_filter_label(marca_filter, marca_options, 'Todas')}",
        f"Canal: {compact_filter_label(canal_filter, canal_options, 'Todos')}",
        f"Estado: {compact_filter_label(estado_filter, estado_options, 'Todos')}",
        f"Periodo: {rango_label}",
    ]
    chips_html = "".join(f'<span class="summary-chip">{item}</span>' for item in summary_chips)
    st.markdown(
        f"""
        <div class="summary-hero-card">
            <div class="summary-hero-title">Resumen comercial ejecutivo</div>
            <div class="summary-hero-subtitle">Vista rapida para valor, conversion, mix de ventas y focos del periodo filtrado.</div>
            <div class="summary-chip-wrap">{chips_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    resumen_kpis = [
        ("Valor total", f"${total_valor:,.0f}", f"{tot} leads analizados"),
        ("Valor ganado", f"${valor_ganado:,.0f}", f"{won} operaciones ganadas"),
        ("Conversion", f"{conv_rate:.1f}%", f"{lost} perdidas y {in_process} en proceso"),
        ("Ticket promedio", f"${ticket_prom:,.0f}", f"Meta cumplida: {cumplimiento_meta:.1f}%"),
    ]
    rk1, rk2, rk3, rk4 = st.columns(4)
    for col, (title, value, note) in zip([rk1, rk2, rk3, rk4], resumen_kpis):
        with col:
            st.markdown(
                f"""
                <div class="summary-kpi-card">
                    <div class="summary-kpi-title">{title}</div>
                    <div class="summary-kpi-value">{value}</div>
                    <div class="summary-kpi-note">{note}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    row1a, row1b = st.columns([1.65, 1.15])
    with row1a:
        st.markdown('<div class="summary-section-label">Evolucion mensual</div>', unsafe_allow_html=True)
        st.markdown('<div class="summary-chart-note">Valor total y leads del periodo para detectar aceleracion o caidas.</div>', unsafe_allow_html=True)
        if not mensual.empty:
            monthly_chart = alt.Chart(mensual).mark_bar(cornerRadiusTopLeft=7, cornerRadiusTopRight=7, color="#60a5fa").encode(
                x=alt.X("MES:N", title=None),
                y=alt.Y("VALOR_TOTAL:Q", title=None),
                tooltip=["MES", "VALOR_TOTAL", "LEADS", "CONVERSION_%"],
            )
            st.altair_chart(dashboard_chart(monthly_chart, 250), use_container_width=True)
        else:
            st.info("No hay fechas suficientes.")
    with row1b:
        st.markdown('<div class="summary-section-label">Top productos vendidos</div>', unsafe_allow_html=True)
        st.markdown('<div class="summary-chart-note">Repuestos con mayor volumen de ventas ganadas.</div>', unsafe_allow_html=True)
        if not top_productos_resumen.empty:
            prod_chart = alt.Chart(top_productos_resumen).mark_bar(color="#38bdf8", cornerRadiusTopRight=6, cornerRadiusBottomRight=6).encode(
                y=alt.Y("REPUESTOS SOLICITADO:N", sort="-x", title=None),
                x=alt.X("CANTIDAD:Q", title=None),
                tooltip=["REPUESTOS SOLICITADO", "CANTIDAD", "VALOR_VENDIDO"],
            )
            st.altair_chart(dashboard_chart(prod_chart, max(250, len(top_productos_resumen) * 28)), use_container_width=True)
        else:
            st.info("No hay productos vendidos.")

    row2a, row2b, row2c = st.columns([1.5, 1, 1])
    with row2a:
        st.markdown('<div class="summary-section-label">Evolucion diaria</div>', unsafe_allow_html=True)
        st.markdown('<div class="summary-chart-note">Ultimos dias del periodo para seguir el ritmo comercial.</div>', unsafe_allow_html=True)
        if not resumen_diario.empty:
            orden_dias = resumen_diario["DIA"].tolist()
            area = alt.Chart(resumen_diario).mark_area(color="#1d4ed8", opacity=0.28).encode(
                x=alt.X("DIA:N", sort=orden_dias, title=None),
                y=alt.Y("VALOR_TOTAL:Q", title=None),
                tooltip=["DIA", "VALOR_TOTAL", "LEADS", "GANADAS"],
            )
            line = alt.Chart(resumen_diario).mark_line(color="#93c5fd", strokeWidth=3, point=True).encode(
                x=alt.X("DIA:N", sort=orden_dias, title=None),
                y=alt.Y("VALOR_TOTAL:Q", title=None),
                tooltip=["DIA", "VALOR_TOTAL", "LEADS", "GANADAS"],
            )
            st.altair_chart(dashboard_chart(area + line, 240), use_container_width=True)
        else:
            st.info("No hay fechas suficientes.")
    with row2b:
        st.markdown('<div class="summary-section-label">Mix por canal</div>', unsafe_allow_html=True)
        st.markdown('<div class="summary-chart-note">Cantidad de casos por origen comercial.</div>', unsafe_allow_html=True)
        canal_chart = donut_chart(resumen_canal, "CANAL", "CASOS", ["#60a5fa", "#2563eb", "#22c55e", "#f59e0b", "#f97316", "#a855f7"])
        if canal_chart is not None:
            st.altair_chart(canal_chart, use_container_width=True)
        else:
            st.info("No hay datos por canal.")
    with row2c:
        st.markdown('<div class="summary-section-label">Estado comercial</div>', unsafe_allow_html=True)
        st.markdown('<div class="summary-chart-note">Distribucion entre ganado, perdido y en proceso.</div>', unsafe_allow_html=True)
        estado_chart = donut_chart(resumen_estado, "COMPRADO", "CASOS", ["#22c55e", "#ef4444", "#f59e0b"])
        if estado_chart is not None:
            st.altair_chart(estado_chart, use_container_width=True)
        else:
            st.info("No hay estados para mostrar.")

    row3a, row3b = st.columns([1.1, 0.9])
    with row3a:
        st.markdown('<div class="summary-section-label">Valor ganado por marca</div>', unsafe_allow_html=True)
        st.markdown('<div class="summary-chart-note">Comparacion de marca basada en ventas cerradas.</div>', unsafe_allow_html=True)
        if not resumen_marca_resumen.empty:
            marca_chart = alt.Chart(resumen_marca_resumen).mark_bar(color="#f59e0b", cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
                x=alt.X("MARCA_ORIG:N", title=None),
                y=alt.Y("VALOR:Q", title=None),
                tooltip=["MARCA_ORIG", "CASOS", "VALOR"],
            )
            st.altair_chart(dashboard_chart(marca_chart, 220), use_container_width=True)
        else:
            st.info("No hay ventas cerradas por marca.")
    with row3b:
        st.markdown('<div class="summary-section-label">Conversion por marca</div>', unsafe_allow_html=True)
        st.markdown('<div class="summary-chart-note">Lectura rapida de tasa de cierre por marca filtrada.</div>', unsafe_allow_html=True)
        if not conversion_marca_resumen.empty:
            st.dataframe(
                conversion_marca_resumen.rename(columns={"MARCA_CAT": "MARCA", "SI": "GANADAS", "NO": "PERDIDAS", "TOTAL": "CASOS", "TASA (%)": "CONVERSION_%"}),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("No hay conversion por marca.")

    st.markdown("---")
    p1, p2 = st.columns(2)
    with p1:
        st.subheader("Top clientes por monto comprado")
        st.dataframe(client_ranking.head(15), use_container_width=True, hide_index=True)
    with p2:
        st.subheader("Top repuestos vendidos")
        st.dataframe(ranking_repuestos_vendidos.head(15), use_container_width=True, hide_index=True)

with tab_map["Panel Ejecutivo"]:
    render_panel_ejecutivo()

with tab_map["Seguros"]:
    s1, s2, s3, s4, s5 = st.columns(5)
    s1.metric("Aseguradoras activas", f"{aseguradoras_activas}")
    s2.metric("Facturas unicas", f"{facturas_seguro}")
    s3.metric("Facturas ganadas", f"{facturas_ganadas_seguro}")
    s4.metric("Facturas perdidas", f"{facturas_perdidas_seguro}")
    s5.metric("Facturas en proceso", f"{facturas_pendientes_seguro}")

    sr1, sr2, sr3, sr4, sr5 = st.columns(5)
    sr1.metric("Facturas mixtas", f"{facturas_mixtas_seguro}")
    sr2.metric("Aseguradora + cliente unicos", f"{aseguradora_cliente_unicos}")
    sr3.metric("Repuestos cotizados", f"{repuestos_seguro}")
    sr4.metric("Repuestos ganados", f"{repuestos_ganados_seguro}")
    sr5.metric("Repuestos perdidos", f"{repuestos_perdidos_seguro}")

    sr6, sr7 = st.columns(2)
    sr6.metric("Repuestos en proceso", f"{repuestos_pendientes_seguro}")
    sr7.metric("Ticket factura ganada", f"${ticket_factura_seguro:,.2f}")

    insurance_brand_dual_summary_display = insurance_brand_dual_summary.rename(
        columns={
            "MARCA_ORIG": "MARCA",
            "FACTURAS_TOTAL": "FACTURAS",
            "FACTURAS_GANADAS": "FACT_GANADAS",
            "FACTURAS_PERDIDAS": "FACT_PERDIDAS",
            "FACTURAS_EN_PROCESO": "FACT_EN_PROCESO",
            "FACTURAS_MIXTAS": "FACT_MIXTAS",
            "REPUESTOS_TOTAL": "REPUESTOS",
            "REPUESTOS_GANADOS": "REP_GANADOS",
            "REPUESTOS_PERDIDOS": "REP_PERDIDOS",
            "REPUESTOS_EN_PROCESO": "REP_EN_PROCESO",
        }
    )
    ranking_talleres_siniestro_display_safe = ranking_talleres_siniestro_display.rename(
        columns={"COMPAÃ‘IA": "COMPANIA"}
    )
    insurer_ticket_summary_display_safe = insurer_ticket_summary_display.rename(
        columns={"COMPAÃ‘IA": "COMPANIA"}
    )
    insurance_brand_summary_display_safe = insurance_brand_summary_display.rename(
        columns={"COMPAÃ‘IA": "COMPANIA", "MARCA_ORIG": "MARCA"}
    )

    st.caption("Factura = Nro. siniestro unico. Repuesto = cada linea del Excel.")
    st.subheader("Mazda vs Kia - facturas y repuestos")
    st.dataframe(insurance_brand_dual_summary_display, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Aseguradora + cliente - repuestos")
        st.dataframe(
            ranking_talleres_siniestro_display_safe[["COMPANIA", "NOMBRE CLIENTE", "REPUESTOS", "REP_GANADOS", "REP_PERDIDOS", "REP_EN_PROCESO", "MARCAS"]]
            if not ranking_talleres_siniestro_display_safe.empty else ranking_talleres_siniestro_display_safe,
            use_container_width=True,
            hide_index=True,
        )
    with c2:
        st.subheader("Top aseguradora + cliente por repuestos ganados")
        if not ranking_talleres_siniestro.empty:
            top_tall = ranking_talleres_siniestro.head(12)[["ETIQUETA", "GANADOS"]].copy()
            st.altair_chart(horizontal_bar(top_tall, "ETIQUETA", "GANADOS"), use_container_width=True)
        else:
            st.info("No hay datos.")

    c3, c4 = st.columns(2)
    with c3:
        st.subheader("Resumen por aseguradora - repuestos")
        st.dataframe(insurer_ticket_summary_display_safe, use_container_width=True, hide_index=True)
    with c4:
        st.subheader("Top aseguradora + marca - repuestos")
        if not insurance_brand_summary.empty:
            top_brand = insurance_brand_summary.head(12)[["ETIQUETA", "GANADOS"]].copy()
            st.altair_chart(horizontal_bar(top_brand, "ETIQUETA", "GANADOS"), use_container_width=True)
        else:
            st.info("No hay datos.")

    st.subheader("Subdivision por marca - repuestos")
    st.dataframe(
        insurance_brand_summary_display_safe[["COMPANIA", "MARCA", "REPUESTOS", "REP_GANADOS", "REP_PERDIDOS", "REP_EN_PROCESO", "VALOR_GANADO"]]
        if not insurance_brand_summary_display_safe.empty else insurance_brand_summary_display_safe,
        use_container_width=True,
        hide_index=True,
    )

with tab_map["Clientes"]:
    cc1, cc2 = st.columns(2)
    cc1.metric("Share Taller Magna", f"{market_share['share_taller_magna']:.1f}%")
    cc2.metric("Share resto de talleres", f"{market_share['share_resto_talleres']:.1f}%")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Ranking por valor comprado")
        if not client_ranking.empty:
            rank_val = client_ranking.head(15)[["NOMBRE CLIENTE", "VALOR_COMPRADO"]].copy()
            st.altair_chart(horizontal_bar(rank_val, "NOMBRE CLIENTE", "VALOR_COMPRADO"), use_container_width=True)
        else:
            st.info("No hay compras ganadas.")
    with c2:
        st.subheader("Ranking por cantidad de compras")
        if not client_ranking.empty:
            rank_cmp = client_ranking.head(15)[["NOMBRE CLIENTE", "COMPRAS"]].copy()
            st.altair_chart(horizontal_bar(rank_cmp, "NOMBRE CLIENTE", "COMPRAS"), use_container_width=True)
        else:
            st.info("No hay compras ganadas.")

    st.subheader("Detalle de clientes")
    st.dataframe(client_ranking, use_container_width=True, hide_index=True)

with tab_map["Repuestos"]:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Repuestos mas vendidos")
        if not ranking_repuestos_vendidos.empty:
            top_sell = ranking_repuestos_vendidos.head(15)[["REPUESTOS SOLICITADO", "CANTIDAD"]].copy()
            st.altair_chart(horizontal_bar(top_sell, "REPUESTOS SOLICITADO", "CANTIDAD"), use_container_width=True)
        else:
            st.info("No hay ventas de repuestos.")
    with c2:
        st.subheader("Valor vendido por repuesto")
        if not ranking_repuestos_vendidos.empty:
            top_value = ranking_repuestos_vendidos.head(15)[["REPUESTOS SOLICITADO", "VALOR_VENDIDO"]].copy()
            st.altair_chart(horizontal_bar(top_value, "REPUESTOS SOLICITADO", "VALOR_VENDIDO"), use_container_width=True)
        else:
            st.info("No hay ventas de repuestos.")

    c3, c4 = st.columns(2)
    with c3:
        st.subheader("Detalle repuestos vendidos")
        st.dataframe(ranking_repuestos_vendidos, use_container_width=True, hide_index=True)
    with c4:
        st.subheader("Repuestos mas perdidos")
        st.dataframe(ranking_repuestos_perdidos, use_container_width=True, hide_index=True)

with tab_map["Perdidas"]:
    perd1, perd2, perd3 = st.columns(3)
    perd1.metric("Monto total perdido", f"${valor_perdido:,.2f}")
    perd2.metric("Clientes que no compraron", f"{bad['NOMBRE CLIENTE'].replace('', pd.NA).dropna().nunique()}")
    perd3.metric("Casos perdidos", f"{len(bad)}")

    p1, p2 = st.columns(2)
    with p1:
        st.subheader("Perdidas por cliente")
        st.dataframe(perdidas_cliente, use_container_width=True, hide_index=True)
    with p2:
        st.subheader("Perdidas por motivo")
        st.dataframe(perdidas_motivo, use_container_width=True, hide_index=True)

    st.subheader("Detalle de clientes que no compraron")
    st.dataframe(detalle_no_compra.sort_values("VALOR", ascending=False), use_container_width=True, hide_index=True)

if "Admin" in tab_map:
    with tab_map["Admin"]:
        st.subheader("Administracion")
        users_df = get_users_df()
        st.markdown("#### Usuarios")
        st.dataframe(users_df, use_container_width=True, hide_index=True)

        st.markdown("#### Gestion rapida de registros")
        delete_options = filtered[["ID", "EMPRESA", "NOMBRE CLIENTE", "CODIGO", "REPUESTOS SOLICITADO", "VALOR"]].copy()
        delete_options["LABEL"] = delete_options.apply(
            lambda r: f"ID {int(r['ID'])} - {r['EMPRESA']} - {r['NOMBRE CLIENTE']} - {r['CODIGO']} - ${float(r['VALOR']):,.2f}",
            axis=1,
        )
        selected_label = st.selectbox("Seleccionar registro para borrar", [""] + delete_options["LABEL"].tolist())
        if selected_label:
            selected_id = int(delete_options.loc[delete_options["LABEL"] == selected_label, "ID"].iloc[0])
            if st.button("Borrar registro seleccionado"):
                delete_lead_by_id(selected_id)
                st.success("Registro eliminado.")
                st.rerun()

with tab_map["Detalle"]:
    detail_cols = [
        "ID", "EMPRESA", "FECHA", "CANAL", "COMPAÃ‘IA", "NÂ° SINIESTRO", "CHASIS", "NOMBRE CLIENTE",
        "CLIENTE_SEGMENTO", "TELEFONO", "MARCA_ORIG", "MARCA_CAT", "MODELO", "CODIGO",
        "REPUESTOS SOLICITADO", "VALOR", "COMPRADO", "MOTIVO", "COMENTARIOS", "CREATED_BY"
    ]
    detail_display = filtered[detail_cols].rename(
        columns={
            "COMPAÃ‘IA": "COMPANIA",
            "NÂ° SINIESTRO": "NRO_SINIESTRO",
            "TELEFONO": "TELEFONO",
        }
    )
    st.dataframe(
        detail_display.sort_values("FECHA", ascending=False, na_position="last"),
        use_container_width=True,
        hide_index=True,
    )

with tab_map["Exportar"]:
    export_df = filtered.copy()
    csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
    xlsx_data = dataframe_to_excel_bytes(export_df, sheet_name="Reporte")
    e1, e2 = st.columns(2)
    with e1:
        st.download_button(
            "Descargar CSV",
            data=csv_data,
            file_name="ali_leads_reporte.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with e2:
        st.download_button(
            "Descargar Excel",
            data=xlsx_data,
            file_name="ali_leads_reporte.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.markdown(
    '<div class="small-note">v6.1: login, base SQLite, importacion Excel, busqueda por siniestro/cliente/codigo, resumen de perdidas, repuestos mas vendidos y pestana Admin solo para administradores.</div>',
    unsafe_allow_html=True,
)


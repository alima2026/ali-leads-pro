import os
import re
import sqlite3
import hashlib
from io import BytesIO
from datetime import datetime
from typing import Optional

import altair as alt
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
TIPOS_MOTIVO = ["", "Precio", "Sin stock", "Demora", "No respondió", "Compró en otro lado", "No le gustó", "Otros"]
TIPOS_COMPANIA = ["", "BSE", "SURA", "Porto Seguro", "SBI", "HDI", "Berkley", "Sancor", "MAPFRE"]
MARCAS_MAGNA = ["MAZDA", "KIA"]

COLUMNAS_BASE = [
    "EMPRESA",
    "FECHA",
    "CANAL",
    "COMPAÑIA",
    "N° SINIESTRO",
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
            record["COMPAÑIA"],
            record["N° SINIESTRO"],
            record["CHASIS"],
            record["NOMBRE CLIENTE"],
            record["TELEFONO"],
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
                r.get("COMPAÑIA", ""),
                r.get("N° SINIESTRO", ""),
                r.get("CHASIS", ""),
                r.get("NOMBRE CLIENTE", ""),
                r.get("TELEFONO", ""),
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
                r.get("COMPAÑIA", ""),
                r.get("N° SINIESTRO", ""),
                r.get("CHASIS", ""),
                r.get("NOMBRE CLIENTE", ""),
                r.get("TELEFONO", ""),
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
    if v in {"SI", "SÍ", "YES", "Y", "1", "TRUE", "OK"}:
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
        return "No respondió"
    if "OTRO LADO" in v or "COMPRO EN OTRO" in v or "COMPRÓ EN OTRO" in v:
        return "Compró en otro lado"
    if "NO LE GUSTO" in v or "NO LE GUSTÓ" in v:
        return "No le gustó"
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
        "TELÉFONO": "TELEFONO",
        "NRO SINIESTRO": "N° SINIESTRO",
        "NRO. SINIESTRO": "N° SINIESTRO",
        "NUMERO SINIESTRO": "N° SINIESTRO",
        "Nº SINIESTRO": "N° SINIESTRO",
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

    for col in ["EMPRESA", "FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "TELEFONO", "MARCA", "MODELO"]:
        df[col] = df[col].replace("", pd.NA)

    for col in ["EMPRESA", "FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "TELEFONO", "MARCA", "MODELO"]:
        df[col] = df[col].ffill()

    df["EMPRESA"] = df["EMPRESA"].fillna("").astype(str).str.upper().str.strip()
    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df["FECHA"] = df["FECHA"].dt.strftime("%Y-%m-%d").fillna("")
    df["CANAL"] = df["CANAL"].apply(normalize_canal)
    df["COMPAÑIA"] = df["COMPAÑIA"].apply(normalize_compania)
    df["N° SINIESTRO"] = df["N° SINIESTRO"].fillna("").astype(str).str.strip()
    df["CHASIS"] = df["CHASIS"].fillna("").astype(str).str.strip()
    df["NOMBRE CLIENTE"] = df["NOMBRE CLIENTE"].fillna("").astype(str).str.strip()
    df["TELEFONO"] = df["TELEFONO"].fillna("").astype(str).str.strip()
    df["MARCA"] = df["MARCA"].fillna("").astype(str).str.strip().str.upper()
    df["MODELO"] = df["MODELO"].fillna("").astype(str).str.strip()
    df["CODIGO"] = df["CODIGO"].fillna("").astype(str).str.strip()
    df["REPUESTOS SOLICITADO"] = df["REPUESTOS SOLICITADO"].fillna("").astype(str).str.strip()
    df["VALOR"] = df["VALOR"].apply(parse_valor)
    df["COMPRADO"] = df["COMPRADO"].apply(normalize_yes_no)
    df["MOTIVO"] = df["MOTIVO"].apply(normalize_motivo)
    df["COMENTARIOS"] = df["COMENTARIOS"].fillna("").astype(str).str.strip()

    df.loc[df["COMPRADO"] == "SI", "MOTIVO"] = ""
    df.loc[df["CANAL"] != "Siniestro", "COMPAÑIA"] = ""

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
            "ID", "EMPRESA", "FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE",
            "TELEFONO", "MARCA_ORIG", "MARCA_CAT", "MODELO", "CODIGO", "REPUESTOS SOLICITADO",
            "VALOR", "COMPRADO", "MOTIVO", "COMENTARIOS", "CLIENTE_SEGMENTO", "CREATED_BY"
        ])

    out = pd.DataFrame({
        "ID": df["id"],
        "EMPRESA": df["empresa"].fillna("").astype(str).str.upper(),
        "FECHA": pd.to_datetime(df["fecha"], errors="coerce"),
        "CANAL": df["canal"].fillna(""),
        "COMPAÑIA": df["compania"].fillna(""),
        "N° SINIESTRO": df["numero_siniestro"].fillna(""),
        "CHASIS": df["chasis"].fillna(""),
        "NOMBRE CLIENTE": df["nombre_cliente"].fillna(""),
        "TELEFONO": df["telefono"].fillna(""),
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


def top_label(series: pd.Series, default_text="Sin datos") -> str:
    if series.empty:
        return default_text
    return str(series.idxmax())


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
    temp["N° SINIESTRO"] = temp["N° SINIESTRO"].fillna("").astype(str).str.strip()
    temp = temp[temp["N° SINIESTRO"] != ""].copy()
    if temp.empty:
        return pd.DataFrame()
    out = (
        temp.groupby("N° SINIESTRO", as_index=False)
        .agg(
            REPUESTOS=("REPUESTOS SOLICITADO", "count"),
            VALOR_TOTAL=("VALOR", "sum"),
            CLIENTE=("NOMBRE CLIENTE", "first"),
            CHASIS=("CHASIS", "first"),
            COMPANIA=("COMPAÑIA", "first"),
            CLIENTE_SEGMENTO=("CLIENTE_SEGMENTO", "first"),
        )
        .sort_values("VALOR_TOTAL", ascending=False)
    )
    return out


def build_taller_siniestro_ranking(solo_siniestros: pd.DataFrame) -> pd.DataFrame:
    if solo_siniestros.empty:
        return pd.DataFrame()
    temp = solo_siniestros.copy()
    temp["N° SINIESTRO"] = temp["N° SINIESTRO"].fillna("").astype(str).str.strip()
    temp["NOMBRE CLIENTE"] = temp["NOMBRE CLIENTE"].fillna("").astype(str).str.strip()
    temp = temp[(temp["N° SINIESTRO"] != "") & (temp["NOMBRE CLIENTE"] != "")].copy()
    if temp.empty:
        return pd.DataFrame()
    unique_cases = temp[["NOMBRE CLIENTE", "CLIENTE_SEGMENTO", "N° SINIESTRO", "VALOR"]].drop_duplicates(
        subset=["NOMBRE CLIENTE", "N° SINIESTRO"]
    )
    out = (
        unique_cases.groupby(["NOMBRE CLIENTE", "CLIENTE_SEGMENTO"], as_index=False)
        .agg(
            SINIESTROS_UNICOS=("N° SINIESTRO", "nunique"),
            VALOR_TOTAL=("VALOR", "sum"),
        )
        .sort_values(["SINIESTROS_UNICOS", "VALOR_TOTAL"], ascending=[False, False])
    )
    return out


def build_insurer_ticket_summary(ranking_siniestros: pd.DataFrame) -> pd.DataFrame:
    if ranking_siniestros.empty:
        return pd.DataFrame()
    temp = ranking_siniestros[ranking_siniestros["COMPANIA"].astype(str).str.strip() != ""].copy()
    if temp.empty:
        return pd.DataFrame()
    out = (
        temp.groupby("COMPANIA", as_index=False)
        .agg(
            SINIESTROS=("N° SINIESTRO", "nunique"),
            VALOR_TOTAL=("VALOR_TOTAL", "sum"),
            TICKET_PROMEDIO=("VALOR_TOTAL", "mean"),
        )
        .sort_values("TICKET_PROMEDIO", ascending=False)
    )
    out["TICKET_PROMEDIO"] = out["TICKET_PROMEDIO"].round(2)
    return out


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


def conversion_status_html(rate: float) -> str:
    if rate >= 70:
        return f'<div class="status-good">Semáforo comercial: Excelente · {rate:.1f}%</div>'
    if rate >= 40:
        return f'<div class="status-mid">Semáforo comercial: Medio · {rate:.1f}%</div>'
    return f'<div class="status-bad">Semáforo comercial: Bajo · {rate:.1f}%</div>'


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
            <p style="color:#475569;">Ingresá para administrar la base y ver el dashboard comercial.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Ingresar", use_container_width=True)
        if submitted:
            user = authenticate_user(username, password)
            if user:
                st.session_state.authenticated = True
                st.session_state.user = user
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")


if not st.session_state.authenticated:
    login_screen()
    st.stop()

user = st.session_state.user


# =========================================================
# SIDEBAR ADMIN / CARGA
# =========================================================
st.sidebar.success(f"Sesión: {user['full_name']} ({user['username']})")
st.sidebar.caption(f"Rol: {user['role']} · Alcance: {user['company_scope']}")

if st.sidebar.button("Cerrar sesión", use_container_width=True):
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
)

if load_mode == "Importar Excel a base de datos":
    uploaded_file = st.sidebar.file_uploader("Subir Excel/CSV", type=["xlsx", "xls", "csv"])
    import_strategy = st.sidebar.radio("Importación", ["Reemplazar empresa", "Agregar registros"], index=0)
    forced_company = st.sidebar.selectbox("Empresa destino", allowed_companies, index=0)

    if uploaded_file is not None:
        try:
            raw_df = read_excel_smart(uploaded_file) if uploaded_file.name.lower().endswith((".xlsx", ".xls")) else pd.read_csv(uploaded_file)
            detected_mode = infer_business_mode_from_filename(uploaded_file.name)
            if detected_mode in allowed_companies:
                forced_company = detected_mode
            clean_df = clean_input_dataframe(raw_df, forced_company=forced_company)
            st.sidebar.info(f"Filas válidas detectadas: {len(clean_df)}")

            with st.sidebar.expander("Vista previa importación", expanded=False):
                st.dataframe(clean_df.head(10), use_container_width=True, hide_index=True)

            if st.sidebar.button("Guardar Excel en base", use_container_width=True):
                if import_strategy == "Reemplazar empresa":
                    replace_company_leads(forced_company, clean_df, user["username"])
                else:
                    append_company_leads(clean_df, user["username"])
                st.sidebar.success("Datos guardados en la base.")
                st.rerun()
        except Exception as e:
            st.sidebar.error(f"No se pudo importar el archivo: {e}")

st.sidebar.markdown("---")
st.sidebar.subheader("Carga manual")
with st.sidebar.expander("Agregar lead manual", expanded=False):
    with st.form("manual_insert_form", clear_on_submit=True):
        empresa = st.selectbox("Empresa", allowed_companies)
        fecha = st.date_input("Fecha", value=datetime.today())
        canal = st.selectbox("Canal", TIPOS_CANAL)
        compania = st.selectbox("Compañía", TIPOS_COMPANIA)
        numero_siniestro = st.text_input("N° Siniestro")
        chasis = st.text_input("Chasis")
        nombre_cliente = st.text_input("Nombre cliente")
        telefono = st.text_input("Teléfono")
        marca = st.text_input("Marca")
        modelo = st.text_input("Modelo")
        codigo = st.text_input("Código")
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
                "COMPAÑIA": compania if canal == "Siniestro" else "",
                "N° SINIESTRO": numero_siniestro,
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
    st.sidebar.subheader("Administración")

    with st.sidebar.expander("Crear usuario", expanded=False):
        with st.form("create_user_form", clear_on_submit=True):
            new_username = st.text_input("Nuevo usuario")
            new_password = st.text_input("Contraseña", type="password")
            full_name = st.text_input("Nombre completo")
            role = st.selectbox("Rol", ["user", "admin"])
            company_scope = st.selectbox("Alcance empresa", ["TODAS"] + TIPOS_EMPRESA)
            create_submit = st.form_submit_button("Crear usuario")
            if create_submit:
                if not new_username or not new_password:
                    st.sidebar.error("Usuario y contraseña son obligatorios.")
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
    st.warning("La base de datos está vacía. Cargá un Excel o agregá leads manualmente.")
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
            Seguimiento de conversión, valor ganado, pérdidas, comparación mensual, alertas, clientes, talleres y siniestros.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)



# =========================================================
# FILTROS ANALITICOS
# =========================================================
st.sidebar.markdown("---")
st.sidebar.subheader("Filtros analíticos")

buscar_siniestro = st.sidebar.text_input("Buscar N° Siniestro")
buscar_cliente = st.sidebar.text_input("Buscar cliente")
buscar_codigo = st.sidebar.text_input("Buscar código / repuesto")

marca_options = sorted(filtered_base["MARCA_CAT"].dropna().astype(str).unique().tolist())
canal_options = sorted(filtered_base["CANAL"].dropna().astype(str).unique().tolist())
estado_options = sorted(filtered_base["COMPRADO"].dropna().astype(str).unique().tolist())
compania_options = sorted([x for x in filtered_base["COMPAÑIA"].dropna().astype(str).unique().tolist() if x])

marca_filter = st.sidebar.multiselect("Marca", marca_options, default=marca_options)
canal_filter = st.sidebar.multiselect("Canal", canal_options, default=canal_options)
estado_filter = st.sidebar.multiselect("Estado", estado_options, default=estado_options)

if "Siniestro" in canal_filter or not canal_filter:
    compania_filter = st.sidebar.multiselect("Compañía", compania_options, default=compania_options)
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
        ((filtered["CANAL"] == "Siniestro") & (filtered["COMPAÑIA"].isin(compania_filter)))
    ]

if buscar_siniestro:
    filtered = filtered[
        filtered["N° SINIESTRO"].astype(str).str.contains(buscar_siniestro.strip(), case=False, na=False)
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
solo_siniestros = filtered[
    (filtered["CANAL"] == "Siniestro")
    & (filtered["COMPAÑIA"].astype(str).str.strip() != "")
    & (filtered["N° SINIESTRO"].astype(str).str.strip() != "")
].copy()

ranking_siniestros = build_siniestro_ranking(solo_siniestros)
ranking_talleres_siniestro = build_taller_siniestro_ranking(solo_siniestros)
insurer_ticket_summary = build_insurer_ticket_summary(ranking_siniestros)
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

top_cliente = top_label(good.groupby("NOMBRE CLIENTE")["VALOR"].sum(), "Sin datos")
top_cliente_siniestros = top_label(
    ranking_talleres_siniestro.set_index("NOMBRE CLIENTE")["SINIESTROS_UNICOS"] if not ranking_talleres_siniestro.empty else pd.Series(dtype=float),
    "Sin datos de siniestros válidos",
)
compania_top = top_label(good[good["CANAL"] == "Siniestro"]["COMPAÑIA"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
motivo_top = top_label(bad["MOTIVO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
canal_top = top_label(good["CANAL"].value_counts(), "Sin ventas")
marca_top = top_label(good["MARCA_ORIG"].value_counts(), "Sin ventas")
aseguradora_ticket_top = top_label(
    insurer_ticket_summary.set_index("COMPANIA")["TICKET_PROMEDIO"] if not insurer_ticket_summary.empty else pd.Series(dtype=float),
    "Sin datos",
)
repuesto_top = top_label(good["REPUESTOS SOLICITADO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
producto_perdido_top = top_label(bad["REPUESTOS SOLICITADO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
cliente_mas_perdido = top_label(
    bad.groupby("NOMBRE CLIENTE")["VALOR"].sum() if not bad.empty else pd.Series(dtype=float),
    "Sin datos",
)

siniestros_unicos = int(solo_siniestros["N° SINIESTRO"].replace("", pd.NA).dropna().nunique()) if not solo_siniestros.empty else 0
valor_promedio_siniestro = round(ranking_siniestros["VALOR_TOTAL"].mean(), 2) if not ranking_siniestros.empty else 0.0
repuestos_promedio_siniestro = round(ranking_siniestros["REPUESTOS"].mean(), 2) if not ranking_siniestros.empty else 0.0

meta_mensual = st.sidebar.number_input("Objetivo mensual ($)", min_value=0.0, value=10000.0, step=100.0)
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
        "COMPAÑIA",
        "N° SINIESTRO",
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
    (r1c1, "Canal más efectivo", canal_top),
    (r1c2, "Marca más vendida", marca_top),
    (r1c3, "Cliente top comprador", top_cliente),
    (r1c4, "Repuesto más vendido", repuesto_top),
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
    (r2c1, "Más siniestros trae", top_cliente_siniestros),
    (r2c2, "Compañía top", compania_top),
    (r2c3, "Producto más perdido", producto_perdido_top),
    (r2c4, "Cliente más perdido", cliente_mas_perdido),
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
cmp1.metric(f"Leads · {actual['MES']}", f"{int(actual['LEADS'])}", delta=f"{delta_leads:+}")
cmp2.metric(f"Valor · {actual['MES']}", f"${float(actual['VALOR_TOTAL']):,.2f}", delta=f"${delta_valor:,.2f}")
cmp3.metric(f"Conversión · {actual['MES']}", f"{float(actual['CONVERSION_%']):.1f}%", delta=f"{delta_conv:+.1f}%")

alert1, alert2, alert3, alert4, alert5 = st.columns(5)
alerts = [
    ("Motivo de pérdida dominante", motivo_top),
    ("Canal en foco", canal_top),
    ("Marca con mejor conversión", marca_top),
    ("Compañía más fuerte", compania_top),
    ("Aseguradora mayor ticket", aseguradora_ticket_top),
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
    st.info(f"Mostrando resultados para N° Siniestro: {buscar_siniestro}")

tab_names = ["📈 Resumen", "🏢 Seguros", "👥 Clientes", "🔧 Repuestos", "📉 Pérdidas", "📋 Detalle", "⬇ Exportar"]
if user["role"] == "admin":
    tab_names.insert(5, "🛠 Admin")
tabs = st.tabs(tab_names)
tab_map = dict(zip(tab_names, tabs))

with tab_map["📈 Resumen"]:
    a, b = st.columns(2)
    with a:
        st.subheader("Conversión por marca")
        st.dataframe(build_conversion_table(filtered, "MARCA_CAT"), use_container_width=True, hide_index=False)
    with b:
        st.subheader("Resumen mensual")
        if not mensual.empty:
            chart = alt.Chart(mensual).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
                x="MES:N", y="VALOR_TOTAL:Q", tooltip=["MES", "VALOR_TOTAL", "LEADS", "CONVERSION_%"]
            )
            st.altair_chart(chart, use_container_width=True)
            st.dataframe(mensual, use_container_width=True, hide_index=True)
        else:
            st.info("No hay fechas suficientes.")

    st.markdown("---")
    st.subheader("Productos y clientes top")
    p1, p2 = st.columns(2)
    with p1:
        st.subheader("Top clientes por monto comprado")
        st.dataframe(client_ranking.head(15), use_container_width=True, hide_index=True)
    with p2:
        st.subheader("Top repuestos vendidos")
        st.dataframe(ranking_repuestos_vendidos.head(15), use_container_width=True, hide_index=True)

with tab_map["🏢 Seguros"]:
    s1, s2, s3 = st.columns(3)
    s1.metric("Siniestros únicos", f"{siniestros_unicos}")
    s2.metric("Valor promedio por siniestro", f"${valor_promedio_siniestro:,.2f}")
    s3.metric("Repuestos promedio por siniestro", f"{repuestos_promedio_siniestro:.2f}")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Ranking de siniestros")
        st.dataframe(ranking_siniestros, use_container_width=True, hide_index=True)
    with c2:
        st.subheader("Top talleres/clientes por siniestro")
        if not ranking_talleres_siniestro.empty:
            top_tall = ranking_talleres_siniestro.head(12)[["NOMBRE CLIENTE", "SINIESTROS_UNICOS"]].copy()
            st.altair_chart(horizontal_bar(top_tall, "NOMBRE CLIENTE", "SINIESTROS_UNICOS"), use_container_width=True)
        else:
            st.info("No hay datos.")

    c3, c4 = st.columns(2)
    with c3:
        st.subheader("Ticket por aseguradora")
        st.dataframe(insurer_ticket_summary, use_container_width=True, hide_index=True)
    with c4:
        st.subheader("Aseguradora con mayor ticket")
        st.metric("Top", aseguradora_ticket_top)

with tab_map["👥 Clientes"]:
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

with tab_map["🔧 Repuestos"]:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Repuestos más vendidos")
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
        st.subheader("Repuestos más perdidos")
        st.dataframe(ranking_repuestos_perdidos, use_container_width=True, hide_index=True)

with tab_map["📉 Pérdidas"]:
    perd1, perd2, perd3 = st.columns(3)
    perd1.metric("Monto total perdido", f"${valor_perdido:,.2f}")
    perd2.metric("Clientes que no compraron", f"{bad['NOMBRE CLIENTE'].replace('', pd.NA).dropna().nunique()}")
    perd3.metric("Casos perdidos", f"{len(bad)}")

    p1, p2 = st.columns(2)
    with p1:
        st.subheader("Pérdidas por cliente")
        st.dataframe(perdidas_cliente, use_container_width=True, hide_index=True)
    with p2:
        st.subheader("Pérdidas por motivo")
        st.dataframe(perdidas_motivo, use_container_width=True, hide_index=True)

    st.subheader("Detalle de clientes que no compraron")
    st.dataframe(detalle_no_compra.sort_values("VALOR", ascending=False), use_container_width=True, hide_index=True)

if "🛠 Admin" in tab_map:
    with tab_map["🛠 Admin"]:
        st.subheader("Administración")
        users_df = get_users_df()
        st.markdown("#### Usuarios")
        st.dataframe(users_df, use_container_width=True, hide_index=True)

        st.markdown("#### Gestión rápida de registros")
        delete_options = filtered[["ID", "EMPRESA", "NOMBRE CLIENTE", "CODIGO", "REPUESTOS SOLICITADO", "VALOR"]].copy()
        delete_options["LABEL"] = delete_options.apply(
            lambda r: f"ID {int(r['ID'])} · {r['EMPRESA']} · {r['NOMBRE CLIENTE']} · {r['CODIGO']} · ${float(r['VALOR']):,.2f}",
            axis=1,
        )
        selected_label = st.selectbox("Seleccionar registro para borrar", [""] + delete_options["LABEL"].tolist())
        if selected_label:
            selected_id = int(delete_options.loc[delete_options["LABEL"] == selected_label, "ID"].iloc[0])
            if st.button("Borrar registro seleccionado"):
                delete_lead_by_id(selected_id)
                st.success("Registro eliminado.")
                st.rerun()

with tab_map["📋 Detalle"]:
    detail_cols = [
        "ID", "EMPRESA", "FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE",
        "CLIENTE_SEGMENTO", "TELEFONO", "MARCA_ORIG", "MARCA_CAT", "MODELO", "CODIGO",
        "REPUESTOS SOLICITADO", "VALOR", "COMPRADO", "MOTIVO", "COMENTARIOS", "CREATED_BY"
    ]
    st.dataframe(
        filtered[detail_cols].sort_values("FECHA", ascending=False, na_position="last"),
        use_container_width=True,
        hide_index=True,
    )

with tab_map["⬇ Exportar"]:
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
    '<div class="small-note">v6.1: login, base SQLite, importación Excel, búsqueda por siniestro/cliente/código, resumen de pérdidas, repuestos más vendidos y pestaña Admin solo para administradores.</div>',
    unsafe_allow_html=True,
)

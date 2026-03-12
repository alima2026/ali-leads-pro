import os
import re
import base64
from io import BytesIO
from typing import Optional

import altair as alt
import pandas as pd
import streamlit as st


# =========================================================
# CONFIGURACION GENERAL
# =========================================================
st.set_page_config(
    page_title="ALI Leads Pro Empresas",
    layout="wide",
    initial_sidebar_state="expanded",
)

TIPOS_CANAL = [
    "Mostrador",
    "Whatsapp",
    "Siniestro",
    "Taller Magna",
    "Taller Particular",
]

TIPOS_COMPRADO = ["SI", "NO"]

TIPOS_MOTIVO = [
    "",
    "Precio",
    "Sin stock",
    "Demora",
    "No respondió",
    "Compró en otro lado",
    "No le gustó",
    "Otros",
]

TIPOS_COMPANIA = [
    "",
    "BSE",
    "SURA",
    "Porto Seguro",
    "SBI",
    "HDI",
    "Berkley",
    "Sancor",
    "MAPFRE",
]

COLUMNAS_BASE = [
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

MARCAS_MAGNA = ["MAZDA", "KIA"]


# =========================================================
# ESTILOS
# =========================================================
def file_to_base64(path: str) -> Optional[str]:
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def set_bg_image(image_file: str) -> None:
    base_dir = os.path.dirname(__file__)
    full_path = os.path.join(base_dir, image_file)
    b64 = file_to_base64(full_path)
    if not b64:
        return

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image:
                linear-gradient(rgba(248,250,252,0.98), rgba(248,250,252,0.98)),
                url("data:image/png;base64,{b64}");
            background-size: cover;
            background-attachment: fixed;
            background-position: center;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 2rem;
        }

        .main-title {
            font-size: 2.8rem;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 0.1rem;
        }

        .subtitle {
            font-size: 1rem;
            color: #475569;
            margin-bottom: 1rem;
        }

        .hero-card {
            background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);
            color: white;
            border-radius: 22px;
            padding: 1.25rem 1.4rem;
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
        }

        .kpi-title {
            font-size: 0.92rem;
            color: #64748b;
            margin-bottom: 0.15rem;
        }

        .kpi-value {
            font-size: 1.7rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.1;
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
            font-size: 1.15rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.2;
        }

        .meta-box {
            background: rgba(255,255,255,0.99);
            border: 1px solid rgba(15,23,42,0.06);
            border-radius: 18px;
            padding: 1rem 1rem;
            box-shadow: 0 4px 12px rgba(15,23,42,0.05);
            margin-bottom: 1rem;
        }

        .mode-good {
            background: #eff6ff;
            color: #1d4ed8;
            border-left: 6px solid #2563eb;
            padding: 0.85rem 1rem;
            border-radius: 14px;
            font-weight: 700;
            margin-bottom: 1rem;
        }

        .mode-neutral {
            background: #f8fafc;
            color: #334155;
            border-left: 6px solid #94a3b8;
            padding: 0.85rem 1rem;
            border-radius: 14px;
            font-weight: 700;
            margin-bottom: 1rem;
        }

        [data-testid="metric-container"] {
            background: rgba(255,255,255,0.99);
            border: 1px solid rgba(15,23,42,0.06);
            padding: 1rem !important;
            border-radius: 18px !important;
            box-shadow: 0 4px 12px rgba(15,23,42,0.05);
        }

        [data-testid="metric-container"] label {
            font-weight: 700 !important;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }

        .stTabs [data-baseweb="tab"] {
            border-radius: 12px;
            padding: 10px 16px;
            background: rgba(255,255,255,0.95);
        }

        .small-note {
            color: #64748b;
            font-size: 0.9rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


set_bg_image("bg1.png")
inject_css()


# =========================================================
# HELPERS GENERALES
# =========================================================
def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def classify_brand(marca: str) -> str:
    marca = normalize_text(marca).upper()
    return marca if marca else "SIN MARCA"


def normalize_yes_no(value: object) -> str:
    v = normalize_text(value).upper()
    if v in {"SI", "SÍ", "YES", "Y", "1", "TRUE", "OK"}:
        return "SI"
    if v in {"NO", "N", "0", "FALSE"}:
        return "NO"
    return "NO"


def normalize_canal(value: object) -> str:
    v = normalize_text(value).upper()

    if not v:
        return "Sin clasificar"
    if "MOSTRADOR" in v:
        return "Mostrador"
    if "WPP" in v or "WSP" in v or "WHATSAPP" in v or "WASAP" in v or v == "WP":
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

    s = str(value).strip().upper()
    s = s.replace("USD", "").replace("U$S", "").replace("$", "")
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)

    try:
        return float(s)
    except Exception:
        return 0.0


# =========================================================
# HELPERS CLIENTES
# =========================================================
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


def build_client_ranking(good_df: pd.DataFrame) -> pd.DataFrame:
    if good_df.empty:
        return pd.DataFrame()

    temp = good_df.copy()
    temp["NOMBRE CLIENTE"] = temp["NOMBRE CLIENTE"].fillna("").astype(str).str.strip()
    temp = temp[temp["NOMBRE CLIENTE"] != ""].copy()

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
    return out.reset_index(drop=True)


def build_client_segment_summary(good_df: pd.DataFrame) -> pd.DataFrame:
    if good_df.empty:
        return pd.DataFrame()

    out = (
        good_df.groupby("CLIENTE_SEGMENTO", as_index=False)
        .agg(
            COMPRAS=("COMPRADO", "size"),
            VALOR_COMPRADO=("VALOR", "sum"),
            CLIENTES_UNICOS=("NOMBRE CLIENTE", lambda s: s.fillna("").replace("", pd.NA).dropna().nunique()),
        )
        .sort_values("VALOR_COMPRADO", ascending=False)
    )
    return out


# =========================================================
# HELPERS SINIESTROS
# =========================================================
def build_siniestro_ranking(solo_siniestros: pd.DataFrame) -> pd.DataFrame:
    if solo_siniestros.empty or "N° SINIESTRO" not in solo_siniestros.columns:
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

    return out.reset_index(drop=True)


def build_taller_siniestro_ranking(solo_siniestros: pd.DataFrame) -> pd.DataFrame:
    if solo_siniestros.empty:
        return pd.DataFrame()

    temp = solo_siniestros.copy()
    temp["N° SINIESTRO"] = temp["N° SINIESTRO"].fillna("").astype(str).str.strip()
    temp["NOMBRE CLIENTE"] = temp["NOMBRE CLIENTE"].fillna("").astype(str).str.strip()
    temp = temp[(temp["N° SINIESTRO"] != "") & (temp["NOMBRE CLIENTE"] != "")].copy()

    if temp.empty:
        return pd.DataFrame()

    # Un siniestro único por cliente/taller
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

    return out.reset_index(drop=True)


def build_insurer_ticket_summary(ranking_siniestros: pd.DataFrame) -> pd.DataFrame:
    if ranking_siniestros.empty or "COMPANIA" not in ranking_siniestros.columns:
        return pd.DataFrame()

    temp = ranking_siniestros.copy()
    temp["COMPANIA"] = temp["COMPANIA"].fillna("").astype(str).str.strip()
    temp = temp[temp["COMPANIA"] != ""].copy()

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
    if good_df.empty:
        return {
            "valor_taller_magna": 0.0,
            "valor_resto_talleres": 0.0,
            "share_taller_magna": 0.0,
            "share_resto_talleres": 0.0,
            "compras_taller_magna": 0,
            "compras_resto_talleres": 0,
        }

    valor_taller_magna = float(good_df.loc[good_df["CLIENTE_SEGMENTO"] == "Taller Magna", "VALOR"].sum())
    valor_resto_talleres = float(good_df.loc[good_df["CLIENTE_SEGMENTO"] == "Mercado Talleres", "VALOR"].sum())

    compras_taller_magna = int((good_df["CLIENTE_SEGMENTO"] == "Taller Magna").sum())
    compras_resto_talleres = int((good_df["CLIENTE_SEGMENTO"] == "Mercado Talleres").sum())

    total_talleres = valor_taller_magna + valor_resto_talleres
    if total_talleres > 0:
        share_taller_magna = round((valor_taller_magna / total_talleres) * 100, 1)
        share_resto_talleres = round((valor_resto_talleres / total_talleres) * 100, 1)
    else:
        share_taller_magna = 0.0
        share_resto_talleres = 0.0

    return {
        "valor_taller_magna": valor_taller_magna,
        "valor_resto_talleres": valor_resto_talleres,
        "share_taller_magna": share_taller_magna,
        "share_resto_talleres": share_resto_talleres,
        "compras_taller_magna": compras_taller_magna,
        "compras_resto_talleres": compras_resto_talleres,
    }


# =========================================================
# LECTURA DE ARCHIVOS
# =========================================================
def detect_header_row_from_excel(excel_file, sheet_name="BASE DE DATOS", max_rows=15):
    preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=max_rows)

    for i in range(len(preview)):
        row_values = [normalize_text(x).upper() for x in preview.iloc[i].tolist()]
        joined = " | ".join(row_values)

        has_fecha = "FECHA" in row_values or "FECHA" in joined
        has_canal = "CANAL" in row_values or "CANAL" in joined
        has_chasis = "CHASIS" in row_values or "CHASIS" in joined

        if has_fecha and has_canal and has_chasis:
            return i

    return 0


def read_excel_smart(uploaded_file, preferred_sheet="BASE DE DATOS"):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = preferred_sheet if preferred_sheet in xls.sheet_names else xls.sheet_names[0]
    header_row = detect_header_row_from_excel(uploaded_file, sheet_name=sheet_name)
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)


def safe_read_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return read_excel_smart(uploaded_file, preferred_sheet="BASE DE DATOS")
    return pd.read_csv(uploaded_file)


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    rename_map = {
        "PRECIO": "VALOR",
        "COMPRADO ": "COMPRADO",
        "TELÉFONO": "TELEFONO",
        "CLIENTE": "CANAL",
        "NRO SINIESTRO": "N° SINIESTRO",
        "NRO. SINIESTRO": "N° SINIESTRO",
        "N° SINIESTRO ": "N° SINIESTRO",
        "NUMERO SINIESTRO": "N° SINIESTRO",
        "Nº SINIESTRO": "N° SINIESTRO",
    }

    for old, new in rename_map.items():
        if old in df.columns and new not in df.columns:
            df = df.rename(columns={old: new})

    drop_cols = [c for c in df.columns if c.startswith("UNNAMED") or c in {"CONCEPTO", "MOTIVOS:"}]
    if drop_cols:
        df = df.drop(columns=drop_cols, errors="ignore")

    for col in COLUMNAS_BASE:
        if col not in df.columns:
            df[col] = ""

    return df


# =========================================================
# MODO DE NEGOCIO
# =========================================================
def detect_business_mode(uploaded_filename: str) -> str:
    name = str(uploaded_filename).strip().lower()

    if "ventas_magna" in name:
        return "MAGNA"
    if "ventas_alimatico" in name:
        return "ALIMATICO"

    return "GENERAL"


def apply_business_rules(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    df = df.copy()

    if "MARCA_ORIG" not in df.columns:
        return df

    df["MARCA_ORIG"] = df["MARCA_ORIG"].fillna("").astype(str).str.upper().str.strip()

    if mode == "MAGNA":
        df = df[df["MARCA_ORIG"].isin(MARCAS_MAGNA)].copy()
        df["MARCA_CAT"] = df["MARCA_ORIG"]

    elif mode == "ALIMATICO":
        df["MARCA_CAT"] = df["MARCA_ORIG"].replace("", "SIN MARCA")

    else:
        df["MARCA_CAT"] = df["MARCA_ORIG"].replace("", "SIN MARCA")

    return df.reset_index(drop=True)


def get_dashboard_titles(mode: str) -> tuple[str, str]:
    if mode == "MAGNA":
        return (
            "📊 ALI Leads Pro · Magna",
            "Dashboard comercial enfocado exclusivamente en Mazda y Kia."
        )
    if mode == "ALIMATICO":
        return (
            "📊 ALI Leads Pro · Alimatico",
            "Dashboard comercial enfocado en todas las marcas cargadas, incluyendo Mazda y Kia."
        )

    return (
        "📊 ALI Leads Pro",
        "Dashboard comercial general para leads, ventas, canales, seguros, repuestos y alertas."
    )


# =========================================================
# CARGA Y LIMPIEZA DE DATA
# =========================================================
@st.cache_data
def load_data(df: pd.DataFrame) -> pd.DataFrame:
    df = standardize_columns(df)
    df = df[[c for c in COLUMNAS_BASE if c in df.columns]].copy()

    for col in ["FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "TELEFONO", "MARCA", "MODELO"]:
        df[col] = df[col].replace("", pd.NA)

    for col in ["FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "TELEFONO", "MARCA", "MODELO"]:
        df[col] = df[col].ffill()

    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
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
    df["CANAL"] = df["CANAL"].apply(normalize_canal)
    df["MOTIVO"] = df["MOTIVO"].apply(normalize_motivo)
    df["COMENTARIOS"] = df["COMENTARIOS"].fillna("").astype(str).str.strip()
    df["COMPAÑIA"] = df["COMPAÑIA"].apply(normalize_compania)

    df.loc[df["COMPRADO"].str.upper() == "SI", "MOTIVO"] = ""
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

    df["MARCA_ORIG"] = df["MARCA"].replace("", "SIN MARCA").astype(str).str.upper()
    df["MARCA_CAT"] = df["MARCA_ORIG"]
    df["CLIENTE_SEGMENTO"] = df.apply(
        lambda r: infer_client_segment(r.get("NOMBRE CLIENTE", ""), r.get("CANAL", "")),
        axis=1,
    )

    return df.reset_index(drop=True)


# =========================================================
# ANALISIS
# =========================================================
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


def conversion_status_html(rate: float) -> str:
    if rate >= 70:
        return f'<div class="status-good">Semáforo comercial: Excelente · {rate:.1f}%</div>'
    if rate >= 40:
        return f'<div class="status-mid">Semáforo comercial: Medio · {rate:.1f}%</div>'
    return f'<div class="status-bad">Semáforo comercial: Bajo · {rate:.1f}%</div>'


def top_label(series: pd.Series, default_text="Sin datos") -> str:
    if series.empty:
        return default_text
    return str(series.idxmax())


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Datos") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "FECHA" not in df.columns:
        return pd.DataFrame()

    temp = df.copy()
    temp = temp[temp["FECHA"].notna()].copy()
    if temp.empty:
        return pd.DataFrame()

    temp["MES"] = temp["FECHA"].dt.to_period("M").astype(str)

    resumen = (
        temp.groupby("MES", as_index=False)
        .agg(
            LEADS=("COMPRADO", "size"),
            VALOR_TOTAL=("VALOR", "sum"),
            GANADAS=("COMPRADO", lambda s: (s.astype(str).str.upper() == "SI").sum()),
        )
    )
    resumen["CONVERSION_%"] = ((resumen["GANADAS"] / resumen["LEADS"]) * 100).round(1)
    return resumen.sort_values("MES")


def get_current_vs_previous_month_metrics(df: pd.DataFrame) -> dict:
    mensual = monthly_summary(df)
    if mensual.empty:
        return {
            "mes_actual": "N/D",
            "mes_anterior": "N/D",
            "leads_actual": 0,
            "leads_anterior": 0,
            "valor_actual": 0.0,
            "valor_anterior": 0.0,
            "conv_actual": 0.0,
            "conv_anterior": 0.0,
        }

    mensual = mensual.reset_index(drop=True)
    actual = mensual.iloc[-1]
    anterior = mensual.iloc[-2] if len(mensual) > 1 else None

    return {
        "mes_actual": actual["MES"],
        "mes_anterior": anterior["MES"] if anterior is not None else "N/D",
        "leads_actual": int(actual["LEADS"]),
        "leads_anterior": int(anterior["LEADS"]) if anterior is not None else 0,
        "valor_actual": float(actual["VALOR_TOTAL"]),
        "valor_anterior": float(anterior["VALOR_TOTAL"]) if anterior is not None else 0.0,
        "conv_actual": float(actual["CONVERSION_%"]),
        "conv_anterior": float(anterior["CONVERSION_%"]) if anterior is not None else 0.0,
    }


def horizontal_bar(df: pd.DataFrame, category_col: str, value_col: str, title: str = ""):
    base = alt.Chart(df).encode(
        y=alt.Y(f"{category_col}:N", sort="-x", title=None),
        x=alt.X(f"{value_col}:Q", title=None),
        tooltip=[category_col, value_col],
    )

    bars = base.mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6)
    text = base.mark_text(align="left", baseline="middle", dx=6).encode(
        text=alt.Text(f"{value_col}:Q", format=",.0f")
    )

    return (bars + text).properties(height=max(220, len(df) * 28), title=title)


def build_alerts(filtered: pd.DataFrame, good: pd.DataFrame, bad: pd.DataFrame) -> dict:
    motivo_top = top_label(bad["MOTIVO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
    marca_conv = build_conversion_table(filtered, "MARCA_ORIG")
    canal_conv = build_conversion_table(filtered, "CANAL")
    comp_conv = build_conversion_table(filtered[filtered["CANAL"] == "Siniestro"], "COMPAÑIA")

    mejor_marca = top_label(marca_conv["TASA (%)"], "Sin datos") if not marca_conv.empty else "Sin datos"
    mejor_canal = top_label(canal_conv["TASA (%)"], "Sin datos") if not canal_conv.empty else "Sin datos"
    mejor_compania = top_label(comp_conv["TASA (%)"], "Sin datos") if not comp_conv.empty else "Sin datos"

    mensual_canal = pd.DataFrame()
    if not filtered.empty and filtered["FECHA"].notna().any():
        temp = filtered.copy()
        temp = temp[temp["FECHA"].notna()].copy()
        temp["MES"] = temp["FECHA"].dt.to_period("M").astype(str)
        mensual_canal = temp.groupby(["MES", "CANAL"]).size().reset_index(name="LEADS")

    canal_en_baja = "Sin datos"
    if not mensual_canal.empty and mensual_canal["MES"].nunique() >= 2:
        meses = sorted(mensual_canal["MES"].unique())
        mes_actual = meses[-1]
        mes_prev = meses[-2]
        piv = mensual_canal.pivot_table(index="CANAL", columns="MES", values="LEADS", fill_value=0)
        if mes_actual in piv.columns and mes_prev in piv.columns:
            piv["DELTA"] = piv[mes_actual] - piv[mes_prev]
            canal_en_baja = str(piv["DELTA"].idxmin()) if not piv.empty else "Sin datos"

    return {
        "motivo_top": motivo_top,
        "mejor_marca": mejor_marca,
        "mejor_canal": mejor_canal,
        "mejor_compania": mejor_compania,
        "canal_en_baja": canal_en_baja,
    }


# =========================================================
# SIDEBAR
# =========================================================
uploaded_file = st.sidebar.file_uploader(
    "Cargar planilla de ventas",
    type=["xlsx", "xls", "csv"],
)

meta_mensual = st.sidebar.number_input(
    "Objetivo mensual de ventas ($)",
    min_value=0.0,
    value=10000.0,
    step=100.0,
)

if not uploaded_file:
    st.sidebar.info("Subí la planilla para comenzar.")
    st.stop()

business_mode = detect_business_mode(uploaded_file.name)

try:
    raw_df = safe_read_file(uploaded_file)
except Exception as e:
    st.sidebar.error(f"Error al leer el archivo: {e}")
    st.stop()

base_data = load_data(raw_df)
data = apply_business_rules(base_data, business_mode)

if data.empty:
    marcas_en_archivo = sorted(
        base_data["MARCA_ORIG"].dropna().astype(str).str.upper().unique().tolist()
    ) if "MARCA_ORIG" in base_data.columns else []

    st.error("La planilla se leyó, pero no se encontraron filas válidas para este modo de negocio.")

    if business_mode == "MAGNA":
        st.info(
            "Modo MAGNA detectado. Este modo solo acepta Mazda y Kia. "
            f"Marcas encontradas en el archivo: {', '.join(marcas_en_archivo) if marcas_en_archivo else 'ninguna'}."
        )
    elif business_mode == "ALIMATICO":
        st.info(
            "Modo ALIMATICO detectado. Este modo analiza todas las marcas cargadas, incluyendo Mazda y Kia. "
            f"Marcas encontradas en el archivo: {', '.join(marcas_en_archivo) if marcas_en_archivo else 'ninguna'}."
        )
    else:
        st.info("Probá renombrando el archivo como ventas_magna.xlsx o ventas_alimatico.xlsx, según corresponda.")

    st.stop()


# =========================================================
# HEADER
# =========================================================
dashboard_title, dashboard_subtitle = get_dashboard_titles(business_mode)

st.markdown(f'<div class="main-title">{dashboard_title}</div>', unsafe_allow_html=True)
st.markdown(f'<div class="subtitle">{dashboard_subtitle}</div>', unsafe_allow_html=True)

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

if business_mode == "MAGNA":
    st.markdown(
        '<div class="mode-good">Modo detectado: MAGNA · Se analizan únicamente Mazda y Kia.</div>',
        unsafe_allow_html=True,
    )
elif business_mode == "ALIMATICO":
    st.markdown(
        '<div class="mode-good">Modo detectado: ALIMATICO · Se analizan todas las marcas cargadas, incluyendo Mazda y Kia.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="mode-neutral">Modo detectado: GENERAL · Se analiza el archivo sin reglas especiales.</div>',
        unsafe_allow_html=True,
    )


# =========================================================
# FILTROS
# =========================================================
st.sidebar.markdown("## Filtros")

marca_options = sorted(data["MARCA_CAT"].dropna().astype(str).unique().tolist())
canal_options = sorted(data["CANAL"].dropna().astype(str).unique().tolist())
estado_options = sorted(data["COMPRADO"].dropna().astype(str).unique().tolist())
compania_options = sorted([x for x in data["COMPAÑIA"].dropna().astype(str).unique().tolist() if x])

marca_filter = st.sidebar.multiselect("Marca", marca_options, default=marca_options)
canal_filter = st.sidebar.multiselect("Canal", canal_options, default=canal_options)
estado_filter = st.sidebar.multiselect("Estado", estado_options, default=estado_options)

mostrar_filtro_compania = "Siniestro" in canal_filter if canal_filter else True
if mostrar_filtro_compania:
    compania_filter = st.sidebar.multiselect("Compañía", compania_options, default=compania_options)
else:
    compania_filter = []

filtered = data.copy()

if filtered["FECHA"].notna().any():
    dmin = filtered["FECHA"].min().date()
    dmax = filtered["FECHA"].max().date()
    fecha_rango = st.sidebar.date_input("Rango de fecha", value=(dmin, dmax))
    if isinstance(fecha_rango, tuple) and len(fecha_rango) == 2:
        fecha_min, fecha_max = fecha_rango
        filtered = filtered[filtered["FECHA"].dt.date.between(fecha_min, fecha_max)]

if marca_filter:
    filtered = filtered[filtered["MARCA_CAT"].isin(marca_filter)]

if canal_filter:
    filtered = filtered[filtered["CANAL"].isin(canal_filter)]

if estado_filter:
    filtered = filtered[filtered["COMPRADO"].isin(estado_filter)]

if compania_filter:
    filtered = filtered[
        (filtered["CANAL"] != "Siniestro") |
        (
            (filtered["CANAL"] == "Siniestro") &
            (filtered["COMPAÑIA"].isin(compania_filter))
        )
    ]

good = filtered[filtered["COMPRADO"].astype(str).str.upper() == "SI"].copy()
bad = filtered[filtered["COMPRADO"].astype(str).str.upper() == "NO"].copy()

solo_siniestros = filtered[filtered["CANAL"] == "Siniestro"].copy()
ranking_siniestros = build_siniestro_ranking(solo_siniestros)
ranking_talleres_siniestro = build_taller_siniestro_ranking(solo_siniestros)
insurer_ticket_summary = build_insurer_ticket_summary(ranking_siniestros)

client_ranking = build_client_ranking(good)
client_segment_summary = build_client_segment_summary(good)

top_cliente = top_label(
    good.groupby("NOMBRE CLIENTE")["VALOR"].sum().sort_values(ascending=False),
    "Sin datos",
)

top_cliente_siniestros = top_label(
    ranking_talleres_siniestro.set_index("NOMBRE CLIENTE")["SINIESTROS_UNICOS"]
    if not ranking_talleres_siniestro.empty else pd.Series(dtype=float),
    "Sin datos",
)

market_share = build_taller_market_share(good)

valor_taller_magna = float(good.loc[good["CLIENTE_SEGMENTO"] == "Taller Magna", "VALOR"].sum())
valor_mercado_talleres = float(good.loc[good["CLIENTE_SEGMENTO"] == "Mercado Talleres", "VALOR"].sum())
valor_clientes_generales = float(good.loc[good["CLIENTE_SEGMENTO"] == "Clientes", "VALOR"].sum())

compras_taller_magna = int((good["CLIENTE_SEGMENTO"] == "Taller Magna").sum())
compras_mercado_talleres = int((good["CLIENTE_SEGMENTO"] == "Mercado Talleres").sum())
compras_clientes_generales = int((good["CLIENTE_SEGMENTO"] == "Clientes").sum())

siniestros_unicos = int(
    solo_siniestros["N° SINIESTRO"].replace("", pd.NA).dropna().nunique()
) if not solo_siniestros.empty else 0

valor_promedio_siniestro = round(
    ranking_siniestros["VALOR_TOTAL"].mean(), 2
) if not ranking_siniestros.empty else 0.0

repuestos_promedio_siniestro = round(
    ranking_siniestros["REPUESTOS"].mean(), 2
) if not ranking_siniestros.empty else 0.0

aseguradora_ticket_top = top_label(
    insurer_ticket_summary.set_index("COMPANIA")["TICKET_PROMEDIO"]
    if not insurer_ticket_summary.empty else pd.Series(dtype=float),
    "Sin datos",
)


# =========================================================
# KPI
# =========================================================
tot = len(filtered)
won = len(good)
lost = len(bad)
conv_rate = round((won / tot) * 100, 1) if tot else 0.0

total_valor = float(filtered["VALOR"].sum())
valor_ganado = float(good["VALOR"].sum())
valor_perdido = float(bad["VALOR"].sum())
ticket_prom = round(valor_ganado / won, 2) if won else 0.0
promedio_perdido = round(valor_perdido / lost, 2) if lost else 0.0

canal_top = top_label(good["CANAL"].value_counts(), "Sin ventas")
marca_top = top_label(good["MARCA_ORIG"].value_counts(), "Sin ventas")
compania_top = top_label(good[good["CANAL"] == "Siniestro"]["COMPAÑIA"].replace("", pd.NA).dropna().value_counts(), "Sin datos")
motivo_top = top_label(bad["MOTIVO"].replace("", pd.NA).dropna().value_counts(), "Sin datos")

cumplimiento_meta = round((valor_ganado / meta_mensual) * 100, 1) if meta_mensual > 0 else 0.0
cumplimiento_meta_cap = min(cumplimiento_meta / 100, 1.0)

comparativa = get_current_vs_previous_month_metrics(filtered)
alertas = build_alerts(filtered, good, bad)


# =========================================================
# SEMAFORO + META
# =========================================================
st.markdown(conversion_status_html(conv_rate), unsafe_allow_html=True)

c_meta1, c_meta2 = st.columns([2, 1])
with c_meta1:
    st.markdown('<div class="meta-box"><b>Progreso contra objetivo mensual</b></div>', unsafe_allow_html=True)
    st.progress(cumplimiento_meta_cap)
    st.caption(
        f"Valor ganado: ${valor_ganado:,.2f} | Meta: ${meta_mensual:,.2f} | Cumplimiento: {cumplimiento_meta:.1f}%"
    )

with c_meta2:
    st.metric("Cumplimiento meta", f"{cumplimiento_meta:.1f}%")
    st.metric("Objetivo mensual", f"${meta_mensual:,.2f}")


# =========================================================
# METRICAS
# =========================================================
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Leads", f"{tot}")
m2.metric("Ventas Ganadas", f"{won}", delta=f"{conv_rate:.1f}%")
m3.metric("Leads Perdidos", f"{lost}")
m4.metric("Ticket Promedio", f"${ticket_prom:,.2f}")

m5, m6, m7, m8 = st.columns(4)
m5.metric("Valor Total Analizado", f"${total_valor:,.2f}")
m6.metric("Valor Ganado", f"${valor_ganado:,.2f}")
m7.metric("Valor Perdido", f"${valor_perdido:,.2f}")
m8.metric("Promedio Perdido", f"${promedio_perdido:,.2f}")


# =========================================================
# TARJETAS EJECUTIVAS
# =========================================================
c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1:
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="kpi-title">Canal más efectivo</div>
            <div class="kpi-value">{canal_top}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="kpi-title">Marca más vendida</div>
            <div class="kpi-value">{marca_top}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with c3:
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="kpi-title">Cliente top comprador</div>
            <div class="kpi-value">{top_cliente}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with c4:
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="kpi-title">Más siniestros trae</div>
            <div class="kpi-value">{top_cliente_siniestros}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with c5:
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="kpi-title">Compañía top</div>
            <div class="kpi-value">{compania_top}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with c6:
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="kpi-title">Motivo principal de pérdida</div>
            <div class="kpi-value">{motivo_top}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# COMPARACION MES VS ANTERIOR
# =========================================================
st.markdown("### Comparación mes actual vs anterior")
cm1, cm2, cm3 = st.columns(3)

with cm1:
    delta_leads = comparativa["leads_actual"] - comparativa["leads_anterior"]
    st.metric(
        f"Leads · {comparativa['mes_actual']}",
        f"{comparativa['leads_actual']}",
        delta=f"{delta_leads:+}"
    )

with cm2:
    delta_valor = comparativa["valor_actual"] - comparativa["valor_anterior"]
    st.metric(
        f"Valor total · {comparativa['mes_actual']}",
        f"${comparativa['valor_actual']:,.2f}",
        delta=f"${delta_valor:,.2f}"
    )

with cm3:
    delta_conv = comparativa["conv_actual"] - comparativa["conv_anterior"]
    st.metric(
        f"Conversión · {comparativa['mes_actual']}",
        f"{comparativa['conv_actual']:.1f}%",
        delta=f"{delta_conv:+.1f}%"
    )


# =========================================================
# ALERTAS
# =========================================================
st.markdown("### Alertas automáticas")
a1, a2, a3, a4, a5 = st.columns(5)

with a1:
    st.markdown(
        f"""
        <div class="alert-card">
            <div class="alert-title">Motivo de pérdida dominante</div>
            <div class="alert-value">{alertas['motivo_top']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with a2:
    st.markdown(
        f"""
        <div class="alert-card">
            <div class="alert-title">Canal en baja</div>
            <div class="alert-value">{alertas['canal_en_baja']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with a3:
    st.markdown(
        f"""
        <div class="alert-card">
            <div class="alert-title">Marca con mejor conversión</div>
            <div class="alert-value">{alertas['mejor_marca']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with a4:
    st.markdown(
        f"""
        <div class="alert-card">
            <div class="alert-title">Compañía más fuerte</div>
            <div class="alert-value">{alertas['mejor_compania']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with a5:
    st.markdown(
        f"""
        <div class="alert-card">
            <div class="alert-title">Aseguradora con mayor ticket</div>
            <div class="alert-value">{aseguradora_ticket_top}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# GRAFICOS SUPERIORES
# =========================================================
g1, g2 = st.columns(2)

with g1:
    st.subheader("Embudo comercial")
    embudo_df = pd.DataFrame(
        {
            "Etapa": ["Leads Totales", "Ganados", "Perdidos"],
            "Cantidad": [tot, won, lost],
        }
    )
    embudo_chart = (
        alt.Chart(embudo_df)
        .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
        .encode(
            x=alt.X("Etapa:N", title="Etapa"),
            y=alt.Y("Cantidad:Q", title="Cantidad"),
            tooltip=["Etapa", "Cantidad"],
        )
    )
    st.altair_chart(embudo_chart, use_container_width=True)

with g2:
    st.subheader("Ganados vs Perdidos")
    donut_df = pd.DataFrame(
        {
            "Estado": ["Ganados", "Perdidos"],
            "Cantidad": [won, lost],
        }
    )
    donut = alt.Chart(donut_df).mark_arc(innerRadius=72).encode(
        theta=alt.Theta(field="Cantidad", type="quantitative"),
        color=alt.Color(field="Estado", type="nominal"),
        tooltip=["Estado", "Cantidad"],
    )
    st.altair_chart(donut, use_container_width=True)


# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(
    ["📈 Resumen", "📅 Mensual", "🧩 Canales", "🏢 Seguros", "🔧 Repuestos", "🏆 Ranking", "👥 Clientes", "📋 Detalle"]
)

with tab1:
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Conversión por Marca")
        conv_marca = build_conversion_table(filtered, "MARCA_CAT")
        if not conv_marca.empty:
            st.dataframe(conv_marca, use_container_width=True)
        else:
            st.info("No hay datos.")

    with c2:
        st.subheader("Evolución de Leads")
        if filtered["FECHA"].notna().any():
            trend = (
                filtered.assign(FECHA_DIA=filtered["FECHA"].dt.date)
                .groupby("FECHA_DIA")
                .size()
                .reset_index(name="Cantidad")
            )
            chart = (
                alt.Chart(trend)
                .mark_line(point=True, strokeWidth=3)
                .encode(
                    x=alt.X("FECHA_DIA:T", title="Fecha"),
                    y=alt.Y("Cantidad:Q", title="Leads"),
                    tooltip=["FECHA_DIA:T", "Cantidad:Q"],
                )
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("No hay fechas válidas.")

    c3, c4 = st.columns(2)

    with c3:
        st.subheader("Marcas más consultadas")
        ranking_marcas = filtered["MARCA_ORIG"].value_counts().head(10).reset_index()
        ranking_marcas.columns = ["Marca", "Cantidad"]
        if not ranking_marcas.empty:
            st.altair_chart(horizontal_bar(ranking_marcas, "Marca", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    with c4:
        st.subheader("Motivos de pérdida")
        motivos = bad["MOTIVO"].replace("", pd.NA).dropna().value_counts().head(10).reset_index()
        motivos.columns = ["Motivo", "Cantidad"]
        if not motivos.empty:
            st.altair_chart(horizontal_bar(motivos, "Motivo", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

with tab2:
    st.subheader("Comparación mensual")
    mensual = monthly_summary(filtered)

    if not mensual.empty:
        c1, c2 = st.columns(2)

        with c1:
            chart1 = (
                alt.Chart(mensual)
                .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
                .encode(
                    x=alt.X("MES:N", title="Mes"),
                    y=alt.Y("LEADS:Q", title="Leads"),
                    tooltip=["MES", "LEADS"],
                )
            )
            st.altair_chart(chart1, use_container_width=True)

        with c2:
            chart2 = (
                alt.Chart(mensual)
                .mark_line(point=True, strokeWidth=3)
                .encode(
                    x=alt.X("MES:N", title="Mes"),
                    y=alt.Y("VALOR_TOTAL:Q", title="Valor Total"),
                    tooltip=["MES", "VALOR_TOTAL"],
                )
            )
            st.altair_chart(chart2, use_container_width=True)

        st.subheader("Tabla mensual")
        st.dataframe(mensual, use_container_width=True, hide_index=True)
    else:
        st.info("No hay suficientes datos con fecha para comparar por mes.")

with tab3:
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Leads por Canal")
        leads_canal = filtered["CANAL"].value_counts().reset_index()
        leads_canal.columns = ["Canal", "Cantidad"]
        if not leads_canal.empty:
            st.altair_chart(horizontal_bar(leads_canal, "Canal", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    with c2:
        st.subheader("Conversión por Canal")
        conv_canal = build_conversion_table(filtered, "CANAL")
        if not conv_canal.empty:
            st.dataframe(conv_canal, use_container_width=True)
        else:
            st.info("Sin datos.")

    c3, c4 = st.columns(2)

    with c3:
        st.subheader("Valor por Canal")
        valor_por_canal = filtered.groupby("CANAL", as_index=False)["VALOR"].sum().sort_values("VALOR", ascending=False)
        if not valor_por_canal.empty:
            st.altair_chart(horizontal_bar(valor_por_canal, "CANAL", "VALOR"), use_container_width=True)
        else:
            st.info("Sin datos.")

    with c4:
        st.subheader("Canales con más ventas")
        ventas_por_canal = good["CANAL"].value_counts().reset_index()
        ventas_por_canal.columns = ["Canal", "Cantidad"]
        if not ventas_por_canal.empty:
            st.altair_chart(horizontal_bar(ventas_por_canal, "Canal", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

with tab4:
    st.subheader("Panel de Seguros")

    if solo_siniestros.empty:
        st.info("No hay datos de siniestros para mostrar.")
    else:
        s1, s2, s3 = st.columns(3)
        s1.metric("Siniestros únicos", f"{siniestros_unicos}")
        s2.metric("Valor promedio por siniestro", f"${valor_promedio_siniestro:,.2f}")
        s3.metric("Repuestos promedio por siniestro", f"{repuestos_promedio_siniestro:,.2f}")

        c1, c2 = st.columns(2)

        with c1:
            st.subheader("Conversión por Compañía")
            conv_compania = build_conversion_table(solo_siniestros, "COMPAÑIA")
            if not conv_compania.empty:
                st.dataframe(conv_compania, use_container_width=True)
            else:
                st.info("Sin datos.")

        with c2:
            st.subheader("Valor por Compañía")
            valor_compania = (
                solo_siniestros.groupby("COMPAÑIA", as_index=False)["VALOR"]
                .sum()
                .sort_values("VALOR", ascending=False)
            )
            if not valor_compania.empty:
                st.altair_chart(horizontal_bar(valor_compania, "COMPAÑIA", "VALOR"), use_container_width=True)
            else:
                st.info("Sin datos.")

        t1, t2 = st.columns(2)

        with t1:
            st.subheader("Ranking de talleres/clientes por siniestro")
            if not ranking_talleres_siniestro.empty:
                st.dataframe(ranking_talleres_siniestro, use_container_width=True, hide_index=True)
            else:
                st.info("No hay suficiente información de talleres/clientes por siniestro.")

        with t2:
            st.subheader("Top talleres/clientes por siniestros")
            if not ranking_talleres_siniestro.empty:
                top_talleres = ranking_talleres_siniestro.head(12)[["NOMBRE CLIENTE", "SINIESTROS_UNICOS"]].copy()
                st.altair_chart(
                    horizontal_bar(top_talleres, "NOMBRE CLIENTE", "SINIESTROS_UNICOS"),
                    use_container_width=True,
                )
            else:
                st.info("No hay suficiente información.")

        i1, i2 = st.columns(2)

        with i1:
            st.subheader("Aseguradoras por ticket promedio")
            if not insurer_ticket_summary.empty:
                st.dataframe(insurer_ticket_summary, use_container_width=True, hide_index=True)
            else:
                st.info("No hay suficiente información de aseguradoras.")

        with i2:
            st.subheader("Top ticket promedio por aseguradora")
            if not insurer_ticket_summary.empty:
                top_insurer_ticket = insurer_ticket_summary.head(10)[["COMPANIA", "TICKET_PROMEDIO"]].copy()
                st.altair_chart(
                    horizontal_bar(top_insurer_ticket, "COMPANIA", "TICKET_PROMEDIO"),
                    use_container_width=True,
                )
            else:
                st.info("No hay suficiente información.")

        st.subheader("Ranking de siniestros")
        if not ranking_siniestros.empty:
            st.dataframe(ranking_siniestros, use_container_width=True, hide_index=True)
        else:
            st.info("No hay siniestros con número cargado.")

        st.subheader("Siniestros más grandes")
        if not ranking_siniestros.empty:
            top_siniestros = ranking_siniestros.head(10)[["N° SINIESTRO", "VALOR_TOTAL"]].copy()
            st.altair_chart(
                horizontal_bar(top_siniestros, "N° SINIESTRO", "VALOR_TOTAL"),
                use_container_width=True,
            )
        else:
            st.info("No hay siniestros con número cargado.")

        st.subheader("Detalle de Seguros")
        st.dataframe(
            solo_siniestros[
                [
                    "FECHA",
                    "COMPAÑIA",
                    "N° SINIESTRO",
                    "CHASIS",
                    "NOMBRE CLIENTE",
                    "CLIENTE_SEGMENTO",
                    "MARCA_ORIG",
                    "MODELO",
                    "CODIGO",
                    "REPUESTOS SOLICITADO",
                    "VALOR",
                    "COMPRADO",
                    "MOTIVO",
                ]
            ].sort_values("FECHA", ascending=False, na_position="last"),
            use_container_width=True,
            hide_index=True,
        )

with tab5:
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Repuestos más solicitados")
        ranking_repuestos = (
            filtered["REPUESTOS SOLICITADO"]
            .replace("", pd.NA)
            .dropna()
            .value_counts()
            .head(15)
            .reset_index()
        )
        ranking_repuestos.columns = ["Repuesto", "Cantidad"]
        if not ranking_repuestos.empty:
            st.altair_chart(horizontal_bar(ranking_repuestos, "Repuesto", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    with c2:
        st.subheader("Códigos más consultados")
        ranking_codigos = (
            filtered["CODIGO"]
            .replace("", pd.NA)
            .dropna()
            .value_counts()
            .head(15)
            .reset_index()
        )
        ranking_codigos.columns = ["Codigo", "Cantidad"]
        if not ranking_codigos.empty:
            st.altair_chart(horizontal_bar(ranking_codigos, "Codigo", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    c3, c4 = st.columns(2)

    with c3:
        st.subheader("Repuestos perdidos")
        repuestos_perdidos = (
            bad["REPUESTOS SOLICITADO"]
            .replace("", pd.NA)
            .dropna()
            .value_counts()
            .head(15)
            .reset_index()
        )
        repuestos_perdidos.columns = ["Repuesto", "Cantidad"]
        if not repuestos_perdidos.empty:
            st.altair_chart(horizontal_bar(repuestos_perdidos, "Repuesto", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    with c4:
        st.subheader("Valor por Repuesto")
        valor_repuesto = (
            filtered.groupby("REPUESTOS SOLICITADO", as_index=False)["VALOR"]
            .sum()
            .sort_values("VALOR", ascending=False)
            .head(15)
        )
        valor_repuesto = valor_repuesto[valor_repuesto["REPUESTOS SOLICITADO"].astype(str).str.strip() != ""]
        if not valor_repuesto.empty:
            st.altair_chart(horizontal_bar(valor_repuesto, "REPUESTOS SOLICITADO", "VALOR"), use_container_width=True)
        else:
            st.info("Sin datos.")

with tab6:
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Top canales")
        top_canales = filtered["CANAL"].value_counts().head(10).reset_index()
        top_canales.columns = ["Canal", "Cantidad"]
        if not top_canales.empty:
            st.altair_chart(horizontal_bar(top_canales, "Canal", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    with c2:
        st.subheader("Top compañías")
        top_companias = filtered["COMPAÑIA"].replace("", pd.NA).dropna().value_counts().head(10).reset_index()
        top_companias.columns = ["Compania", "Cantidad"]
        if not top_companias.empty:
            st.altair_chart(horizontal_bar(top_companias, "Compania", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    c3, c4 = st.columns(2)

    with c3:
        st.subheader("Top marcas")
        top_marcas = filtered["MARCA_ORIG"].value_counts().head(10).reset_index()
        top_marcas.columns = ["Marca", "Cantidad"]
        if not top_marcas.empty:
            st.altair_chart(horizontal_bar(top_marcas, "Marca", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

    with c4:
        st.subheader("Top códigos")
        top_codigos = filtered["CODIGO"].replace("", pd.NA).dropna().value_counts().head(10).reset_index()
        top_codigos.columns = ["Codigo", "Cantidad"]
        if not top_codigos.empty:
            st.altair_chart(horizontal_bar(top_codigos, "Codigo", "Cantidad"), use_container_width=True)
        else:
            st.info("Sin datos.")

with tab7:
    st.subheader("Clientes que más compran")

    cc1, cc2, cc3 = st.columns(3)
    cc1.metric("Valor Taller Magna", f"${valor_taller_magna:,.2f}")
    cc2.metric("Valor Mercado Talleres", f"${valor_mercado_talleres:,.2f}")
    cc3.metric("Valor Clientes Generales", f"${valor_clientes_generales:,.2f}")

    cc4, cc5, cc6 = st.columns(3)
    cc4.metric("Compras Taller Magna", f"{compras_taller_magna}")
    cc5.metric("Compras Mercado Talleres", f"{compras_mercado_talleres}")
    cc6.metric("Compras Clientes Generales", f"{compras_clientes_generales}")

    mc1, mc2 = st.columns(2)
    mc1.metric("Share Taller Magna vs talleres", f"{market_share['share_taller_magna']:.1f}%")
    mc2.metric("Share resto de talleres", f"{market_share['share_resto_talleres']:.1f}%")

    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Ranking por valor comprado")
        if not client_ranking.empty:
            ranking_valor = client_ranking.head(15)[["NOMBRE CLIENTE", "VALOR_COMPRADO"]].copy()
            st.altair_chart(
                horizontal_bar(ranking_valor, "NOMBRE CLIENTE", "VALOR_COMPRADO"),
                use_container_width=True,
            )
        else:
            st.info("No hay compras ganadas para analizar.")

    with c2:
        st.subheader("Ranking por cantidad de compras")
        if not client_ranking.empty:
            ranking_compras = client_ranking.head(15)[["NOMBRE CLIENTE", "COMPRAS"]].copy()
            st.altair_chart(
                horizontal_bar(ranking_compras, "NOMBRE CLIENTE", "COMPRAS"),
                use_container_width=True,
            )
        else:
            st.info("No hay compras ganadas para analizar.")

    st.subheader("Comparación por segmento de cliente")
    if not client_segment_summary.empty:
        c3, c4 = st.columns(2)

        with c3:
            st.altair_chart(
                horizontal_bar(client_segment_summary, "CLIENTE_SEGMENTO", "VALOR_COMPRADO"),
                use_container_width=True,
            )

        with c4:
            st.altair_chart(
                horizontal_bar(client_segment_summary, "CLIENTE_SEGMENTO", "COMPRAS"),
                use_container_width=True,
            )
    else:
        st.info("No hay suficiente información de clientes para comparar segmentos.")

    st.subheader("Detalle de ranking de clientes")
    if not client_ranking.empty:
        st.dataframe(
            client_ranking,
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No hay ranking de clientes disponible.")

with tab8:
    st.subheader("Detalle general de leads")
    detail_cols = [
        "FECHA",
        "CANAL",
        "COMPAÑIA",
        "N° SINIESTRO",
        "CHASIS",
        "NOMBRE CLIENTE",
        "CLIENTE_SEGMENTO",
        "TELEFONO",
        "MARCA_ORIG",
        "MARCA_CAT",
        "MODELO",
        "CODIGO",
        "REPUESTOS SOLICITADO",
        "VALOR",
        "COMPRADO",
        "MOTIVO",
        "COMENTARIOS",
    ]
    if not filtered.empty:
        st.dataframe(
            filtered[detail_cols].sort_values("FECHA", ascending=False, na_position="last"),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.warning("No hay datos con los filtros seleccionados.")


# =========================================================
# EDICION
# =========================================================
with st.expander("🧾 Editar base comercial", expanded=False):
    st.caption("Podés corregir canal, comprado, motivo, valor, comentarios y compañía.")

    editor_cols = [
        "FECHA",
        "CANAL",
        "COMPAÑIA",
        "N° SINIESTRO",
        "NOMBRE CLIENTE",
        "MARCA",
        "MODELO",
        "CODIGO",
        "REPUESTOS SOLICITADO",
        "VALOR",
        "COMPRADO",
        "MOTIVO",
        "COMENTARIOS",
        "CHASIS",
        "TELEFONO",
    ]
    editable_df = data[editor_cols].copy()

    st.data_editor(
        editable_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "CANAL": st.column_config.SelectboxColumn("Canal", options=["Sin clasificar"] + TIPOS_CANAL, required=True),
            "COMPRADO": st.column_config.SelectboxColumn("Comprado", options=TIPOS_COMPRADO, required=True),
            "MOTIVO": st.column_config.SelectboxColumn("Motivo", options=TIPOS_MOTIVO),
            "COMPAÑIA": st.column_config.SelectboxColumn("Compañía", options=TIPOS_COMPANIA),
            "FECHA": st.column_config.DatetimeColumn("Fecha", format="DD/MM/YYYY"),
            "VALOR": st.column_config.NumberColumn("Valor", format="%.2f"),
        },
        disabled=[
            c for c in editable_df.columns
            if c not in ["CANAL", "COMPRADO", "MOTIVO", "VALOR", "COMENTARIOS", "COMPAÑIA"]
        ],
    )

    st.caption("Los cambios hechos en el editor impactan en la vista actual mientras esta sesión esté abierta.")


# =========================================================
# EXPORTACION
# =========================================================
st.markdown("### ⬇ Exportación")

export_base_cols = [
    "FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "TELEFONO",
    "MARCA", "MODELO", "CODIGO", "REPUESTOS SOLICITADO", "VALOR",
    "COMPRADO", "MOTIVO", "COMENTARIOS"
]

base_export = data[export_base_cols].copy()
filtered_export = filtered[
    [
        "FECHA", "CANAL", "COMPAÑIA", "N° SINIESTRO", "CHASIS", "NOMBRE CLIENTE", "CLIENTE_SEGMENTO", "TELEFONO",
        "MARCA_ORIG", "MARCA_CAT", "MODELO", "CODIGO", "REPUESTOS SOLICITADO", "VALOR",
        "COMPRADO", "MOTIVO", "COMENTARIOS"
    ]
].copy()

csv_base = base_export.to_csv(index=False).encode("utf-8-sig")
csv_filtered = filtered_export.to_csv(index=False).encode("utf-8-sig")

xlsx_base = dataframe_to_excel_bytes(base_export, sheet_name="Base Corregida")
xlsx_filtered = dataframe_to_excel_bytes(filtered_export, sheet_name="Reporte Filtrado")

e1, e2, e3, e4 = st.columns(4)

with e1:
    st.download_button(
        label="Base corregida CSV",
        data=csv_base,
        file_name="ali_leads_base_corregida.csv",
        mime="text/csv",
    )

with e2:
    st.download_button(
        label="Reporte filtrado CSV",
        data=csv_filtered,
        file_name="ali_leads_reporte_filtrado.csv",
        mime="text/csv",
    )

with e3:
    st.download_button(
        label="Base corregida Excel",
        data=xlsx_base,
        file_name="ali_leads_base_corregida.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with e4:
    st.download_button(
        label="Reporte filtrado Excel",
        data=xlsx_filtered,
        file_name="ali_leads_reporte_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown(
    '<div class="small-note">Modo Magna: solo Mazda y Kia. Modo Alimatico: todas las marcas, incluyendo Mazda y Kia. COMPRADO = SI → ganado | COMPRADO = NO → perdido. El filtro de compañía afecta solo al canal Siniestro. Incluye análisis de clientes, talleres, siniestros y ticket por aseguradora.</div>',
    unsafe_allow_html=True,
)

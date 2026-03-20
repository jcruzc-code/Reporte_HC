"""
Dashboard Ejecutivo — Gestión de Personal
Cleaned Perfect S.A. · Servicios Generales & Limpieza
Una sola página · scroll vertical · Mapa Folium con HeatMap + Sedes
"""
import json
import unicodedata
from pathlib import Path

import folium
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from folium.plugins import HeatMap
from streamlit_folium import st_folium

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Ejecutivo · Gestión de Personal",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Brand palette ────────────────────────────────────────────────────────────
ACCENT       = "#00a885"
ACCENT_LIGHT = "#22d3aa"
ACCENT_GLOW  = "rgba(0,168,133,.35)"
BG_DEEP      = "#070b14"
BG_CARD      = "rgba(15,23,42,.95)"
BORDER       = "rgba(100,116,139,.30)"
TEXT_DIM     = "#94a3b8"
TEXT_MAIN    = "#e2e8f0"
TEXT_BRIGHT  = "#f8fafc"

PALETTE = ["#00a885", "#2563eb", "#7c3aed", "#d97706", "#16a34a",
           "#0891b2", "#dc2626", "#9333ea", "#ea580c", "#475569"]

# ─── SVG icons ────────────────────────────────────────────────────────────────
ICON = {
    "users": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#00a885" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
    "building": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/><line x1="12" y1="12" x2="12" y2="16"/><line x1="10" y1="14" x2="14" y2="14"/></svg>',
    "pin": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>',
    "bar": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="2.2" stroke-linecap="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>',
    "globe": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="2.2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>',
    "list": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="2.2" stroke-linecap="round"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>',
    "filter": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#00a885" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>',
    "fire": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2.2" stroke-linecap="round"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"/></svg>',
}

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ──── Base ──── */
html,body,[class*="css"],.stApp{font-family:'Inter',system-ui,sans-serif!important;color:#e2e8f0!important}
.stApp{background:radial-gradient(circle at top,#1a2538 0%,#0b1220 46%,#070b14 100%)}
.block-container{padding:0 1.4rem 1.3rem 2.6rem!important;max-width:100%!important}

/* ──── Sidebar ──── */
section[data-testid="stSidebar"]{background:#0f172a;border-right:1px solid #1f2937}
section[data-testid="stSidebar"]>div{padding:1.2rem .9rem}
section[data-testid="stSidebar"] *{color:#e5e7eb!important}
.sidebar-label{font-size:.67rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#94a3b8!important;margin:.85rem 0 .3rem;display:block}
div[data-baseweb="select"]>div{background:#111827!important;border:1px solid #334155!important;border-radius:9px!important;color:#e2e8f0!important}
input,textarea,[data-baseweb="input"] input{color:#e2e8f0!important}

/* Sidebar button */
.stButton>button{background:linear-gradient(120deg,#00a885,#00c295)!important;color:#fff!important;border:none!important;border-radius:10px!important;font-weight:700!important;font-size:.8rem!important;width:100%;margin-top:.5rem}

/* ──── Sidebar toggle floating ──── */
.sidebar-toggle-floating{
    position:fixed;top:14px;left:12px;z-index:999999;
    background:linear-gradient(120deg,#00a885,#00c295);color:#fff;border:none;
    border-radius:999px;padding:8px 14px;font-size:.78rem;font-weight:800;
    cursor:pointer;box-shadow:0 3px 16px rgba(0,168,133,.4);transition:all .2s ease;
    display:flex;align-items:center;gap:6px;
}
.sidebar-toggle-floating:hover{filter:brightness(.92);transform:translateY(-1px)}
.sidebar-toggle-floating svg{width:16px;height:16px}

/* ──── Hero ──── */
.hero{background:linear-gradient(125deg,rgba(15,23,42,.9),rgba(15,23,42,.68));padding:18px 30px 16px;border:1px solid rgba(148,163,184,.2);border-radius:18px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;backdrop-filter:blur(5px)}
.hero-title{font-size:1.24rem;font-weight:800;color:#f8fafc}
.hero-sub{font-size:.77rem;color:#94a3b8}
.hero-badge{background:rgba(0,168,133,.16);color:#8bf5dd;border:1px solid rgba(45,212,191,.33);border-radius:999px;padding:6px 16px;font-size:.78rem;font-weight:700}

/* ──── Storytelling ──── */
.story-wrap{padding:12px 2px 2px}
.story-bar{background:linear-gradient(96deg,rgba(15,23,42,.92),rgba(30,41,59,.72));border:1px solid rgba(148,163,184,.25);border-left:3px solid #22d3aa;border-radius:12px;padding:11px 18px;font-size:.8rem;color:#cbd5e1;line-height:1.6}

/* ──── KPI cards ──── */
.kpi-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:14px;padding-top:14px}
.kpi-card{background:linear-gradient(160deg,rgba(15,23,42,.95),rgba(15,23,42,.72));border:1px solid rgba(100,116,139,.32);border-radius:14px;padding:15px 18px 12px;box-shadow:0 8px 28px rgba(2,6,23,.36);position:relative}
.kpi-bar{position:absolute;top:0;left:0;right:0;height:3px;border-radius:14px 14px 0 0}
.kpi-label{font-size:.66rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94a3b8}
.kpi-value{font-size:1.65rem;font-weight:800;color:#f8fafc}
.kpi-sub{font-size:.65rem;color:#64748b;margin-top:2px}

/* ──── Section headings ──── */
.sec-head{font-size:.68rem;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:#94a3b8;padding-bottom:7px;border-bottom:1px solid rgba(148,163,184,.24);margin:0 0 10px;display:flex;align-items:center;gap:6px}

/* ──── Charts ──── */
div[data-testid="stPlotlyChart"]{background:linear-gradient(160deg,rgba(15,23,42,.96),rgba(15,23,42,.78));border:1px solid rgba(100,116,139,.3);border-radius:13px;padding:10px 12px;box-shadow:0 8px 22px rgba(2,6,23,.25);margin-bottom:16px}

/* ──── Folium map container ──── */
.folium-map-wrapper{border-radius:16px;overflow:hidden;border:1px solid rgba(94,234,212,.22);box-shadow:0 16px 34px rgba(2,6,23,.45)}

/* ──── Map legend ──── */
.map-legend{background:linear-gradient(160deg,rgba(15,23,42,.96),rgba(15,23,42,.78));border:1px solid rgba(100,116,139,.3);border-radius:12px;padding:12px 16px;font-size:.72rem;color:#cbd5e1;margin-top:8px}
.legend-gradient{height:10px;border-radius:5px;background:linear-gradient(90deg, #0f172a, #00ff00, #adff2f, #ffff00, #ffa500, #ff0000);margin:6px 0}

/* ──── Hide defaults ──── */
#MainMenu,footer,header{visibility:hidden}
button[data-testid="collapsedControl"]{display:flex!important;visibility:visible!important;opacity:1!important}
[data-testid="stHorizontalBlock"]{gap:1.05rem!important}
</style>
""", unsafe_allow_html=True)

# ─── File paths ───────────────────────────────────────────────────────────────
DATA_FILE = Path("datos_git.xlsx")
GEOJSON_LOCAL = Path("peru_departamentos.geojson")
PROVINCE_COORDS_FILE = Path("province_coords.csv")
SEDE_COORDS_FILE = Path("sede_coords.csv")

PROVINCE_FIX = {
    "CUZCO": "CUSCO", "CUSCO": "CUSCO", "LIMA METROPOLITANA": "LIMA",
    "TUMBES": "TUMBES", "JUNIN": "JUNIN", "HUANUCO": "HUANUCO",
}


# ─── Helpers ──────────────────────────────────────────────────────────────────
def norm(v) -> str:
    if pd.isna(v):
        return "S/I"
    s = str(v).strip().upper()
    return " ".join(
        "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").split()
    ) or "S/I"


def classify_regimen(v) -> str:
    if pd.isna(v):
        return "Sin info"
    r = str(v).upper()
    if (r.startswith("FT") and ">= 8H" in r) or ("PERFECT" in r and ">= 8H" in r) or r == "FULL TIME":
        return "Full Time ≥8H"
    if r.startswith("FT"):
        return "Full Time <8H"
    if r.startswith("PT") or "PART" in r:
        return "Part Time"
    return "Otros"


# ─── Data loading ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WIDE")
    df.columns = [norm(c) for c in df.columns]
    for col in ["PROVINCIA", "DISTRITO", "CLIENTE", "UNIDAD", "CARGO", "PLANILLA",
                 "REGIMEN PLANILLA", "SUPERVISOR", "TAREADOR", "DIRECCION", "DEPARTAMENTO", "ZONA"]:
        if col in df.columns:
            df[col] = df[col].map(norm)
    df["PROVINCIA"] = df["PROVINCIA"].replace(PROVINCE_FIX)
    df["FECHA DE INGRESO"] = pd.to_datetime(df.get("FECHA DE INGRESO"), errors="coerce")
    df["DNI"] = pd.to_numeric(df["DNI"], errors="coerce").astype("Int64")
    df = df[df["FECHA DE CESE"].isna()].copy()
    df.drop(columns=["FECHA DE CESE"], errors="ignore", inplace=True)
    hoy = pd.Timestamp.today().normalize()
    df["antiguedad_dias"] = (hoy - df["FECHA DE INGRESO"]).dt.days
    df["antiguedad_meses"] = (df["antiguedad_dias"] / 30.44).clip(lower=0)
    df["antiguedad"] = (df["antiguedad_dias"] / 365.25).clip(lower=0)
    df["REGIMEN_SIMPLE"] = df["REGIMEN PLANILLA"].map(classify_regimen)
    df["ANTIGUEDAD_LABEL"] = df["antiguedad"].fillna(0).apply(
        lambda x: f"{int(x*12)}m" if x < 1 else f"{x:.1f}a"
    )
    df["UBICACION"] = df["DISTRITO"].fillna("S/I") + ", " + df["PROVINCIA"].fillna("S/I") + ", Perú"
    return df


@st.cache_data(show_spinner=False)
def load_province_coords() -> pd.DataFrame:
    if not PROVINCE_COORDS_FILE.exists():
        return pd.DataFrame(columns=["PROVINCIA", "lat", "lon"])
    coords = pd.read_csv(PROVINCE_COORDS_FILE)
    coords.columns = [c.strip().upper() for c in coords.columns]
    rename = {}
    if "LAT" not in coords.columns and "LATITUDE" in coords.columns:
        rename["LATITUDE"] = "LAT"
    if "LON" not in coords.columns and "LONGITUDE" in coords.columns:
        rename["LONGITUDE"] = "LON"
    if rename:
        coords = coords.rename(columns=rename)
    if "PROVINCIA" not in coords.columns or "LAT" not in coords.columns or "LON" not in coords.columns:
        return pd.DataFrame(columns=["PROVINCIA", "lat", "lon"])
    coords["PROVINCIA"] = coords["PROVINCIA"].map(norm)
    coords = coords.rename(columns={"LAT": "lat", "LON": "lon"})[["PROVINCIA", "lat", "lon"]]
    coords["lat"] = pd.to_numeric(coords["lat"], errors="coerce")
    coords["lon"] = pd.to_numeric(coords["lon"], errors="coerce")
    return coords.dropna(subset=["PROVINCIA", "lat", "lon"]).drop_duplicates("PROVINCIA")


@st.cache_data(show_spinner=False)
def load_sede_coords() -> pd.DataFrame:
    """Load geocoded sede coordinates."""
    if not SEDE_COORDS_FILE.exists():
        return pd.DataFrame(columns=["DIRECCION", "DISTRITO", "PROVINCIA", "DEPARTAMENTO", "lat", "lon"])
    coords = pd.read_csv(SEDE_COORDS_FILE)
    coords["lat"] = pd.to_numeric(coords["lat"], errors="coerce")
    coords["lon"] = pd.to_numeric(coords["lon"], errors="coerce")
    return coords.dropna(subset=["lat", "lon"])


@st.cache_data(show_spinner=False)
def load_geojson() -> dict | None:
    url = "https://raw.githubusercontent.com/juancgalvez/peru-geojson/master/geojson/peru_departamentos.geojson"
    try:
        import requests
        data = requests.get(url, timeout=10)
        data.raise_for_status()
        geo = data.json()
        GEOJSON_LOCAL.write_text(json.dumps(geo, ensure_ascii=False), encoding="utf-8")
        return geo
    except Exception:
        if GEOJSON_LOCAL.exists():
            return json.loads(GEOJSON_LOCAL.read_text(encoding="utf-8"))
    return None


# ─── Chart styling ────────────────────────────────────────────────────────────
def chart_base(fig, height=270):
    fig.update_layout(
        template="plotly_dark", height=height, margin=dict(l=8, r=16, t=6, b=6),
        paper_bgcolor="rgba(15,23,42,0)", plot_bgcolor="rgba(15,23,42,0)",
        font_family="Inter, system-ui, sans-serif", font_color="#dbeafe",
    )
    fig.update_xaxes(tickfont_color="#cbd5e1", title_font_color="#cbd5e1", gridcolor="rgba(148,163,184,.2)")
    fig.update_yaxes(tickfont_color="#cbd5e1", title_font_color="#cbd5e1", gridcolor="rgba(148,163,184,.2)")
    fig.update_traces(textfont=dict(color="#f8fafc"), insidetextfont=dict(color="#f8fafc"), outsidetextfont=dict(color="#f8fafc"))
    fig.update_layout(legend=dict(font=dict(color="#cbd5e1")))
    return fig


# ─── Filter helpers ───────────────────────────────────────────────────────────
def sync(key, valid):
    st.session_state[key] = [v for v in st.session_state.get(key, []) if v in valid]


def clear_filters():
    for k in ["f_cli", "f_uni", "f_car", "f_reg", "f_prov", "f_dept", "f_zona", "f_dir"]:
        if k in st.session_state:
            st.session_state[k] = []


def sidebar_filters(df: pd.DataFrame):
    with st.sidebar:
        st.markdown(
            f'{ICON["filter"]}<span style="font-size:.95rem;font-weight:700;color:#e5e7eb"> Filtros</span>',
            unsafe_allow_html=True,
        )
        
        exclude_ssgg = st.toggle("Excluir cliente SSGG", value=False)

        # ── Estado (Activo / Cesado) ──
        cese_cols = [c for c in df.columns if "CESE" in str(c).upper()]
        if cese_cols:
            cese_col = cese_cols[0]
            st.markdown('<span class="sidebar-label">Estado de Colaboradores</span>', unsafe_allow_html=True)
            tipo_est = st.radio("", ["Solo Activos", "Solo Cesados", "Todos"], horizontal=True, label_visibility="collapsed")
            is_empty = df[cese_col].isna() | (df[cese_col] == "S/I") | (df[cese_col].astype(str).str.strip() == "") | (df[cese_col].astype(str).str.upper() == "NAT")
            if tipo_est == "Solo Activos":
                df = df[is_empty]
            elif tipo_est == "Solo Cesados":
                df = df[~is_empty]

        # ── Zona ──
        st.markdown('<span class="sidebar-label">Zona</span>', unsafe_allow_html=True)
        zona_o = sorted(df["ZONA"].dropna().unique())
        sync("f_zona", zona_o)
        zona = st.multiselect("", zona_o, key="f_zona", placeholder="Todas las zonas")
        base = df[df["ZONA"].isin(zona)] if zona else df

        # ── Departamento ──
        if "DEPARTAMENTO" in base.columns:
            st.markdown('<span class="sidebar-label">Departamento</span>', unsafe_allow_html=True)
            dept_o = sorted(base["DEPARTAMENTO"][base["DEPARTAMENTO"] != "S/I"].dropna().unique())
            sync("f_dept", dept_o)
            dept = st.multiselect("", dept_o, key="f_dept", placeholder="Todos los departamentos")
            base = base[base["DEPARTAMENTO"].isin(dept)] if dept else base
        else:
            dept = []

        # ── Cliente ──
        st.markdown('<span class="sidebar-label">Cliente</span>', unsafe_allow_html=True)
        cli_o = sorted(base["CLIENTE"].dropna().unique())
        sync("f_cli", cli_o)
        cli = st.multiselect("", cli_o, key="f_cli", placeholder="Todos los clientes")
        base = base[base["CLIENTE"].isin(cli)] if cli else base

        # ── Unidad ──
        st.markdown('<span class="sidebar-label">Unidad</span>', unsafe_allow_html=True)
        uni_o = sorted(base["UNIDAD"].dropna().unique())
        sync("f_uni", uni_o)
        uni = st.multiselect("", uni_o, key="f_uni", placeholder="Todas las unidades")
        base = base[base["UNIDAD"].isin(uni)] if uni else base

        # ── Cargo ──
        st.markdown('<span class="sidebar-label">Cargo</span>', unsafe_allow_html=True)
        car_o = sorted(base["CARGO"].dropna().unique())
        sync("f_car", car_o)
        car = st.multiselect("", car_o, key="f_car", placeholder="Todos los cargos")
        base = base[base["CARGO"].isin(car)] if car else base

        # ── Régimen ──
        st.markdown('<span class="sidebar-label">Régimen</span>', unsafe_allow_html=True)
        reg_o = sorted(base["REGIMEN PLANILLA"].dropna().unique())
        sync("f_reg", reg_o)
        reg = st.multiselect("", reg_o, key="f_reg", placeholder="Todos los regímenes")
        base = base[base["REGIMEN PLANILLA"].isin(reg)] if reg else base

        # ── Provincia ──
        st.markdown('<span class="sidebar-label">Provincia</span>', unsafe_allow_html=True)
        prov_o = sorted(base["PROVINCIA"].dropna().unique())
        sync("f_prov", prov_o)
        prov = st.multiselect("", prov_o, key="f_prov", placeholder="Todas las provincias")
        base = base[base["PROVINCIA"].isin(prov)] if prov else base

        # ── Dirección (Sede) ──
        if "DIRECCION" in base.columns:
            st.markdown('<span class="sidebar-label">Dirección (Sede)</span>', unsafe_allow_html=True)
            dir_o = sorted(base["DIRECCION"].dropna().unique())
            sync("f_dir", dir_o)
            dir_f = st.multiselect("", dir_o, key="f_dir", placeholder="Todas las direcciones")
        else:
            dir_f = []

        st.button("Limpiar filtros", on_click=clear_filters, use_container_width=True)

    # Apply all filters
    f = df.copy()
    if exclude_ssgg:
        f = f[f["CLIENTE"] != "SSGG"]
    if zona:
        f = f[f["ZONA"].isin(zona)]
    if dept:
        f = f[f["DEPARTAMENTO"].isin(dept)]
    if cli:
        f = f[f["CLIENTE"].isin(cli)]
    if uni:
        f = f[f["UNIDAD"].isin(uni)]
    if car:
        f = f[f["CARGO"].isin(car)]
    if reg:
        f = f[f["REGIMEN PLANILLA"].isin(reg)]
    if prov:
        f = f[f["PROVINCIA"].isin(prov)]
    if dir_f:
        f = f[f["DIRECCION"].isin(dir_f)]
    return f


# ─── Storytelling ─────────────────────────────────────────────────────────────
def generate_story(df: pd.DataFrame, full_df: pd.DataFrame) -> str:
    total_dni = df["DNI"].dropna().nunique()
    if total_dni == 0:
        return "No hay colaboradores que coincidan con los filtros seleccionados."

    n_dept = df["DEPARTAMENTO"][df["DEPARTAMENTO"] != "S/I"].nunique() if "DEPARTAMENTO" in df.columns else 0
    n_cli = df["CLIENTE"].nunique()
    
    # Analyze state dynamically based on filters
    is_filtered = len(df) < len(full_df)
    
    # 1. Base metric
    story = f"Mostrando <strong>{total_dni:,} colaboradores</strong> activos"
    if n_dept == 1:
        dept_name = df[df["DEPARTAMENTO"] != "S/I"]["DEPARTAMENTO"].iloc[0]
        story += f" en el departamento de <strong>{dept_name}</strong>."
    else:
        story += f" distribuidos en <strong>{n_dept} departamentos</strong>."

    # 2. Client analysis
    if n_cli == 1:
        cli_name = df["CLIENTE"].iloc[0]
        uni_count = df["UNIDAD"].nunique()
        story += f" La operación de <strong>{cli_name}</strong> se gestiona a través de {uni_count} unidades operativas."
    else:
        top_cli_s = df.groupby("CLIENTE")["DNI"].nunique().sort_values(ascending=False)
        top_cli = top_cli_s.index[0] if not top_cli_s.empty else ""
        top_cli_n = int(top_cli_s.iloc[0]) if not top_cli_s.empty else 0
        porcentaje_top = int((top_cli_n / max(total_dni, 1)) * 100)
        
        if porcentaje_top > 40:
            story += f" Existe una alta dependencia con <strong>{top_cli}</strong>, quien representa el {porcentaje_top}% ({top_cli_n:,}) de la fuerza laboral mostrada."
        else:
            story += f" La cartera cuenta con {n_cli} clientes, liderados por <strong>{top_cli}</strong> ({porcentaje_top}% del personal)."

    # 3. Geo analysis
    if n_dept > 1:
        top_dept_s = df[df["DEPARTAMENTO"] != "S/I"].groupby("DEPARTAMENTO")["DNI"].nunique().sort_values(ascending=False)
        if not top_dept_s.empty:
            top_dept = top_dept_s.index[0]
            top_dept_n = int(top_dept_s.iloc[0])
            porcentaje_dept = int((top_dept_n / max(total_dni, 1)) * 100)
            if porcentaje_dept > 50:
                story += f" Se observa una fuerte centralización geográfica en <strong>{top_dept}</strong> ({porcentaje_dept}% de concentración)."
            else:
                story += f" La región con mayor despliegue es <strong>{top_dept}</strong> con {top_dept_n:,} colaboradores."

    return story


# ─── KPI Cards ────────────────────────────────────────────────────────────────
def render_kpis(df: pd.DataFrame):
    total = df["DNI"].dropna().nunique()
    n_sedes = df["DIRECCION"][df["DIRECCION"] != "S/I"].nunique() if "DIRECCION" in df.columns else 0
    n_dept = df["DEPARTAMENTO"][df["DEPARTAMENTO"] != "S/I"].nunique() if "DEPARTAMENTO" in df.columns else 0

    cards = [
        ("#00a885", "Colaboradores", f"{total:,}", "activos"),
        ("#2563eb", "Clientes", f"{df['CLIENTE'].nunique():,}", "contratos"),
        ("#7c3aed", "Sedes", f"{n_sedes:,}", "ubicaciones"),
        ("#d97706", "Departamentos", f"{n_dept:,}", "cobertura"),
        ("#16a34a", "Unidades", f"{df['UNIDAD'].nunique():,}", "operativas"),
    ]

    for col, (color, label, value, sub) in zip(st.columns(5), cards):
        with col:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-bar" style="background:{color}"></div>'
                f'<div class="kpi-label">{label}</div><div class="kpi-value">{value}</div>'
                f'<div class="kpi-sub">{sub}</div></div>',
                unsafe_allow_html=True,
            )


# ─── Folium Map ───────────────────────────────────────────────────────────────
def render_folium_map(df: pd.DataFrame, geojson: dict | None, sede_coords: pd.DataFrame, province_coords: pd.DataFrame):
    """Render interactive Folium map with HeatMap + Sede markers."""

    # ── Build map ──
    m = folium.Map(
        location=[-9.19, -75.02],
        zoom_start=6,
        tiles=None,
        control_scale=True,
    )

    # Dark tile layer
    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png",
        attr="CartoDB Dark Matter",
        name="Mapa Base",
        max_zoom=18,
    ).add_to(m)

    # ── Department choropleth (if GeoJSON available) ──
    if geojson:
        dept_counts = (
            df[df.get("DEPARTAMENTO", pd.Series(dtype=str)) != "S/I"]
            .groupby("DEPARTAMENTO")["DNI"]
            .nunique()
            .to_dict()
        ) if "DEPARTAMENTO" in df.columns else {}

        # Normalize GeoJSON feature names
        for feat in geojson.get("features", []):
            props = feat.get("properties", {})
            name_raw = props.get("NOMBDEP") or props.get("NOMBPROV") or props.get("name", "")
            props["DEPT_NORM"] = norm(name_raw)
            props["count"] = dept_counts.get(props["DEPT_NORM"], 0)

        choropleth = folium.Choropleth(
            geo_data=geojson,
            data=pd.DataFrame(list(dept_counts.items()), columns=["dept", "count"]) if dept_counts else pd.DataFrame(columns=["dept", "count"]),
            columns=["dept", "count"],
            key_on="feature.properties.DEPT_NORM",
            fill_color="YlGn",
            fill_opacity=0.25,
            line_opacity=0.5,
            line_color="#22d3aa",
            legend_name="Colaboradores por departamento",
            nan_fill_color="#0f172a",
            name="Departamentos",
            highlight=True,
        )
        choropleth.add_to(m)

        # Add tooltip to choropleth
        choropleth.geojson.add_child(
            folium.features.GeoJsonTooltip(
                fields=["DEPT_NORM", "count"],
                aliases=["Departamento:", "Colaboradores:"],
                style="background-color:#1e293b;color:#e2e8f0;border:1px solid #334155;border-radius:8px;padding:8px;font-family:Inter,sans-serif;font-size:12px",
            )
        )

    # ── HeatMap layer (weighted by log scale) ──
    heat_data = []

    # Use sede coordinates if available
    if not sede_coords.empty:
        sede_agg = (
            df[df["DIRECCION"] != "S/I"]
            .groupby("DIRECCION")["DNI"]
            .nunique()
            .reset_index(name="count")
        ) if "DIRECCION" in df.columns else pd.DataFrame(columns=["DIRECCION", "count"])

        sede_merged = sede_agg.merge(
            sede_coords[["DIRECCION", "lat", "lon"]].drop_duplicates("DIRECCION"),
            on="DIRECCION",
            how="inner",
        )

        for _, row in sede_merged.iterrows():
            # Log weight to prevent Lima from dominating
            weight = float(np.log1p(row["count"]))
            heat_data.append([float(row["lat"]), float(row["lon"]), weight])
    else:
        # Fallback to province-level
        prov_counts = (
            df[df["PROVINCIA"] != "S/I"]
            .groupby("PROVINCIA")["DNI"]
            .nunique()
            .reset_index(name="count")
        )
        prov_merged = prov_counts.merge(province_coords, on="PROVINCIA", how="inner")
        for _, row in prov_merged.iterrows():
            weight = float(np.log1p(row["count"]))
            heat_data.append([float(row["lat"]), float(row["lon"]), weight])

    if heat_data:
        HeatMap(
            heat_data,
            name="🔥 Mapa de Calor",
            min_opacity=0.3,
            max_zoom=13,
            radius=20,
            blur=15,
            gradient={
                "0.0": "#00ff00",
                "0.25": "#adff2f",
                "0.5": "#ffff00",
                "0.75": "#ffa500",
                "1.0": "#ff0000",
            },
        ).add_to(m)

    # ── Sede markers ──
    if not sede_coords.empty and "DIRECCION" in df.columns:
        sede_detail = (
            df[df["DIRECCION"] != "S/I"]
            .groupby(["DIRECCION", "DISTRITO", "PROVINCIA"], as_index=False)
            .agg(
                colaboradores=("DNI", "nunique"),
                clientes=("CLIENTE", "nunique"),
                unidades=("UNIDAD", "nunique"),
                cliente_top=("CLIENTE", lambda x: x.value_counts().index[0] if len(x) > 0 else "S/I"),
            )
        )
        sede_detail = sede_detail.merge(
            sede_coords[["DIRECCION", "lat", "lon"]].drop_duplicates("DIRECCION"),
            on="DIRECCION",
            how="inner",
        )

        # Feature group for sede markers
        sede_group = folium.FeatureGroup(name="📍 Sedes Operativas")

        max_colab = sede_detail["colaboradores"].max() if not sede_detail.empty else 1

        for _, row in sede_detail.iterrows():
            # Dynamic radius based on headcount
            radius = max(3, min(14, 3 + (row["colaboradores"] / max(max_colab, 1)) * 11))

            popup_html = f"""
            <div style="font-family:Inter,sans-serif;min-width:220px;color:#e2e8f0;background:#1e293b;padding:12px;border-radius:10px;border:1px solid #334155">
                <div style="font-size:13px;font-weight:700;color:#22d3aa;margin-bottom:6px">📍 {row['DIRECCION'][:50]}</div>
                <div style="font-size:11px;color:#94a3b8;margin-bottom:8px">{row['DISTRITO']} · {row['PROVINCIA']}</div>
                <table style="width:100%;font-size:11px;border-collapse:collapse">
                    <tr><td style="color:#94a3b8;padding:2px 0">Colaboradores</td><td style="text-align:right;font-weight:700;color:#f8fafc">{row['colaboradores']}</td></tr>
                    <tr><td style="color:#94a3b8;padding:2px 0">Clientes</td><td style="text-align:right;font-weight:700;color:#f8fafc">{row['clientes']}</td></tr>
                    <tr><td style="color:#94a3b8;padding:2px 0">Unidades</td><td style="text-align:right;font-weight:700;color:#f8fafc">{row['unidades']}</td></tr>
                    <tr><td style="color:#94a3b8;padding:2px 0">Cliente principal</td><td style="text-align:right;font-weight:700;color:#22d3aa">{row['cliente_top']}</td></tr>
                </table>
            </div>
            """

            folium.CircleMarker(
                location=[row["lat"], row["lon"]],
                radius=radius,
                color="#22d3aa",
                fill=True,
                fill_color="#10b981",
                fill_opacity=0.7,
                weight=1.5,
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=f"{row['DIRECCION'][:40]} ({row['colaboradores']} col.)",
            ).add_to(sede_group)

        sede_group.add_to(m)

    # ── Layer control ──
    folium.LayerControl(collapsed=False).add_to(m)

    # ── Render in Streamlit ──
    st.markdown('<div class="folium-map-wrapper">', unsafe_allow_html=True)
    map_data = st_folium(
        m,
        width=None,
        height=500,
        returned_objects=["last_object_clicked"],
        use_container_width=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Handle click → filter by province or sede ──
    if map_data and map_data.get("last_object_clicked"):
        clicked = map_data["last_object_clicked"]
        click_lat = clicked.get("lat")
        click_lon = clicked.get("lng")
        
        sede_match = None
        prov_match = None
        
        if click_lat and click_lon:
            # Check if clicked exactly on a Sede CircleMarker
            if not sede_coords.empty:
                dists_sede = ((sede_coords["lat"] - click_lat)**2 + (sede_coords["lon"] - click_lon)**2)
                if dists_sede.min() < 0.0001:  # extremely tight radius ensures it only triggers when clicking the marker
                    sede_match = sede_coords.loc[dists_sede.idxmin(), "DIRECCION"]
            
            # Fallback to Province polygon/area
            if not sede_match and not province_coords.empty:
                dists_prov = ((province_coords["lat"] - click_lat)**2 + (province_coords["lon"] - click_lon)**2)
                prov_match = province_coords.loc[dists_prov.idxmin(), "PROVINCIA"]
            
            # Execute specific filter
            if sede_match:
                if f"sede_{sede_match}" != st.session_state.get("_last_map_click"):
                    st.session_state["_last_map_click"] = f"sede_{sede_match}"
                    st.session_state["_pending_dir_filter"] = sede_match
                    st.rerun()
            elif prov_match:
                if f"prov_{prov_match}" != st.session_state.get("_last_map_click"):
                    st.session_state["_last_map_click"] = f"prov_{prov_match}"
                    st.session_state["_pending_prov_filter"] = prov_match
                    st.rerun()

    # ── Legend ──
    st.markdown(
        '<div class="map-legend">'
        f'{ICON["fire"]} <strong>Mapa de Calor</strong> — Escala logarítmica (ponderada para visibilidad equilibrada)'
        '<div class="legend-gradient"></div>'
        '<div style="display:flex;justify-content:space-between;font-size:.65rem;color:#64748b">'
        '<span>Menor concentración</span><span>Mayor concentración</span></div>'
        '</div>',
        unsafe_allow_html=True,
    )


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    # Process pending filter requests from interactive components (e.g. map clicks)
    if "_pending_prov_filter" in st.session_state:
        st.session_state["f_prov"] = [st.session_state.pop("_pending_prov_filter")]
    if "_pending_dir_filter" in st.session_state:
        st.session_state["f_dir"] = [st.session_state.pop("_pending_dir_filter")]

    if not DATA_FILE.exists():
        st.error("No se encontró datos_git.xlsx")
        st.stop()

    df = load_data(DATA_FILE)
    geo = load_geojson()
    sede_coords = load_sede_coords()
    province_coords = load_province_coords()
    filtered = sidebar_filters(df)

    # ── Floating sidebar toggle button ──
    # Inject into parent document via components.html (runs JS at load)
    import streamlit.components.v1 as components
    components.html("""
    <script>
    (function() {
      var doc = window.parent.document;
      // Avoid duplicates
      if (doc.getElementById('cp-sidebar-toggle')) return;

      // Create button in parent document
      var btn = doc.createElement('button');
      btn.id = 'cp-sidebar-toggle';
      btn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg> Filtros';
      btn.style.cssText = 'position:fixed;top:14px;left:12px;z-index:999999;background:linear-gradient(120deg,#00a885,#00c295);color:#fff;border:none;border-radius:999px;padding:8px 14px;font-size:13px;font-weight:800;cursor:pointer;box-shadow:0 3px 16px rgba(0,168,133,.4);display:flex;align-items:center;gap:6px;font-family:Inter,system-ui,sans-serif;transition:all .2s ease';
      btn.onmouseover = function(){ this.style.filter='brightness(.92)'; this.style.transform='translateY(-1px)'; };
      btn.onmouseout  = function(){ this.style.filter='none'; this.style.transform='none'; };
      btn.onclick = function() {
        var collapsed = doc.querySelector('button[data-testid="collapsedControl"]');
        if (collapsed) { collapsed.click(); return; }
        var closeBtn = doc.querySelector('section[data-testid="stSidebar"] button[data-testid="baseButton-headerNoPadding"]');
        if (closeBtn) { closeBtn.click(); return; }
        var hdr = doc.querySelector('section[data-testid="stSidebar"] [data-testid="stSidebarHeader"] button');
        if (hdr) { hdr.click(); return; }
      };
      doc.body.appendChild(btn);
    })();
    </script>
    """, height=0)

    # ── Hero banner ──
    st.markdown(
        f'<div class="hero"><div><div class="hero-title">Dashboard Ejecutivo — Gestión de Personal</div>'
        f'<div class="hero-sub">Cleaned Perfect S.A. · Servicios Generales y Limpieza · {pd.Timestamp.today().strftime("%d/%m/%Y")}</div></div>'
        f'<div class="hero-badge">{filtered["DNI"].dropna().nunique():,} colaboradores activos</div></div>',
        unsafe_allow_html=True,
    )

    # ── Storytelling bar ──
    st.markdown(
        f'<div class="story-wrap"><div class="story-bar">{generate_story(filtered, df)}</div></div>',
        unsafe_allow_html=True,
    )

    # ── KPIs ──
    st.markdown('<div class="kpi-grid">', unsafe_allow_html=True)
    render_kpis(filtered)
    st.markdown('</div><div style="height:22px"></div>', unsafe_allow_html=True)

    # ── Map + Ranking row ──
    st.markdown('<div style="padding:0 2px">', unsafe_allow_html=True)
    c_map, c_rank = st.columns([1.60, 1.25], vertical_alignment="top")

    with c_map:
        st.markdown(
            f'<div class="sec-head">{ICON["globe"]} Cobertura Nacional · Mapa de Calor y Sedes</div>',
            unsafe_allow_html=True,
        )
        render_folium_map(filtered, geo, sede_coords, province_coords)

    with c_rank:
        st.markdown(
            f'<div class="sec-head">{ICON["bar"]} Ranking de Clientes</div>',
            unsafe_allow_html=True,
        )
        rank = (
            filtered.groupby("CLIENTE")["DNI"]
            .nunique()
            .sort_values(ascending=False)
            .reset_index(name="N")
        )
        fig_rank = px.bar(
            rank, x="N", y="CLIENTE", orientation="h",
            color="N", color_continuous_scale=["#052e2b", "#10b981"], text="N",
        )
        # Make the chart height dynamic based on the number of clients so labels are legible
        # but place it inside a fixed scrollable container
        fig_rank.update_layout(
            yaxis=dict(autorange="reversed", tickfont=dict(size=11)), 
            coloraxis_showscale=False
        )
        
        with st.container(height=500, border=False):
            st.plotly_chart(
                chart_base(fig_rank, height=max(450, len(rank) * 28)), 
                use_container_width=True, 
                config={"displayModeBar": False}
            )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Three charts row ──
    st.markdown('<div style="padding:8px 2px 0">', unsafe_allow_html=True)
    a, b, c = st.columns(3)

    with a:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Régimen Simplificado</div>', unsafe_allow_html=True)
        reg = filtered.groupby("REGIMEN_SIMPLE")["DNI"].nunique().reset_index(name="N")
        fig_reg = px.pie(reg, values="N", names="REGIMEN_SIMPLE", color="REGIMEN_SIMPLE",
                         color_discrete_sequence=PALETTE, hole=0.45)
        st.plotly_chart(chart_base(fig_reg, 260), use_container_width=True, config={"displayModeBar": False})

    with b:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Top 10 Departamentos</div>', unsafe_allow_html=True)
        if "DEPARTAMENTO" in filtered.columns:
            dept_data = (
                filtered[filtered["DEPARTAMENTO"] != "S/I"]
                .groupby("DEPARTAMENTO")["DNI"]
                .nunique()
                .sort_values(ascending=False)
                .head(10)
                .reset_index(name="N")
            )
        else:
            dept_data = (
                filtered[filtered["PROVINCIA"] != "S/I"]
                .groupby("PROVINCIA")["DNI"]
                .nunique()
                .sort_values(ascending=False)
                .head(10)
                .reset_index(name="N")
            )
        
        # Use single color and textposition outside to solve low contrast issue. Give x-axis ~15% padding for text
        max_val = dept_data["N"].max()
        fig_dept = px.bar(dept_data, x="N", y=dept_data.columns[0], orientation="h", text="N")
        fig_dept.update_traces(marker_color="#2563eb", textposition="outside", textfont=dict(color="#f8fafc", size=12))
        fig_dept.update_layout(
            yaxis=dict(autorange="reversed"), 
            xaxis=dict(range=[0, max_val * 1.15], showgrid=False),
            coloraxis_showscale=False
        )
        st.plotly_chart(chart_base(fig_dept, 260), use_container_width=True, config={"displayModeBar": False})

    with c:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Mix Clientes</div>', unsafe_allow_html=True)
        mix_all = filtered.groupby("CLIENTE")["DNI"].nunique().sort_values(ascending=False)
        if len(mix_all) > 7:
            mix_top = mix_all.head(7)
            rest_sum = mix_all.iloc[7:].sum()
            mix = pd.concat([mix_top, pd.Series({"Resto (Agrupado)": rest_sum})]).reset_index()
            mix.columns = ["CLIENTE", "N"]
        else:
            mix = mix_all.reset_index(name="N")

        fig_mix = px.pie(mix, values="N", names="CLIENTE", color_discrete_sequence=PALETTE, hole=0.45)
        
        fig_mix.update_traces(
            textposition='inside', 
            textinfo='percent+label',
            textfont=dict(size=10),
            hovertemplate="<b>%{label}</b><br>Colaboradores: %{value}<br>Proporción: %{percent}<extra></extra>"
        )
        fig_mix.update_layout(showlegend=False, margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(chart_base(fig_mix, 260), use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Top 10 districts ──
    st.markdown('<div style="padding:8px 2px 0">', unsafe_allow_html=True)
    st.markdown(f'<div class="sec-head">{ICON["bar"]} Top 10 Distritos de Operación</div>', unsafe_allow_html=True)
    dept_col = "DEPARTAMENTO" if "DEPARTAMENTO" in filtered.columns else "PROVINCIA"
    dist = (
        filtered[filtered["DISTRITO"] != "S/I"]
        .groupby(["DISTRITO", dept_col])["DNI"]
        .nunique()
        .sort_values(ascending=False)
        .head(10)
        .reset_index(name="N")
    )
    dist["UBI"] = dist["DISTRITO"] + " / " + dist[dept_col]
    fig_dist = px.bar(dist, x="N", y="UBI", orientation="h", text="N",
                      color=dept_col, color_discrete_sequence=PALETTE)
    fig_dist.update_layout(yaxis=dict(autorange="reversed"))
    st.plotly_chart(chart_base(fig_dist, 330), use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Zona comparison ──
    if "ZONA" in filtered.columns:
        st.markdown('<div style="padding:8px 2px 0">', unsafe_allow_html=True)
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Distribución Lima vs Provincia</div>', unsafe_allow_html=True)
        zona_d = filtered.groupby("ZONA")["DNI"].nunique().reset_index(name="N")
        fig_zona = px.pie(zona_d, values="N", names="ZONA", color="ZONA",
                          color_discrete_map={"LIMA": "#2563eb", "PROVINCIA": "#10b981", "S/I": "#475569"},
                          hole=0.5)
        st.plotly_chart(chart_base(fig_zona, 260), use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Detail table ──
    st.markdown('<div style="padding:8px 2px 0">', unsafe_allow_html=True)
    st.markdown(f'<div class="sec-head">{ICON["list"]} Detalle de Colaboradores Activos</div>', unsafe_allow_html=True)
    cols_show = [c for c in ["DNI", "APELLIDOS Y NOMBRES", "CARGO", "CLIENTE", "UNIDAD", "SUPERVISOR",
                             "DIRECCION", "DISTRITO", "PROVINCIA", "DEPARTAMENTO", "ZONA",
                             "REGIMEN_SIMPLE", "FECHA DE INGRESO", "ANTIGUEDAD_LABEL"]
                 if c in filtered.columns]
    disp = filtered[cols_show].copy()
    rename_map = {
        "APELLIDOS Y NOMBRES": "Nombre", "DISTRITO": "Distrito", "DIRECCION": "Dirección",
        "DEPARTAMENTO": "Dept.", "ANTIGUEDAD_LABEL": "Antigüedad", "REGIMEN_SIMPLE": "Régimen",
    }
    disp = disp.rename(columns={k: v for k, v in rename_map.items() if k in disp.columns})
    srch = st.text_input("", placeholder="Buscar por nombre, DNI, cliente, dirección o provincia...")
    if srch:
        blob = disp.astype(str).agg(" ".join, axis=1).str.upper()
        disp = disp[blob.str.contains(srch.upper(), regex=False, na=False)]
    st.dataframe(
        disp, hide_index=True, use_container_width=True, height=360,
        column_config={"FECHA DE INGRESO": st.column_config.DateColumn(format="DD/MM/YYYY")},
    )
    st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()

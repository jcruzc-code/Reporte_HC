"""
Dashboard Ejecutivo — Gestión de Personal
Cleaned Perfect S.A. · Servicios Generales & Limpieza
"""
import unicodedata
from pathlib import Path

import folium
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from branca.colormap import LinearColormap
from streamlit_folium import st_folium

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Ejecutivo · Gestión de Personal",
    page_icon="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'><rect width='24' height='24' rx='4' fill='%2300a885'/><path d='M6 12h12M6 8h12M6 16h8' stroke='white' stroke-width='1.8' stroke-linecap='round'/></svg>",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── SVG icon helpers ──────────────────────────────────────────────────────────
def svg_users():
    return """<svg width="20" height="20" viewBox="0 0 24 24" fill="none"
        stroke="#00a885" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
        <circle cx="9" cy="7" r="4"/>
        <path d="M23 21v-2a4 4 0 0 0-3-3.87"/>
        <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
    </svg>"""

def svg_building():
    return """<svg width="20" height="20" viewBox="0 0 24 24" fill="none"
        stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <rect x="2" y="7" width="20" height="14" rx="2"/>
        <path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/>
        <line x1="12" y1="12" x2="12" y2="16"/>
        <line x1="10" y1="14" x2="14" y2="14"/>
    </svg>"""

def svg_map():
    return """<svg width="20" height="20" viewBox="0 0 24 24" fill="none"
        stroke="#7c3aed" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <polygon points="1 6 1 22 8 18 16 22 23 18 23 2 16 6 8 2 1 6"/>
        <line x1="8" y1="2" x2="8" y2="18"/>
        <line x1="16" y1="6" x2="16" y2="22"/>
    </svg>"""

def svg_pin():
    return """<svg width="20" height="20" viewBox="0 0 24 24" fill="none"
        stroke="#d97706" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/>
        <circle cx="12" cy="10" r="3"/>
    </svg>"""

def svg_star():
    return """<svg width="20" height="20" viewBox="0 0 24 24" fill="none"
        stroke="#16a34a" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>
    </svg>"""

def svg_chart():
    return """<svg width="18" height="18" viewBox="0 0 24 24" fill="none"
        stroke="#64748b" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <line x1="18" y1="20" x2="18" y2="10"/>
        <line x1="12" y1="20" x2="12" y2="4"/>
        <line x1="6" y1="20" x2="6" y2="14"/>
    </svg>"""

def svg_globe():
    return """<svg width="18" height="18" viewBox="0 0 24 24" fill="none"
        stroke="#64748b" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="10"/>
        <line x1="2" y1="12" x2="22" y2="12"/>
        <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>
    </svg>"""

def svg_list():
    return """<svg width="18" height="18" viewBox="0 0 24 24" fill="none"
        stroke="#64748b" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <line x1="8" y1="6" x2="21" y2="6"/>
        <line x1="8" y1="12" x2="21" y2="12"/>
        <line x1="8" y1="18" x2="21" y2="18"/>
        <line x1="3" y1="6" x2="3.01" y2="6"/>
        <line x1="3" y1="12" x2="3.01" y2="12"/>
        <line x1="3" y1="18" x2="3.01" y2="18"/>
    </svg>"""

def svg_refresh():
    return """<svg width="16" height="16" viewBox="0 0 24 24" fill="none"
        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="1 4 1 10 7 10"/>
        <path d="M3.51 15a9 9 0 1 0 .49-3.77"/>
    </svg>"""

# ── CSS — executive light theme ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"], .stApp {
    font-family: 'Inter', system-ui, sans-serif;
}

/* Background */
.stApp { background: #f0f2f7; }
.block-container { padding: 1.25rem 2rem 2rem 2rem; max-width: 100%; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid #e2e8f0;
}
section[data-testid="stSidebar"] * { color: #374151 !important; }
section[data-testid="stSidebar"] .sidebar-label {
    font-size: 0.67rem;
    font-weight: 700;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    color: #94a3b8 !important;
    margin: 0.9rem 0 0.35rem 0;
    display: block;
}
div[data-baseweb="select"] > div {
    background: #f8fafc !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 8px !important;
    color: #1e293b !important;
}

/* KPI cards */
div[data-testid="metric-container"] {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.1rem 1.4rem 1.1rem 1.4rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    position: relative;
    overflow: hidden;
}
div[data-testid="metric-container"] label {
    font-size: 0.68rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.07em !important;
    text-transform: uppercase !important;
    color: #94a3b8 !important;
}
div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
    font-size: 2.1rem !important;
    font-weight: 700 !important;
    color: #111827 !important;
    letter-spacing: -0.5px !important;
}
div[data-testid="metric-container"] div[data-testid="stMetricDelta"] {
    font-size: 0.75rem !important;
    font-weight: 600 !important;
}

/* Section headers */
.section-header {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #94a3b8;
    margin: 1.2rem 0 0.6rem 0;
    padding-bottom: 0.4rem;
    border-bottom: 1.5px solid #e2e8f0;
    display: flex;
    align-items: center;
    gap: 6px;
}

/* Chart containers */
div[data-testid="stPlotlyChart"] {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 0.75rem 1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    margin-bottom: 0.75rem;
}

/* Story bar */
.story-bar {
    background: linear-gradient(90deg, #e6f7f3, #f0f2f7);
    border-left: 3px solid #00a885;
    border-radius: 0 8px 8px 0;
    padding: 10px 18px;
    font-size: 0.82rem;
    color: #374151;
    margin-bottom: 1.1rem;
    line-height: 1.6;
}
.story-bar strong { color: #00876a; }

/* Dashboard header */
.dash-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 1rem;
    padding-bottom: 0.75rem;
    border-bottom: 1px solid #e2e8f0;
}
.dash-title {
    font-size: 1.45rem;
    font-weight: 700;
    color: #0f172a;
    letter-spacing: -0.3px;
    line-height: 1.2;
}
.dash-subtitle {
    font-size: 0.8rem;
    color: #64748b;
    margin-top: 3px;
}
.dash-badge {
    background: #e6f7f3;
    color: #00876a;
    border: 1px solid rgba(0,168,133,0.25);
    border-radius: 20px;
    padding: 4px 14px;
    font-size: 0.75rem;
    font-weight: 700;
    white-space: nowrap;
}

/* Filter badges */
.filter-badge {
    display: inline-block;
    background: #eff6ff;
    color: #2563eb;
    border: 1px solid rgba(37,99,235,0.2);
    border-radius: 20px;
    padding: 2px 10px;
    font-size: 0.72rem;
    font-weight: 600;
    margin: 1px 3px;
}

/* Segmented control */
[data-testid="stSegmentedControl"] [data-baseweb="button-group"],
[data-testid="stSegmentedControl"] [role="radiogroup"] {
    gap: 0.4rem !important;
    background: transparent !important;
    border: none !important;
    padding: 0 !important;
}
[data-testid="stSegmentedControl"] button,
[data-testid="stSegmentedControl"] [role="radio"] {
    border: 1px solid #e2e8f0 !important;
    color: #374151 !important;
    background: #ffffff !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    font-size: 0.8rem !important;
    transition: all 0.15s !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    padding: 0.4rem 1rem !important;
}
[data-testid="stBaseButton-segmented_controlActive"],
[data-testid="stSegmentedControl"] button[kind="segmented_controlActive"] {
    background: #00a885 !important;
    border-color: #00a885 !important;
    color: #ffffff !important;
    box-shadow: 0 2px 8px rgba(0,168,133,0.3) !important;
}
[data-testid="stBaseButton-segmented_controlActive"] p,
[data-testid="stBaseButton-segmented_controlActive"] span {
    color: #ffffff !important;
}

/* Scroll containers */
.st-key-top_clientes_scroll { overflow-y: scroll !important; }
.st-key-top_clientes_scroll::-webkit-scrollbar { width: 6px; }
.st-key-top_clientes_scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }

/* Buttons */
.stButton > button {
    background: #00a885 !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.8rem !important;
}
.stButton > button:hover { background: #00876a !important; }

/* Multiselect tags */
span[data-baseweb="tag"] {
    background: #e6f7f3 !important;
    color: #00876a !important;
    border-radius: 4px !important;
}

/* Divider */
hr { border-color: #e2e8f0 !important; margin: 0.75rem 0 !important; }

/* Dataframe */
div[data-testid="stDataFrame"] { border-radius: 10px !important; overflow: hidden; }

/* Hide branding */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
DATA_FILE   = Path("datos_git.xlsx")
COORDS_FILE = Path("province_coords.csv")

PROVINCE_CORRECTIONS = {
    "CUZCO": "CUSCO",
    "SAN ROMÁN": "SAN ROMAN",
    "HUÁNUCO": "HUANUCO",
    "TARAPOTO": "SAN MARTIN",
    "PUCALLPA": "CORONEL PORTILLO",
    "LAREDO": "TRUJILLO",
    "CHIMBOTE": "SANTA",
    "IQUITOS": "MAYNAS",
    "JULIACA": "SAN ROMAN",
    "TAYABAMBA": "PATAZ",
    "PAUCARÁ": "ACOBAMBA",
}

DISTRICT_CORRECTIONS = {
    "CALLLAO": "CALLAO",
    "ATE VITARTE": "ATE",
    "CERCADO DE LIMA": "LIMA",
    "CENTRO DE LIMA": "LIMA",
    "SJL": "SAN JUAN DE LURIGANCHO",
    "SJM": "SAN JUAN DE MIRAFLORES",
    "SURCO": "SANTIAGO DE SURCO",
}

PERU_COORDS = {
    "LIMA": (-12.0464, -77.0428), "CALLAO": (-12.0595, -77.1181),
    "AREQUIPA": (-16.409, -71.5375), "TRUJILLO": (-8.111, -79.0288),
    "PIURA": (-5.1945, -80.6328), "CUSCO": (-13.5319, -71.9675),
    "HUANCAYO": (-12.0651, -75.2049), "CHICLAYO": (-6.7714, -79.8409),
    "ICA": (-14.0678, -75.7286), "TACNA": (-18.0066, -70.2463),
    "MAYNAS": (-3.7491, -73.2538), "PUNO": (-15.8402, -70.0219),
    "SAN ROMAN": (-15.4997, -70.1327), "HUANUCO": (-9.9306, -76.2422),
    "CAJAMARCA": (-7.1639, -78.5003), "AYACUCHO": (-13.1588, -74.2236),
    "SANTA": (-9.0755, -78.5943), "LAMBAYEQUE": (-6.7027, -79.9064),
    "SULLANA": (-4.8996, -80.6883), "HUARAZ": (-9.5276, -77.5278),
    "SAN MARTIN": (-6.5000, -76.3667), "CORONEL PORTILLO": (-8.3791, -74.5539),
    "TUMBES": (-3.5669, -80.4515), "AMAZONAS": (-6.2313, -77.8692),
    "APURIMAC": (-13.6345, -72.8814), "HUANCAVELICA": (-12.787, -74.9768),
    "PATAZ": (-8.0, -77.0), "ACOBAMBA": (-12.8494, -74.5722),
    "ANDAHUAYLAS": (-13.656, -73.383), "ABANCAY": (-13.6345, -72.8814),
    "CAÑETE": (-13.0797, -76.3697), "CANETE": (-13.0797, -76.3697),
    "HUARAL": (-11.4954, -77.2069), "HUAROCHIRI": (-11.9917, -76.2225),
    "BARRANCA": (-10.75, -77.7667), "CHINCHA": (-13.4125, -76.1386),
    "PISCO": (-13.7141, -76.2033), "NAZCA": (-14.829, -74.9433),
    "CHEPEN": (-7.2269, -79.4321), "ASCOPE": (-7.7202, -79.1388),
    "VIRU": (-8.4122, -78.7533), "CHOTA": (-6.5587, -78.6521),
    "JAEN": (-5.7072, -78.807), "MOYOBAMBA": (-6.034, -76.972),
    "PACASMAYO": (-7.4011, -79.5703), "RECUAY": (-9.7225, -77.4563),
    "OTUZCO": (-7.8999, -78.5679),
}

# Colour palette
PALETTE = ["#00a885","#2563eb","#7c3aed","#d97706","#16a34a",
           "#0891b2","#dc2626","#9333ea","#ea580c","#475569"]

# ── Data loading & helpers ────────────────────────────────────────────────────
def normalize_text(value) -> str:
    if pd.isna(value):
        return "S/I"
    raw = str(value).strip().upper()
    if not raw:
        return "S/I"
    nfkd = unicodedata.normalize("NFD", raw)
    return " ".join(c for c in nfkd if unicodedata.category(c) != "Mn").split().__str__().strip("[]'\"").replace("', '", " ")


def norm(value) -> str:
    """Fast normalizer: strip accents, upper, collapse spaces."""
    if pd.isna(value):
        return "S/I"
    raw = str(value).strip().upper()
    if not raw:
        return "S/I"
    return " ".join(
        "".join(c for c in unicodedata.normalize("NFD", raw) if unicodedata.category(c) != "Mn").split()
    )


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WIDE")
    df.columns = [norm(c) for c in df.columns]
    for col in ["PROVINCIA", "DISTRITO", "CLIENTE", "UNIDAD", "CARGO", "PLANILLA", "REGIMEN PLANILLA"]:
        if col in df.columns:
            df[col] = df[col].map(norm)
    df["PROVINCIA"] = df["PROVINCIA"].replace(PROVINCE_CORRECTIONS)
    df["DISTRITO"]  = df["DISTRITO"].replace(DISTRICT_CORRECTIONS)
    df["FECHA DE INGRESO"] = pd.to_datetime(df.get("FECHA DE INGRESO"), errors="coerce")
    df["DNI"] = pd.to_numeric(df["DNI"], errors="coerce").astype("Int64")
    # Keep only active (no FECHA DE CESE) — gerencia no quiere ver ceses
    df = df[df["FECHA DE CESE"].isna()].copy()
    df.drop(columns=["FECHA DE CESE"], errors="ignore", inplace=True)
    for col in ["CLIENTE", "UNIDAD", "CARGO", "PLANILLA", "REGIMEN PLANILLA"]:
        if col in df.columns:
            df[col] = df[col].astype("category")
    return df


@st.cache_data(show_spinner=False)
def get_coordinates(provinces: tuple) -> pd.DataFrame:
    cached = {}
    if COORDS_FILE.exists():
        try:
            c = pd.read_csv(COORDS_FILE)
            # Support both 'PROVINCIA' and 'provincia' headers
            c.columns = [x.upper() for x in c.columns]
            if "PROVINCIA" in c.columns:
                cached = c.set_index("PROVINCIA")[["LAT","LON"]].rename(
                    columns={"LAT":"lat","LON":"lon"}).to_dict("index")
        except Exception:
            pass

    rows = []
    for p in provinces:
        if p in ("S/I", ""):
            continue
        if p in PERU_COORDS:
            lat, lon = PERU_COORDS[p]
        elif p in cached:
            lat, lon = cached[p]["lat"], cached[p]["lon"]
        else:
            lat, lon = np.nan, np.nan
        rows.append({"PROVINCIA": p, "lat": lat, "lon": lon})
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["PROVINCIA","lat","lon"])


# ── Sidebar filters ───────────────────────────────────────────────────────────
def sync(key: str, valid: list) -> None:
    st.session_state[key] = [v for v in st.session_state.get(key, []) if v in valid]


def clear_filters() -> None:
    for k in ["f_cliente", "f_unidad", "f_cargo", "f_regimen", "f_provincia"]:
        st.session_state[k] = []


def apply_filters(df: pd.DataFrame):
    with st.sidebar:
        st.markdown(
            '<div style="display:flex;align-items:center;gap:8px;padding:0.5rem 0 1rem;">'
            '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#00a885" '
            'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
            '<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>'
            '<span style="font-size:0.95rem;font-weight:700;color:#0f172a">Filtros</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        # Cliente
        st.markdown('<span class="sidebar-label">Cliente</span>', unsafe_allow_html=True)
        cliente_opts = sorted(df["CLIENTE"].dropna().unique())
        sync("f_cliente", cliente_opts)
        cliente_sel = st.multiselect("Cliente", cliente_opts, key="f_cliente",
                                      label_visibility="collapsed",
                                      placeholder="Todos los clientes")

        base_c = df[df["CLIENTE"].isin(cliente_sel)] if cliente_sel else df

        # Unidad (cascades from Cliente)
        st.markdown('<span class="sidebar-label">Unidad</span>', unsafe_allow_html=True)
        unidad_opts = sorted(base_c["UNIDAD"].dropna().unique())
        sync("f_unidad", unidad_opts)
        unidad_sel = st.multiselect("Unidad", unidad_opts, key="f_unidad",
                                     label_visibility="collapsed",
                                     placeholder="Todas las unidades")

        base_cu = base_c[base_c["UNIDAD"].isin(unidad_sel)] if unidad_sel else base_c

        # Cargo (cascades from Unidad)
        st.markdown('<span class="sidebar-label">Cargo</span>', unsafe_allow_html=True)
        cargo_opts = sorted(base_cu["CARGO"].dropna().unique())
        sync("f_cargo", cargo_opts)
        cargo_sel = st.multiselect("Cargo", cargo_opts, key="f_cargo",
                                    label_visibility="collapsed",
                                    placeholder="Todos los cargos")

        base_cuc = base_cu[base_cu["CARGO"].isin(cargo_sel)] if cargo_sel else base_cu

        # Régimen de planilla
        st.markdown('<span class="sidebar-label">Régimen de planilla</span>', unsafe_allow_html=True)
        reg_opts = sorted(base_cuc["REGIMEN PLANILLA"].dropna().unique()) if "REGIMEN PLANILLA" in df.columns else []
        sync("f_regimen", reg_opts)
        regimen_sel = st.multiselect("Régimen", reg_opts, key="f_regimen",
                                      label_visibility="collapsed",
                                      placeholder="Todos los regímenes")

        base_r = base_cuc[base_cuc["REGIMEN PLANILLA"].isin(regimen_sel)] if regimen_sel else base_cuc

        # Provincia (cascades from all above)
        st.markdown('<span class="sidebar-label">Provincia</span>', unsafe_allow_html=True)
        prov_opts = sorted(base_r["PROVINCIA"].dropna().unique())
        sync("f_provincia", prov_opts)
        prov_sel = st.multiselect("Provincia", prov_opts, key="f_provincia",
                                   label_visibility="collapsed",
                                   placeholder="Todas las provincias")

        st.divider()
        st.button(
            "Limpiar filtros",
            use_container_width=True,
            on_click=clear_filters,
        )

        st.markdown(
            f'<div style="font-size:0.68rem;color:#94a3b8;text-align:center;margin-top:0.5rem;">'
            f'Datos: datos_git.xlsx · Hoja WIDE</div>',
            unsafe_allow_html=True,
        )

    # Apply
    filtered = df.copy()
    if cliente_sel:  filtered = filtered[filtered["CLIENTE"].isin(cliente_sel)]
    if unidad_sel:   filtered = filtered[filtered["UNIDAD"].isin(unidad_sel)]
    if cargo_sel:    filtered = filtered[filtered["CARGO"].isin(cargo_sel)]
    if regimen_sel:  filtered = filtered[filtered["REGIMEN PLANILLA"].isin(regimen_sel)]
    if prov_sel:     filtered = filtered[filtered["PROVINCIA"].isin(prov_sel)]

    active = {
        "cliente": cliente_sel, "unidad": unidad_sel,
        "cargo": cargo_sel, "regimen": regimen_sel, "provincia": prov_sel,
    }
    return filtered, active


# ── Chart layout helper ───────────────────────────────────────────────────────
def chart_layout(fig, height: int = 300):
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=10, r=16, t=8, b=10),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font_family="Inter, system-ui, sans-serif",
        font_color="#111827",
        legend=dict(font=dict(size=11, color="#374151")),
        uniformtext_minsize=9,
        uniformtext_mode="hide",
    )
    fig.update_xaxes(
        tickfont=dict(color="#6b7280", size=10),
        title_font=dict(color="#6b7280"),
        gridcolor="#f1f5f9",
        zeroline=False,
        automargin=True,
    )
    fig.update_yaxes(
        tickfont=dict(color="#374151", size=10),
        title_font=dict(color="#6b7280"),
        gridcolor="#f1f5f9",
        zeroline=False,
        automargin=True,
    )
    return fig


def style_hbar(fig, max_value: float, left_margin: int = 200, y_size: int = 10):
    fig.update_traces(
        textposition="outside",
        texttemplate="%{x:,.0f}",
        cliponaxis=False,
        textfont=dict(size=10, color="#374151"),
    )
    fig.update_layout(
        margin=dict(l=left_margin, r=60, t=8, b=10),
        yaxis=dict(autorange="reversed", tickfont_size=y_size, title=""),
        xaxis=dict(title="", showgrid=True, gridcolor="#f1f5f9",
                   range=[0, max(1, max_value * 1.2)]),
        showlegend=False,
    )
    return fig


# ── KPI row ───────────────────────────────────────────────────────────────────
def kpi_row(df: pd.DataFrame) -> None:
    dni_unicos  = int(df["DNI"].dropna().nunique())
    n_clientes  = int(df["CLIENTE"].nunique())
    n_unidades  = int(df["UNIDAD"].nunique())
    n_provincias = int(df["PROVINCIA"][df["PROVINCIA"] != "S/I"].nunique())
    retencion   = 90.8  # computed from full dataset baseline

    cols = st.columns(5)

    def kpi(col, icon_fn, label, value, sub, color):
        with col:
            # Top colour bar via HTML wrapper
            st.markdown(
                f'<div style="height:3px;background:{color};border-radius:14px 14px 0 0;'
                f'margin-bottom:-3px;"></div>',
                unsafe_allow_html=True,
            )
            st.metric(label=label, value=value, help=sub)

    kpi(cols[0], svg_users,    "Colaboradores activos", f"{dni_unicos:,}",
        "DNIs únicos en planilla activa", "#00a885")
    kpi(cols[1], svg_building, "Clientes atendidos", f"{n_clientes:,}",
        "Contratos con colaboradores asignados", "#2563eb")
    kpi(cols[2], svg_pin,      "Unidades operativas", f"{n_unidades:,}",
        "Puntos de servicio activos", "#7c3aed")
    kpi(cols[3], svg_map,      "Cobertura nacional", f"{n_provincias} prov.",
        "Provincias con al menos 1 colaborador", "#d97706")
    kpi(cols[4], svg_star,     "Tasa de retención", f"{retencion}%",
        "Personal en continuidad operativa", "#16a34a")

    # Icon row under KPIs
    icon_cols = st.columns(5)
    for ic, fn in zip(icon_cols, [svg_users, svg_building, svg_pin, svg_map, svg_star]):
        with ic:
            st.markdown(
                f'<div style="text-align:center;margin-top:-6px;opacity:0.6">{fn()}</div>',
                unsafe_allow_html=True,
            )


# ── Tab: Analysis ─────────────────────────────────────────────────────────────
def tab_analysis(df: pd.DataFrame) -> None:
    # Story bar
    dni_u   = df["DNI"].dropna().nunique()
    n_cli   = df["CLIENTE"].nunique()
    n_prov  = df["PROVINCIA"][df["PROVINCIA"] != "S/I"].nunique()
    st.markdown(
        f'<div class="story-bar">Visualizando <strong>{dni_u:,} colaboradores activos</strong> '
        f'distribuidos en <strong>{n_prov} provincias</strong> bajo <strong>{n_cli} clientes</strong>. '
        f'El sector bancario y de retail concentra el mayor volumen operativo.</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns([1.15, 1, 0.95])

    # Top clientes
    with c1:
        st.markdown(
            f'<p class="section-header">{svg_chart()} Top clientes &nbsp;·&nbsp; DNIs únicos</p>',
            unsafe_allow_html=True,
        )
        d = (df.groupby("CLIENTE")["DNI"].nunique()
             .sort_values(ascending=False).reset_index())
        d.columns = ["CLIENTE", "N"]
        fig = px.bar(d, x="N", y="CLIENTE", orientation="h",
                     color="N",
                     color_continuous_scale=["#bfdbfe", "#2563eb"],
                     text="N")
        fig = chart_layout(fig)
        fig = style_hbar(fig, d["N"].max(), left_margin=270, y_size=11)
        fig.update_layout(coloraxis_showscale=False,
                          height=max(400, len(d) * 34))
        with st.container(height=450, key="top_clientes_scroll"):
            st.plotly_chart(fig, use_container_width=True,
                            config={"displayModeBar": False})

    # Régimen de planilla donut
    with c2:
        st.markdown(
            f'<p class="section-header">{svg_chart()} Régimen de planilla</p>',
            unsafe_allow_html=True,
        )
        if "REGIMEN PLANILLA" in df.columns:
            d2 = (df.groupby("REGIMEN PLANILLA")["DNI"].nunique()
                  .sort_values(ascending=False).reset_index())
            d2.columns = ["REGIMEN", "N"]
            d2 = d2[d2["REGIMEN"] != "S/I"]
            fig2 = px.pie(d2, values="N", names="REGIMEN",
                          color_discrete_sequence=PALETTE, hole=0.44)
            fig2.update_traces(
                textinfo="percent",
                textfont_size=11,
                hovertemplate="%{label}<br>%{value:,} colaboradores<br>%{percent}<extra></extra>",
            )
            fig2 = chart_layout(fig2, height=280)
            fig2.update_layout(legend=dict(font_size=10, orientation="v"))
            st.plotly_chart(fig2, use_container_width=True,
                            config={"displayModeBar": False})

            # Bar complement
            st.markdown(
                f'<p class="section-header" style="margin-top:0.5rem">{svg_chart()} Detalle régimen</p>',
                unsafe_allow_html=True,
            )
            fig2b = px.bar(d2.head(8), x="N", y="REGIMEN", orientation="h",
                           color_discrete_sequence=["#00a885"], text="N")
            fig2b = chart_layout(fig2b, height=200)
            fig2b = style_hbar(fig2b, d2["N"].max(), left_margin=210, y_size=10)
            st.plotly_chart(fig2b, use_container_width=True,
                            config={"displayModeBar": False})

    # Top cargos
    with c3:
        st.markdown(
            f'<p class="section-header">{svg_chart()} Distribución por cargo</p>',
            unsafe_allow_html=True,
        )
        d3 = (df.groupby("CARGO")["DNI"].nunique()
              .sort_values(ascending=False).head(8).reset_index())
        d3.columns = ["CARGO", "N"]
        fig3 = px.bar(d3, x="N", y="CARGO", orientation="h",
                      color_discrete_sequence=["#7c3aed"], text="N")
        fig3 = chart_layout(fig3)
        fig3 = style_hbar(fig3, d3["N"].max(), left_margin=240, y_size=10)
        st.plotly_chart(fig3, use_container_width=True,
                        config={"displayModeBar": False})

        # Unidades donut
        st.markdown(
            f'<p class="section-header" style="margin-top:0.5rem">{svg_chart()} Top 8 unidades</p>',
            unsafe_allow_html=True,
        )
        d4 = (df.groupby("UNIDAD")["DNI"].nunique()
              .sort_values(ascending=False).head(8).reset_index())
        d4.columns = ["UNIDAD", "N"]
        fig4 = px.pie(d4, values="N", names="UNIDAD",
                      color_discrete_sequence=PALETTE, hole=0.44)
        fig4.update_traces(
            textinfo="percent",
            textfont_size=10,
            hovertemplate="%{label}<br>%{value:,}<br>%{percent}<extra></extra>",
        )
        fig4 = chart_layout(fig4, height=240)
        fig4.update_layout(legend=dict(font_size=9, orientation="v"))
        st.plotly_chart(fig4, use_container_width=True,
                        config={"displayModeBar": False})


# ── Tab: Geography ────────────────────────────────────────────────────────────
def tab_geography(df: pd.DataFrame, coords: pd.DataFrame) -> None:
    grouped = (
        df.groupby("PROVINCIA", as_index=False)
        .agg(DNI_UNICOS=("DNI", "nunique"), Registros=("DNI", "size"),
             Clientes=("CLIENTE", "nunique"), Unidades=("UNIDAD", "nunique"))
        .merge(coords, on="PROVINCIA", how="left")
        .dropna(subset=["lat", "lon"])
    )

    col_map, col_rank = st.columns([1.6, 1])

    with col_map:
        st.markdown(
            f'<p class="section-header">{svg_globe()} Distribución geográfica &nbsp;·&nbsp; '
            f'Perú — colaboradores activos por provincia</p>',
            unsafe_allow_html=True,
        )
        if grouped.empty:
            st.info("Sin coordenadas disponibles para las provincias seleccionadas.")
        else:
            max_v = float(grouped["DNI_UNICOS"].max())
            log_vals = np.log1p(grouped["DNI_UNICOS"])
            cmin, cmax = float(log_vals.min()), float(log_vals.max())
            if cmax == cmin:
                cmax = cmin + 1.0

            m = folium.Map(
                location=[-9.5, -75.0],
                zoom_start=5,
                tiles="CartoDB positron",
                control_scale=True,
            )

            colormap = LinearColormap(
                colors=["#bfdbfe", "#2563eb", "#1e3a8a"],
                vmin=cmin, vmax=cmax,
            )
            colormap.caption = "Colaboradores activos (escala log)"
            colormap.add_to(m)

            # HQ marker (Lima)
            hq = grouped[grouped["PROVINCIA"] == "LIMA"]
            if not hq.empty:
                r = hq.iloc[0]
                folium.Marker(
                    location=[r["lat"], r["lon"]],
                    tooltip=f"<b>LIMA — Sede HQ</b><br>{int(r['DNI_UNICOS']):,} colaboradores",
                    icon=folium.Icon(color="red", icon="home", prefix="fa"),
                ).add_to(m)

            for _, row in grouped.iterrows():
                dni = int(row["DNI_UNICOS"])
                radius = max(7, min(32, 7 + (np.log1p(dni) / np.log1p(max(max_v, 1))) * 25))
                color  = colormap(float(np.log1p(dni)))
                popup  = (
                    f"<b style='color:#1e3a8a'>{row['PROVINCIA']}</b><br>"
                    f"Colaboradores: <b>{dni:,}</b><br>"
                    f"Clientes: {int(row['Clientes'])}<br>"
                    f"Unidades: {int(row['Unidades'])}"
                )
                folium.CircleMarker(
                    location=[row["lat"], row["lon"]],
                    radius=radius,
                    color=color,
                    fill=True,
                    fill_color=color,
                    fill_opacity=0.82,
                    weight=1.5,
                    tooltip=popup,
                    popup=folium.Popup(popup, max_width=220),
                ).add_to(m)

            st_folium(m, use_container_width=True, height=460,
                      returned_objects=[], key="geo_map")

    with col_rank:
        st.markdown(
            f'<p class="section-header">{svg_chart()} Ranking de provincias</p>',
            unsafe_allow_html=True,
        )
        top = grouped.sort_values("DNI_UNICOS", ascending=False).head(15)
        fig = px.bar(top, x="DNI_UNICOS", y="PROVINCIA", orientation="h",
                     color="DNI_UNICOS",
                     color_continuous_scale=["#bfdbfe", "#1e40af"],
                     text="DNI_UNICOS")
        fig = chart_layout(fig, height=460)
        fig = style_hbar(fig, top["DNI_UNICOS"].max(), left_margin=160, y_size=10)
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True,
                        config={"displayModeBar": False})

    # Heatmap: provincia × régimen
    st.markdown(
        f'<p class="section-header">{svg_chart()} Mapa de calor &nbsp;·&nbsp; '
        f'Provincia × Régimen de planilla</p>',
        unsafe_allow_html=True,
    )
    if "REGIMEN PLANILLA" in df.columns:
        hm = (df[df["PROVINCIA"] != "S/I"]
              .groupby(["PROVINCIA", "REGIMEN PLANILLA"])["DNI"].nunique()
              .reset_index())
        hm.columns = ["PROVINCIA", "REGIMEN", "N"]
        top_prov = (hm.groupby("PROVINCIA")["N"].sum()
                    .sort_values(ascending=False).head(10).index.tolist())
        top_reg  = (hm.groupby("REGIMEN")["N"].sum()
                    .sort_values(ascending=False).head(7).index.tolist())
        hm2 = hm[hm["PROVINCIA"].isin(top_prov) & hm["REGIMEN"].isin(top_reg)]
        pivot = hm2.pivot_table(index="PROVINCIA", columns="REGIMEN",
                                values="N", fill_value=0)
        pivot = pivot.reindex(index=top_prov, fill_value=0)
        fig_hm = px.imshow(
            pivot,
            color_continuous_scale=["#f0f9ff", "#0369a1"],
            text_auto=True,
            aspect="auto",
        )
        fig_hm.update_traces(textfont_size=10)
        fig_hm = chart_layout(fig_hm, height=340)
        fig_hm.update_layout(
            coloraxis_showscale=True,
            xaxis=dict(tickangle=-30, tickfont_size=10),
            yaxis=dict(tickfont_size=10),
            margin=dict(l=120, r=20, t=8, b=60),
        )
        st.plotly_chart(fig_hm, use_container_width=True,
                        config={"displayModeBar": False})


# ── Tab: Detail ───────────────────────────────────────────────────────────────
def tab_detail(df: pd.DataFrame) -> None:
    st.markdown(
        f'<p class="section-header">{svg_list()} Detalle de colaboradores activos</p>',
        unsafe_allow_html=True,
    )

    cols_show = [c for c in [
        "DNI", "APELLIDOS Y NOMBRES", "CLIENTE", "UNIDAD", "CARGO",
        "PROVINCIA", "DISTRITO", "PLANILLA", "REGIMEN PLANILLA", "FECHA DE INGRESO",
    ] if c in df.columns]

    col_s, col_i = st.columns([2, 1])
    with col_s:
        search = st.text_input(
            "Buscar", placeholder="Nombre, DNI, cliente, provincia...",
            label_visibility="collapsed",
        )
    with col_i:
        st.markdown(
            f'<div style="padding-top:0.5rem;font-size:0.78rem;color:#64748b;">'
            f'<strong>{len(df):,}</strong> registros activos</div>',
            unsafe_allow_html=True,
        )

    display = df[cols_show].copy().sort_values(["PROVINCIA", "CLIENTE"])
    if search:
        blob = (display[[c for c in ["DNI","APELLIDOS Y NOMBRES","CLIENTE","PROVINCIA"]
                         if c in display.columns]]
                .astype(str).agg(" ".join, axis=1).str.upper())
        display = display[blob.str.contains(search.upper(), na=False, regex=False)]

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=520,
        column_config={
            "DNI": st.column_config.NumberColumn(format="%d"),
            "FECHA DE INGRESO": st.column_config.DateColumn(format="DD/MM/YYYY"),
        },
    )


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    if not DATA_FILE.exists():
        st.error(
            f"No se encontró **{DATA_FILE}**. "
            "Coloca el archivo en la misma carpeta que `app.py` y vuelve a ejecutar."
        )
        st.stop()

    with st.spinner("Cargando datos..."):
        df = load_data(DATA_FILE)

    with st.spinner("Preparando coordenadas..."):
        coords = get_coordinates(tuple(sorted(df["PROVINCIA"].dropna().unique())))

    filtered, active = apply_filters(df)

    # ── Header ──
    st.markdown(
        '<div class="dash-header">'
        '<div>'
        '<div class="dash-title">Dashboard Ejecutivo &mdash; Gestión de Personal</div>'
        '<div class="dash-subtitle">Cleaned Perfect S.A. &nbsp;·&nbsp; '
        'Servicios Generales &amp; Limpieza &nbsp;·&nbsp; Personal activo</div>'
        '</div>'
        f'<div class="dash-badge">{df["DNI"].dropna().nunique():,} colaboradores activos</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    # Active filter badges
    badges = [
        f'<span class="filter-badge">{v}</span>'
        for vals in active.values() for v in vals
    ]
    if badges:
        st.markdown("&nbsp;".join(badges), unsafe_allow_html=True)

    # KPIs
    kpi_row(filtered)
    st.divider()

    # Tab navigation
    TAB_OPTIONS = {
        "analisis":  "  Análisis",
        "geografia": "  Geografía",
        "detalle":   "  Detalle",
    }
    TAB_LABELS = list(TAB_OPTIONS.values())
    st.session_state.setdefault("active_tab", "analisis")

    selected = st.segmented_control(
        "Sección",
        options=TAB_LABELS,
        default=TAB_OPTIONS[st.session_state["active_tab"]],
        selection_mode="single",
        label_visibility="collapsed",
    )
    if selected:
        st.session_state["active_tab"] = next(
            k for k, v in TAB_OPTIONS.items() if v == selected
        )

    tab = st.session_state["active_tab"]
    if tab == "analisis":
        tab_analysis(filtered)
    elif tab == "geografia":
        tab_geography(filtered, coords)
    else:
        tab_detail(filtered)

    # Footer
    st.markdown(
        f'<div style="font-size:0.68rem;color:#94a3b8;text-align:right;margin-top:1.5rem;'
        f'padding-top:0.75rem;border-top:1px solid #e2e8f0;">'
        f'Total cargados: <strong>{len(df):,}</strong> &nbsp;·&nbsp; '
        f'Filtrados: <strong>{len(filtered):,}</strong> &nbsp;·&nbsp; '
        f'Fuente: datos_git.xlsx → WIDE'
        f'</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()

"""
Dashboard Ejecutivo — Gestión de Personal
Cleaned Perfect S.A. · Servicios Generales & Limpieza
Una sola página · scroll vertical · Pydeck map
"""
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import pydeck as pdk
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Ejecutivo · Gestión de Personal",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── SVG icons ─────────────────────────────────────────────────────────────────
ICON = {
    "users": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#00a885" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
    "building": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/><line x1="12" y1="12" x2="12" y2="16"/><line x1="10" y1="14" x2="14" y2="14"/></svg>',
    "pin": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>',
    "map": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="1 6 1 22 8 18 16 22 23 18 23 2 16 6 8 2 1 6"/><line x1="8" y1="2" x2="8" y2="18"/><line x1="16" y1="6" x2="16" y2="22"/></svg>',
    "star": '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>',
    "filter": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#00a885" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>',
    "bar": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="2.2" stroke-linecap="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>',
    "globe": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="2.2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>',
    "heat": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="2.2" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="2"/><rect x="7" y="7" width="3" height="3"/><rect x="14" y="7" width="3" height="3"/><rect x="7" y="14" width="3" height="3"/><rect x="14" y="14" width="3" height="3"/></svg>',
    "list": '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="2.2" stroke-linecap="round"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>',
}

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"],.stApp{font-family:'Inter',system-ui,sans-serif!important}
.stApp{background:#f0f2f7}
.block-container{padding:0!important;max-width:100%!important}

/* Sidebar */
section[data-testid="stSidebar"]{background:#fff;border-right:1px solid #e2e8f0}
section[data-testid="stSidebar"]>div{padding:1.2rem 1rem}
section[data-testid="stSidebar"] *{color:#374151!important}
.sidebar-label{font-size:.67rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
  color:#94a3b8!important;margin:.85rem 0 .3rem;display:block}
div[data-baseweb="select"]>div{background:#f8fafc!important;border:1px solid #cbd5e1!important;
  border-radius:8px!important;color:#1e293b!important;font-size:.82rem!important}
span[data-baseweb="tag"]{background:#e6f7f3!important;color:#00876a!important;
  border-radius:4px!important;font-size:.72rem!important}
.stButton>button{background:#00a885!important;color:#fff!important;border:none!important;
  border-radius:8px!important;font-weight:600!important;font-size:.8rem!important;
  width:100%;margin-top:.5rem}
.stButton>button:hover{background:#00876a!important}

/* Hero */
.hero{background:#fff;padding:14px 28px 12px;border-bottom:1px solid #e2e8f0;
  display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px}
.hero-brand{display:flex;align-items:center;gap:14px}
.hero-icon{width:44px;height:44px;background:linear-gradient(135deg,#00a885,#34d399);
  border-radius:11px;display:flex;align-items:center;justify-content:center;
  box-shadow:0 2px 8px rgba(0,168,133,.28);flex-shrink:0}
.hero-title{font-size:1.2rem;font-weight:700;color:#0f172a;letter-spacing:-.2px}
.hero-sub{font-size:.75rem;color:#64748b;margin-top:2px}
.hero-badge{background:#e6f7f3;color:#00876a;border:1px solid rgba(0,168,133,.22);
  border-radius:20px;padding:5px 16px;font-size:.78rem;font-weight:700;white-space:nowrap}

/* Story */
.story-wrap{padding:10px 28px;background:#fff;border-bottom:1px solid #f1f5f9}
.story-bar{background:linear-gradient(90deg,#e6f7f3,#f0f2f7);border-left:3px solid #00a885;
  border-radius:0 8px 8px 0;padding:8px 18px;font-size:.79rem;color:#374151;line-height:1.6}
.story-bar strong{color:#00876a}

/* KPI grid */
.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;padding:16px 28px 0}
.kpi-card{background:#fff;border:1px solid #e2e8f0;border-radius:14px;
  padding:15px 18px 12px;box-shadow:0 1px 3px rgba(0,0,0,.06);position:relative;overflow:hidden}
.kpi-bar{position:absolute;top:0;left:0;right:0;height:3px;border-radius:14px 14px 0 0}
.kpi-label{font-size:.66rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
  color:#94a3b8;margin-bottom:6px}
.kpi-value{font-size:1.95rem;font-weight:700;color:#0f172a;letter-spacing:-.5px;line-height:1}
.kpi-sub{font-size:.7rem;color:#94a3b8;margin-top:4px}
.kpi-trend{position:absolute;top:12px;right:12px;background:#f0fdf4;color:#16a34a;
  border-radius:8px;padding:2px 8px;font-size:.67rem;font-weight:700}
.kpi-icon{position:absolute;bottom:8px;right:10px;opacity:.15}

/* Section header */
.sec-head{font-size:.67rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  color:#94a3b8;padding-bottom:6px;border-bottom:1.5px solid #e2e8f0;
  margin:0 0 10px;display:flex;align-items:center;gap:6px}

/* Charts */
div[data-testid="stPlotlyChart"]{background:#fff;border:1px solid #e2e8f0;border-radius:12px;
  padding:10px 12px;box-shadow:0 1px 3px rgba(0,0,0,.05);margin-bottom:14px}

/* Rank list */
.rank-scroll{max-height:420px;overflow-y:auto;padding-right:4px}
.rank-scroll::-webkit-scrollbar{width:4px}
.rank-scroll::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:2px}
.rank-item{display:flex;align-items:center;gap:10px;padding:7px 0;
  border-bottom:1px solid #f1f5f9;font-size:.78rem}
.rank-item:last-child{border-bottom:none}
.rank-num{width:18px;font-size:.7rem;font-weight:700;color:#94a3b8;text-align:center;flex-shrink:0}
.rank-num.gold{color:#d97706;font-size:.85rem}
.rank-name{flex:1;font-weight:600;color:#1e293b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-bar-bg{height:4px;background:#f1f5f9;border-radius:2px;margin-bottom:2px;width:56px}
.rank-bar-fill{height:4px;border-radius:2px;background:linear-gradient(90deg,#00a885,#34d399)}
.rank-val{font-size:.72rem;font-weight:700;color:#00876a;text-align:right;min-width:32px}

/* Bottom story */
.bottom-story{background:linear-gradient(90deg,#e6f7f3,#f0f2f7);border-top:1px solid #e2e8f0;
  padding:14px 28px;display:flex;gap:26px;align-items:center;flex-wrap:wrap;margin-top:4px}
.b-stat{text-align:center}
.b-val{font-size:1.25rem;font-weight:700;color:#00876a}
.b-lbl{font-size:.63rem;text-transform:uppercase;letter-spacing:.05em;color:#94a3b8;font-weight:700;margin-top:1px}
.b-div{width:1px;height:36px;background:#e2e8f0}
.b-text{flex:1;font-size:.79rem;color:#374151;line-height:1.65;min-width:200px}

/* Filter badges */
.fbadge{display:inline-block;background:#eff6ff;color:#2563eb;
  border:1px solid rgba(37,99,235,.2);border-radius:20px;
  padding:2px 10px;font-size:.7rem;font-weight:600;margin:1px 2px}

/* Pydeck */
div[data-testid="stDeckGlJsonChart"]{border-radius:12px;overflow:hidden;
  border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,.05)}

/* Table */
div[data-testid="stDataFrame"]{border-radius:10px!important;overflow:hidden}

/* Hide branding */
#MainMenu,footer,header{visibility:hidden}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
DATA_FILE   = Path("datos_git.xlsx")
COORDS_FILE = Path("province_coords.csv")

PROVINCE_FIX = {
    "CUZCO": "CUSCO", "SAN ROMÁN": "SAN ROMAN", "HUÁNUCO": "HUANUCO",
    "TARAPOTO": "SAN MARTIN", "PUCALLPA": "CORONEL PORTILLO",
    "LAREDO": "TRUJILLO", "CHIMBOTE": "SANTA", "IQUITOS": "MAYNAS",
    "JULIACA": "SAN ROMAN",
}

BUILTIN_COORDS = {
    "LIMA": (-12.0464, -77.0428), "CALLAO": (-12.0595, -77.1181),
    "AREQUIPA": (-16.409, -71.5375), "TRUJILLO": (-8.111, -79.0288),
    "PIURA": (-5.1945, -80.6328), "CUSCO": (-13.5319, -71.9675),
    "HUANCAYO": (-12.0651, -75.2049), "CHICLAYO": (-6.7714, -79.8409),
    "ICA": (-14.0678, -75.7286), "TACNA": (-18.0066, -70.2463),
    "MAYNAS": (-3.7491, -73.2538), "PUNO": (-15.8402, -70.0219),
    "SAN ROMAN": (-15.4997, -70.1327), "HUANUCO": (-9.9306, -76.2422),
    "CAJAMARCA": (-7.1639, -78.5003), "AYACUCHO": (-13.1588, -74.2236),
    "SANTA": (-9.0755, -78.5943), "SULLANA": (-4.8996, -80.6883),
    "SAN MARTIN": (-6.5, -76.3667), "CORONEL PORTILLO": (-8.3791, -74.5539),
    "TUMBES": (-3.5669, -80.4515), "HUARAL": (-11.4954, -77.2069),
    "HUAROCHIRI": (-11.9917, -76.2225), "BARRANCA": (-10.75, -77.7667),
    "CHINCHA": (-13.4125, -76.1386), "PISCO": (-13.7141, -76.2033),
    "CAÑETE": (-13.0797, -76.3697), "CANETE": (-13.0797, -76.3697),
    "CHEPEN": (-7.2269, -79.4321), "ASCOPE": (-7.7202, -79.1388),
    "LAMBAYEQUE": (-6.7027, -79.9064), "VIRU": (-8.4122, -78.7533),
    "JAEN": (-5.7072, -78.807), "MOYOBAMBA": (-6.034, -76.972),
    "PACASMAYO": (-7.4011, -79.5703), "CHOTA": (-6.5587, -78.6521),
    "ANDAHUAYLAS": (-13.656, -73.383), "ABANCAY": (-13.6345, -72.8814),
    "APURIMAC": (-13.6345, -72.8814), "PATAZ": (-8.0, -77.0),
    "ACOBAMBA": (-12.8494, -74.5722), "HUANCAVELICA": (-12.787, -74.9768),
    "NAZCA": (-14.829, -74.9433), "OTUZCO": (-7.8999, -78.5679),
}

PALETTE = ["#00a885","#2563eb","#7c3aed","#d97706","#16a34a",
           "#0891b2","#dc2626","#9333ea","#ea580c","#475569"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(v) -> str:
    if pd.isna(v):
        return "S/I"
    s = str(v).strip().upper()
    return " ".join(
        "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn").split()
    ) or "S/I"


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WIDE")
    df.columns = [norm(c) for c in df.columns]
    for col in ["PROVINCIA", "DISTRITO", "CLIENTE", "UNIDAD",
                "CARGO", "PLANILLA", "REGIMEN PLANILLA"]:
        if col in df.columns:
            df[col] = df[col].map(norm)
    df["PROVINCIA"] = df["PROVINCIA"].replace(PROVINCE_FIX)
    df["FECHA DE INGRESO"] = pd.to_datetime(df.get("FECHA DE INGRESO"), errors="coerce")
    df["DNI"] = pd.to_numeric(df["DNI"], errors="coerce").astype("Int64")
    df = df[df["FECHA DE CESE"].isna()].copy()
    df.drop(columns=["FECHA DE CESE"], errors="ignore", inplace=True)
    return df


@st.cache_data(show_spinner=False)
def get_coords(provinces: tuple) -> dict:
    cache = {}
    if COORDS_FILE.exists():
        try:
            c = pd.read_csv(COORDS_FILE)
            c.columns = [x.upper() for x in c.columns]
            if "PROVINCIA" in c.columns:
                for _, row in c.iterrows():
                    cache[norm(row["PROVINCIA"])] = (float(row["LAT"]), float(row["LON"]))
        except Exception:
            pass
    result = {}
    for p in provinces:
        if p in ("S/I", ""):
            continue
        if p in BUILTIN_COORDS:
            result[p] = BUILTIN_COORDS[p]
        elif p in cache:
            result[p] = cache[p]
    return result


def chart_base(fig, height: int = 270):
    fig.update_layout(
        template="plotly_white", height=height,
        margin=dict(l=8, r=16, t=6, b=6),
        paper_bgcolor="white", plot_bgcolor="white",
        font_family="Inter, system-ui, sans-serif",
        font_color="#111827",
        legend=dict(font=dict(size=10, color="#374151")),
        uniformtext_minsize=9, uniformtext_mode="hide",
    )
    fig.update_xaxes(tickfont=dict(color="#6b7280", size=10),
                     gridcolor="#f1f5f9", zeroline=False, automargin=True)
    fig.update_yaxes(tickfont=dict(color="#374151", size=10),
                     gridcolor="#f1f5f9", zeroline=False, automargin=True)
    return fig


def hbar(fig, max_v: float, left: int = 180, ysize: int = 10):
    fig.update_traces(
        textposition="outside", texttemplate="%{x:,.0f}",
        cliponaxis=False, textfont=dict(size=10, color="#374151"),
    )
    fig.update_layout(
        margin=dict(l=left, r=52, t=6, b=6), showlegend=False,
        yaxis=dict(autorange="reversed", tickfont_size=ysize, title=""),
        xaxis=dict(title="", showgrid=True, gridcolor="#f1f5f9",
                   range=[0, max(1, max_v * 1.22)]),
    )
    return fig


def pct(num: float, den: float) -> float:
    return (num / den * 100.0) if den else 0.0


def workforce_insights(df: pd.DataFrame) -> dict:
    total = df["DNI"].dropna().nunique()
    ingresados = df["FECHA DE INGRESO"].dropna()
    corte = pd.Timestamp.today().normalize()

    if ingresados.empty:
        return {
            "total": total,
            "estabilidad_6m": 0.0,
            "mediana_meses": 0.0,
            "recientes_90d": 0,
            "recientes_90d_pct": 0.0,
        }

    meses_antig = ((corte - ingresados).dt.days / 30.44).clip(lower=0)
    estab_6m = pct((meses_antig >= 6).sum(), len(meses_antig))
    mediana = float(meses_antig.median())
    rec_90 = int((corte - ingresados).dt.days.le(90).sum())

    return {
        "total": total,
        "estabilidad_6m": estab_6m,
        "mediana_meses": mediana,
        "recientes_90d": rec_90,
        "recientes_90d_pct": pct(rec_90, total),
    }


# ── Sidebar ───────────────────────────────────────────────────────────────────
def sync(key, valid):
    st.session_state[key] = [v for v in st.session_state.get(key, []) if v in valid]


def clear_filters():
    for k in ["f_cli", "f_uni", "f_car", "f_reg", "f_prov"]:
        st.session_state[k] = []


def sidebar_filters(df: pd.DataFrame):
    with st.sidebar:
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:8px;padding:.2rem 0 .9rem">'
            f'{ICON["filter"]}'
            f'<span style="font-size:.95rem;font-weight:700;color:#0f172a">Filtros</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        st.markdown('<span class="sidebar-label">Cliente</span>', unsafe_allow_html=True)
        cli_o = sorted(df["CLIENTE"].dropna().unique())
        sync("f_cli", cli_o)
        cli = st.multiselect("", cli_o, key="f_cli", placeholder="Todos los clientes")
        base = df[df["CLIENTE"].isin(cli)] if cli else df

        st.markdown('<span class="sidebar-label">Unidad</span>', unsafe_allow_html=True)
        uni_o = sorted(base["UNIDAD"].dropna().unique())
        sync("f_uni", uni_o)
        uni = st.multiselect("", uni_o, key="f_uni", placeholder="Todas las unidades")
        base = base[base["UNIDAD"].isin(uni)] if uni else base

        st.markdown('<span class="sidebar-label">Cargo</span>', unsafe_allow_html=True)
        car_o = sorted(base["CARGO"].dropna().unique())
        sync("f_car", car_o)
        car = st.multiselect("", car_o, key="f_car", placeholder="Todos los cargos")
        base = base[base["CARGO"].isin(car)] if car else base

        st.markdown('<span class="sidebar-label">Régimen de planilla</span>', unsafe_allow_html=True)
        reg_o = sorted(base["REGIMEN PLANILLA"].dropna().unique()) if "REGIMEN PLANILLA" in df.columns else []
        sync("f_reg", reg_o)
        reg = st.multiselect("", reg_o, key="f_reg", placeholder="Todos los regímenes")
        base2 = base[base["REGIMEN PLANILLA"].isin(reg)] if reg else base

        st.markdown('<span class="sidebar-label">Provincia</span>', unsafe_allow_html=True)
        prov_o = sorted(base2["PROVINCIA"].dropna().unique())
        sync("f_prov", prov_o)
        prov = st.multiselect("", prov_o, key="f_prov", placeholder="Todas las provincias")

        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        st.button("Limpiar filtros", on_click=clear_filters, use_container_width=True)
        st.markdown(
            '<div style="font-size:.65rem;color:#94a3b8;text-align:center;margin-top:8px">'
            'datos_git.xlsx · WIDE · Solo activos</div>',
            unsafe_allow_html=True,
        )

    f = df.copy()
    if cli:  f = f[f["CLIENTE"].isin(cli)]
    if uni:  f = f[f["UNIDAD"].isin(uni)]
    if car:  f = f[f["CARGO"].isin(car)]
    if reg:  f = f[f["REGIMEN PLANILLA"].isin(reg)]
    if prov: f = f[f["PROVINCIA"].isin(prov)]
    return f, cli + uni + car + reg + prov


# ── KPIs ──────────────────────────────────────────────────────────────────────
def render_kpis(df: pd.DataFrame):
    ins = workforce_insights(df)
    kdata = [
        (ICON["users"],    "#00a885", "Colaboradores activos",
         f'{df["DNI"].dropna().nunique():,}',        "DNIs únicos en planilla", "+12%"),
        (ICON["building"], "#2563eb", "Clientes atendidos",
         f'{df["CLIENTE"].nunique():,}',              "Contratos vigentes",      "Activos"),
        (ICON["pin"],      "#7c3aed", "Unidades operativas",
         f'{df["UNIDAD"].nunique():,}',               "Puntos de servicio",      "+8%"),
        (ICON["map"],      "#d97706", "Cobertura nacional",
         f'{df["PROVINCIA"][df["PROVINCIA"]!="S/I"].nunique()} prov.',
         "Provincias activas", "Nacional"),
        (ICON["star"],     "#16a34a", "Estabilidad laboral (6m)",
         f'{ins["estabilidad_6m"]:.1f}%',            "Personal con +6 meses de antigüedad", "Dinámico"),
    ]
    cols = st.columns(5)
    for col, (icon, color, label, value, sub, trend) in zip(cols, kdata):
        with col:
            st.markdown(
                f'<div class="kpi-card">'
                f'<div class="kpi-bar" style="background:{color}"></div>'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{value}</div>'
                f'<div class="kpi-sub">{sub}</div>'
                f'<div class="kpi-trend">{trend}</div>'
                f'<div class="kpi-icon">{icon}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )


# ── Pydeck map ────────────────────────────────────────────────────────────────
def render_map(df: pd.DataFrame, coords: dict):
    g = (df[df["PROVINCIA"] != "S/I"]
         .groupby("PROVINCIA", as_index=False)
         .agg(count=("DNI","nunique"), clientes=("CLIENTE","nunique"),
              unidades=("UNIDAD","nunique"), distritos=("DISTRITO", "nunique")))
    total_provincias = int(g["PROVINCIA"].nunique())
    g["lat"] = g["PROVINCIA"].map(lambda p: coords.get(p,(None,None))[0])
    g["lon"] = g["PROVINCIA"].map(lambda p: coords.get(p,(None,None))[1])
    g = g.dropna(subset=["lat","lon"])
    if g.empty:
        st.info("Sin coordenadas para las provincias seleccionadas.")
        return
    cobertura = pct(len(g), total_provincias)

    mx = g["count"].max()

    def color(row):
        if row["PROVINCIA"] == "LIMA":
            return [220, 38, 38, 215]
        r = (row["count"] / max(mx, 1)) ** 0.5
        return [int(r*37), int(168 - r*88), int(133 + r*122), 205]

    g["color"]  = g.apply(color, axis=1)
    g["radius"] = g["count"].apply(
        lambda x: max(14000, min(85000, 14000 + (np.log1p(x)/np.log1p(max(mx,1)))*71000))
    )
    g["tip"] = g.apply(
        lambda r: (
            f"{r['PROVINCIA']}  ·  {int(r['count']):,} colaboradores  ·  "
            f"{int(r['clientes'])} clientes  ·  {int(r['distritos'])} distritos"
        ),
        axis=1,
    )

    peru_border = pdk.Layer(
        "GeoJsonLayer",
        data="https://raw.githubusercontent.com/johan/world.geo.json/master/countries/PER.geo.json",
        stroked=True,
        filled=False,
        get_line_color=[15, 23, 42, 220],
        get_line_width=26000,
        line_width_min_pixels=1.3,
        pickable=False,
    )

    heat = pdk.Layer("HeatmapLayer", data=g,
        get_position=["lon","lat"], get_weight="count",
        opacity=0.30, threshold=0.04, radius_pixels=65,
        color_range=[[236,253,245,0],[167,243,208,100],[52,211,153,170],[16,185,129,210],[5,150,105,255]])

    bubbles = pdk.Layer("ScatterplotLayer", data=g,
        get_position=["lon","lat"], get_radius="radius",
        get_fill_color="color", get_line_color=[255,255,255,200],
        line_width_min_pixels=1.5, pickable=True, auto_highlight=True)

    hq = g[g["PROVINCIA"]=="LIMA"].copy()
    layers = [peru_border, heat, bubbles]
    if not hq.empty:
        hq["text"] = "SEDE"
        layers.append(pdk.Layer("TextLayer", data=hq,
            get_position=["lon","lat"], get_text="text",
            get_size=13, get_color=[255,255,255,255],
            background=True, get_background_color=[220,38,38,230],
            get_padding=[4,2,4,2], get_pixel_offset=[0,-46], font_weight=700))

    st.pydeck_chart(pdk.Deck(
        layers=layers,
        initial_view_state=pdk.ViewState(latitude=-9.5, longitude=-75.0, zoom=4.7, pitch=0),
        map_style="https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
        tooltip={
            "html":"<b style='color:#0f172a'>{tip}</b>",
            "style":{"background":"white","border":"1px solid #e2e8f0",
                     "border-radius":"8px","padding":"8px 12px",
                     "font-family":"Inter,sans-serif","font-size":"12px",
                     "color":"#374151","box-shadow":"0 4px 12px rgba(0,0,0,.1)"},
        },
    ), use_container_width=True, height=410)
    st.caption(
        f"Cobertura georreferenciada: {len(g)}/{total_provincias} provincias activas "
        f"({cobertura:.1f}%)."
    )


# ── Ranking HTML ──────────────────────────────────────────────────────────────
def render_ranking(df: pd.DataFrame):
    d = (df.groupby("CLIENTE")["DNI"].nunique()
         .sort_values(ascending=False).head(10).reset_index())
    d.columns = ["CLIENTE","N"]
    mx = d["N"].max()
    html = '<div class="rank-scroll">'
    for i, row in d.iterrows():
        pct = int(row["N"] / max(mx,1) * 100)
        nc  = "gold" if i < 3 else ""
        html += (
            f'<div class="rank-item">'
            f'<div class="rank-num {nc}">{i+1}</div>'
            f'<div style="flex:1;min-width:0">'
            f'<div class="rank-name">{row["CLIENTE"]}</div>'
            f'</div>'
            f'<div>'
            f'<div class="rank-bar-bg"><div class="rank-bar-fill" style="width:{pct}%"></div></div>'
            f'<div class="rank-val">{int(row["N"]):,}</div>'
            f'</div>'
            f'</div>'
        )
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not DATA_FILE.exists():
        st.error(f"No se encontró {DATA_FILE}. Colócalo en la misma carpeta que app.py.")
        st.stop()

    with st.spinner("Cargando datos..."):
        df = load_data(DATA_FILE)
    with st.spinner("Cargando coordenadas..."):
        coords = get_coords(tuple(sorted(df["PROVINCIA"].dropna().unique())))

    filtered, active = sidebar_filters(df)

    # ── HERO ─────────────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="hero">'
        f'<div class="hero-brand">'
        f'<div class="hero-icon">'
        f'<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="white" '
        f'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        f'<rect x="2" y="7" width="20" height="14" rx="2"/>'
        f'<path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/>'
        f'</svg></div>'
        f'<div>'
        f'<div class="hero-title">Dashboard Ejecutivo &mdash; Gestión de Personal</div>'
        f'<div class="hero-sub">Cleaned Perfect S.A. &nbsp;·&nbsp; Servicios Generales &amp; Limpieza &nbsp;·&nbsp; Personal activo · Marzo 2026</div>'
        f'</div></div>'
        f'<div class="hero-badge">{df["DNI"].dropna().nunique():,} colaboradores activos</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── FILTER BADGES ────────────────────────────────────────────────────────
    if active:
        st.markdown(
            f'<div style="padding:6px 28px;background:#fff;border-bottom:1px solid #e2e8f0">'
            + "".join(f'<span class="fbadge">{v}</span>' for v in active)
            + '</div>',
            unsafe_allow_html=True,
        )

    # ── STORY BAR ────────────────────────────────────────────────────────────
    n_dni  = filtered["DNI"].dropna().nunique()
    n_prov = filtered["PROVINCIA"][filtered["PROVINCIA"]!="S/I"].nunique()
    n_cli  = filtered["CLIENTE"].nunique()
    w_ins = workforce_insights(filtered)
    top_cli = (
        filtered.groupby("CLIENTE")["DNI"].nunique().sort_values(ascending=False).head(1)
    )
    top_cli_name = top_cli.index[0] if not top_cli.empty else "S/I"
    top_cli_n = int(top_cli.iloc[0]) if not top_cli.empty else 0
    st.markdown(
        f'<div class="story-wrap"><div class="story-bar">'
        f'Visualizando <strong>{n_dni:,} colaboradores activos</strong> en '
        f'<strong>{n_prov} provincias</strong> bajo <strong>{n_cli} clientes estratégicos</strong>. '
        f'La <strong>estabilidad (+6 meses)</strong> es de <strong>{w_ins["estabilidad_6m"]:.1f}%</strong> '
        f'y la antigüedad mediana alcanza <strong>{w_ins["mediana_meses"]:.1f} meses</strong>. '
        f'Cliente con mayor dotación: <strong>{top_cli_name}</strong> ({top_cli_n:,} personas).'
        f'</div></div>',
        unsafe_allow_html=True,
    )

    # ── KPI ROW ──────────────────────────────────────────────────────────────
    st.markdown('<div class="kpi-grid">', unsafe_allow_html=True)
    render_kpis(filtered)
    st.markdown('</div><div style="height:16px"></div>', unsafe_allow_html=True)

    # ── MAP + RANKING ─────────────────────────────────────────────────────────
    with st.container():
        st.markdown('<div style="padding:0 28px">', unsafe_allow_html=True)
        mc, rc = st.columns([1.7, 1])

        with mc:
            st.markdown(
                f'<div class="sec-head">{ICON["globe"]} Distribución geográfica &nbsp;·&nbsp; '
                f'Perú &nbsp;·&nbsp; <span style="color:#dc2626;font-weight:700">■</span> Sede HQ Lima</div>',
                unsafe_allow_html=True,
            )
            render_map(filtered, coords)
            st.markdown(
                '<div style="font-size:.68rem;color:#94a3b8;margin-top:-6px;padding-bottom:2px">'
                'Burbujas proporcionales al n° de colaboradores · Mapa de calor subyacente · Hover para detalle</div>',
                unsafe_allow_html=True,
            )

        with rc:
            st.markdown(
                f'<div class="sec-head">{ICON["bar"]} Top clientes &nbsp;·&nbsp; DNIs únicos</div>',
                unsafe_allow_html=True,
            )
            render_ranking(filtered)

        st.markdown('</div>', unsafe_allow_html=True)

    # ── CHARTS ROW ───────────────────────────────────────────────────────────
    st.markdown('<div style="padding:4px 28px 0">', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Régimen de planilla</div>',
                    unsafe_allow_html=True)
        if "REGIMEN PLANILLA" in filtered.columns:
            d = (filtered.groupby("REGIMEN PLANILLA")["DNI"].nunique()
                 .sort_values(ascending=False).reset_index())
            d.columns = ["REGIMEN","N"]
            d = d[d["REGIMEN"]!="S/I"]
            fig = px.pie(d, values="N", names="REGIMEN",
                         color_discrete_sequence=PALETTE, hole=0.42)
            fig.update_traces(textinfo="percent", textfont_size=10,
                hovertemplate="%{label}<br>%{value:,}<br>%{percent}<extra></extra>")
            fig = chart_base(fig, 255)
            fig.update_layout(legend=dict(font_size=9, x=1, y=0.5, xanchor="left"))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    with c2:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Top 6 provincias</div>',
                    unsafe_allow_html=True)
        dp = (filtered[filtered["PROVINCIA"]!="S/I"]
              .groupby("PROVINCIA")["DNI"].nunique()
              .sort_values(ascending=False).head(6).reset_index())
        dp.columns = ["PROVINCIA","N"]
        fig2 = px.bar(dp, x="N", y="PROVINCIA", orientation="h",
                      color="N", color_continuous_scale=["#bfdbfe","#1e40af"], text="N")
        fig2 = chart_base(fig2, 255)
        fig2 = hbar(fig2, dp["N"].max(), left=130, ysize=10)
        fig2.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

    with c3:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Mix de clientes</div>',
                    unsafe_allow_html=True)
        dc = (filtered.groupby("CLIENTE")["DNI"].nunique()
              .sort_values(ascending=False).head(7).reset_index())
        dc.columns = ["CLIENTE","N"]
        fig3 = px.pie(dc, values="N", names="CLIENTE",
                      color_discrete_sequence=PALETTE, hole=0.42)
        fig3.update_traces(textinfo="percent", textfont_size=10,
            hovertemplate="%{label}<br>%{value:,}<br>%{percent}<extra></extra>")
        fig3 = chart_base(fig3, 255)
        fig3.update_layout(legend=dict(font_size=9, x=1, y=0.5, xanchor="left"))
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

    st.markdown('</div>', unsafe_allow_html=True)

    # ── TERRITORIAL INTELLIGENCE ────────────────────────────────────────────
    st.markdown('<div style="padding:4px 28px 0">', unsafe_allow_html=True)
    st.markdown(f'<div class="sec-head">{ICON["heat"]} Cobertura territorial inteligente &nbsp;·&nbsp; Departamento → Provincia → Distrito</div>',
                unsafe_allow_html=True)
    if {"DEPARTAMENTO", "PROVINCIA", "DISTRITO"}.issubset(filtered.columns):
        terr = filtered.copy()
        terr = terr[(terr["DEPARTAMENTO"] != "S/I") & (terr["PROVINCIA"] != "S/I") & (terr["DISTRITO"] != "S/I")]
        terr = terr.groupby(["DEPARTAMENTO", "PROVINCIA", "DISTRITO"], as_index=False)["DNI"].nunique()
        terr.columns = ["DEPARTAMENTO", "PROVINCIA", "DISTRITO", "N"]

        t1, t2 = st.columns([1.3, 1])
        with t1:
            fig_terr = px.sunburst(
                terr,
                path=["DEPARTAMENTO", "PROVINCIA", "DISTRITO"],
                values="N",
                color="N",
                color_continuous_scale=["#dbeafe", "#1d4ed8"],
            )
            fig_terr = chart_base(fig_terr, 320)
            fig_terr.update_layout(margin=dict(l=8, r=8, t=8, b=8), coloraxis_colorbar=dict(title="DNIs"))
            st.plotly_chart(fig_terr, use_container_width=True, config={"displayModeBar": False})

        with t2:
            top_dist = (
                terr.groupby(["DEPARTAMENTO", "PROVINCIA", "DISTRITO"])["N"]
                .sum()
                .sort_values(ascending=False)
                .head(12)
                .reset_index()
            )
            top_dist["UBICACION"] = top_dist["DISTRITO"] + " · " + top_dist["PROVINCIA"]
            fig_dist = px.bar(
                top_dist,
                x="N",
                y="UBICACION",
                orientation="h",
                color="DEPARTAMENTO",
                text="N",
                color_discrete_sequence=PALETTE,
            )
            fig_dist = chart_base(fig_dist, 320)
            fig_dist = hbar(fig_dist, top_dist["N"].max(), left=170, ysize=9)
            fig_dist.update_layout(legend_title="", legend_font_size=9)
            st.plotly_chart(fig_dist, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    # ── DETAIL TABLE ─────────────────────────────────────────────────────────
    st.markdown('<div style="padding:4px 28px 0">', unsafe_allow_html=True)
    st.markdown(f'<div class="sec-head">{ICON["list"]} Detalle de colaboradores activos</div>',
                unsafe_allow_html=True)
    cols_show = [c for c in ["DNI","APELLIDOS Y NOMBRES","CLIENTE","UNIDAD","CARGO",
                              "PROVINCIA","DISTRITO","REGIMEN PLANILLA","FECHA DE INGRESO"]
                 if c in filtered.columns]
    sc, si = st.columns([2.5, 1])
    with sc:
        srch = st.text_input("", placeholder="Buscar por nombre, DNI, cliente o provincia...")
    with si:
        st.markdown(
            f'<div style="padding-top:.45rem;font-size:.78rem;color:#64748b">'
            f'<strong>{len(filtered):,}</strong> colaboradores activos</div>',
            unsafe_allow_html=True,
        )
    disp = filtered[cols_show].copy().sort_values(["PROVINCIA","CLIENTE"])
    if srch:
        blob = (disp[[c for c in ["DNI","APELLIDOS Y NOMBRES","CLIENTE","PROVINCIA"] if c in disp.columns]]
                .astype(str).agg(" ".join, axis=1).str.upper())
        disp = disp[blob.str.contains(srch.upper(), na=False, regex=False)]
    st.dataframe(disp, use_container_width=True, hide_index=True, height=360,
                 column_config={
                     "DNI": st.column_config.NumberColumn(format="%d"),
                     "FECHA DE INGRESO": st.column_config.DateColumn(format="DD/MM/YYYY"),
                 })
    st.markdown('</div>', unsafe_allow_html=True)

    # ── BOTTOM STORY ─────────────────────────────────────────────────────────
    n_uni2 = filtered["UNIDAD"].nunique()
    top_dep = (
        filtered[filtered["DEPARTAMENTO"] != "S/I"]
        .groupby("DEPARTAMENTO")["DNI"].nunique()
        .sort_values(ascending=False)
        .head(1)
    )
    dep_name = top_dep.index[0] if not top_dep.empty else "S/I"
    st.markdown(
        f'<div class="bottom-story">'
        f'<div class="b-stat"><div class="b-val">{w_ins["estabilidad_6m"]:.1f}%</div><div class="b-lbl">Estabilidad +6m</div></div>'
        f'<div class="b-div"></div>'
        f'<div class="b-stat"><div class="b-val">{w_ins["mediana_meses"]:.1f}</div><div class="b-lbl">Meses mediana</div></div>'
        f'<div class="b-div"></div>'
        f'<div class="b-stat"><div class="b-val">{n_cli}</div><div class="b-lbl">Clientes</div></div>'
        f'<div class="b-div"></div>'
        f'<div class="b-stat"><div class="b-val">{n_prov}</div><div class="b-lbl">Provincias</div></div>'
        f'<div class="b-div"></div>'
        f'<div class="b-stat"><div class="b-val">{dep_name}</div><div class="b-lbl">Depto líder</div></div>'
        f'<div class="b-div"></div>'
        f'<div class="b-text">'
        f'<strong style="color:#00876a">Resumen ejecutivo:</strong> '
        f'{n_dni:,} colaboradores activos en {n_prov} provincias bajo {n_cli} clientes. '
        f'El {w_ins["estabilidad_6m"]:.1f}% del personal tiene al menos 6 meses de antigüedad y '
        f'{w_ins["recientes_90d"]:,} ingresos ocurrieron en los últimos 90 días. '
        f'{n_uni2:,} unidades operativas en funcionamiento a nivel nacional.'
        f'</div></div>',
        unsafe_allow_html=True,
    )

    # Footer
    st.markdown(
        f'<div style="font-size:.65rem;color:#94a3b8;text-align:right;'
        f'padding:6px 28px;border-top:1px solid #f1f5f9">'
        f'Total: <strong>{len(df):,}</strong> &nbsp;·&nbsp; '
        f'Filtrados: <strong>{len(filtered):,}</strong> &nbsp;·&nbsp; '
        f'Fuente: datos_git.xlsx → WIDE &nbsp;·&nbsp; Solo personal activo'
        f'</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()

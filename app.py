"""
Dashboard Ejecutivo — Gestión de Personal
Cleaned Perfect S.A. · Servicios Generales & Limpieza
Una sola página · scroll vertical · Mapa coroplético Perú
"""
import json
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import pydeck as pdk
import requests
import streamlit as st

st.set_page_config(
    page_title="Dashboard Ejecutivo · Gestión de Personal",
    layout="wide",
    initial_sidebar_state="collapsed",
)

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

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"],.stApp{font-family:'Inter',system-ui,sans-serif!important;color:#0f172a!important}
.stApp{background:#f0f2f7}.block-container{padding:0 0 0 1.5rem!important;max-width:100%!important}
section[data-testid="stSidebar"]{background:#fff;border-right:1px solid #e2e8f0}
section[data-testid="stSidebar"]>div{padding:1.2rem 1rem}
.sidebar-label{font-size:.67rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#0f172a!important;margin:.85rem 0 .3rem;display:block}
div[data-baseweb="select"]>div{background:#f8fafc!important;border:1px solid #cbd5e1!important;border-radius:8px!important;color:#0f172a!important}
section[data-testid="stSidebar"] *{color:#0f172a!important}
input,textarea,[data-baseweb="input"] input{color:#0f172a!important}
.stButton>button{background:#00a885!important;color:#fff!important;border:none!important;border-radius:8px!important;font-weight:600!important;font-size:.8rem!important;width:100%;margin-top:.5rem}
.hero{background:#fff;padding:14px 28px 12px;border-bottom:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px}
.hero-title{font-size:1.2rem;font-weight:700;color:#0f172a}.hero-sub{font-size:.75rem;color:#64748b}.hero-badge{background:#e6f7f3;color:#00876a;border:1px solid rgba(0,168,133,.22);border-radius:20px;padding:5px 16px;font-size:.78rem;font-weight:700}
.story-wrap{padding:10px 28px;background:#fff;border-bottom:1px solid #f1f5f9}.story-bar{background:linear-gradient(90deg,#e6f7f3,#f0f2f7);border-left:3px solid #00a885;border-radius:0 8px 8px 0;padding:8px 18px;font-size:.79rem;color:#374151;line-height:1.6}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;padding:16px 28px 0}.kpi-card{background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:15px 18px 12px;box-shadow:0 1px 3px rgba(0,0,0,.06);position:relative}
.kpi-bar{position:absolute;top:0;left:0;right:0;height:3px;border-radius:14px 14px 0 0}.kpi-label{font-size:.66rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#94a3b8}.kpi-value{font-size:1.7rem;font-weight:700;color:#0f172a}
.sec-head{font-size:.67rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94a3b8;padding-bottom:6px;border-bottom:1.5px solid #e2e8f0;margin:0 0 10px;display:flex;align-items:center;gap:6px}
div[data-testid="stPlotlyChart"]{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:10px 12px;box-shadow:0 1px 3px rgba(0,0,0,.05);margin-bottom:14px}
div[data-testid="stDeckGlJsonChart"]{border-radius:12px;overflow:hidden;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,.05)}
#MainMenu,footer,header{visibility:hidden}
button[data-testid="collapsedControl"]{display:flex!important;visibility:visible!important;opacity:1!important}
.sidebar-toggle-floating{
  position:fixed;top:14px;left:12px;z-index:9999;
  background:#00a885;color:#fff;border:none;border-radius:999px;
  padding:8px 12px;font-size:.78rem;font-weight:700;cursor:pointer;
  box-shadow:0 3px 10px rgba(0,0,0,.16);transition:all .2s ease
}
.sidebar-toggle-floating:hover{filter:brightness(.95);transform:translateY(-1px)}
</style>
""", unsafe_allow_html=True)

DATA_FILE = Path("datos_git.xlsx")
GEOJSON_LOCAL = Path("peru_departamentos.geojson")
PROVINCE_COORDS_FILE = Path("province_coords.csv")

PROVINCE_FIX = {
    "CUZCO": "CUSCO", "CUSCO": "CUSCO", "LIMA METROPOLITANA": "LIMA",
    "TUMBES": "TUMBES", "JUNIN": "JUNIN", "HUANUCO": "HUANUCO",
}

PALETTE = ["#00a885", "#2563eb", "#7c3aed", "#d97706", "#16a34a",
           "#0891b2", "#dc2626", "#9333ea", "#ea580c", "#475569"]


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
    if ((r.startswith("FT") and ">= 8H" in r) or ("PERFECT" in r and ">= 8H" in r) or r == "FULL TIME"):
        return "Full Time ≥8H"
    if r.startswith("FT"):
        return "Full Time <8H"
    if r.startswith("PT") or "PART" in r:
        return "Part Time"
    return "Otros"


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WIDE")
    df.columns = [norm(c) for c in df.columns]
    for col in ["PROVINCIA", "DISTRITO", "CLIENTE", "UNIDAD", "CARGO", "PLANILLA", "REGIMEN PLANILLA", "SUPERVISOR", "TAREADOR"]:
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
    df["ANTIGUEDAD_LABEL"] = df["antiguedad"].fillna(0).apply(lambda x: f"{int(x*12)}m" if x < 1 else f"{x:.1f}a")
    df["UBICACION"] = (
        df["DISTRITO"].fillna("S/I") + ", " + df["PROVINCIA"].fillna("S/I") + ", Perú"
    )
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
def load_geojson() -> dict | None:
    url = "https://raw.githubusercontent.com/juancgalvez/peru-geojson/master/geojson/peru_departamentos.geojson"
    try:
        data = requests.get(url, timeout=10)
        data.raise_for_status()
        geo = data.json()
        GEOJSON_LOCAL.write_text(json.dumps(geo, ensure_ascii=False), encoding="utf-8")
        return geo
    except Exception:
        if GEOJSON_LOCAL.exists():
            return json.loads(GEOJSON_LOCAL.read_text(encoding="utf-8"))
    return None


def chart_base(fig, height=270):
    fig.update_layout(
        template="plotly_white", height=height, margin=dict(l=8, r=16, t=6, b=6),
        paper_bgcolor="white", plot_bgcolor="white",
        font_family="Inter, system-ui, sans-serif", font_color="#0f172a"
    )
    fig.update_xaxes(tickfont_color="#0f172a", title_font_color="#0f172a")
    fig.update_yaxes(tickfont_color="#0f172a", title_font_color="#0f172a")
    fig.update_traces(
        textfont=dict(color="#0f172a"),
        insidetextfont=dict(color="#0f172a"),
        outsidetextfont=dict(color="#0f172a"),
    )
    fig.update_layout(legend=dict(font=dict(color="#0f172a")))
    return fig


def sync(key, valid):
    st.session_state[key] = [v for v in st.session_state.get(key, []) if v in valid]


def clear_filters():
    for k in ["f_cli", "f_uni", "f_car", "f_reg", "f_prov"]:
        st.session_state[k] = []


def sidebar_filters(df: pd.DataFrame):
    with st.sidebar:
        st.markdown(f'{ICON["filter"]}<span style="font-size:.95rem;font-weight:700;color:#0f172a"> Filtros</span>', unsafe_allow_html=True)
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

        st.markdown('<span class="sidebar-label">Régimen</span>', unsafe_allow_html=True)
        reg_o = sorted(base["REGIMEN PLANILLA"].dropna().unique())
        sync("f_reg", reg_o)
        reg = st.multiselect("", reg_o, key="f_reg", placeholder="Todos los regímenes")
        base = base[base["REGIMEN PLANILLA"].isin(reg)] if reg else base

        st.markdown('<span class="sidebar-label">Provincia</span>', unsafe_allow_html=True)
        prov_o = sorted(base["PROVINCIA"].dropna().unique())
        sync("f_prov", prov_o)
        prov = st.multiselect("", prov_o, key="f_prov", placeholder="Todas las provincias")

        st.button("Limpiar filtros", on_click=clear_filters, use_container_width=True)

    f = df.copy()
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
    return f


def generate_story(df: pd.DataFrame, full_df: pd.DataFrame) -> str:
    total_dni = df["DNI"].dropna().nunique()
    total_full = max(full_df["DNI"].dropna().nunique(), 1)
    pct_del_total = total_dni / total_full * 100
    n_prov = df["PROVINCIA"][df["PROVINCIA"] != "S/I"].nunique()
    n_cli = df["CLIENTE"].nunique()
    n_uni = df["UNIDAD"].nunique()

    if total_dni == 0:
        return "Sin registros para los filtros seleccionados."

    top_cli_series = df.groupby("CLIENTE")["DNI"].nunique().sort_values(ascending=False)
    top_prov_series = df[df["PROVINCIA"] != "S/I"].groupby("PROVINCIA")["DNI"].nunique().sort_values(ascending=False)
    cli_geo_series = df[df["PROVINCIA"] != "S/I"].groupby("CLIENTE")["PROVINCIA"].nunique().sort_values(ascending=False)

    top_cli = top_cli_series.index[0] if not top_cli_series.empty else "S/I"
    top_prov = top_prov_series.index[0] if not top_prov_series.empty else "S/I"
    top_prov_n = int(top_prov_series.iloc[0]) if not top_prov_series.empty else 0
    partes = [
        (f"<strong>{total_dni:,} colaboradores activos</strong>" if pct_del_total >= 99 else f"Filtro activo: <strong>{total_dni:,}</strong> ({pct_del_total:.1f}% del total)."),
        f"Operación en <strong>{n_prov} provincias</strong> bajo <strong>{n_cli} clientes</strong> y <strong>{n_uni:,} unidades</strong>.",
        f"Cliente líder: <strong>{top_cli}</strong> con <strong>{int(top_cli_series.iloc[0]) if not top_cli_series.empty else 0}</strong> colaboradores.",
        f"<strong>{top_prov}</strong> concentra <strong>{(top_prov_n / max(total_dni,1) * 100):.1f}%</strong> del headcount.",
    ]

    if n_cli > 1 and not cli_geo_series.empty:
        partes.append(f"Mayor cobertura geográfica: <strong>{cli_geo_series.index[0]}</strong> con <strong>{int(cli_geo_series.iloc[0])} provincias</strong>.")

    return " &nbsp;·&nbsp; ".join(partes)


def render_kpis(df: pd.DataFrame):
    total = df["DNI"].dropna().nunique()

    cards = [
        ("#00a885", "Colaboradores", f"{total:,}"),
        ("#2563eb", "Clientes", f"{df['CLIENTE'].nunique():,}"),
        ("#7c3aed", "Unidades", f"{df['UNIDAD'].nunique():,}"),
        ("#d97706", "Provincias", f"{df[df['PROVINCIA']!='S/I']['PROVINCIA'].nunique():,}"),
    ]

    for col, (color, label, value) in zip(st.columns(4), cards):
        with col:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-bar" style="background:{color}"></div>'
                f'<div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div>',
                unsafe_allow_html=True,
            )


def render_map(df: pd.DataFrame, geojson: dict | None):
    if not geojson:
        st.warning("No se pudo cargar el GeoJSON de Perú.")
        return

    d = df[df["PROVINCIA"] != "S/I"].groupby("PROVINCIA", as_index=False).agg(
        count=("DNI", "nunique"), clientes=("CLIENTE", "nunique"), unidades=("UNIDAD", "nunique")
    )
    counts = {r["PROVINCIA"]: int(r["count"]) for _, r in d.iterrows()}
    max_count = max(counts.values()) if counts else 1

    features = []
    for feat in geojson.get("features", []):
        props = feat.get("properties", {})
        name = norm(props.get("NOMBDEP") or props.get("NOMBPROV") or props.get("name") or "S/I")
        n = counts.get(name, 0)
        color_ratio = (np.log1p(n) / np.log1p(max_count)) if max_count > 0 else 0
        start = np.array([15, 23, 42, 180])
        mid = np.array([0, 168, 133, 220])
        end = np.array([52, 211, 153, 240])
        if color_ratio <= 0.5:
            t = color_ratio / 0.5
            fill = (start + (mid - start) * t).astype(int).tolist()
        else:
            t = (color_ratio - 0.5) / 0.5
            fill = (mid + (end - mid) * t).astype(int).tolist()
        new_props = {
            "DEPT": name,
            "PROVINCIA": name,
            "count": n,
            "clientes": int(d.loc[d["PROVINCIA"] == name, "clientes"].max()) if name in counts else 0,
            "unidades": int(d.loc[d["PROVINCIA"] == name, "unidades"].max()) if name in counts else 0,
            "fill_color": fill,
            "label": f"{name} ({n})",
            "layer_type": "department",
        }
        feat2 = {"type": "Feature", "geometry": feat.get("geometry"), "properties": new_props}
        features.append(feat2)

    province_counts = (
        df[df["PROVINCIA"] != "S/I"]
        .groupby(["PROVINCIA", "DISTRITO"], as_index=False)["DNI"]
        .nunique()
        .rename(columns={"DNI": "colaboradores"})
    )
    top_district = (
        province_counts.sort_values(["PROVINCIA", "colaboradores"], ascending=[True, False])
        .drop_duplicates("PROVINCIA")
        .rename(columns={"DISTRITO": "distrito_principal"})
    )
    top_client = (
        df[df["PROVINCIA"] != "S/I"]
        .groupby(["PROVINCIA", "CLIENTE"], as_index=False)["DNI"]
        .nunique()
        .rename(columns={"DNI": "n_cliente"})
        .sort_values(["PROVINCIA", "n_cliente"], ascending=[True, False])
        .drop_duplicates("PROVINCIA")
        .rename(columns={"CLIENTE": "cliente_top"})
    )
    points = (
        d[["PROVINCIA", "count"]]
        .merge(top_district[["PROVINCIA", "distrito_principal"]], on="PROVINCIA", how="left")
        .merge(top_client[["PROVINCIA", "cliente_top", "n_cliente"]], on="PROVINCIA", how="left")
    )
    coords = load_province_coords()
    points = points.merge(coords, on="PROVINCIA", how="left").dropna(subset=["lat", "lon"])
    points["layer_type"] = "province_point"
    points["radius_pixels"] = np.where(points["PROVINCIA"] == "LIMA", 4.5, 3.5)
    points["label"] = points["PROVINCIA"] + " (" + points["count"].astype(int).astype(str) + ")"
    points["cliente_top_det"] = (
        points["cliente_top"].fillna("S/I") + " (" + points["n_cliente"].fillna(0).astype(int).astype(str) + ")"
    )

    label_data = points[points["count"] >= 50].copy()
    geo_data = {"type": "FeatureCollection", "features": features}
    geo_layer = pdk.Layer(
        "GeoJsonLayer",
        id="geo_dept",
        data=geo_data,
        get_fill_color="properties.fill_color",
        get_line_color=[100, 200, 160, 180],
        line_width_min_pixels=1.2,
        pickable=True,
        stroked=True,
        filled=True,
        auto_highlight=True,
        update_triggers={"get_fill_color": [max_count]},
    )
    point_layer = pdk.Layer(
        "ScatterplotLayer",
        id="province_points",
        data=points.to_dict("records"),
        get_position="[lon, lat]",
        get_radius="radius_pixels",
        radius_units="pixels",
        radius_min_pixels=3,
        radius_max_pixels=8,
        radius_scale=1,
        get_fill_color=[251, 191, 36, 220],
        get_line_color=[255, 230, 160, 230],
        line_width_min_pixels=0.8,
        pickable=True,
    )
    text_layer = pdk.Layer(
        "TextLayer",
        id="dept_labels",
        data=label_data.to_dict("records"),
        get_position="[lon, lat]",
        get_text="label",
        get_color=[255, 255, 255, 220],
        get_size=12,
        get_pixel_offset=[0, 0],
        get_text_anchor="'middle'",
        get_alignment_baseline="'center'",
        get_font_weight=700,
        background=False,
        pickable=False,
    )

    def on_select(selection):
        objects = selection.get("selection", {}).get("objects", {})
        selected_obj = []
        for layer_key, items in objects.items():
            if "geo" in layer_key.lower() and items:
                selected_obj = items
                break
        if not selected_obj:
            return
        dept = selected_obj[0].get("properties", {}).get("DEPT")
        if dept:
            st.session_state["f_prov"] = [dept]
            st.rerun()

    st.pydeck_chart(
        pdk.Deck(
            layers=[geo_layer, point_layer, text_layer],
            initial_view_state=pdk.ViewState(
                latitude=-9.5, longitude=-74.8, zoom=4.9, min_zoom=4.0, max_zoom=9, pitch=0
            ),
            map_style="https://basemaps.cartocdn.com/gl/dark-matter-gl-style/style.json",
            tooltip={
                "html": (
                    "<b>{PROVINCIA}</b><br/>"
                    "Colaboradores: {count}<br/>"
                    "Clientes: {clientes}<br/>"
                    "Unidades: {unidades}<br/>"
                    "Distrito principal: {distrito_principal}<br/>"
                    "Cliente líder local: {cliente_top_det}"
                )
            },
        ),
        use_container_width=True,
        height=430,
        on_select=on_select,
        selection_mode="single-object",
    )


def main():
    if not DATA_FILE.exists():
        st.error("No se encontró datos_git.xlsx")
        st.stop()

    df = load_data(DATA_FILE)
    geo = load_geojson()
    filtered = sidebar_filters(df)

    st.markdown("""
    <button
      class="sidebar-toggle-floating"
      title="Mostrar u ocultar filtros"
      aria-label="Mostrar u ocultar filtros"
      onclick="
        const hostDoc = window.parent?.document || document;
        const btn = hostDoc.querySelector('button[data-testid=\\"collapsedControl\\"]');
        if (btn) { btn.click(); }
      "
    >☰ Filtros</button>
    """, unsafe_allow_html=True)

    st.markdown(
        f'<div class="hero"><div><div class="hero-title">Dashboard Ejecutivo — Gestión de Personal</div>'
        f'<div class="hero-sub">Cleaned Perfect S.A. · Servicios Generales y Limpieza · {pd.Timestamp.today().strftime("%d/%m/%Y")}</div></div>'
        f'<div class="hero-badge">{df["DNI"].dropna().nunique():,} colaboradores activos</div></div>',
        unsafe_allow_html=True,
    )

    st.markdown(f'<div class="story-wrap"><div class="story-bar">{generate_story(filtered, df)}</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="kpi-grid">', unsafe_allow_html=True)
    render_kpis(filtered)
    st.markdown('</div><div style="height:14px"></div>', unsafe_allow_html=True)

    st.markdown('<div style="padding:0 28px">', unsafe_allow_html=True)
    c_map, c_rank = st.columns([1.7, 1])
    with c_map:
        st.markdown(f'<div class="sec-head">{ICON["globe"]} Mapa nacional coroplético</div>', unsafe_allow_html=True)
        render_map(filtered, geo)
    with c_rank:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Top clientes</div>', unsafe_allow_html=True)
        rank = filtered.groupby("CLIENTE")["DNI"].nunique().sort_values(ascending=False).head(10).reset_index(name="N")
        fig_rank = px.bar(rank, x="N", y="CLIENTE", orientation="h", color="N", color_continuous_scale=["#dcfce7", "#166534"], text="N")
        fig_rank.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
        st.plotly_chart(chart_base(fig_rank, 430), use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div style="padding:4px 28px 0">', unsafe_allow_html=True)
    a, b, c = st.columns(3)
    with a:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Régimen simplificado</div>', unsafe_allow_html=True)
        reg = filtered.groupby("REGIMEN_SIMPLE")["DNI"].nunique().reset_index(name="N")
        fig_reg = px.pie(reg, values="N", names="REGIMEN_SIMPLE", color="REGIMEN_SIMPLE", color_discrete_sequence=PALETTE, hole=0.45)
        st.plotly_chart(chart_base(fig_reg, 260), use_container_width=True, config={"displayModeBar": False})
    with b:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Top 6 provincias</div>', unsafe_allow_html=True)
        prov = filtered[filtered["PROVINCIA"] != "S/I"].groupby("PROVINCIA")["DNI"].nunique().sort_values(ascending=False).head(6).reset_index(name="N")
        fig_prov = px.bar(prov, x="N", y="PROVINCIA", orientation="h", text="N", color="N", color_continuous_scale=["#bfdbfe", "#1e40af"])
        fig_prov.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
        st.plotly_chart(chart_base(fig_prov, 260), use_container_width=True, config={"displayModeBar": False})
    with c:
        st.markdown(f'<div class="sec-head">{ICON["bar"]} Mix clientes</div>', unsafe_allow_html=True)
        mix = filtered.groupby("CLIENTE")["DNI"].nunique().sort_values(ascending=False).head(7).reset_index(name="N")
        fig_mix = px.pie(mix, values="N", names="CLIENTE", color_discrete_sequence=PALETTE, hole=0.45)
        st.plotly_chart(chart_base(fig_mix, 260), use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div style="padding:4px 28px 0">', unsafe_allow_html=True)
    st.markdown(f'<div class="sec-head">{ICON["bar"]} Top 10 distritos de operación</div>', unsafe_allow_html=True)
    dist = (filtered[filtered["DISTRITO"] != "S/I"].groupby(["DISTRITO", "PROVINCIA"])["DNI"].nunique().sort_values(ascending=False).head(10).reset_index(name="N"))
    dist["UBI"] = dist["DISTRITO"] + " / " + dist["PROVINCIA"]
    fig_dist = px.bar(dist, x="N", y="UBI", orientation="h", text="N", color="PROVINCIA", color_discrete_sequence=PALETTE)
    fig_dist.update_layout(yaxis=dict(autorange="reversed"))
    st.plotly_chart(chart_base(fig_dist, 330), use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div style="padding:4px 28px 0">', unsafe_allow_html=True)
    st.markdown(f'<div class="sec-head">{ICON["list"]} Detalle de colaboradores activos</div>', unsafe_allow_html=True)
    cols_show = ["DNI", "APELLIDOS Y NOMBRES", "CARGO", "CLIENTE", "UNIDAD", "SUPERVISOR", "PROVINCIA", "DISTRITO", "REGIMEN_SIMPLE", "FECHA DE INGRESO", "ANTIGUEDAD_LABEL"]
    disp = filtered[cols_show].copy()
    disp = disp.rename(columns={"APELLIDOS Y NOMBRES": "Nombre", "DISTRITO": "Distrito", "ANTIGUEDAD_LABEL": "Antigüedad", "REGIMEN_SIMPLE": "Régimen"})
    srch = st.text_input("", placeholder="Buscar por nombre, DNI, cliente o provincia...")
    if srch:
        blob = disp.astype(str).agg(" ".join, axis=1).str.upper()
        disp = disp[blob.str.contains(srch.upper(), regex=False, na=False)]
    st.dataframe(disp, hide_index=True, use_container_width=True, height=360, column_config={"FECHA DE INGRESO": st.column_config.DateColumn(format="DD/MM/YYYY")})
    st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()

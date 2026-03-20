"""
Geocodifica las direcciones únicas de datos_git.xlsx
y genera sede_coords.csv con lat/lon.

Estrategia:
1. Agrupa por DIRECCION + DISTRITO + PROVINCIA + DEPARTAMENTO (sede única)
2. Intenta geocodificar con Nominatim usando la dirección completa
3. Si falla, usa DISTRITO + PROVINCIA + DEPARTAMENTO como fallback
4. Si falla, usa las coordenadas de province_coords.csv como último recurso
5. Guarda resultados incrementalmente (no re-geocodifica lo ya resuelto)
"""
import time
import unicodedata
from pathlib import Path

import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

DATA_FILE = Path("datos_git.xlsx")
OUTPUT_FILE = Path("sede_coords.csv")
PROVINCE_FILE = Path("province_coords.csv")


def norm(v) -> str:
    if pd.isna(v):
        return "S/I"
    s = str(v).strip().upper()
    return " ".join(
        "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").split()
    ) or "S/I"


def load_existing():
    if OUTPUT_FILE.exists():
        df = pd.read_csv(OUTPUT_FILE)
        return set(df["DIRECCION"].astype(str).values)
    return set()


def load_province_fallback():
    if PROVINCE_FILE.exists():
        p = pd.read_csv(PROVINCE_FILE)
        p.columns = [c.strip().upper() for c in p.columns]
        if "LATITUDE" in p.columns:
            p = p.rename(columns={"LATITUDE": "LAT", "LONGITUDE": "LON"})
        return dict(zip(p["PROVINCIA"].str.strip().str.upper(), zip(p["LAT"], p["LON"])))
    return {}


def geocode_address(geolocator, address, distrito, provincia, departamento, prov_fallback):
    """Try geocoding with multiple strategies."""
    queries = [
        f"{address}, Peru",
        f"{address}, {provincia}, Peru",
        f"{distrito}, {provincia}, {departamento}, Peru",
        f"{distrito}, {provincia}, Peru",
        f"{provincia}, {departamento}, Peru",
    ]

    for query in queries:
        if "S/I" in query:
            continue
        try:
            location = geolocator.geocode(query, timeout=10, country_codes="pe")
            if location:
                # Basic sanity check: must be in Peru (lat: -18.5 to -0.0, lon: -81.5 to -68.5)
                if -18.5 <= location.latitude <= 0.0 and -81.5 <= location.longitude <= -68.5:
                    return location.latitude, location.longitude
        except (GeocoderTimedOut, GeocoderServiceError):
            time.sleep(2)
        except Exception:
            pass
        time.sleep(1.1)

    # Fallback to province coordinates
    prov_key = norm(provincia)
    if prov_key in prov_fallback:
        lat, lon = prov_fallback[prov_key]
        return lat, lon

    return None, None


def main():
    print("Loading data...")
    df = pd.read_excel(DATA_FILE, sheet_name="WIDE")
    df.columns = [norm(c) for c in df.columns]

    # Filter active employees only
    df = df[df["FECHA DE CESE"].isna()].copy()

    for col in ["DIRECCION", "DISTRITO", "PROVINCIA", "DEPARTAMENTO"]:
        if col in df.columns:
            df[col] = df[col].map(norm)

    # Get unique sedes (address + context)
    sedes = (
        df.groupby(["DIRECCION", "DISTRITO", "PROVINCIA", "DEPARTAMENTO"], as_index=False)
        .agg(colaboradores=("DNI", "nunique"))
        .sort_values("colaboradores", ascending=False)
    )
    sedes = sedes[sedes["DIRECCION"] != "S/I"]

    print(f"Total unique sedes: {len(sedes)}")

    # Load existing results
    existing = load_existing()
    prov_fallback = load_province_fallback()

    # Filter out already geocoded
    to_geocode = sedes[~sedes["DIRECCION"].isin(existing)]
    print(f"Already geocoded: {len(existing)}, remaining: {len(to_geocode)}")

    if len(to_geocode) == 0:
        print("All addresses already geocoded!")
        return

    geolocator = Nominatim(
        user_agent="cleaned_perfect_dashboard_v1",
        timeout=10
    )

    results = []
    # Load existing results to append
    if OUTPUT_FILE.exists():
        results = pd.read_csv(OUTPUT_FILE).to_dict("records")

    total = len(to_geocode)
    for i, (_, row) in enumerate(to_geocode.iterrows()):
        addr = row["DIRECCION"]
        dist = row["DISTRITO"]
        prov = row["PROVINCIA"]
        dept = row["DEPARTAMENTO"]
        n = row["colaboradores"]

        print(f"[{i+1}/{total}] ({n} emp) {addr[:60]}...", end=" ")

        lat, lon = geocode_address(geolocator, addr, dist, prov, dept, prov_fallback)

        if lat is not None:
            print(f"OK ({lat:.4f}, {lon:.4f})")
            results.append({
                "DIRECCION": addr,
                "DISTRITO": dist,
                "PROVINCIA": prov,
                "DEPARTAMENTO": dept,
                "colaboradores": n,
                "lat": lat,
                "lon": lon,
            })
        else:
            print("FAILED")
            results.append({
                "DIRECCION": addr,
                "DISTRITO": dist,
                "PROVINCIA": prov,
                "DEPARTAMENTO": dept,
                "colaboradores": n,
                "lat": None,
                "lon": None,
            })

        # Save every 20 results
        if (i + 1) % 20 == 0:
            pd.DataFrame(results).to_csv(OUTPUT_FILE, index=False)
            print(f"  [saved {len(results)} records]")

    # Final save
    pd.DataFrame(results).to_csv(OUTPUT_FILE, index=False)
    final = pd.DataFrame(results)
    found = final["lat"].notna().sum()
    print(f"\nDone! {found}/{len(final)} addresses geocoded successfully.")
    print(f"Saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

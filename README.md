# Dashboard Ejecutivo — Gestión de Personal
**Cleaned Perfect S.A. · Servicios Generales & Limpieza**

Aplicación Streamlit de nivel gerencial para visualizar la fuerza laboral activa.

## Estructura del repositorio

```
├── app.py                  ← Aplicación principal
├── datos_git.xlsx          ← Fuente de datos (hoja: WIDE)
├── province_coords.csv     ← Coordenadas de provincias (lat, lon)
├── requirements.txt        ← Dependencias Python
└── README.md
```

## Funcionalidades

- **KPIs ejecutivos** — Colaboradores activos, clientes, unidades, cobertura y retención.
- **Filtros en cascada** — Cliente → Unidad → Cargo → Régimen → Provincia (sidebar).
- **Mapa geográfico Folium** — Burbujas proporcionales al volumen. Sede HQ marcada. Heatmap de calor por intensidad.
- **Análisis** — Top clientes, distribución por régimen, cargos y unidades.
- **Mapa de calor** — Provincia × Régimen (heatmap interactivo).
- **Detalle** — Tabla filtrable con búsqueda libre.

> Solo muestra **personal activo** (sin fecha de cese).

## Despliegue en Streamlit Cloud

1. Sube este repositorio a GitHub.
2. En [share.streamlit.io](https://share.streamlit.io) → **New app** → selecciona el repo.
3. Main file: `app.py`.
4. Listo — se actualiza automáticamente con cada `git push`.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Fuente de datos

| Campo | Descripción |
|---|---|
| `WIDE` | Hoja principal del Excel |
| `DNI` | Identificador único de colaborador |
| `CLIENTE` | Empresa cliente |
| `UNIDAD` | Unidad operativa dentro del cliente |
| `CARGO` | Puesto del colaborador |
| `REGIMEN PLANILLA` | Tipo de jornada |
| `PROVINCIA` | Provincia de operación |
| `FECHA DE INGRESO` | Fecha de alta en planilla |

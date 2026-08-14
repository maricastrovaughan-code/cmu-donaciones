
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="CMU | Tablero de Donaciones",
    page_icon="🧭",
    layout="wide"
)

st.title("🧭 Centro de Mando Unificado — Donaciones")
st.caption("Seguimiento de necesidades, acopio, inventario, despachos, recepciones e incidencias.")

REQUIRED_SHEETS = [
    "Catalogo_Productos",
    "Necesidades",
    "Entradas",
    "Despachos",
    "Recepciones",
    "Incidencias",
]

def read_sheet(file, sheet_name):
    df = pd.read_excel(file, sheet_name=sheet_name, header=2)
    df = df.dropna(how="all")
    return df

def clean_numeric(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df

def clean_date(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def pct_label(v):
    if pd.isna(v):
        return "0%"
    return f"{v:.1f}%"

uploaded = st.sidebar.file_uploader(
    "Cargar archivo Excel del CMU",
    type=["xlsx"],
    help="Use la plantilla Plantilla_CMU_Donaciones.xlsx"
)

st.sidebar.markdown("### Filtros")
selected_categories = st.sidebar.multiselect(
    "Categoría",
    ["Alimentos", "Medicamentos", "Aseo"],
    default=["Alimentos", "Medicamentos", "Aseo"]
)

if uploaded is None:
    st.info("Cargue el archivo Excel del CMU para visualizar el tablero.")
    st.stop()

try:
    xl = pd.ExcelFile(uploaded)
    missing = [s for s in REQUIRED_SHEETS if s not in xl.sheet_names]
    if missing:
        st.error("Faltan hojas requeridas: " + ", ".join(missing))
        st.stop()

    necesidades = read_sheet(uploaded, "Necesidades")
    entradas = read_sheet(uploaded, "Entradas")
    despachos = read_sheet(uploaded, "Despachos")
    recepciones = read_sheet(uploaded, "Recepciones")
    incidencias = read_sheet(uploaded, "Incidencias")
    catalogo = read_sheet(uploaded, "Catalogo_Productos")

except Exception as e:
    st.error(f"No fue posible leer el archivo: {e}")
    st.stop()

# Limpieza
necesidades = clean_numeric(necesidades, ["Cantidad_requerida"])
entradas = clean_numeric(entradas, ["Cantidad_recibida", "Cantidad_aprobada", "Cantidad_rechazada"])
despachos = clean_numeric(despachos, ["Cantidad_despachada"])
recepciones = clean_numeric(recepciones, ["Cantidad_recibida_conforme"])

necesidades = clean_date(necesidades, ["Fecha_reporte", "Fecha_limite"])
entradas = clean_date(entradas, ["Fecha_hora", "Fecha_vencimiento"])
despachos = clean_date(despachos, ["Fecha_hora_salida", "ETA"])
recepciones = clean_date(recepciones, ["Fecha_hora_recepcion"])
incidencias = clean_date(incidencias, ["Fecha_hora", "Fecha_cierre"])

for df in [necesidades, entradas, despachos]:
    if "Categoria" in df.columns:
        df["Categoria"] = df["Categoria"].astype(str)

# Filtro por categoría
necesidades_f = necesidades[necesidades["Categoria"].isin(selected_categories)].copy() if "Categoria" in necesidades else necesidades.copy()
entradas_f = entradas[entradas["Categoria"].isin(selected_categories)].copy() if "Categoria" in entradas else entradas.copy()
despachos_f = despachos[despachos["Categoria"].isin(selected_categories)].copy() if "Categoria" in despachos else despachos.copy()

# Solo necesidades activas para cobertura
activas = necesidades_f[
    necesidades_f["Estado"].astype(str).str.lower().eq("activa")
].copy() if "Estado" in necesidades_f else necesidades_f.copy()

# Sumas por necesidad
approved = (
    entradas_f.groupby("ID_Necesidad_asignada", dropna=False)["Cantidad_aprobada"]
    .sum()
    .rename("Cantidad_acopiada")
) if "ID_Necesidad_asignada" in entradas_f else pd.Series(dtype=float)

sent = (
    despachos_f.groupby("ID_Necesidad", dropna=False)["Cantidad_despachada"]
    .sum()
    .rename("Cantidad_despachada")
) if "ID_Necesidad" in despachos_f else pd.Series(dtype=float)

# Recepciones se vinculan por despacho -> necesidad
if not recepciones.empty and not despachos_f.empty and "ID_Despacho" in recepciones and "ID_Despacho" in despachos_f:
    rec_map = recepciones.merge(
        despachos_f[["ID_Despacho", "ID_Necesidad"]],
        on="ID_Despacho",
        how="left"
    )
    received = (
        rec_map.groupby("ID_Necesidad", dropna=False)["Cantidad_recibida_conforme"]
        .sum()
        .rename("Cantidad_recibida_destino")
    )
else:
    received = pd.Series(dtype=float)

resumen = activas.copy()
if not resumen.empty:
    resumen = resumen.set_index("ID_Necesidad")
    resumen = resumen.join(approved, how="left").join(sent, how="left").join(received, how="left")
    for c in ["Cantidad_acopiada", "Cantidad_despachada", "Cantidad_recibida_destino"]:
        resumen[c] = resumen[c].fillna(0)

    resumen["Inventario_disponible"] = (resumen["Cantidad_acopiada"] - resumen["Cantidad_despachada"]).clip(lower=0)
    resumen["Faltante_acopio"] = (resumen["Cantidad_requerida"] - resumen["Cantidad_acopiada"]).clip(lower=0)
    resumen["Cobertura_acopio_pct"] = np.where(
        resumen["Cantidad_requerida"] > 0,
        resumen["Cantidad_acopiada"] / resumen["Cantidad_requerida"] * 100,
        0
    )
    resumen["Cobertura_destino_pct"] = np.where(
        resumen["Cantidad_requerida"] > 0,
        resumen["Cantidad_recibida_destino"] / resumen["Cantidad_requerida"] * 100,
        0
    )
    resumen["Semaforo"] = pd.cut(
        resumen["Cobertura_acopio_pct"],
        bins=[-np.inf, 50, 80, 100, np.inf],
        labels=["CRÍTICO", "ALERTA", "PRÓXIMO A META", "CUBIERTO"],
        right=False
    )
    resumen = resumen.reset_index()
else:
    resumen = pd.DataFrame()

# KPIs
c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("Necesidades activas", int(len(resumen)))

cubiertas = int((resumen["Cobertura_acopio_pct"] >= 100).sum()) if not resumen.empty else 0
c2.metric("Necesidades cubiertas", cubiertas)

criticas = int((resumen["Cobertura_acopio_pct"] < 50).sum()) if not resumen.empty else 0
c3.metric("Necesidades críticas", criticas)

en_transito = 0
if not despachos_f.empty and "Estado_despacho" in despachos_f:
    en_transito = int(despachos_f["Estado_despacho"].astype(str).str.lower().eq("en tránsito").sum())
c4.metric("Despachos en tránsito", en_transito)

inc_abiertas = 0
if not incidencias.empty and "Estado" in incidencias:
    inc_abiertas = int(incidencias["Estado"].astype(str).str.lower().isin(["abierta", "en gestión"]).sum())
c5.metric("Incidencias abiertas", inc_abiertas)

st.divider()

# Resumen por categoría
left, right = st.columns([1.1, 1])

with left:
    st.subheader("Cobertura promedio por categoría")
    if not resumen.empty:
        cat = (
            resumen.groupby("Categoria", as_index=False)["Cobertura_acopio_pct"]
            .mean()
            .sort_values("Cobertura_acopio_pct", ascending=True)
        )
        st.bar_chart(cat.set_index("Categoria"))
        st.caption("Promedio simple del porcentaje de cobertura de las líneas de necesidad; no suma unidades incompatibles.")
    else:
        st.info("No hay necesidades activas para mostrar.")

with right:
    st.subheader("Estado de las necesidades")
    if not resumen.empty:
        sem = resumen["Semaforo"].value_counts().reindex(
            ["CRÍTICO", "ALERTA", "PRÓXIMO A META", "CUBIERTO"], fill_value=0
        )
        st.bar_chart(sem)
    else:
        st.info("Sin datos.")

st.divider()

# Tabla operativa principal
st.subheader("Prioridades operativas")
if not resumen.empty:
    cols = [
        "ID_Necesidad","Destino","Categoria","Producto","Unidad_operativa","Prioridad",
        "Cantidad_requerida","Cantidad_acopiada","Inventario_disponible",
        "Cantidad_despachada","Cantidad_recibida_destino","Faltante_acopio",
        "Cobertura_acopio_pct","Cobertura_destino_pct","Semaforo","Fecha_limite"
    ]
    cols = [c for c in cols if c in resumen.columns]
    show = resumen[cols].copy()

    prio_rank = {"Alta": 1, "Media": 2, "Baja": 3}
    if "Prioridad" in show:
        show["_prio"] = show["Prioridad"].map(prio_rank).fillna(9)
    else:
        show["_prio"] = 9

    show = show.sort_values(
        ["_prio", "Cobertura_acopio_pct", "Fecha_limite"],
        ascending=[True, True, True]
    ).drop(columns="_prio")

    st.dataframe(
        show,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Cobertura_acopio_pct": st.column_config.ProgressColumn(
                "Cobertura acopio",
                format="%.1f%%",
                min_value=0,
                max_value=100,
            ),
            "Cobertura_destino_pct": st.column_config.ProgressColumn(
                "Cobertura recibida destino",
                format="%.1f%%",
                min_value=0,
                max_value=100,
            ),
            "Fecha_limite": st.column_config.DateColumn("Fecha límite", format="YYYY-MM-DD"),
        }
    )
else:
    st.info("No hay necesidades activas.")

# Inventario
st.divider()
st.subheader("Inventario disponible por producto")
if not resumen.empty:
    inv = (
        resumen.groupby(["Categoria","ID_Producto","Producto","Unidad_operativa"], as_index=False)
        .agg(
            Acopiado=("Cantidad_acopiada","sum"),
            Despachado=("Cantidad_despachada","sum"),
            Disponible=("Inventario_disponible","sum")
        )
    )
    st.dataframe(inv, use_container_width=True, hide_index=True)
else:
    st.info("Sin información de inventario.")

# Tendencia de entradas
st.divider()
left, right = st.columns(2)
with left:
    st.subheader("Entradas aprobadas por día")
    if not entradas_f.empty and "Fecha_hora" in entradas_f:
        trend = entradas_f.dropna(subset=["Fecha_hora"]).copy()
        if not trend.empty:
            trend["Fecha"] = trend["Fecha_hora"].dt.date
            trend = trend.groupby("Fecha", as_index=False)["Cantidad_aprobada"].sum()
            st.line_chart(trend.set_index("Fecha"))
        else:
            st.info("Sin fechas de entrada registradas.")
    else:
        st.info("Sin entradas.")

with right:
    st.subheader("Rechazos por categoría")
    if not entradas_f.empty:
        rej = entradas_f.groupby("Categoria", as_index=False)["Cantidad_rechazada"].sum()
        st.bar_chart(rej.set_index("Categoria"))
    else:
        st.info("Sin entradas.")

# Despachos
st.divider()
st.subheader("Despachos")
if not despachos_f.empty:
    disp_cols = [
        "ID_Despacho","Fecha_hora_salida","Destino","Categoria","Producto",
        "Cantidad_despachada","Unidad_operativa","Vehiculo","Placa",
        "Responsable_CMU","Receptor_previsto","Estado_despacho","ETA"
    ]
    disp_cols = [c for c in disp_cols if c in despachos_f.columns]
    st.dataframe(
        despachos_f[disp_cols].sort_values("Fecha_hora_salida", ascending=False),
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("No hay despachos registrados.")

# Incidencias
st.divider()
st.subheader("Incidencias activas")
if not incidencias.empty and "Estado" in incidencias:
    active_inc = incidencias[
        incidencias["Estado"].astype(str).str.lower().isin(["abierta","en gestión"])
    ].copy()
    if not active_inc.empty:
        sev_rank = {"Crítica":1,"Alta":2,"Media":3,"Baja":4}
        active_inc["_sev"] = active_inc["Severidad"].map(sev_rank).fillna(9)
        active_inc = active_inc.sort_values(["_sev","Fecha_hora"]).drop(columns="_sev")
        st.dataframe(active_inc, use_container_width=True, hide_index=True)
    else:
        st.success("No hay incidencias abiertas.")
else:
    st.success("No hay incidencias abiertas.")

# Alertas automáticas
st.divider()
st.subheader("Alertas automáticas")
alerts = []

if not resumen.empty:
    crit = resumen[resumen["Cobertura_acopio_pct"] < 50]
    for _, r in crit.iterrows():
        alerts.append(
            f"🔴 {r.get('ID_Necesidad','')} — {r.get('Producto','')}: "
            f"cobertura de acopio {r.get('Cobertura_acopio_pct',0):.1f}%."
        )

    overdue = resumen[
        resumen["Fecha_limite"].notna() &
        (resumen["Fecha_limite"] < pd.Timestamp.today().normalize()) &
        (resumen["Cobertura_acopio_pct"] < 100)
    ]
    for _, r in overdue.iterrows():
        alerts.append(
            f"⏰ {r.get('ID_Necesidad','')} — fecha límite vencida y necesidad aún no cubierta."
        )

if not entradas_f.empty and "Fecha_vencimiento" in entradas_f:
    soon = entradas_f[
        entradas_f["Fecha_vencimiento"].notna() &
        (entradas_f["Fecha_vencimiento"] <= pd.Timestamp.today().normalize() + pd.Timedelta(days=90))
    ]
    for _, r in soon.head(20).iterrows():
        alerts.append(
            f"💊/📦 {r.get('ID_Entrada','')} — revisar vencimiento de {r.get('Producto','')} "
            f"({r.get('Fecha_vencimiento').date() if pd.notna(r.get('Fecha_vencimiento')) else ''})."
        )

if alerts:
    for a in alerts:
        st.warning(a)
else:
    st.success("No se detectan alertas automáticas con las reglas actuales.")

st.caption(
    "Nota: el tablero es una herramienta de apoyo operativo. "
    "Las decisiones sobre medicamentos, seguridad, transporte y entrega deben seguir "
    "los protocolos institucionales y de las autoridades competentes."
)

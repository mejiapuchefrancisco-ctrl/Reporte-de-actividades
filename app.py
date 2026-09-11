"""
Dashboard Ejecutivo — AIRPLAN T1 JMC
MVP de control de costos: Mano de Obra + Equipos

CÓMO CORRERLO (en tu PC, con Visual Studio Code):
  1. pip install streamlit pandas plotly requests
  2. En la terminal: streamlit run app.py
  3. Cada vez que le des "refrescar" en el navegador, Mano de Obra se trae
     AUTOMÁTICAMENTE los registros más recientes que los maestros hayan
     guardado desde la app de campo — no requiere ningún paso manual.

ESTADO DE LOS DATOS:
  - Mano de Obra: EN VIVO desde el 7 de septiembre de 2026 en adelante,
    alimentada directo por la app de campo (registro_mano_obra.html). El
    costo se calcula al vuelo cruzando tarifas (en vivo desde Google Drive)
    y recargos (recargos.csv).
  - Tarifas de personal: EN VIVO desde Maestro_Personal_AIRPLAN.xlsx — si
    agregas a alguien nuevo ahí (con su cargo), aparece solo en el próximo
    refresco del dashboard, sin que nadie tenga que exportar nada.
  - Equipos: el inventario general (83 registros) sigue siendo un snapshot
    fijo de agosto 2026 — pero los movimientos/solicitudes de equipo YA
    están en vivo, alimentados por el mismo interruptor "Equipos" de la
    app de campo.
  - Eficiencia: NUEVO — compara la cantidad ejecutada (que el maestro reporta
    por bloque en la app) contra la cantidad esperada según el rendimiento
    presupuestado de Ayudantes/Oficiales. Requiere que el Sheet tenga la
    columna 'loteId' (ver instrucciones abajo).
"""
import re
import unicodedata
import streamlit as st
import pandas as pd
import plotly.express as px

# ============ CONFIGURACIÓN ============
URL_SHEET_MANO_OBRA = (
    "https://docs.google.com/spreadsheets/d/"
    "1Sbt-r9yR_pcyluP_JLX9nA2nwxpW3cU2ATVPMwvNP-I/export?format=csv&gid=0"
)
URL_SHEET_EQUIPOS = (
    "https://docs.google.com/spreadsheets/d/"
    "1Sbt-r9yR_pcyluP_JLX9nA2nwxpW3cU2ATVPMwvNP-I/gviz/tq?tqx=out:csv&sheet=Equipos"
)
URL_TARIFAS = (
    "https://drive.google.com/uc?export=download&id="
    "1gUGj51qvXvmwZvyhhgy-05CY5INWHSST"
)
RUTA_RECARGOS = "recargos.csv"
RUTA_EQUIPOS_INVENTARIO = "hechos_equipos_inventario.csv"
RUTA_LOGO = "aia.png"

# ============ PALETA CORPORATIVA AIA ============
# Verde extraído directamente del logo (#1A5632). El resto son tonos
# complementarios elegidos para que combinen bien con ese verde en gráficas
# con varias categorías — no son colores de marca oficiales, así que si
# tienes una guía de marca con más colores, dímelos y los cambiamos.
VERDE_AIA = "#1A5632"
VERDE_OSCURO = "#0F331D"
VERDE_CLARO = "#4C8A64"
DORADO_ACENTO = "#C9A227"
GRIS_NEUTRO = "#7A8B82"

# Paleta para gráficas con varias categorías (pies, barras por proveedor, etc.)
PALETA_CATEGORICA = [VERDE_AIA, DORADO_ACENTO, GRIS_NEUTRO, VERDE_CLARO, "#8C5E3C", "#3C5A64"]
# Escala para gráficas de un solo valor ordenado (ranking de barras)
ESCALA_VERDE = [VERDE_CLARO, VERDE_AIA, VERDE_OSCURO]

px.defaults.color_discrete_sequence = PALETA_CATEGORICA
px.defaults.color_continuous_scale = ESCALA_VERDE
# ==================================================

st.set_page_config(page_title="Dashboard Ejecutivo AIA", layout="wide", page_icon=RUTA_LOGO)

st.markdown(f"""
<style>
    [data-testid="stMetricValue"] {{ color: {VERDE_AIA}; }}
    .stTabs [aria-selected="true"] {{ color: {VERDE_AIA} !important; }}
    .stTabs [data-baseweb="tab-highlight"] {{ background-color: {VERDE_AIA} !important; }}
    h1, h2, h3 {{ color: {VERDE_OSCURO}; }}
</style>
""", unsafe_allow_html=True)


def normalizar_nombre(nombre):
    if not isinstance(nombre, str):
        return ""
    n = nombre.strip().upper()
    n = unicodedata.normalize("NFKD", n).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", n)


@st.cache_data(ttl=60)
def cargar_datos():
    # --- Mano de Obra: en vivo desde el Google Sheet de la app de campo ---
    mo = pd.read_csv(URL_SHEET_MANO_OBRA)
    mo = mo.dropna(subset=["nombre"]).copy()
    # La columna loteId siempre es la ÚLTIMA columna del Sheet (así la escribe
    # Codigo.gs, en orden fijo) — la identificamos por posición, no por el
    # texto del encabezado, para no depender de que esté escrito perfecto
    # (l minúscula e I mayúscula son casi indistinguibles en muchas fuentes).
    ultima_col = mo.columns[-1]
    if ultima_col != "loteId":
        mo = mo.rename(columns={ultima_col: "loteId"})
    # El encabezado real del Sheet quedó escrito "codcat" en vez de "codact"
    # desde el principio — lo normalizamos aquí para que el resto del código
    # (y el cálculo de Eficiencia) no dependa de que alguien lo corrija allá.
    if "codact" not in mo.columns and "codcat" in mo.columns:
        mo = mo.rename(columns={"codcat": "codact"})
    mo["fecha"] = pd.to_datetime(mo["fecha"], errors="coerce")
    mo["horas"] = pd.to_numeric(mo["horas"], errors="coerce")
    mo["nombre_norm"] = mo["nombre"].apply(normalizar_nombre)

    tarifas = pd.read_excel(URL_TARIFAS, sheet_name="Maestro Personal", header=None)
    header_idx = tarifas[tarifas[0] == "Documento"].index[0]
    tarifas.columns = tarifas.iloc[header_idx]
    tarifas = tarifas.iloc[header_idx + 1:].dropna(subset=["NombreNormalizado"]).copy()
    tarifas["nombre_norm"] = tarifas["NombreNormalizado"].apply(normalizar_nombre)
    tarifas["TarifaHoraReal"] = pd.to_numeric(tarifas["TarifaHoraReal"], errors="coerce")

    recargos = pd.read_csv(RUTA_RECARGOS)

    mo = mo.merge(tarifas[["nombre_norm", "TarifaHoraReal"]], on="nombre_norm", how="left")
    mo = mo.merge(recargos[["TipoHora", "Factor"]], left_on="tipoHora", right_on="TipoHora", how="left")
    mo["match_tarifa"] = mo["TarifaHoraReal"].notna()
    mo["costo_calculado"] = mo["horas"] * mo["TarifaHoraReal"] * mo["Factor"]
    mo = mo.rename(columns={
        "frenteNombre": "frente", "tipoHora": "tipo_hora", "TarifaHoraReal": "tarifa_hora",
    })

    # --- Equipos: inventario fijo (agosto) + movimientos EN VIVO (app unificada) ---
    inv = pd.read_csv(RUTA_EQUIPOS_INVENTARIO, parse_dates=["fecha"])
    mov = pd.read_csv(URL_SHEET_EQUIPOS)
    mov = mov.dropna(subset=["equipo"]).copy()
    mov["cantidad"] = pd.to_numeric(mov["cantidad"], errors="coerce")
    mov["valorUnitario"] = pd.to_numeric(mov["valorUnitario"], errors="coerce")
    mov["costo_estimado"] = mov["cantidad"] * mov["valorUnitario"]
    mov = mov.rename(columns={"frenteNombre": "frente"})
    return mo, inv, mov


def rol_generico(cargo):
    """Traduce un cargo real (ej. 'Maestro Primero_Electricista') al rol
    genérico del catálogo de rendimientos (Ayudante / Oficial). Maestro no
    tiene tasa propia en el catálogo, así que no se traduce (queda fuera
    del cálculo de productividad)."""
    if not isinstance(cargo, str):
        return None
    c = cargo.lower()
    if "maestro" in c:
        return None
    if "ayudante" in c or "auxiliar" in c:
        return "Ayudante"
    if "oficial" in c:
        return "Oficial"
    return None


@st.cache_data(ttl=60)
def calcular_eficiencia(mo_df):
    if "loteId" not in mo_df.columns or "cantidad" not in mo_df.columns:
        return None  # el Sheet todavía no tiene las columnas nuevas

    rendimientos = pd.read_csv("rendimientos.csv")

    df = mo_df.dropna(subset=["loteId", "cantidad"]).copy()
    if df.empty:
        return df

    df["rol"] = df["cargo"].apply(rol_generico)
    df["horas"] = pd.to_numeric(df["horas"], errors="coerce")
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce")

    # Cruce por actividad + rol genérico para traer el rendimiento presupuestado
    df = df.merge(rendimientos, left_on=["codact", "rol"], right_on=["codact", "rol"], how="left")

    df["cantidad_esperada_persona"] = df["horas"] * df["rendimiento"]
    df["horas_productivas"] = df["horas"].where(df["rol"].notna(), 0)

    # Agregamos por bloque (loteId): la cantidad ejecutada es la misma en
    # todas las filas del bloque, así que tomamos una sola vez (first);
    # la cantidad esperada sí se suma persona por persona (solo Ayudante/Oficial).
    agg = df.groupby("loteId").agg(
        fecha=("fecha", "first"),
        frente=("frente", "first"),
        descripcion_actividad=("actDesc", "first"),
        unidad=("actUnidad", "first"),
        cantidad_ejecutada=("cantidad", "first"),
        cantidad_esperada=("cantidad_esperada_persona", "sum"),
        horas_productivas=("horas_productivas", "sum"),
    ).reset_index()

    agg = agg[agg["cantidad_esperada"] > 0].copy()
    agg["eficiencia_pct"] = 100 * agg["cantidad_ejecutada"] / agg["cantidad_esperada"]
    return agg


mo, inv, mov = cargar_datos()


col_logo, col_titulo = st.columns([1, 6])
with col_logo:
    st.image(RUTA_LOGO, width=120)
with col_titulo:

    st.title("Dashboard Ejecutivo — Control de Costos")
    st.caption("AIRPLAN T1 JMC · Mano de Obra + Equipos")

with st.sidebar:
    st.header("Filtros")
    frentes_mo = sorted(mo["frente"].dropna().unique())
    frente_sel = st.multiselect("Frente de trabajo", frentes_mo, default=frentes_mo)
    st.divider()
    st.info(
        "🟢 **Mano de Obra: datos en vivo**\n\n"
        f"{len(mo)} registro(s) guardados por los maestros desde la app de campo "
        "(a partir del 7 de septiembre de 2026). Se actualiza solo al refrescar."
    )

mo_f = mo[mo["frente"].isin(frente_sel)] if frente_sel else mo

# ============ KPIs PRINCIPALES ============
col1, col2, col3, col4 = st.columns(4)
col1.metric("Costo M.O. calculado", f"$ {mo_f['costo_calculado'].sum():,.0f}")
col2.metric("Horas registradas", f"{mo_f['horas'].sum():,.0f} h")
col3.metric("Personal activo", mo_f["nombre"].nunique())
col4.metric(
    "Valor inventariado Equipos",
    f"$ {(inv['cantidad'] * inv['valorUnitario']).sum():,.0f}",
)

st.divider()

tab1, tab2, tab3, tab4 = st.tabs(
    ["👷 Mano de Obra", "🚜 Equipos", "📈 Eficiencia", "🧹 Calidad de datos"])

# ============ TAB MANO DE OBRA ============
with tab1:
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Costo por frente de trabajo")
        costo_frente = (
            mo_f.groupby("frente", as_index=False)["costo_calculado"]
            .sum()
            .sort_values("costo_calculado", ascending=True)
        )
        fig = px.bar(
            costo_frente, x="costo_calculado", y="frente", orientation="h",
            labels={"costo_calculado": "Costo ($)", "frente": "Frente"},
            color="costo_calculado", color_continuous_scale=ESCALA_VERDE,
        )
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, width='stretch')

    with c2:
        st.subheader("Horas por tipo de hora")
        horas_tipo = mo_f.groupby("tipo_hora", as_index=False)["horas"].sum()
        fig2 = px.pie(horas_tipo, names="tipo_hora", values="horas", hole=0.4,
                       color_discrete_sequence=PALETA_CATEGORICA)
        st.plotly_chart(fig2, width='stretch')

    st.subheader("Costo acumulado en el tiempo")
    acumulado = (
        mo_f.sort_values("fecha")
        .groupby("fecha", as_index=False)["costo_calculado"].sum()
    )
    acumulado["costo_acumulado"] = acumulado["costo_calculado"].cumsum()
    fig3 = px.line(acumulado, x="fecha", y="costo_acumulado", markers=True,
                    color_discrete_sequence=[VERDE_AIA])
    fig3.update_traces(line=dict(width=3), marker=dict(size=8, color=VERDE_OSCURO))
    st.plotly_chart(fig3, width='stretch')

    st.subheader("Detalle de registros")
    st.dataframe(
        mo_f[["fecha", "nombre", "cargo", "frente", "horas", "tipo_hora",
              "tarifa_hora", "costo_calculado"]].sort_values("fecha", ascending=False),
        width='stretch',
    )

# ============ TAB EQUIPOS ============
with tab2:
    st.caption(
        "📋 Inventario y gráficas de categoría: snapshot fijo de agosto 2026. "
        "🟢 Movimientos/solicitudes (abajo): en vivo desde la app de campo."
    )
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Equipos por categoría")
        cat = inv.groupby("categoriaEquipo", as_index=False)["cantidad"].sum()
        fig4 = px.bar(cat, x="categoriaEquipo", y="cantidad",
                       labels={"categoriaEquipo": "Categoría", "cantidad": "Cantidad"},
                       color="categoriaEquipo", color_discrete_sequence=PALETA_CATEGORICA)
        fig4.update_layout(showlegend=False)
        st.plotly_chart(fig4, width='stretch')

    with c2:
        st.subheader("Valor inventariado por proveedor")
        inv["valor_total"] = inv["cantidad"] * inv["valorUnitario"]
        prov = inv.groupby("proveedor", as_index=False)["valor_total"].sum()
        fig5 = px.pie(prov, names="proveedor", values="valor_total", hole=0.4,
                       color_discrete_sequence=PALETA_CATEGORICA)
        st.plotly_chart(fig5, width='stretch')

    st.subheader("Inventario por frente")
    frente_inv = inv.groupby("frenteNombre", as_index=False)["cantidad"].sum().sort_values(
        "cantidad", ascending=True)
    fig6 = px.bar(frente_inv, x="cantidad", y="frenteNombre", orientation="h",
                   color="cantidad", color_continuous_scale=ESCALA_VERDE)
    fig6.update_layout(coloraxis_showscale=False)
    st.plotly_chart(fig6, width='stretch')

    st.subheader("Detalle de inventario")
    st.dataframe(inv, width='stretch')

    st.subheader("🟢 Movimientos / solicitudes de equipo (en vivo)")
    st.metric("Costo estimado en movimientos registrados", f"$ {mov['costo_estimado'].sum(skipna=True):,.0f}")
    st.dataframe(
        mov[["fecha", "solicitante", "frente", "proveedor", "equipo", "unidad",
             "cantidad", "valorUnitario", "costo_estimado"]].sort_values("fecha", ascending=False),
        width='stretch',
    )

# ============ TAB EFICIENCIA ============
with tab3:
    st.subheader("Cumplimiento de productividad por actividad")
    st.caption(
        "Compara la cantidad ejecutada (reportada por el maestro) contra la cantidad "
        "esperada según el rendimiento presupuestado de Ayudantes y Oficiales en esa "
        "actividad. **Las horas de Maestro no se incluyen en el cálculo** — el catálogo "
        "de rendimientos no tiene una tasa productiva específica para ese cargo, ya que "
        "su rol es principalmente de supervisión."
    )

    efi = calcular_eficiencia(mo)

    if efi is None:
        st.warning(
            "⚠️ Todavía no has agregado la columna **loteId** a la hoja 'Registros' "
            "de tu Google Sheet. Agrégala en la primera fila (después de 'registradoPor') "
            "para que este cálculo empiece a funcionar."
        )
    elif efi.empty:
        st.info(
            "Todavía no hay registros con 'cantidad ejecutada' diligenciada desde la app, "
            "o ninguno tiene personal de cargo Ayudante/Oficial para calcular el rendimiento "
            "esperado."
        )
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Bloques con dato de producción", len(efi))
        c2.metric("Eficiencia promedio", f"{efi['eficiencia_pct'].mean():.0f}%")
        c3.metric("Bloques por debajo del 80%", int((efi["eficiencia_pct"] < 80).sum()))

        st.subheader("Eficiencia por actividad")
        por_actividad = efi.groupby("descripcion_actividad", as_index=False).agg(
            cantidad_ejecutada=("cantidad_ejecutada", "sum"),
            cantidad_esperada=("cantidad_esperada", "sum"),
            horas_productivas=("horas_productivas", "sum"),
        )
        por_actividad["eficiencia_pct"] = (
            100 * por_actividad["cantidad_ejecutada"] / por_actividad["cantidad_esperada"]
        )
        fig_efi = px.bar(
            por_actividad.sort_values("eficiencia_pct"),
            x="eficiencia_pct", y="descripcion_actividad", orientation="h",
            labels={"eficiencia_pct": "Eficiencia (%)", "descripcion_actividad": "Actividad"},
            color="eficiencia_pct", color_continuous_scale=[DORADO_ACENTO, VERDE_CLARO, VERDE_AIA],
        )
        fig_efi.add_vline(x=100, line_dash="dash", line_color=GRIS_NEUTRO)
        fig_efi.update_layout(coloraxis_showscale=False, yaxis={"tickfont": {"size": 10}})
        st.plotly_chart(fig_efi, width='stretch')

        st.subheader("Eficiencia por frente")
        por_frente = efi.groupby("frente", as_index=False).agg(
            cantidad_ejecutada=("cantidad_ejecutada", "sum"),
            cantidad_esperada=("cantidad_esperada", "sum"),
        )
        por_frente["eficiencia_pct"] = 100 * por_frente["cantidad_ejecutada"] / por_frente["cantidad_esperada"]
        fig_frente = px.bar(
            por_frente.sort_values("eficiencia_pct"),
            x="eficiencia_pct", y="frente", orientation="h",
            color="eficiencia_pct", color_continuous_scale=[DORADO_ACENTO, VERDE_CLARO, VERDE_AIA],
        )
        fig_frente.add_vline(x=100, line_dash="dash", line_color=GRIS_NEUTRO)
        fig_frente.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig_frente, width='stretch')

        st.subheader("Detalle por bloque")
        st.dataframe(
            efi[["fecha", "frente", "descripcion_actividad", "unidad", "cantidad_ejecutada",
                 "cantidad_esperada", "horas_productivas", "eficiencia_pct"]]
            .sort_values("fecha", ascending=False),
            width='stretch',
        )

# ============ TAB CALIDAD DE DATOS ============
with tab4:
    st.subheader("Personal sin tarifa confirmada")
    sin_tarifa = mo.loc[~mo["match_tarifa"], ["nombre", "cargo", "frente", "horas"]]
    if len(sin_tarifa):
        st.dataframe(sin_tarifa, width='stretch')
        st.info(
            "Estos nombres no cruzaron con la tabla maestra de personal. "
            "Revisa si hay errores de digitación o personal nuevo sin tarifa asignada."
        )
    else:
        st.success("Todos los registros tienen tarifa confirmada.")

    st.subheader("Resumen de cobertura")
    st.metric("Total de registros en vivo", len(mo))

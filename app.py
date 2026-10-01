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
  - Rendimiento real por categoría: NUEVO (23-sep-2026) — cruza las horas de
    Registros con las cantidades de la pestaña Mediciones (Oficiales), por
    semana + frente + categoría, usando categorias_sinco.csv.
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
URL_SHEET_MEDICIONES = (
    "https://docs.google.com/spreadsheets/d/"
    "1Sbt-r9yR_pcyluP_JLX9nA2nwxpW3cU2ATVPMwvNP-I/gviz/tq?tqx=out:csv&headers=1&sheet=Mediciones"
)
RUTA_CATEGORIAS = "categorias_sinco.csv"  # código SINCO -> categoría A-K / NM / FASE2
# Las 11 categorías del formulario de Mediciones (mismo nombre exacto que envía
# medicion_cantidades.html, porque el Sheet guarda el nombre y no la letra)
CATEGORIAS_MEDICION = {
    "A": ("Vaciado de concreto", "m³"),
    "B": ("Mamposteria (muro en bloque + dovelas)", "m²"),
    "C": ("Revoque", "m²"),
    "D": ("Pintura", "m²"),
    "E": ("Enchape / Revestimiento ceramico", "m²"),
    "F": ("Instalacion de piso (baldosa/granito/porcelanato)", "m²"),
    "G": ("Instalacion de cielo raso", "m²"),
    "H": ("Muro en drywall / Superboard", "m²"),
    "I": ("Demolicion (muros y pisos)", "m²"),
    "J": ("Excavacion", "m³"),
    "K": ("Retiro de escombros / Trasiego", "m³"),
}
NOMBRES_CATEGORIA = {k: v[0] for k, v in CATEGORIAS_MEDICION.items()}
NOMBRES_CATEGORIA.update({
    "L": "Aseo y orden", "M": "Apoyo y logistica (cargue, acarreo)", "N": "Desmonte y retiro",
    "O": "Instalaciones provisionales", "P": "Otra",
})
# Referencia de productividad por frente × categoría (HH/unidad). Sale de la columna
# "HH/unidad FINAL" de Referencias_Productividad_v2_Publicas.xlsx (APU SINCO ponderado
# con rendimientos públicos). Si cambias el Excel, se vuelve a exportar este CSV.
RUTA_REFERENCIAS = "referencias_productividad.csv"
# Semáforo: real / referencia
SEMAFORO_AMARILLO = 1.20  # hasta 20 % por encima de la referencia
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


def etiqueta_corta_frente(nombre_frente):
    """Extrae la etiqueta corta real que usan en campo para nombrar el
    frente (ej. '3A' de '3A - ALTERNATIVA SISTEMA...'), en vez del código
    interno secuencial (01, 02...) que no coincide con esa numeración."""
    if not isinstance(nombre_frente, str):
        return nombre_frente
    m = re.match(r"^(.*?)\s*-\s*", nombre_frente)
    return m.group(1) if m else nombre_frente


def formato_pesos_millones(valor):
    """Formatea un valor en pesos colombianos, abreviado en millones para
    que sea fácil de leer de un vistazo (ej. $26.6M en vez de $26.639.450)."""
    if pd.isna(valor):
        return "$ 0"
    if abs(valor) >= 1_000_000:
        return f"$ {valor/1_000_000:,.1f}M"
    if abs(valor) >= 1_000:
        return f"$ {valor/1_000:,.0f}K"
    return f"$ {valor:,.0f}"


def etiqueta_semana(fecha):
    """Devuelve la semana laboral (lunes a SÁBADO, ya que sábado es día
    laboral normal en este proyecto) de una fecha, como 'DD Mon - DD Mon',
    junto con el lunes de esa semana (para ordenar cronológicamente)."""
    if pd.isna(fecha):
        return pd.Series({"semana_lunes": pd.NaT, "semana_etiqueta": None})
    lunes = fecha - pd.Timedelta(days=fecha.weekday())
    sabado = lunes + pd.Timedelta(days=5)
    meses = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
             7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
    etiqueta = f"{lunes.day:02d} {meses[lunes.month]} - {sabado.day:02d} {meses[sabado.month]}"
    return pd.Series({"semana_lunes": lunes, "semana_etiqueta": etiqueta})


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
    # El Sheet exporta los decimales con coma ("7,5"): sin esto se perdían las horas con decimales
    mo["horas"] = pd.to_numeric(mo["horas"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
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
    mo[["semana_lunes", "semana_etiqueta"]] = mo["fecha"].apply(etiqueta_semana)
    mo["frente_corto"] = mo["frente"].apply(etiqueta_corta_frente)

    # --- Equipos: inventario fijo (agosto) + movimientos EN VIVO (app unificada) ---
    inv = pd.read_csv(RUTA_EQUIPOS_INVENTARIO, parse_dates=["fecha"])
    mov = pd.read_csv(URL_SHEET_EQUIPOS)
    mov = mov.dropna(subset=["equipo"]).copy()
    mov["cantidad"] = pd.to_numeric(mov["cantidad"], errors="coerce")
    mov["valorUnitario"] = pd.to_numeric(mov["valorUnitario"], errors="coerce")
    mov["costo_estimado"] = mov["cantidad"] * mov["valorUnitario"]
    mov = mov.rename(columns={"frenteNombre": "frente"})
    mov["fecha"] = pd.to_datetime(mov["fecha"], errors="coerce")
    mov[["semana_lunes", "semana_etiqueta"]] = mov["fecha"].apply(etiqueta_semana)
    mov["frente_corto"] = mov["frente"].apply(etiqueta_corta_frente)
    return mo, inv, mov, tarifas


def rol_generico(cargo):
    """Traduce un cargo real (ej. 'Maestro Primero_Electricista') al rol
    genérico del catálogo de rendimientos (Ayudante / Oficial / Otro).
    Maestro usa la tasa 'Otro' (trabajo de supervisión/mixto) en vez de
    quedar excluido del cálculo de productividad."""
    if not isinstance(cargo, str):
        return None
    c = cargo.lower()
    if "ayudante" in c or "auxiliar" in c:
        return "Ayudante"
    if "oficial" in c:
        return "Oficial"
    if "maestro" in c:
        return "Otro"
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


@st.cache_data(ttl=60)
def cargar_mediciones():
    """Mediciones diarias de los Oficiales (pestaña Mediciones del Sheet, en vivo).
    Devuelve un DataFrame vacío (con columnas) si todavía no hay mediciones."""
    cols = ["fecha", "frente_cod", "cat", "cantidad", "oficial"]
    try:
        med = pd.read_csv(URL_SHEET_MEDICIONES)
    except Exception:
        return pd.DataFrame(columns=cols)
    if "frenteCod" not in med.columns or "categoria" not in med.columns or med.empty:
        return pd.DataFrame(columns=cols)
    nombre_a_letra = {v[0]: k for k, v in CATEGORIAS_MEDICION.items()}
    med["cat"] = med["categoria"].map(nombre_a_letra)
    med["cantidad"] = pd.to_numeric(
        med["cantidad"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    med["frente_cod"] = (pd.to_numeric(med["frenteCod"], errors="coerce")
                         .astype("Int64").astype(str).str.zfill(2))
    med["fecha"] = pd.to_datetime(med["fecha"], errors="coerce")
    med = med.dropna(subset=["fecha", "cat", "cantidad"])
    return med[cols].copy()


def cruce_rendimiento(mo_df, med_df):
    """Une horas-hombre (Registros) con cantidades medidas (Mediciones) por
    semana + frente + categoría. El código SINCO de cada hora se traduce a su
    categoría con categorias_sinco.csv; la cantidad la pone el Oficial."""
    cat = pd.read_csv(RUTA_CATEGORIAS, dtype=str)
    h = mo_df.dropna(subset=["codact", "fecha"]).copy()
    h["codact"] = h["codact"].astype(str).str.strip()
    cat = cat.rename(columns={"frente": "frente_cod", "categoria": "cat"})
    # Actividades no presupuestadas: el código ya trae frente y categoría (NP-05-C)
    np_cods = pd.Series(mo_df["codact"].dropna().astype(str).str.strip().unique())
    np_cods = np_cods[np_cods.str.match(r"^NP-\d{2}-[A-Z]$")]
    if not np_cods.empty:
        cat = pd.concat([cat, pd.DataFrame({"codact": np_cods, "frente_cod": np_cods.str[3:5],
                                            "cat": np_cods.str[6]})], ignore_index=True)
    h = h.merge(cat[["codact", "frente_cod", "cat"]], on="codact", how="left")
    h = h[h["cat"].isin(CATEGORIAS_MEDICION.keys())]
    horas = h.groupby(["semana_lunes", "semana_etiqueta", "frente_cod", "cat"], as_index=False).agg(
        HH=("horas", "sum"), costo=("costo_calculado", "sum"), personas=("nombre", "nunique"))

    m = med_df.copy()
    if not m.empty:
        m[["semana_lunes", "semana_etiqueta"]] = m["fecha"].apply(etiqueta_semana)
    cant = (m.groupby(["semana_lunes", "semana_etiqueta", "frente_cod", "cat"], as_index=False)
            .agg(cantidad=("cantidad", "sum"), mediciones=("cantidad", "size"))
            if not m.empty else
            pd.DataFrame(columns=["semana_lunes", "semana_etiqueta", "frente_cod", "cat", "cantidad", "mediciones"]))

    x = horas.merge(cant, on=["semana_lunes", "semana_etiqueta", "frente_cod", "cat"], how="outer")
    x["HH"] = pd.to_numeric(x["HH"], errors="coerce").fillna(0)
    x["costo"] = pd.to_numeric(x["costo"], errors="coerce").fillna(0)
    x["cantidad"] = pd.to_numeric(x["cantidad"], errors="coerce").fillna(0)
    # Combinaciones sin horas y sin cantidad (p. ej. registros con 0 h) no aportan nada
    x = x[(x["HH"] > 0) | (x["cantidad"] > 0)].copy()
    x["estado"] = "⏳ Horas sin medición"
    x.loc[(x["HH"] > 0) & (x["cantidad"] > 0), "estado"] = "✅ Cruce completo"
    x.loc[(x["HH"] <= 0) & (x["cantidad"] > 0), "estado"] = "⚠️ Medición sin horas"
    ok = x["estado"] == "✅ Cruce completo"
    x["HH_por_unidad"] = (x["HH"] / x["cantidad"]).where(ok).astype(float)
    x["costo_por_unidad"] = (x["costo"] / x["cantidad"]).where(ok).astype(float)
    x["categoria"] = x["cat"].map(lambda c: f"{c} - {CATEGORIAS_MEDICION[c][0]}")
    x["unidad"] = x["cat"].map(lambda c: CATEGORIAS_MEDICION[c][1])
    x = agregar_referencia(x)
    return x.sort_values(["semana_lunes", "frente_cod", "cat"])


@st.cache_data(ttl=3600)
def cargar_referencias():
    try:
        ref = pd.read_csv(RUTA_REFERENCIAS, dtype={"frente_cod": str, "cat": str})
    except Exception:
        return pd.DataFrame(columns=["frente_cod", "cat", "ref_hh_unidad"])
    ref["frente_cod"] = ref["frente_cod"].str.zfill(2)
    ref["ref_hh_unidad"] = pd.to_numeric(ref["ref_hh_unidad"], errors="coerce")
    return ref[["frente_cod", "cat", "ref_hh_unidad"]]


def semaforo(indice):
    if pd.isna(indice):
        return "—"
    if indice <= 1:
        return "🟢"
    if indice <= SEMAFORO_AMARILLO:
        return "🟡"
    return "🔴"


def agregar_referencia(x):
    """Agrega la referencia (HH/unidad) de su frente × categoría y el semáforo.
    Índice = HH/unidad real ÷ referencia (1,00 = igual a la referencia; menor = mejor).
    $ referencia/unidad usa la tarifa real promedio de esas horas (con recargos)."""
    x = x.drop(columns=[c for c in ["ref_hh_unidad"] if c in x.columns])
    x = x.merge(cargar_referencias(), on=["frente_cod", "cat"], how="left")
    x["indice"] = (x["HH_por_unidad"] / x["ref_hh_unidad"]).astype(float)
    x["semaforo"] = x["indice"].map(semaforo)
    tarifa_prom = (x["costo"] / x["HH"]).where(x["HH"] > 0)
    x["costo_ref_por_unidad"] = (x["ref_hh_unidad"] * tarifa_prom).where(x["HH_por_unidad"].notna())
    x["diferencia_pesos"] = (x["costo"] - x["costo_ref_por_unidad"] * x["cantidad"]).where(x["indice"].notna())
    return x


mo, inv, mov, tarifas = cargar_datos()
med = cargar_mediciones()


col_logo, col_titulo = st.columns([1, 6])
with col_logo:
    st.image(RUTA_LOGO, width=120)
with col_titulo:

    st.title("Dashboard Ejecutivo — Control de Costos")
    st.caption("AIRPLAN T1 JMC · Mano de Obra + Equipos")

with st.sidebar:
    st.header("Filtros")
    frentes_mo = sorted(mo["frente"].dropna().unique())
    frente_sel = st.multiselect("Frente de trabajo / intervención", frentes_mo, default=frentes_mo)

    semanas_disponibles = (
        mo.dropna(subset=["semana_lunes", "semana_etiqueta"])
        .drop_duplicates(subset=["semana_etiqueta"])
        .sort_values("semana_lunes")["semana_etiqueta"].tolist()
    )
    semana_sel = st.multiselect("Semana (lunes a sábado)", semanas_disponibles, default=semanas_disponibles)

    fecha_min = mo["fecha"].min()
    fecha_max = mo["fecha"].max()
    rango_fecha = st.date_input(
        "Rango de fechas (o un solo día)",
        value=(fecha_min.date(), fecha_max.date()) if pd.notna(fecha_min) else None,
        min_value=fecha_min.date() if pd.notna(fecha_min) else None,
        max_value=fecha_max.date() if pd.notna(fecha_max) else None,
    )
    st.divider()
    st.info(
        "🟢 **Mano de Obra: datos en vivo**\n\n"
        f"{len(mo)} registro(s) guardados por los maestros desde la app de campo "
        "(a partir del 7 de septiembre de 2026). Se actualiza solo al refrescar."
    )

mo_f = mo[mo["frente"].isin(frente_sel)] if frente_sel else mo
mo_f = mo_f[mo_f["semana_etiqueta"].isin(semana_sel)] if semana_sel else mo_f
mov_f = mov.copy()
mov_f = mov_f[mov_f["semana_etiqueta"].isin(semana_sel)] if semana_sel else mov_f
if isinstance(rango_fecha, tuple) and len(rango_fecha) == 2:
    d_ini, d_fin = rango_fecha
    mo_f = mo_f[(mo_f["fecha"].dt.date >= d_ini) & (mo_f["fecha"].dt.date <= d_fin)]
    mov_f = mov_f[(mov_f["fecha"].dt.date >= d_ini) & (mov_f["fecha"].dt.date <= d_fin)]
elif isinstance(rango_fecha, tuple) and len(rango_fecha) == 1:
    mo_f = mo_f[mo_f["fecha"].dt.date == rango_fecha[0]]
    mov_f = mov_f[mov_f["fecha"].dt.date == rango_fecha[0]]

# Mediciones: mismos filtros de fecha/semana/frente que Mano de Obra
med_f = med.copy()
if not med_f.empty:
    med_f[["semana_lunes", "semana_etiqueta"]] = med_f["fecha"].apply(etiqueta_semana)
    med_f = med_f[med_f["semana_etiqueta"].isin(semana_sel)] if semana_sel else med_f
    if isinstance(rango_fecha, tuple) and len(rango_fecha) == 2:
        med_f = med_f[(med_f["fecha"].dt.date >= rango_fecha[0]) & (med_f["fecha"].dt.date <= rango_fecha[1])]
    elif isinstance(rango_fecha, tuple) and len(rango_fecha) == 1:
        med_f = med_f[med_f["fecha"].dt.date == rango_fecha[0]]
    if frente_sel:
        codigos_sel = (pd.to_numeric(mo.loc[mo["frente"].isin(frente_sel), "codfrente"], errors="coerce")
                       .dropna().astype(int).astype(str).str.zfill(2).unique())
        med_f = med_f[med_f["frente_cod"].isin(codigos_sel)]
    med_f = med_f.drop(columns=["semana_lunes", "semana_etiqueta"])

# ============ KPIs PRINCIPALES (Mano de Obra) ============
col1, col2, col3 = st.columns(3)
col1.metric("Costo M.O. calculado", f"$ {mo_f['costo_calculado'].sum():,.0f}")
col2.metric("Horas registradas", f"{mo_f['horas'].sum():,.0f} h")
col3.metric("Personal activo", mo_f["nombre"].nunique())

st.divider()

# Pestañas visibles (24-sep-2026): Resumen semanal (fusiona Resumen Ejecutivo +
# Mano de Obra), Códigos por Persona y Productividad. Equipos y Calidad de datos
# quedan ocultas; para volver a mostrarlas cambia estas banderas a True.
MOSTRAR_EQUIPOS = False
MOSTRAR_CALIDAD = False
tab0, tab8, tab5, tab3, tab7 = st.tabs(["📋 Resumen semanal", "✅ Cobertura de registro", "🔎 Códigos por Persona", "📈 Productividad", "⚠️ Cargo vs. tarea"])

# ============ TAB RESUMEN EJECUTIVO SEMANAL ============
with tab0:
    st.subheader("Resumen ejecutivo — periodo filtrado")
    st.caption(
        "Misma combinacion del informe semanal en Word, en vivo: respeta los "
        "filtros de frente, semana y fecha de la barra lateral."
    )

    if mo_f.empty:
        st.info("No hay registros de mano de obra para los filtros seleccionados.")
    else:
        horas_totales = mo_f["horas"].sum()
        costo_total = mo_f["costo_calculado"].sum()
        personal_activo = mo_f["nombre"].nunique()
        frentes_trabajados = mo_f["frente_corto"].nunique()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Horas totales", f"{horas_totales:,.1f} h")
        k2.metric("Personal activo", personal_activo)
        k3.metric("Frentes trabajados", frentes_trabajados)
        k4.metric("Costo estimado", f"$ {costo_total:,.0f}")

        # --- Alerta de recargo ---
        costo_ordinaria = mo_f.loc[mo_f["tipo_hora"] == "Ordinaria Diurna", "costo_calculado"].sum()
        costo_extra = costo_total - costo_ordinaria
        pct_extra = (costo_extra / costo_total * 100) if costo_total else 0
        horas_extra = mo_f.loc[mo_f["tipo_hora"] != "Ordinaria Diurna", "horas"].sum()

        if pct_extra >= 50:
            st.error(
                f"⚠️ **Recargo elevado / supera la jornada ordinaria:** el **{pct_extra:,.1f}%** "
                f"del costo (${costo_extra:,.0f}) corresponde a horas extra y nocturnas "
                f"({horas_extra:,.0f} h), frente a un {100-pct_extra:,.1f}% en jornada ordinaria."
            )
        elif pct_extra >= 35:
            st.warning(
                f"⚠️ **Recargo alto:** el **{pct_extra:,.1f}%** del costo (${costo_extra:,.0f}) "
                f"corresponde a horas extra y nocturnas ({horas_extra:,.0f} h)."
            )
        else:
            st.success(
                f"✅ Recargo dentro de rango normal: {pct_extra:,.1f}% del costo en horas extra/nocturnas."
            )

        st.divider()

        ce1, ce2 = st.columns(2)
        with ce1:
            st.markdown("**Distribucion por Frente**")
            pf = (
                mo_f.groupby("frente", as_index=False)
                .agg(horas=("horas", "sum"), costo=("costo_calculado", "sum"))
                .sort_values("costo", ascending=True)
            )
            pf["pct"] = pf["costo"] / pf["costo"].sum() * 100
            figpf = px.bar(
                pf, x="costo", y="frente", orientation="h", text=pf["pct"].map(lambda x: f"{x:.1f}%"),
                labels={"costo": "Costo ($)", "frente": ""},
                color="costo", color_continuous_scale=ESCALA_VERDE,
            )
            figpf.update_layout(coloraxis_showscale=False)
            figpf.update_traces(textposition="outside")
            st.plotly_chart(figpf, width='stretch')

        with ce2:
            st.markdown("**Distribucion por Tipo de Hora**")
            pt = (
                mo_f.groupby("tipo_hora", as_index=False)
                .agg(horas=("horas", "sum"), costo=("costo_calculado", "sum"))
                .sort_values("costo", ascending=True)
            )
            pt["pct"] = pt["costo"] / pt["costo"].sum() * 100
            colores_tipo = ["#B00000" if t != "Ordinaria Diurna" else VERDE_AIA for t in pt["tipo_hora"]]
            figpt = px.bar(
                pt, x="costo", y="tipo_hora", orientation="h", text=pt["pct"].map(lambda x: f"{x:.1f}%"),
                labels={"costo": "Costo ($)", "tipo_hora": ""},
            )
            figpt.update_traces(marker_color=colores_tipo, textposition="outside")
            st.plotly_chart(figpt, width='stretch')

        st.divider()

        st.markdown("**¿En que se esta yendo el costo? — Personal mas representativo**")
        top_n_personas = 6
        tp = (
            mo_f.groupby("nombre", as_index=False)
            .agg(cargo=("cargo", "first"), horas=("horas", "sum"), costo=("costo_calculado", "sum"))
            .sort_values("costo", ascending=False)
            .head(top_n_personas)
        )
        recargo_pct = mo_f.groupby("nombre").apply(
            lambda g: (g.loc[g["tipo_hora"] != "Ordinaria Diurna", "costo_calculado"].sum()
                       / g["costo_calculado"].sum() * 100) if g["costo_calculado"].sum() else 0
        )
        tp["% recargo"] = tp["nombre"].map(recargo_pct).round(0).astype(int)
        pct_top_personas = tp["costo"].sum() / costo_total * 100 if costo_total else 0
        st.caption(f"Las {top_n_personas} personas de mayor costo concentran el {pct_top_personas:.0f}% del gasto del periodo.")
        st.dataframe(
            tp.rename(columns={"nombre": "Nombre", "cargo": "Cargo", "horas": "Horas", "costo": "Costo"})
            .style.format({"Costo": "${:,.0f}", "Horas": "{:,.0f}", "% recargo": "{:d}%"})
            .map(lambda v: "color: #B00000; font-weight: bold" if isinstance(v, (int, float)) and v >= 50 else "",
                 subset=["% recargo"]),
            width='stretch', hide_index=True,
        )

        st.markdown("**Actividades mas representativas del costo**")
        top_n_act = 6
        ta = (
            mo_f.groupby(["frente_corto", "actDesc"], as_index=False)
            .agg(horas=("horas", "sum"), costo=("costo_calculado", "sum"))
            .sort_values("costo", ascending=False)
            .head(top_n_act)
        )
        ta["actDesc"] = ta["actDesc"].str.split("(").str[0].str.strip().str.slice(0, 70)
        pct_top_act = ta["costo"].sum() / costo_total * 100 if costo_total else 0
        st.caption(f"Las {top_n_act} actividades de mayor costo concentran el {pct_top_act:.0f}% del gasto del periodo.")
        st.dataframe(
            ta.rename(columns={"frente_corto": "Frente", "actDesc": "Actividad", "horas": "Horas", "costo": "Costo"})
            .style.format({"Costo": "${:,.0f}", "Horas": "{:,.0f}"}),
            width='stretch', hide_index=True,
        )

    st.divider()
    st.subheader("Asistencia de personal de obra (Maestro, Oficial, Ayudante)")
    st.caption(
        "Compara el roster completo del Maestro de Personal contra quien "
        "registro actividad en los dias presentes en el periodo filtrado "
        "(barra lateral). Solo incluye cargos de obra, no almacen/servicios generales."
    )

    dias_periodo = sorted(mo_f["fecha"].dropna().dt.date.unique())
    roster_obra = tarifas[
        tarifas["CargoReal"].astype(str).str.contains("Maestro|Oficial|Ayudante", case=False, na=False)
        & ~tarifas["CargoReal"].astype(str).str.contains("Almacen|Servicios Generales", case=False, na=False)
    ].copy()

    if not dias_periodo or roster_obra.empty:
        st.info("No hay suficientes datos en el periodo filtrado para calcular asistencia.")
    else:
        registros_por_persona = (
            mo_f.dropna(subset=["fecha"])
            .groupby("nombre_norm")["fecha"]
            .apply(lambda s: set(s.dt.date))
        )

        filas_asist = []
        for _, row in roster_obra.iterrows():
            dias_reg = registros_por_persona.get(row["nombre_norm"], set())
            dias_reg_en_periodo = dias_reg & set(dias_periodo)
            faltantes = sorted(set(dias_periodo) - dias_reg_en_periodo)
            if len(dias_reg_en_periodo) == len(dias_periodo):
                estado = "Completo"
            elif len(dias_reg_en_periodo) == 0:
                estado = "Total"
            else:
                estado = "Parcial"
            filas_asist.append({
                "nombre": row["NombreCompleto"], "cargo": row["CargoReal"],
                "estado": estado, "dias_con_registro": len(dias_reg_en_periodo),
                "dias_totales": len(dias_periodo),
                "dias_faltantes": ", ".join(d.strftime("%a %d").capitalize() for d in faltantes),
            })
        asist = pd.DataFrame(filas_asist)

        n_completo = (asist["estado"] == "Completo").sum()
        n_parcial = (asist["estado"] == "Parcial").sum()
        n_total = (asist["estado"] == "Total").sum()

        a1, a2, a3 = st.columns(3)
        a1.metric("✅ Registro completo", n_completo)
        a2.metric("🟡 Registro parcial", n_parcial)
        a3.metric("🔴 Sin ningún registro", n_total)

        # dia con mas ausencias entre quienes registraron algo (patron tipo "miercoles")
        con_algo = asist[asist["estado"].isin(["Parcial"])]
        if not con_algo.empty:
            from collections import Counter
            conteo_dias = Counter()
            for _, row in roster_obra.iterrows():
                dias_reg = registros_por_persona.get(row["nombre_norm"], set())
                dias_reg_en_periodo = dias_reg & set(dias_periodo)
                if 0 < len(dias_reg_en_periodo) < len(dias_periodo):
                    for d in set(dias_periodo) - dias_reg_en_periodo:
                        conteo_dias[d] += 1
            if conteo_dias:
                dia_top, veces = conteo_dias.most_common(1)[0]
                st.warning(
                    f"📌 **{dia_top.strftime('%A %d').capitalize()}** es el dia que mas se repite sin "
                    f"registro entre quienes si trabajaron el resto de la semana ({veces} personas)."
                )

        with st.expander(f"🔴 Sin ningún registro ({n_total})"):
            st.dataframe(
                asist.loc[asist["estado"] == "Total", ["nombre", "cargo"]],
                width='stretch', hide_index=True,
            )
        with st.expander(f"🟡 Registro parcial ({n_parcial})"):
            st.dataframe(
                asist.loc[asist["estado"] == "Parcial", ["nombre", "cargo", "dias_con_registro", "dias_totales", "dias_faltantes"]],
                width='stretch', hide_index=True,
            )
        with st.expander(f"✅ Registro completo ({n_completo})"):
            st.dataframe(
                asist.loc[asist["estado"] == "Completo", ["nombre", "cargo"]],
                width='stretch', hide_index=True,
            )

# ============ (antes TAB MANO DE OBRA) -> continúa en Resumen semanal ============
with tab0:
    st.divider()
    # ---------- NUEVO: actividades no presupuestadas (códigos NP-<frente>-<categoría>) ----------
    st.subheader("Actividades no presupuestadas")
    np_f = mo_f[mo_f["codact"].astype(str).str.startswith("NP-")].copy()
    if np_f.empty:
        st.caption("No hay actividades no presupuestadas en el periodo filtrado.")
    else:
        np_f["cat"] = np_f["codact"].str.split("-").str[2]
        np_f["categoria"] = np_f["cat"].map(lambda c: f"{c} - {NOMBRES_CATEGORIA.get(c, c)}")
        np_f["descripcion"] = np_f["actDesc"].astype(str).str.replace("[NO PRESUPUESTADA] ", "", regex=False)
        n1, n2, n3 = st.columns(3)
        n1.metric("Costo M.O. no presupuestado", f"$ {np_f['costo_calculado'].sum():,.0f}")
        n2.metric("Horas no presupuestadas", f"{np_f['horas'].sum():,.1f} h")
        n3.metric("% del costo del periodo", f"{100 * np_f['costo_calculado'].sum() / max(mo_f['costo_calculado'].sum(), 1):.1f}%")
        st.caption("Soporte de horas y costo para sustentar mayores cantidades u obras adicionales.")
        tabla_np = (np_f.groupby(["frente_corto", "categoria", "descripcion"], as_index=False)
                    .agg(horas=("horas", "sum"), costo=("costo_calculado", "sum"), personas=("nombre", "nunique"))
                    .sort_values("costo", ascending=False)
                    .rename(columns={"frente_corto": "Frente", "categoria": "Categoría", "descripcion": "Descripción",
                                     "horas": "Horas", "costo": "Costo M.O. ($)", "personas": "Personas"}))
        st.dataframe(tabla_np.style.format({"Horas": "{:,.1f}", "Costo M.O. ($)": "${:,.0f}"}),
                     width='stretch', hide_index=True)

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

    st.subheader("📅 Resumen semanal por frente (lunes a sábado)")
    st.caption(
        "Para el corte de fin de semana: cuánto costó cada frente/intervención "
        "en cada semana laboral. Usa el filtro de fecha en la barra lateral para "
        "ver un día puntual en vez de la semana completa."
    )
    semanal_mo = (
        mo_f.dropna(subset=["semana_lunes"])
        .groupby(["semana_lunes", "semana_etiqueta", "frente", "frente_corto"], as_index=False)
        .agg(costo_semana=("costo_calculado", "sum"), horas_semana=("horas", "sum"))
        .sort_values("semana_lunes")
    )
    if semanal_mo.empty:
        st.info("Todavía no hay suficientes fechas para armar el resumen semanal.")
    else:
        orden_semanas = semanal_mo.sort_values("semana_lunes")["semana_etiqueta"].unique()
        figs = px.bar(
            semanal_mo, x="semana_etiqueta", y="costo_semana", color="frente",
            labels={"semana_etiqueta": "Semana", "costo_semana": "Costo ($)", "frente": "Frente"},
            barmode="group", color_discrete_sequence=PALETA_CATEGORICA,
        )
        figs.update_xaxes(categoryorder="array", categoryarray=orden_semanas)
        st.plotly_chart(figs, width='stretch')

        st.markdown("**Matriz costo por semana y frente** _(columnas = código de frente)_")
        matriz_mo = semanal_mo.pivot_table(
            index="semana_etiqueta", columns="frente_corto", values="costo_semana",
            aggfunc="sum", fill_value=0,
        ).reindex(orden_semanas)
        st.dataframe(matriz_mo.style.format(formato_pesos_millones), width='stretch')

        st.markdown("**Horas por semana y frente** _(columnas = código de frente)_")
        matriz_horas = semanal_mo.pivot_table(
            index="semana_etiqueta", columns="frente_corto", values="horas_semana",
            aggfunc="sum", fill_value=0,
        ).reindex(orden_semanas)
        st.dataframe(matriz_horas.style.format("{:,.1f}"), width='stretch')

    with st.expander("Detalle de registros"):
        st.dataframe(
            mo_f[["fecha", "nombre", "cargo", "frente", "horas", "tipo_hora",
                  "tarifa_hora", "costo_calculado", "semana_etiqueta"]].sort_values("fecha", ascending=False),
            width='stretch',
        )

# ============ TAB EQUIPOS ============
if MOSTRAR_EQUIPOS:
    st.metric(
        "Valor inventariado Equipos",
        f"$ {(inv['cantidad'] * inv['valorUnitario']).sum():,.0f}",
    )
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
    st.metric("Costo estimado en movimientos registrados", f"$ {mov_f['costo_estimado'].sum(skipna=True):,.0f}")

    st.markdown("**📅 Resumen semanal por frente (lunes a sábado)**")
    semanal_eq = (
        mov_f.dropna(subset=["semana_lunes", "frente"])
        .groupby(["semana_lunes", "semana_etiqueta", "frente", "frente_corto"], as_index=False)
        .agg(costo_semana=("costo_estimado", "sum"), movimientos_semana=("equipo", "count"))
        .sort_values("semana_lunes")
    )
    if semanal_eq.empty:
        st.info("Todavía no hay suficientes fechas para armar el resumen semanal de equipos.")
    else:
        figeq = px.bar(
            semanal_eq, x="semana_etiqueta", y="costo_semana", color="frente",
            labels={"semana_etiqueta": "Semana", "costo_semana": "Costo ($)", "frente": "Frente"},
            barmode="group", color_discrete_sequence=PALETA_CATEGORICA,
        )
        figeq.update_xaxes(categoryorder="array", categoryarray=sorted(semanal_eq["semana_etiqueta"].unique(),
                            key=lambda e: semanal_eq.loc[semanal_eq["semana_etiqueta"] == e, "semana_lunes"].iloc[0]))
        st.plotly_chart(figeq, width='stretch')

        st.markdown("_(columnas = código de frente)_")
        matriz_eq = semanal_eq.pivot_table(
            index="semana_etiqueta", columns="frente_corto", values="costo_semana",
            aggfunc="sum", fill_value=0,
        ).reindex(semanal_eq.sort_values("semana_lunes")["semana_etiqueta"].unique())
        st.dataframe(matriz_eq.style.format(formato_pesos_millones), width='stretch')

    st.dataframe(
        mov_f[["fecha", "solicitante", "frente", "proveedor", "equipo", "unidad",
             "cantidad", "valorUnitario", "costo_estimado", "semana_etiqueta"]].sort_values("fecha", ascending=False),
        width='stretch',
    )

# ============ TAB EFICIENCIA ============
with tab3:
    # ---------- NUEVO: rendimiento real por categoría (Mediciones × Horas) ----------
    st.subheader("Rendimiento real por categoría — Mediciones de Oficiales × Horas de Maestros")
    st.caption(
        "Cada hora registrada se asigna a una categoría (A-K) según su código SINCO "
        "(categorias_sinco.csv) y se cruza con la cantidad que midieron los Oficiales, "
        "por **semana + frente + categoría**. El rendimiento es **integral**: incluye todas "
        "las horas de la categoría (p. ej. en Concreto también el acero de refuerzo). "
        "Ojo: si hay personal sin registrar, las horas salen bajas y el rendimiento sale "
        "mejor de lo que es."
    )
    cruce = cruce_rendimiento(mo_f, med_f)
    completos = cruce[cruce["estado"] == "✅ Cruce completo"]
    if med_f.empty:
        st.info(
            "Todavía no hay mediciones de los Oficiales en el periodo/frentes filtrados. "
            "Mientras tanto, esta tabla muestra las horas-hombre que ya están listas para cruzarse."
        )
    # Resumen del periodo filtrado por frente × categoría (suma todas las semanas)
    periodo = (completos.groupby(["frente_cod", "cat", "categoria", "unidad"], as_index=False)
               .agg(HH=("HH", "sum"), costo=("costo", "sum"), cantidad=("cantidad", "sum")))
    if not periodo.empty:
        periodo["HH_por_unidad"] = periodo["HH"] / periodo["cantidad"]
        periodo["costo_por_unidad"] = periodo["costo"] / periodo["cantidad"]
        periodo = agregar_referencia(periodo)

    k1, k2, k3, k4 = st.columns(4)
    n_sem = periodo["semaforo"].value_counts() if not periodo.empty else pd.Series(dtype=int)
    k1.metric("🟢 Igual o mejor que la referencia", int(n_sem.get("🟢", 0)))
    k2.metric("🟡 Hasta 20 % por encima", int(n_sem.get("🟡", 0)))
    k3.metric("🔴 Más de 20 % por encima", int(n_sem.get("🔴", 0)))
    dif = periodo["diferencia_pesos"].sum() if not periodo.empty else 0
    k4.metric("Mayor costo (+) / ahorro (−) vs. referencia", f"$ {dif:,.0f}")
    st.caption(
        "Referencia = HH por unidad de cada frente y categoría (APU SINCO ponderado con "
        "rendimientos públicos y factor de terminal en operación). 🟢 real ≤ referencia · "
        f"🟡 hasta {int((SEMAFORO_AMARILLO - 1) * 100)} % por encima · 🔴 más que eso. "
        "Menos HH por unidad = más productivo."
    )
    k5, k6 = st.columns(2)
    k5.metric("Horas sin medición", f"{cruce.loc[cruce['estado'] == '⏳ Horas sin medición', 'HH'].sum():,.0f} h")
    k6.metric("Mediciones sin horas", int((cruce["estado"] == "⚠️ Medición sin horas").sum()))

    if not periodo.empty and periodo["ref_hh_unidad"].notna().any():
        g = periodo[periodo["ref_hh_unidad"].notna()].copy()
        g["etiqueta"] = g["semaforo"] + " " + g["frente_cod"] + " · " + g["categoria"] + " (" + g["unidad"] + ")"
        g = g.sort_values("indice")
        largo = g.melt(id_vars=["etiqueta"], value_vars=["HH_por_unidad", "ref_hh_unidad"],
                       var_name="serie", value_name="HH/unidad")
        largo["serie"] = largo["serie"].map({"HH_por_unidad": "Real", "ref_hh_unidad": "Referencia"})
        fig_ref = px.bar(
            largo, x="HH/unidad", y="etiqueta", color="serie", barmode="group", orientation="h",
            color_discrete_map={"Real": VERDE_AIA, "Referencia": GRIS_NEUTRO},
            labels={"etiqueta": "Frente · Categoría", "serie": ""},
            title="Horas-hombre por unidad: real vs. referencia (periodo filtrado)",
        )
        fig_ref.update_layout(height=max(320, 60 * len(g)), yaxis={"categoryorder": "array",
                              "categoryarray": g["etiqueta"].tolist()[::-1]})
        st.plotly_chart(fig_ref, width='stretch')

        st.markdown("**Resumen del periodo por frente y categoría**")
        tabla_p = g[["semaforo", "frente_cod", "categoria", "unidad", "HH", "cantidad", "HH_por_unidad",
                     "ref_hh_unidad", "indice", "costo_por_unidad", "costo_ref_por_unidad",
                     "diferencia_pesos"]].rename(columns={
            "semaforo": "", "frente_cod": "Frente", "categoria": "Categoría", "unidad": "Unidad",
            "HH": "Horas-hombre", "cantidad": "Cantidad medida", "HH_por_unidad": "HH/u real",
            "ref_hh_unidad": "HH/u referencia", "indice": "Real ÷ ref.", "costo_por_unidad": "$/u real",
            "costo_ref_por_unidad": "$/u referencia", "diferencia_pesos": "Mayor costo (+) / ahorro (−)"})
        st.dataframe(
            tabla_p.style.format({
                "Horas-hombre": "{:,.1f}", "Cantidad medida": "{:,.2f}", "HH/u real": "{:,.2f}",
                "HH/u referencia": "{:,.2f}", "Real ÷ ref.": "{:,.2f}", "$/u real": "${:,.0f}",
                "$/u referencia": "${:,.0f}", "Mayor costo (+) / ahorro (−)": "${:,.0f}",
            }, na_rep="—"),
            width='stretch', hide_index=True,
        )

    st.markdown("**Detalle por semana**")
    tabla = cruce[["semaforo", "semana_etiqueta", "frente_cod", "categoria", "unidad", "HH", "costo",
                   "cantidad", "HH_por_unidad", "ref_hh_unidad", "costo_por_unidad", "estado"]].rename(columns={
        "semaforo": "", "semana_etiqueta": "Semana", "frente_cod": "Frente", "categoria": "Categoría",
        "unidad": "Unidad", "HH": "Horas-hombre", "costo": "Costo M.O. ($)",
        "cantidad": "Cantidad medida", "HH_por_unidad": "HH/u real", "ref_hh_unidad": "HH/u referencia",
        "costo_por_unidad": "$ / unidad", "estado": "Estado"})
    st.dataframe(
        tabla.style.format({
            "Horas-hombre": "{:,.1f}", "Costo M.O. ($)": "${:,.0f}", "Cantidad medida": "{:,.2f}",
            "HH/u real": "{:,.2f}", "HH/u referencia": "{:,.2f}", "$ / unidad": "${:,.0f}",
        }, na_rep="—"),
        width='stretch', hide_index=True,
    )

    st.divider()
    st.subheader("Cumplimiento de productividad por actividad")
    st.caption(
        "Compara la cantidad ejecutada (reportada por el maestro) contra la cantidad "
        "esperada según el rendimiento presupuestado de Ayudantes, Oficiales y Maestros "
        "en esa actividad. **Las horas de Maestro se comparan con la tasa 'Otro'** del "
        "catálogo de rendimientos — su trabajo suele ser más de supervisión, así que "
        "esta tasa es más general que la de Ayudante/Oficial."
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

# ============ TAB COBERTURA DE REGISTRO (1-oct-2026) ============
# Quién del personal de obra (maestros, oficiales y ayudantes del Maestro de Personal) tiene
# registro completo, parcial o ninguno en la semana. Reemplaza la sección que iba en el PDF del lunes.
FESTIVOS_2026 = {"2026-01-01", "2026-01-12", "2026-03-23", "2026-04-02", "2026-04-03", "2026-05-01", "2026-05-18",
                 "2026-06-08", "2026-06-15", "2026-06-29", "2026-07-20", "2026-08-07", "2026-08-17", "2026-10-12",
                 "2026-11-02", "2026-11-16", "2026-12-08", "2026-12-25"}
with tab8:
    st.subheader("Cobertura de registro del personal de obra")
    semanas_cob = (mo.dropna(subset=["semana_lunes", "semana_etiqueta"]).drop_duplicates("semana_etiqueta")
                   .sort_values("semana_lunes", ascending=False))
    if semanas_cob.empty:
        st.info("No hay registros todavía.")
    else:
        hoy = pd.Timestamp.now(tz="America/Bogota").tz_localize(None).normalize()
        opciones = semanas_cob["semana_etiqueta"].tolist()
        lunes_list = semanas_cob["semana_lunes"].tolist()
        # por defecto la última semana ya cerrada (si la actual va empezando, la anterior)
        idx = 1 if len(opciones) > 1 and (hoy - pd.Timestamp(lunes_list[0])).days < 2 else 0
        c1, c2 = st.columns([2, 1])
        sem = c1.selectbox("Semana", opciones, index=idx)
        incluir_sab = c2.checkbox("Contar el sábado", value=True)
        lunes = pd.Timestamp(lunes_list[opciones.index(sem)])
        dias = [lunes + pd.Timedelta(days=i) for i in range(6 if incluir_sab else 5)]
        dias = [d for d in dias if d.strftime("%Y-%m-%d") not in FESTIVOS_2026 and d < hoy]
        roster = tarifas[tarifas["CargoReal"].astype(str).str.contains("Maestro|Oficial|Ayudante", case=False, na=False)]
        roster = roster[["NombreCompleto" if "NombreCompleto" in roster.columns else "NombreNormalizado", "CargoReal", "nombre_norm"]]
        roster.columns = ["Persona", "Cargo", "nombre_norm"]
        reg = mo[mo["fecha"].isin(dias)]
        hechos = reg.groupby("nombre_norm")["fecha"].apply(lambda s: set(s.dt.normalize())).to_dict()
        horas_dia = reg.groupby(["nombre_norm", reg["fecha"].dt.normalize()])["horas"].sum().to_dict()
        etiquetas = {d: ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb"][d.dayofweek] + f" {d.day}" for d in dias}
        filas = []
        for p_ in roster.itertuples():
            tiene = hechos.get(p_.nombre_norm, set())
            n = sum(1 for d in dias if d in tiene)
            estado = "Completo" if dias and n == len(dias) else ("Parcial" if n else "Sin registro")
            fila = {"Estado": estado, "Persona": p_.Persona, "Cargo": str(p_.Cargo).replace("_", " ")}
            for d in dias:
                h = horas_dia.get((p_.nombre_norm, d))
                fila[etiquetas[d]] = f"{h:g} h" if h else "—"
            fila["Días"] = f"{n}/{len(dias)}"
            filas.append(fila)
        cob = pd.DataFrame(filas)
        if cob.empty or not dias:
            st.info("Sin días hábiles cerrados en esa semana.")
        else:
            orden_est = {"Sin registro": 0, "Parcial": 1, "Completo": 2}
            cob = cob.sort_values(["Estado", "Persona"], key=lambda s: s.map(orden_est) if s.name == "Estado" else s)
            n_c, n_p, n_s = [(cob["Estado"] == e).sum() for e in ("Completo", "Parcial", "Sin registro")]
            k1, k2, k3 = st.columns(3)
            k1.metric("🟢 Registro completo", f"{n_c} de {len(cob)}")
            k2.metric("🟡 Registro parcial", n_p)
            k3.metric("🔴 Sin ningún registro", n_s)
            st.caption(f"Personal de obra del Maestro de Personal (maestros, oficiales y ayudantes) · {len(dias)} día(s) hábil(es) "
                       "ya cerrados, sin festivos. Cada casilla muestra las horas registradas ese día.")
            filtro = st.multiselect("Mostrar", ["Sin registro", "Parcial", "Completo"], default=["Sin registro", "Parcial", "Completo"])
            vista = cob[cob["Estado"].isin(filtro)]
            colores = {"Sin registro": "background-color:#FDEBEA;color:#B4453A;font-weight:600",
                       "Parcial": "background-color:#FFF3CD;color:#7A5E00;font-weight:600",
                       "Completo": "background-color:#E2EFDA;color:#1A5632;font-weight:600"}
            cols_dia = [etiquetas[d] for d in dias]
            st.dataframe(
                vista.style.map(lambda v: colores.get(v, ""), subset=["Estado"])
                .map(lambda v: "color:#B4453A" if v == "—" else "color:#1A5632", subset=cols_dia),
                width='stretch', hide_index=True, height=min(38 * (len(vista) + 1), 900))
            st.download_button("⬇️ Descargar cobertura (CSV)", cob.to_csv(index=False).encode("utf-8"),
                               file_name=f"cobertura_{lunes:%Y-%m-%d}.csv", mime="text/csv")

# ============ TAB CARGO VS. TAREA (1-oct-2026) ============
# Oficiales y maestros haciendo trabajo de ayudante (demolición, excavación, escombros, aseo,
# acarreo) y control de la jornada de supervisión de los maestros.
CAT_AYUDANTE = {"I", "J", "K", "L", "M"}
RX_AYUDANTE = re.compile(r"escombr|trasieg|botad|cargue|acarreo|realizaci[oó]n de aseo|aseo durante|orden y aseo|"
                         r"aseo general|limpieza|demolic|excavaci|llenos? con material", re.I)

def es_tarea_ayudante(codact, desc, obs):
    desc = str(desc or ""); obs = str(obs or "")
    if desc.startswith("[SUPERVISIÓN]") or obs.startswith("SUPERVISIÓN"):
        return False
    cod = str(codact or "").strip()
    cat = MAPA_CAT.get(cod)
    if not cat and cod.startswith("NP-"):
        cat = cod.split("-")[-1]
    if cat in CAT_AYUDANTE:
        return True
    return bool(RX_AYUDANTE.search(desc.split("(")[0]))

try:
    MAPA_CAT = dict(pd.read_csv(RUTA_CATEGORIAS, dtype=str)[["codact", "categoria"]].values)
except Exception:
    MAPA_CAT = {}

with tab7:
    st.subheader("Oficiales y maestros en tareas de ayudante")
    st.caption("Horas de oficiales y maestros en demolición, excavación, escombros/trasiego, aseo y acarreo. "
               "**Sobrecosto** = horas × (tarifa de la persona − tarifa promedio de ayudante) × factor del tipo de hora. "
               "La supervisión registrada por los maestros no cuenta. Respeta los filtros de la barra lateral.")
    t_ay = tarifas.loc[tarifas["CargoReal"].astype(str).str.contains("Ayudante", case=False, na=False), "TarifaHoraReal"]
    tarifa_ayud = float(t_ay[t_ay > 0].mean()) if len(t_ay) else 0.0
    cargo_real = dict(zip(tarifas["nombre_norm"], tarifas["CargoReal"].astype(str)))
    d = mo_f.copy()
    d["cargo_real"] = d["nombre_norm"].map(cargo_real).fillna(d["cargo"].astype(str))
    d = d[d["cargo_real"].str.contains("Oficial|Maestro", case=False, na=False)]
    d = d[[es_tarea_ayudante(c, a, o) for c, a, o in zip(d["codact"], d["actDesc"], d["observaciones"])]]
    d["sobrecosto"] = d["horas"] * (d["tarifa_hora"] - tarifa_ayud).clip(lower=0) * d["Factor"]
    if d.empty:
        st.success("Sin oficiales ni maestros en tareas de ayudante en el periodo filtrado.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Personas", d["nombre"].nunique())
        c2.metric("Horas en tareas de ayudante", f"{d['horas'].sum():,.0f} h")
        c3.metric("Sobrecosto vs. hacerlo con ayudantes", f"$ {d['sobrecosto'].sum():,.0f}")
        st.caption(f"Tarifa promedio de ayudante: $ {tarifa_ayud:,.0f}/h (Maestro de Personal).")
        por_p = (d.groupby(["nombre", "cargo_real"], as_index=False)
                 .agg(horas=("horas", "sum"), sobrecosto=("sobrecosto", "sum"), frentes=("codfrente", lambda s: ", ".join(sorted(set(s.astype(str)))))))
        por_p = por_p.sort_values("sobrecosto", ascending=False)
        fig_ay = px.bar(por_p.head(15).sort_values("sobrecosto"), x="sobrecosto", y="nombre", orientation="h",
                        color_discrete_sequence=[DORADO_ACENTO],
                        labels={"sobrecosto": "Sobrecosto ($)", "nombre": ""}, title="Sobrecosto por persona")
        st.plotly_chart(fig_ay, width='stretch')
        st.dataframe(por_p.rename(columns={"nombre": "Persona", "cargo_real": "Cargo", "horas": "Horas",
                                           "sobrecosto": "Sobrecosto ($)", "frentes": "Frentes"})
                     .style.format({"Horas": "{:,.1f}", "Sobrecosto ($)": "${:,.0f}"}), width='stretch', hide_index=True)
        st.markdown("**Detalle por semana y actividad**")
        det = (d.assign(actividad=d["actDesc"].astype(str).str.replace("[NO PRESUPUESTADA] ", "", regex=False).str.split("(").str[0].str.strip().str[:70])
               .groupby(["semana_etiqueta", "nombre", "codact", "actividad"], as_index=False)
               .agg(horas=("horas", "sum"), sobrecosto=("sobrecosto", "sum"))
               .sort_values(["semana_etiqueta", "sobrecosto"], ascending=[True, False]))
        st.dataframe(det.rename(columns={"semana_etiqueta": "Semana", "nombre": "Persona", "codact": "Código",
                                         "actividad": "Actividad", "horas": "Horas", "sobrecosto": "Sobrecosto ($)"})
                     .style.format({"Horas": "{:,.1f}", "Sobrecosto ($)": "${:,.0f}"}), width='stretch', hide_index=True)

    st.divider()
    st.subheader("Jornada de supervisión de los maestros")
    st.caption("Registros hechos con la opción 'Hoy solo supervisé' de la app. Su costo ya está repartido entre las actividades "
               "de la cuadrilla. Aquí se revisan horas, extras y lo que hicieron fuera del horario de la cuadrilla.")
    sup = mo_f[mo_f["observaciones"].astype(str).str.startswith("SUPERVISIÓN")].copy()
    if sup.empty:
        st.info("Todavía no hay jornadas registradas como supervisión (la opción llega con la versión nueva de la app).")
    else:
        sup["es_extra"] = sup["tipo_hora"].astype(str).str.contains("Extra")
        sup["horario"] = sup["observaciones"].str.extract(r"SUPERVISIÓN (\d{2}:\d{2}-\d{2}:\d{2})")[0]
        sup["motivo"] = sup["observaciones"].str.extract(r"Fuera del horario de la cuadrilla: (.*)$")[0]
        dia = (sup.groupby(["nombre", "fecha"], as_index=False)
               .agg(horario=("horario", "first"), horas=("horas", "sum"),
                    extras=("horas", lambda s: s[sup.loc[s.index, "es_extra"]].sum()),
                    costo=("costo_calculado", "sum"), motivo=("motivo", "first")))
        res = (dia.groupby("nombre", as_index=False)
               .agg(dias=("fecha", "nunique"), horas=("horas", "sum"), extras=("extras", "sum"), costo=("costo", "sum"),
                    dias_fuera=("motivo", lambda s: s.notna().sum())))
        st.dataframe(res.rename(columns={"nombre": "Maestro", "dias": "Días", "horas": "Horas", "extras": "Horas extra",
                                         "costo": "Costo ($)", "dias_fuera": "Días con horas fuera de la cuadrilla"})
                     .style.format({"Horas": "{:,.1f}", "Horas extra": "{:,.1f}", "Costo ($)": "${:,.0f}"}),
                     width='stretch', hide_index=True)
        st.markdown("**Detalle por día**")
        st.dataframe(dia.sort_values(["fecha", "nombre"], ascending=[False, True])
                     .rename(columns={"nombre": "Maestro", "fecha": "Fecha", "horario": "Horario", "horas": "Horas",
                                      "extras": "Extra", "costo": "Costo ($)", "motivo": "Fuera del horario de la cuadrilla"})
                     .style.format({"Horas": "{:,.1f}", "Extra": "{:,.1f}", "Costo ($)": "${:,.0f}", "Fecha": "{:%d/%m/%Y}"}),
                     width='stretch', hide_index=True)

# ============ TAB CALIDAD DE DATOS ============
if MOSTRAR_CALIDAD:
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

# ============ TAB CÓDIGOS POR PERSONA ============
with tab5:
    st.subheader("Códigos de actividad por persona")
    st.caption(
        "Selecciona una persona para ver el resumen de horas por frente y "
        "actividad — listos para asignar en el programa de nómina. Respeta "
        "los filtros de la barra lateral (frente, semana, fecha), así que "
        "puedes acotar primero al periodo de nómina que necesites."
    )

    nombres_disp = sorted(mo_f["nombre"].dropna().unique())
    if not nombres_disp:
        st.info("No hay registros para los filtros seleccionados en la barra lateral.")
    else:
        persona_sel = st.selectbox("Persona", nombres_disp)
        datos_persona = mo_f[mo_f["nombre"] == persona_sel]

        c1, c2, c3 = st.columns(3)
        c1.metric("Total horas", f"{datos_persona['horas'].sum():,.1f} h")
        c2.metric("Frentes distintos", datos_persona["frente_corto"].nunique())
        c3.metric("Actividades distintas", datos_persona["codact"].nunique())

        # --- Horas extra / recargo por tipo de hora ---
        horas_ordinaria = datos_persona.loc[
            datos_persona["tipo_hora"] == "Ordinaria Diurna", "horas"
        ].sum()
        horas_extra_total = datos_persona["horas"].sum() - horas_ordinaria

        c4, c5 = st.columns(2)
        c4.metric("Horas extra / recargo", f"{horas_extra_total:,.1f} h")
        pct_extra = (horas_extra_total / datos_persona["horas"].sum() * 100) if datos_persona["horas"].sum() else 0
        c5.metric("% del total en recargo", f"{pct_extra:,.0f}%")

        st.markdown("**Desglose por tipo de hora**")
        resumen_tipo_hora = (
            datos_persona
            .groupby("tipo_hora", as_index=False)
            .agg(horas=("horas", "sum"), costo=("costo_calculado", "sum"), registros=("id", "count"))
            .sort_values("horas", ascending=False)
            .rename(columns={
                "tipo_hora": "Tipo de Hora", "horas": "Horas",
                "costo": "Costo estimado", "registros": "N° registros",
            })
        )
        st.dataframe(
            resumen_tipo_hora.style.format({"Costo estimado": "${:,.0f}"}),
            width='stretch', hide_index=True,
        )
        if horas_extra_total > 0:
            tipos_extra = sorted(
                datos_persona.loc[datos_persona["tipo_hora"] != "Ordinaria Diurna", "tipo_hora"].unique()
            )
            st.caption("Tipos de recargo presentes: " + ", ".join(tipos_extra))

        resumen_persona = (
            datos_persona
            .groupby(["frente_corto", "frente", "codact", "actDesc"], as_index=False)
            .agg(horas=("horas", "sum"), registros=("id", "count"))
            .sort_values(["frente_corto", "codact"])
            .rename(columns={
                "frente_corto": "Frente", "frente": "Nombre completo del frente", "codact": "Cód. Actividad",
                "actDesc": "Actividad", "horas": "Horas", "registros": "N° registros",
            })
        )
        st.dataframe(resumen_persona, width='stretch', hide_index=True)

        st.download_button(
            "⬇️ Descargar este resumen (CSV)",
            resumen_persona.to_csv(index=False).encode("utf-8"),
            file_name=f"codigos_{persona_sel.replace(' ', '_')}.csv",
            mime="text/csv",
        )

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
import numpy as np
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


DIAS_ES = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
DIAS_ES_LARGO = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
COLS_ASIST = {"nombre": "Nombre", "cargo": "Cargo", "dias_con_registro": "Días con registro",
              "dias_totales": "Días del periodo", "dias_faltantes": "Días sin registro"}


def dia_es(d, largo=False):
    """Día en español: 'Mar 29' o 'Martes 29' (no depende del idioma del servidor)."""
    return f"{(DIAS_ES_LARGO if largo else DIAS_ES)[d.weekday()]} {d.day:02d}"


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
    vista = st.radio("Vista", ["Gerencia", "Control interno"], horizontal=True,
                     help="Gerencia: resumen, alertas y presupuesto. Control interno: cobertura de registro, códigos por persona, productividad y cargo vs. tarea.")
    st.divider()
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

VISTA_GERENCIA = (vista == "Gerencia")

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
# ============ ALERTAS DE LA SEMANA (7-oct-2026) ============
# Parámetros (proyección de la A, MO proyectada SINCO, umbrales, residentes) en Drive:
URL_PARAMS = (
    "https://drive.google.com/uc?export=download&id="
    "1c31oKbuKnXO2XV2AC3n4VOqhQmxm6FTr"
)
INC_HE = {  # horas equivalentes de recargo por tipo de hora (base para el % de horas extra)
    "Ordinaria Diurna": 0.0, "Recargo Nocturno (Ordinaria)": 0.35, "Extra Diurna": 1.25, "Extra Nocturna": 1.75,
    "Dominical/Festivo Ordinaria Diurna": 1.9, "Dominical/Festivo Extra Diurna": 2.15, "Dominical/Festivo Extra Nocturna": 2.65,
}


@st.cache_data(ttl=600)
def cargar_parametros():
    x = pd.read_excel(URL_PARAMS, sheet_name=None, dtype={"Código": str, "Cód. frente": str})
    pa = x["Proyeccion_A"].copy()
    mop = x["MO_proyectada"].copy()
    um = x["Umbrales"]
    rs = x["Residentes"].copy()
    U = dict(zip(um["Parámetro"].astype(str), um["Valor"]))
    for k, v in list(U.items()):
        try:
            U[k] = float(v)
        except (TypeError, ValueError):
            U[k] = str(v)
    pa["nombre_norm"] = pa["Nombre en la app (Maestro de Personal)"].fillna("").apply(normalizar_nombre)
    for c in ["Inicio", "Fin"]:
        pa[c] = pd.to_datetime(pa[c], errors="coerce")
    mop["Código"] = mop["Código"].astype(str).str.strip()
    rs["Cód. frente"] = rs["Cód. frente"].astype(str).str.strip().str.zfill(2)
    return pa, mop, U, rs


URL_SEGUIMIENTO = (
    "https://docs.google.com/spreadsheets/d/"
    "1Sbt-r9yR_pcyluP_JLX9nA2nwxpW3cU2ATVPMwvNP-I/gviz/tq?tqx=out:csv&headers=1&sheet=Alertas_Seguimiento"
)
COLS_SEG = ["Semana (lunes)", "Tipo de alerta", "Frente", "Persona o código", "Descripción de la alerta", "Plan de mejora",
            "Responsable", "Fecha compromiso", "Estado", "Resultado / comentario"]


@st.cache_data(ttl=60)
def cargar_seguimiento():
    try:
        sg = pd.read_csv(URL_SEGUIMIENTO, dtype=str)
    except Exception:
        return pd.DataFrame(columns=COLS_SEG + ["tipo_n", "ref", "lunes"])
    if not set(COLS_SEG[:4] + ["Estado"]).issubset(sg.columns):   # la pestaña aún no existe
        return pd.DataFrame(columns=COLS_SEG + ["tipo_n", "ref", "lunes"])
    sg = sg.dropna(subset=["Tipo de alerta"]).copy()
    sg["tipo_n"] = sg["Tipo de alerta"].astype(str).str.extract(r"^(\d)")[0]
    sg["ref"] = sg["Persona o código"].fillna("").apply(normalizar_nombre)
    sg["lunes"] = pd.to_datetime(sg["Semana (lunes)"], errors="coerce").dt.normalize()
    sg["Fecha compromiso"] = pd.to_datetime(sg["Fecha compromiso"], errors="coerce")
    sg["Estado"] = sg["Estado"].fillna("Abierta")
    return sg


def marcar_plan(df, tipo_n, col_ref, sg, lunes):
    """Columna 'Plan de mejora': estado del plan de esta semana o, si no hay, el plan abierto más reciente de semanas anteriores."""
    df = df.copy()
    s_ = sg[(sg["tipo_n"] == str(tipo_n)) & (sg["lunes"] <= lunes)].sort_values("lunes")
    est = {}
    for _, r in s_.iterrows():
        if r["lunes"] == lunes:
            est[r["ref"]] = r["Estado"]
        elif r["Estado"] != "Cerrada":
            est[r["ref"]] = f"{r['Estado']} desde {r['lunes']:%d-%b}"
    df["Plan de mejora"] = df[col_ref].fillna("").astype(str).apply(normalizar_nombre).map(est).fillna("⚪ Sin plan")
    return df


def semaforo_txt(valor, amarillo, rojo):
    if pd.isna(valor):
        return ""
    return "🔴 Rojo" if valor > rojo else ("🟡 Amarillo" if valor > amarillo else "🟢 Verde")


def mostrar_alertas(mo_base, tarifas):
    st.subheader("🚨 Alertas de la semana")
    try:
        pa, mop, U, rs = cargar_parametros()
    except Exception as e:
        st.error(f"No pude leer Parametros_Alertas.xlsx en Drive ({e}). Revisa que siga compartido con 'Cualquier persona con el enlace'.")
        return
    RES = dict(zip(rs["Cód. frente"], rs["Residente(s)"].fillna("—")))
    d = mo_base.dropna(subset=["fecha"]).copy()
    if d.empty:
        st.info("No hay registros para los frentes seleccionados.")
        return
    d["lunes"] = (d["fecha"] - pd.to_timedelta(d["fecha"].dt.dayofweek, unit="D")).dt.normalize()
    d["cod_fr"] = pd.to_numeric(d["codfrente"], errors="coerce").fillna(0).astype(int).astype(str).str.zfill(2)
    d["residente_fr"] = d["cod_fr"].map(RES).fillna("—")
    cargo_real = dict(zip(tarifas["nombre_norm"], tarifas["CargoReal"].astype(str)))
    d["cargo_real"] = d["nombre_norm"].map(cargo_real).fillna(d["cargo"].astype(str))
    d["es_maestro"] = d["cargo_real"].str.upper().str.startswith("MAESTRO")
    d["heq"] = d["horas"] * d["tipo_hora"].map(INC_HE).fillna(0)
    desc = d["actDesc"].fillna("").astype(str).str.upper()
    d["campamento"] = (desc.str.contains(str(U.get("palabras_campamento", "CAMPAMENT")), regex=True)
                       & ~desc.str.contains(str(U.get("palabras_no_campamento", "TRASLAD|RETIR|DESMONT|TRASIEG|ACARRE")), regex=True)
                       & ~d["es_maestro"])
    d["codact"] = d["codact"].astype(str).str.strip()

    lunes_disp = sorted(d["lunes"].unique())
    hoy = pd.Timestamp.now(tz="America/Bogota").tz_localize(None).normalize()
    completas = [l for l in lunes_disp if pd.Timestamp(l) + pd.Timedelta(days=6) < hoy]
    defecto = completas[-1] if completas else lunes_disp[-1]
    etiq = {l: f"{pd.Timestamp(l):%d-%b} a {pd.Timestamp(l) + pd.Timedelta(days=6):%d-%b-%Y}" for l in lunes_disp}
    sem = st.selectbox("Semana (lunes a domingo)", lunes_disp[::-1], index=lunes_disp[::-1].index(defecto),
                       format_func=lambda l: etiq[l], key="sem_alertas")
    ini, fin = pd.Timestamp(sem), pd.Timestamp(sem) + pd.Timedelta(days=6)
    w = d[(d["fecha"] >= ini) & (d["fecha"] <= fin)]
    hasta = d[d["fecha"] <= fin]
    st.caption("Valores con la tarifa del Maestro de Personal (factor 1,4914). Umbrales y proyección: archivo Parametros_Alertas en Drive.")

    # 1. % de horas extra por persona
    base_sem = U.get("horas_base_semana", 49)
    g = w.groupby("nombre_norm")
    per = pd.DataFrame({
        "Nombre": g["nombre"].first(), "Cargo": g["cargo_real"].first(),
        "Días": g["fecha"].apply(lambda s: s[s.dt.dayofweek < 6].dt.date.nunique()),
        "Horas": g["horas"].sum(), "Horas extra/recargo": g.apply(lambda t: t.loc[t["heq"] > 0, "horas"].sum()),
        "heq": g["heq"].sum(), "tarifa": g["tarifa_hora"].first(),
        "Frente principal": g.apply(lambda t: t.groupby("cod_fr")["horas"].sum().idxmax()),
    })
    per["Residente"] = per["Frente principal"].map(RES).fillna("—")
    per["base"] = base_sem * (per["Días"] / 6).clip(upper=1)
    per["% HE"] = (per["heq"] / per["base"]).where(per["base"] > 0)
    per["Semáforo"] = per["% HE"].apply(lambda v: semaforo_txt(v, U.get("he_amarillo", .4), U.get("he_rojo", .45)))
    per["Exceso sobre lo proyectado ($)"] = ((per["% HE"] - U.get("he_amarillo", .4)).clip(lower=0) * per["base"] * per["tarifa"]).fillna(0)
    he_alert = per[per["% HE"] > U.get("he_amarillo", .4)].sort_values("% HE", ascending=False)
    sg = cargar_seguimiento()
    he_alert = marcar_plan(he_alert, 1, "Nombre", sg, ini)

    # 2. Jornadas excesivas
    dia = w.groupby(["nombre", "fecha"])["horas"].sum().reset_index()
    dia_ex = dia[dia["horas"] > U.get("horas_dia_max", 12)]
    semh = w.groupby("nombre")["horas"].sum()
    sem_ex = semh[semh > U.get("horas_semana_max", 60)]

    fr_persona = w.groupby(["nombre_norm", "cod_fr"])["horas"].sum().reset_index().sort_values("horas").drop_duplicates("nombre_norm", keep="last").set_index("nombre_norm")["cod_fr"]

    # 3. A de maestros vs. proyección
    fac = 7 / 30
    activos = pa[(pa["Inicio"].isna() | (pa["Inicio"] <= fin)) & (pa["Fin"].isna() | (pa["Fin"] >= ini))]
    pm = activos[activos["Cargo en la proyección"].astype(str).str.upper().str.startswith("MAESTRO") & (activos["¿Proyectado en A?"] == "Sí")]
    real_m = w[w["es_maestro"]].groupby("nombre_norm").agg(Nombre=("nombre", "first"), Real=("costo_calculado", "sum"))
    proy_m = pm.assign(P=pm["Salario total mes proyectado (con HE)"] * fac).groupby("nombre_norm")["P"].sum()
    a_m = real_m.join(proy_m.rename("Proyectado"), how="outer")
    vac = a_m.index == ""
    a_m.loc[vac, "Nombre"] = f"Vacantes de maestro sin persona ({int((pm['nombre_norm'] == '').sum())})"
    a_m["Nombre"] = a_m["Nombre"].fillna(pd.Series(dict(zip(pm["nombre_norm"], pm["Nombre en la proyección"]))))
    a_m = a_m.fillna({"Real": 0, "Proyectado": 0})
    a_m["Diferencia (real − proy.)"] = a_m["Real"] - a_m["Proyectado"]
    a_m["Frente"] = a_m.index.map(fr_persona).fillna("—")
    a_m["Residente"] = a_m["Frente"].map(RES).fillna("—")
    a_m["Lectura"] = np.where(a_m["Proyectado"] == 0, "Hace de maestro pero no está proyectado como maestro en la A",
                     np.where(a_m["Real"] == 0, "Proyectado sin registros esta semana", ""))
    a_real, a_proy = a_m["Real"].sum(), a_m["Proyectado"].sum()

    # 4. Oficiales / ayudantes proyectados en A  y  5. campamento
    po = activos[~activos["Cargo en la proyección"].astype(str).str.upper().str.startswith("MAESTRO")
                 & (activos["¿Proyectado en A?"] == "Sí") & (activos["nombre_norm"] != "")]
    wo = w[w["nombre_norm"].isin(po["nombre_norm"])]
    ofi = pd.DataFrame({
        "Nombre": wo.groupby("nombre_norm")["nombre"].first(),
        "A (campamento)": wo[wo["campamento"]].groupby("nombre_norm")["costo_calculado"].sum(),
        "APU (actividades)": wo[~wo["campamento"]].groupby("nombre_norm")["costo_calculado"].sum(),
    }).join(po.assign(P=po["Salario total mes proyectado (con HE)"] * fac).groupby("nombre_norm")["P"].sum().rename("Proyectado en A"), how="outer").fillna(0)
    ofi["Nombre"] = ofi["Nombre"].replace(0, np.nan).fillna(pd.Series(dict(zip(po["nombre_norm"], po["Nombre en la proyección"]))))
    ofi["A sobrestimada"] = ofi["Proyectado en A"] - ofi["A (campamento)"]
    ofi["Frente"] = ofi.index.map(fr_persona).fillna("—")
    ofi["Residente"] = ofi["Frente"].map(RES).fillna("—")
    camp = w[w["campamento"]].groupby(["codact", "actDesc", "cod_fr"]).agg(Horas=("horas", "sum"), Personas=("nombre", "nunique"), Costo=("costo_calculado", "sum")).reset_index()
    camp["Residente"] = camp["cod_fr"].map(RES).fillna("—")

    # 6. Sobrecosto de MO (acumulado hasta la semana) y posible doble pago
    acum = hasta.groupby("codact").agg(Costo_acum=("costo_calculado", "sum"))
    act_sem = w.groupby("codact").agg(Costo_sem=("costo_calculado", "sum"), Horas_sem=("horas", "sum"))
    c = act_sem.join(acum).join(mop.set_index("Código")[["Descripción", "Proy. mano de obra", "Proy. subcontratos"]], how="left")
    c = c[~c.index.str.startswith(("SUP-", "NP-"))]
    c["Frente"] = c.index.str[:2]
    c["Residente"] = c["Frente"].map(RES).fillna("—")
    c["% MO consumida"] = (c["Costo_acum"] / c["Proy. mano de obra"]).where(c["Proy. mano de obra"] > 0)
    sobre = c[c["% MO consumida"] >= U.get("mo_amarillo", .8)].copy()
    sobre["Semáforo"] = sobre["% MO consumida"].apply(lambda v: semaforo_txt(v, U.get("mo_amarillo", .8) - 1e-9, U.get("mo_rojo", 1.0)))
    sobre["Exceso ($)"] = (sobre["Costo_acum"] - sobre["Proy. mano de obra"]).clip(lower=0)
    doble = c[(c["Proy. mano de obra"].fillna(0) == 0) & (c["Proy. subcontratos"].fillna(0) > 0)].copy()

    # 7. No presupuestadas
    npw = w[w["codact"].str.startswith("NP-")].groupby(["codact", "actDesc", "cod_fr"]).agg(Horas=("horas", "sum"), Costo=("costo_calculado", "sum")).reset_index()
    npw["Residente"] = npw["cod_fr"].map(RES).fillna("—")

    # ---------- Resumen ----------
    k = st.columns(6)
    cop = lambda v: f"$ {v:,.0f}".replace(",", ".")
    dif_a = a_real - a_proy
    tarjetas = [
        ("Personas sobre 40 % HE", f"{len(he_alert)}", f"🔴 {cop(he_alert['Exceso sobre lo proyectado ($)'].sum())} por encima del 40 % proyectado"),
        ("Jornadas excesivas", f"{len(dia_ex) + len(sem_ex)}", f"{len(dia_ex)} días de más de {U.get('horas_dia_max', 12):.0f} h · {len(sem_ex)} personas con más de {U.get('horas_semana_max', 60):.0f} h/semana"),
        ("A maestros (real / proy.)", f"{(a_real / a_proy if a_proy else 0):.0%}",
         (f"🔴 {cop(dif_a)} por encima de lo proyectado" if dif_a > 0 else f"{cop(-dif_a)} por debajo de lo proyectado (ver detalle: vacantes / sin registro)")),
        ("Actividades con MO ≥ 80 %", f"{len(sobre)}", f"🔴 {cop(sobre['Exceso ($)'].sum())} ya por encima de la MO proyectada"),
        ("Posible doble pago", cop(doble['Costo_sem'].sum()), f"MO propia en {len(doble)} actividades con subcontrato"),
        ("No presupuestadas", cop(npw['Costo'].sum()), f"{npw['codact'].nunique()} códigos · candidatas a adicional"),
    ]
    k = st.columns(6)
    for col_, (lab, val, nota) in zip(k, tarjetas):
        col_.metric(lab, val)
        col_.caption(nota)

    fmt_p = {"% HE": "{:.0%}", "Horas": "{:,.1f}", "Horas extra/recargo": "{:,.1f}", "Exceso sobre lo proyectado ($)": "$ {:,.0f}"}
    with st.expander(f"1 · Horas extra por persona — {len(he_alert)} sobre {U.get('he_amarillo', .4):.0%} (proyectado en la A: 40 %)", expanded=True):
        st.dataframe(he_alert[["Nombre", "Cargo", "Frente principal", "Residente", "Días", "Horas", "Horas extra/recargo", "% HE", "Semáforo", "Exceso sobre lo proyectado ($)", "Plan de mejora"]]
                     .style.format(fmt_p), hide_index=True, width='stretch')
        st.caption("% HE = horas equivalentes de recargo (extra diurna 1,25; nocturna 1,75; recargo nocturno 0,35; dominical 1,90/2,15/2,65) ÷ 49 h base de la semana.")
    with st.expander(f"2 · Jornadas excesivas — {len(dia_ex)} días de más de {U.get('horas_dia_max', 12):.0f} h, {len(sem_ex)} personas con más de {U.get('horas_semana_max', 60):.0f} h"):
        c1, c2 = st.columns(2)
        c1.dataframe(marcar_plan(dia_ex.rename(columns={"nombre": "Nombre", "fecha": "Fecha", "horas": "Horas"}), 2, "Nombre", sg, ini).sort_values("Horas", ascending=False)
                     .style.format({"Horas": "{:,.1f}", "Fecha": "{:%d/%m/%Y}"}), hide_index=True, width='stretch')
        c2.dataframe(marcar_plan(sem_ex.rename("Horas en la semana").reset_index().rename(columns={"nombre": "Nombre"}), 2, "Nombre", sg, ini).sort_values("Horas en la semana", ascending=False)
                     .style.format({"Horas en la semana": "{:,.1f}"}), hide_index=True, width='stretch')
    with st.expander(f"3 · Administración (A) de maestros — real \\$ {a_real:,.0f} vs proyectado \\$ {a_proy:,.0f}"):
        st.dataframe(marcar_plan(a_m.reset_index(drop=True), 3, "Nombre", sg, ini)[["Nombre", "Frente", "Residente", "Proyectado", "Real", "Diferencia (real − proy.)", "Lectura", "Plan de mejora"]]
                     .sort_values("Diferencia (real − proy.)", ascending=False)
                     .style.format({"Proyectado": "$ {:,.0f}", "Real": "$ {:,.0f}", "Diferencia (real − proy.)": "$ {:,.0f}"}), hide_index=True, width='stretch')
        st.caption("Proyectado de la semana = salario total mes proyectado (con 40 % HE) × 7/30. Regla: todo lo que registran los maestros va a la A.")
    with st.expander(f"4 · Oficiales y ayudantes proyectados en la A — A sobrestimada \\$ {ofi['A sobrestimada'].sum():,.0f} esta semana"):
        st.dataframe(marcar_plan(ofi.reset_index(drop=True), 4, "Nombre", sg, ini)[["Nombre", "Frente", "Residente", "Proyectado en A", "A (campamento)", "APU (actividades)", "A sobrestimada", "Plan de mejora"]]
                     .sort_values("A sobrestimada", ascending=False)
                     .style.format({c_: "$ {:,.0f}" for c_ in ["Proyectado en A", "A (campamento)", "APU (actividades)", "A sobrestimada"]}), hide_index=True, width='stretch')
        st.caption("Por regla, la A solo asume el trabajo de oficiales y ayudantes DENTRO del campamento; el resto se carga al APU de la actividad.")
    with st.expander(f"5 · Trabajo dentro del campamento (va a la A) — \\$ {camp['Costo'].sum():,.0f}"):
        st.dataframe(marcar_plan(camp.rename(columns={"codact": "Código", "actDesc": "Actividad", "cod_fr": "Frente"}), 5, "Código", sg, ini)
                     .style.format({"Horas": "{:,.1f}", "Costo": "$ {:,.0f}"}), hide_index=True, width='stretch')
    with st.expander(f"6 · Sobrecosto de mano de obra — {len(sobre)} actividades con 80 % o más de la MO proyectada consumida"):
        st.dataframe(marcar_plan(sobre.reset_index().rename(columns={"codact": "Código", "Costo_sem": "Costo semana", "Costo_acum": "Costo acumulado"}), 6, "Código", sg, ini)
                     [["Código", "Descripción", "Frente", "Residente", "Proy. mano de obra", "Costo acumulado", "% MO consumida", "Semáforo", "Exceso ($)", "Costo semana", "Plan de mejora"]]
                     .sort_values("% MO consumida", ascending=False)
                     .style.format({"Proy. mano de obra": "$ {:,.0f}", "Costo acumulado": "$ {:,.0f}", "% MO consumida": "{:.0%}", "Exceso ($)": "$ {:,.0f}", "Costo semana": "$ {:,.0f}"}),
                     hide_index=True, width='stretch')
        st.caption("Ojo: muchas alertas aquí vienen de códigos mal asignados en la app; revisar con el maestro antes de concluir sobrecosto.")
    with st.expander(f"7 · Posible doble pago — MO propia en actividades contratadas por subcontrato: \\$ {doble['Costo_sem'].sum():,.0f}"):
        st.dataframe(marcar_plan(doble.reset_index().rename(columns={"codact": "Código", "Costo_sem": "MO propia semana", "Horas_sem": "Horas semana"}), 7, "Código", sg, ini)
                     [["Código", "Descripción", "Frente", "Residente", "Proy. subcontratos", "Horas semana", "MO propia semana", "Plan de mejora"]].sort_values("MO propia semana", ascending=False)
                     .style.format({"Proy. subcontratos": "$ {:,.0f}", "Horas semana": "{:,.1f}", "MO propia semana": "$ {:,.0f}"}), hide_index=True, width='stretch')
    with st.expander(f"8 · No presupuestadas de la semana — \\$ {npw['Costo'].sum():,.0f} (candidatas a cobro como adicional)"):
        st.dataframe(marcar_plan(npw.rename(columns={"codact": "Código", "actDesc": "Actividad", "cod_fr": "Frente"}), 8, "Código", sg, ini).sort_values("Costo", ascending=False)
                     .style.format({"Horas": "{:,.1f}", "Costo": "$ {:,.0f}"}), hide_index=True, width='stretch')

    # ---------- Informe semanal para enviar (lo administra control de costos) ----------
    st.divider()
    st.markdown("#### 📤 Informe semanal para gerencia")
    rango = f"{ini:%d-%b} a {fin:%d-%b-%Y}"
    def top(df, col_n, col_v, fmt, n=3):
        if df.empty:
            return "ninguna"
        x = df.sort_values(col_v, ascending=False).head(n)
        return "; ".join(f"{r[col_n]} ({fmt(r[col_v])})" for _, r in x.iterrows())
    pesos = lambda v: f"$ {v:,.0f}".replace(",", ".")
    texto = (
        f"Asunto: Alertas de control de costos AIRPLAN T1 JMC — semana {rango}\n\n"
        f"Buen día. Comparto las alertas de la semana {rango} (detalle en el dashboard, vista Gerencia > Alertas de la semana, y en el Excel adjunto):\n\n"
        f"1. Horas extra: {len(he_alert)} personas superaron el 40 % proyectado; exceso de {pesos(he_alert['Exceso sobre lo proyectado ($)'].sum())}. "
        f"Mayores: {top(he_alert, 'Nombre', '% HE', lambda v: f'{v:.0%}')}.\n"
        f"2. Jornadas excesivas: {len(dia_ex)} días de más de {U.get('horas_dia_max', 12):.0f} h y {len(sem_ex)} personas con más de {U.get('horas_semana_max', 60):.0f} h en la semana.\n"
        f"3. Administración (A) de maestros: real {pesos(a_real)} vs proyectado {pesos(a_proy)} ({(a_real / a_proy if a_proy else 0):.0%}).\n"
        f"4. Oficiales y ayudantes proyectados en la A: {pesos(ofi['A sobrestimada'].sum())} de la A no se consumieron porque su trabajo fue al APU.\n"
        f"5. Trabajo dentro del campamento (cargado a la A): {pesos(camp['Costo'].sum())}.\n"
        f"6. Sobrecosto de mano de obra: {len(sobre)} actividades con 80 % o más de la MO proyectada consumida; {pesos(sobre['Exceso ($)'].sum())} por encima. "
        f"Mayores: {top(sobre.reset_index(), 'codact', '% MO consumida', lambda v: f'{v:.0%}')}.\n"
        f"7. Posible doble pago (MO propia en actividades con subcontrato): {pesos(doble['Costo_sem'].sum())} en {len(doble)} actividades.\n"
        f"8. No presupuestadas: {pesos(npw['Costo'].sum())} en {npw['codact'].nunique()} códigos (candidatas a cobro como adicional).\n\n"
        f"Planes de mejora: {int((sg['lunes'] == ini).sum()) if not sg.empty else 0} registrados para esta semana; "
        f"{int(((sg['Estado'] != 'Cerrada') & sg['Fecha compromiso'].notna() & (sg['Fecha compromiso'] < hoy)).sum()) if not sg.empty else 0} vencidos.\n\n"
        f"Quedo atento a los planes de cada frente en la pestaña Alertas_Seguimiento."
    )
    st.caption("Texto listo para pegar en el correo (revísalo y ajústalo antes de enviarlo):")
    st.text_area("Texto del correo", texto, height=330, label_visibility="collapsed", key="texto_informe")
    import io
    buf = io.BytesIO()
    hojas = {
        "Horas extra": he_alert[["Nombre", "Cargo", "Frente principal", "Residente", "Días", "Horas", "Horas extra/recargo", "% HE", "Semáforo", "Exceso sobre lo proyectado ($)", "Plan de mejora"]],
        "Jornadas >12h": dia_ex.rename(columns={"nombre": "Nombre", "fecha": "Fecha", "horas": "Horas"}),
        "Semanas >60h": sem_ex.rename("Horas en la semana").reset_index().rename(columns={"nombre": "Nombre"}),
        "A maestros": a_m.reset_index(drop=True)[["Nombre", "Frente", "Residente", "Proyectado", "Real", "Diferencia (real − proy.)", "Lectura"]],
        "Of-ayud en A": ofi.reset_index(drop=True)[["Nombre", "Frente", "Residente", "Proyectado en A", "A (campamento)", "APU (actividades)", "A sobrestimada"]],
        "Campamento": camp.rename(columns={"codact": "Código", "actDesc": "Actividad", "cod_fr": "Frente"}),
        "Sobrecosto MO": sobre.reset_index().rename(columns={"codact": "Código", "Costo_sem": "Costo semana", "Costo_acum": "Costo acumulado"})[["Código", "Descripción", "Frente", "Residente", "Proy. mano de obra", "Costo acumulado", "% MO consumida", "Semáforo", "Exceso ($)", "Costo semana"]],
        "Doble pago": doble.reset_index().rename(columns={"codact": "Código", "Costo_sem": "MO propia semana", "Horas_sem": "Horas semana"})[["Código", "Descripción", "Frente", "Residente", "Proy. subcontratos", "Horas semana", "MO propia semana"]],
        "No presupuestadas": npw.rename(columns={"codact": "Código", "actDesc": "Actividad", "cod_fr": "Frente"}),
        "Planes de mejora": sg[[c_ for c_ in COLS_SEG if c_ in sg.columns]] if not sg.empty else pd.DataFrame(columns=COLS_SEG),
    }
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({"Informe": texto.split("\n")}).to_excel(xw, sheet_name="Resumen", index=False)
        for nom, df_ in hojas.items():
            df_.to_excel(xw, sheet_name=nom[:31], index=False)
        for ws_ in xw.book.worksheets:
            for col in ws_.columns:
                ws_.column_dimensions[col[0].column_letter].width = min(60, max(10, max(len(str(c_.value or "")) for c_ in col[:50]) + 2))
    st.download_button("⬇️ Descargar informe de la semana (Excel)", buf.getvalue(),
                       file_name=f"Alertas_semana_{ini:%Y-%m-%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_informe")

    # ---------- 9. Planes de mejora (seguimiento antes / ahora) ----------
    def indicadores(l0):
        """Valor de cada alerta por persona/código en la semana que empieza el lunes l0 (para comparar antes vs. ahora)."""
        a_, b_ = pd.Timestamp(l0), pd.Timestamp(l0) + pd.Timedelta(days=6)
        ww = d[(d["fecha"] >= a_) & (d["fecha"] <= b_)]
        out = {}
        if ww.empty:
            return out
        gg = ww.groupby("nombre_norm")
        dias = gg["fecha"].apply(lambda x: x[x.dt.dayofweek < 6].dt.date.nunique())
        base = base_sem * (dias / 6).clip(upper=1)
        for nn_, v in (gg["heq"].sum() / base).items():
            out[("1", nn_)] = v
        for nn_, v in gg["horas"].sum().items():
            out[("2", nn_)] = v
        pm_all = pa[pa["Cargo en la proyección"].astype(str).str.upper().str.startswith("MAESTRO") & (pa["¿Proyectado en A?"] == "Sí")]
        pp = pm_all.groupby("nombre_norm")["Salario total mes proyectado (con HE)"].sum() * fac
        for nn_, v in ww[ww["es_maestro"]].groupby("nombre_norm")["costo_calculado"].sum().items():
            out[("3", nn_)] = v / pp[nn_] if pp.get(nn_, 0) else np.nan
        for nn_, v in ww[ww["nombre_norm"].isin(po["nombre_norm"])].groupby("nombre_norm")["costo_calculado"].sum().items():
            out[("4", nn_)] = v
        cod = ww.copy()
        cod["ref"] = cod["codact"].apply(normalizar_nombre)
        for t_, m_ in (("5", cod["campamento"]), ("7", pd.Series(True, index=cod.index)), ("8", cod["codact"].str.startswith("NP-"))):
            for r_, v in cod[m_].groupby("ref")["costo_calculado"].sum().items():
                out[(t_, r_)] = v
        acum_ = d[d["fecha"] <= b_].groupby("codact")["costo_calculado"].sum()
        pmo = mop.set_index("Código")["Proy. mano de obra"]
        for c_, v in acum_.items():
            if pmo.get(c_, 0) and pmo.get(c_, 0) > 0:
                out[("6", normalizar_nombre(c_))] = v / pmo[c_]
        out["_sem6"] = cod.groupby("ref")["costo_calculado"].sum().to_dict()
        return out

    FMT_T = {"1": lambda v: f"{v:.0%}", "2": lambda v: f"{v:,.1f} h", "3": lambda v: f"{v:.0%} de lo proy.", "4": lambda v: cop(v),
             "5": lambda v: cop(v), "6": lambda v: f"{v:.0%} consumido", "7": lambda v: cop(v), "8": lambda v: cop(v)}

    def resultado(t_, antes, ahora, ref_, ind_now):
        if t_ == "6":
            nuevo = ind_now.get("_sem6", {}).get(ref_, 0)
            return "🟢 Sin consumo nuevo esta semana" if nuevo == 0 else f"🔴 Sigue consumiendo ({cop(nuevo)} esta semana)"
        if pd.isna(ahora):
            return "🟢 No aparece esta semana — se puede cerrar" if t_ in ("1", "2", "5", "7", "8") else "⚪ Sin registros esta semana"
        if t_ == "1" and ahora <= U.get("he_amarillo", .4):
            return "🟢 Mejoró, bajo el 40 % — se puede cerrar"
        if t_ == "2" and ahora <= U.get("horas_semana_max", 60):
            return "🟢 Mejoró, bajo el límite — se puede cerrar"
        if pd.isna(antes) or antes == 0:
            return "⚪ Sin dato de la semana del plan"
        r_ = ahora / antes
        return "🟢 Mejoró" if r_ < 0.9 else ("🔴 Empeoró" if r_ > 1.1 else "🟡 Sigue igual")

    st.divider()
    st.markdown("#### 📝 Planes de mejora — seguimiento")
    if sg.empty:
        st.info("Aún no hay planes registrados. Se escriben en el Google Sheet de Registros, pestaña **Alertas_Seguimiento** "
                "(semana, tipo de alerta, frente, persona o código tal como aparece en esta pestaña, plan, responsable, fecha compromiso y estado).")
    else:
        abiertos = sg[sg["Estado"] != "Cerrada"]
        vencidos = abiertos[abiertos["Fecha compromiso"].notna() & (abiertos["Fecha compromiso"] < hoy)]
        q = st.columns(4)
        q[0].metric("Planes registrados esta semana", int((sg["lunes"] == ini).sum()))
        q[1].metric("Abiertos o en plan", len(abiertos))
        q[2].metric("Vencidos", len(vencidos))
        q[3].metric("Cerrados", int((sg["Estado"] == "Cerrada").sum()))
        ver = sg[(sg["lunes"] <= ini) & ((sg["lunes"] == ini) | (sg["Estado"] != "Cerrada"))].copy()
        ind_now = indicadores(ini)
        cache_ind = {}
        antes_l, ahora_l, res_l = [], [], []
        for _, r in ver.iterrows():
            t_, ref_ = str(r["tipo_n"]), r["ref"]
            if pd.isna(r["lunes"]) or t_ not in FMT_T:
                antes_l.append(""); ahora_l.append(""); res_l.append(""); continue
            if r["lunes"] not in cache_ind:
                cache_ind[r["lunes"]] = indicadores(r["lunes"])
            antes = cache_ind[r["lunes"]].get((t_, ref_), np.nan)
            ahora = ind_now.get((t_, ref_), np.nan)
            antes_l.append("" if pd.isna(antes) else FMT_T[t_](antes))
            ahora_l.append("" if pd.isna(ahora) else FMT_T[t_](ahora))
            res_l.append("🆕 Registrado esta semana" if r["lunes"] == ini else resultado(t_, antes, ahora, ref_, ind_now))
        ver["Semana del plan"] = antes_l
        ver[f"Semana {ini:%d-%b}"] = ahora_l
        ver["Resultado"] = res_l
        ver["Vencido"] = np.where(ver["Fecha compromiso"].notna() & (ver["Fecha compromiso"] < hoy) & (ver["Estado"] != "Cerrada"), "⏰ Sí", "")
        st.dataframe(ver.sort_values(["lunes", "Tipo de alerta"])[["Semana (lunes)", "Tipo de alerta", "Frente", "Persona o código", "Plan de mejora", "Responsable",
                                                                  "Fecha compromiso", "Estado", "Vencido", "Semana del plan", f"Semana {ini:%d-%b}", "Resultado", "Resultado / comentario"]]
                     .style.format({"Fecha compromiso": lambda v: "" if pd.isna(v) else f"{v:%d/%m/%Y}"}), hide_index=True, width='stretch')
        st.caption("Semana del plan = valor de la alerta cuando se registró el plan; la columna siguiente = valor en la semana elegida arriba. "
                   "Si dice 'se puede cerrar', cambia el estado a 'Cerrada' en Alertas_Seguimiento y anota el resultado.")
    st.caption("Para registrar un plan: en el Google Sheet de Registros, pestaña Alertas_Seguimiento, copia la semana (lunes), el tipo de alerta "
               "y la persona o código exactamente como aparecen aquí. El plan sigue apareciendo las semanas siguientes hasta que lo marques 'Cerrada'.")


# ============ PRESUPUESTO VS. EJECUTADO (7-oct-2026) ============
# Fuentes (Drive, "cualquier persona con el enlace"; se actualizan con "Subir nueva versión"):
#   - Programación / curva S (hoja AVANCE SEMANA A SEMANA)
#   - Cortes para SINCO (hoja PPTO): venta y facturado por ítem
#   - Parametros_Alertas: pestaña Cortes (fecha de cada corte) y MO_proyectada (costo proyectado y MO del APU)
URL_PROGRAMACION = "https://drive.google.com/uc?export=download&id=1VZr9GvsMjxG5rmw58Tb9GGIpNBHytGnx"
URL_CORTES = "https://drive.google.com/uc?export=download&id=1sI0-Hct9xogJHFhjbT4HDRhFK5xsMiLw"
# Intervención de la programación "(6) ..." -> código de frente SINCO
INTERV_A_FRENTE = {"3": "01", "3A": "02", "3B": "03", "4": "04", "5": "05", "5A": "06", "5B": "07", "6": "08",
                   "7": "09", "8": "10", "9": "11", "10": "12", "11": "13", "13": "14", "14": "15", "15": "16", "16": "17"}
NOMBRE_FRENTE = {"01": "3 - Check in", "02": "3A - Equipaje saliendo BHS", "03": "3B - Filtro seg. internacional",
                 "04": "4 - Filtro seg. nacional", "05": "5 - Sala embarque remota", "06": "5A - Equipaje llegando",
                 "07": "5B - Sala de espera actual", "08": "6 - Plataforma fase 1", "09": "7 - Nuevo centro conexiones",
                 "10": "8 - Inmigración", "11": "9 - Emigración", "12": "10 - Plataforma fase 2",
                 "13": "11 y 12 - Segundo filtro", "14": "13 - Dobles lectoras", "15": "14 - Guía socioambiental",
                 "16": "15 - Guía SST-SMS", "17": "16 - PMT"}
FRENTES_GUIAS = {"15", "16", "17"}  # su MO es personal profesional que no se registra en la app


_MES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def _f(d, anio=False):
    """Fecha corta en español: 24-sep (o 24-sep-2026)."""
    if d is None or pd.isna(d):
        return "—"
    return f"{d.day:02d}-{_MES[d.month-1]}" + (f"-{d.year}" if anio else "")


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan


def _leer_xlsx(url, hoja):
    import io, urllib.request, openpyxl
    with urllib.request.urlopen(url, timeout=120) as r:
        b = r.read()
    wb = openpyxl.load_workbook(io.BytesIO(b), read_only=True, data_only=True)
    return list(wb[hoja].iter_rows(values_only=True))


@st.cache_data(ttl=3600, show_spinner="Leyendo la programación (curva S)...")
def cargar_programacion():
    rows = _leer_xlsx(URL_PROGRAMACION, "AVANCE SEMANA A SEMANA")
    h2 = rows[1]
    semanas = []
    c = 10
    while c < len(h2) and h2[c] is not None:
        f = pd.Timestamp(h2[c]).normalize()
        semanas.append((f, c))
        c += 8
    regs = []
    for r in rows[5:]:
        if r[1] is None:
            continue
        wbs = str(r[1]).strip()
        nombre = str(r[2] if r[2] is not None else (r[3] or "")).strip()
        for f, c in semanas:
            regs.append((wbs, nombre, _num(r[5]), f, _num(r[c + 3]), _num(r[c + 7]), _num(r[c + 2]), _num(r[c + 6])))
    P = pd.DataFrame(regs, columns=["wbs", "nombre", "costo", "semana", "PV", "EV", "PV_fis", "EV_fis"])
    P["nivel"] = np.where(P["wbs"] == "0", 0, P["wbs"].str.count(r"\.") + 1)
    interv = P["nombre"].str.extract(r"^\((\w+)\)")[0]
    P["fr"] = np.where(P["nivel"] == 1, interv.map(INTERV_A_FRENTE), None)
    return P


@st.cache_data(ttl=3600, show_spinner="Leyendo los cortes de obra...")
def cargar_cortes():
    rows = _leer_xlsx(URL_CORTES, "PPTO")
    h1, h3 = rows[0], rows[2]
    cols = {}
    for i, v in enumerate(h1):
        if isinstance(v, str) and v.strip().upper().startswith("CORTE"):
            k = int(re.findall(r"\d+", v)[0])
            cc = cs = None
            for j in range(i, min(i + 6, len(h3))):
                t = str(h3[j] or "").strip().upper()
                if t == "CANT" and cc is None:
                    cc = j
                elif t == "SUBTOTAL" and cs is None:
                    cs = j
            if cc is not None and cs is not None:
                cols[k] = (cc, cs)
    total_fila = None
    regs = []
    for r in rows[4:]:
        cod = str(r[0]).strip() if r[0] is not None else ""
        if total_fila is None and cod and not re.match(r"^\d\d(\.\d\d)+$", cod):
            total_fila = r
        if not re.match(r"^\d\d(\.\d\d)+$", cod) or str(r[1] or "").strip().upper() == "TITULOS":
            continue
        d = {"cod": cod, "desc": r[2], "und": r[4], "cant": _num(r[3]), "pu": _num(r[5]), "venta": _num(r[6])}
        for k, (cc, cs) in cols.items():
            d[f"q{k}"] = _num(r[cc])
            d[f"v{k}"] = _num(r[cs])
        regs.append(d)
    C = pd.DataFrame(regs).fillna({c: 0 for c in [f"v{k}" for k in cols] + [f"q{k}" for k in cols] + ["venta", "cant"]})
    # solo ítems sin hijos (las filas título con valor duplicarían)
    cods = set(C["cod"])
    C = C[~C["cod"].apply(lambda x: any(o.startswith(x + ".") for o in cods))].copy()
    total_archivo = _num(total_fila[6]) if total_fila is not None else np.nan
    return C, sorted(cols), total_archivo


def _fechas_cortes():
    try:
        x = pd.read_excel(URL_PARAMS, sheet_name="Cortes")
        x = x.rename(columns={x.columns[0]: "Corte", x.columns[1]: "Fecha"})
        x["Corte"] = pd.to_numeric(x["Corte"], errors="coerce")
        x["Fecha"] = pd.to_datetime(x["Fecha"], errors="coerce", dayfirst=True)
        x = x.dropna().astype({"Corte": int}).sort_values("Corte")
        return dict(zip(x["Corte"], x["Fecha"]))
    except Exception:
        return {}


def mostrar_presupuesto(mo_base):
    try:
        P = cargar_programacion()
        C, cortes, total_archivo = cargar_cortes()
    except Exception as e:
        st.error(f"No se pudieron leer los archivos de programación o de cortes en Drive ({e}). "
                 "Revisa que sigan compartidos como 'Cualquier persona con el enlace'.")
        return
    _, mop, _, _ = cargar_parametros()
    FEC = _fechas_cortes()

    # ---- semana de corte de la programación: la última con ejecutado, sin pasar de hoy
    hoy = pd.Timestamp.today().normalize()
    top = P[P["wbs"] == "0"].set_index("semana").sort_index()
    con_ev = top[(top["EV"] > 0) & (top.index <= hoy)]
    if con_ev.empty:
        st.warning("La programación no tiene semanas con ejecutado.")
        return
    sem = con_ev.index.max()
    L1 = P[(P["nivel"] == 1) & P["fr"].notna()]
    fr_sem = L1[L1["semana"] == sem].groupby("fr")[["costo", "PV", "EV"]].sum()

    # ---- cortes: facturado por ítem y por frente; último corte con valor
    cortes_con_valor = [k for k in cortes if C[f"v{k}"].abs().sum() > 0]
    ult = max(cortes_con_valor) if cortes_con_valor else None
    C["facturado"] = C[[f"v{k}" for k in cortes_con_valor]].sum(axis=1) if cortes_con_valor else 0.0
    C["fr"] = C["cod"].str[:2]
    M = mop.rename(columns={"Código": "cod", "Proyectado total": "proy", "Proy. mano de obra": "pMO", "Frente": "frS"})[["cod", "frS", "proy", "pMO"]]
    M["frS"] = M["frS"].astype(str).str.strip().str.zfill(2)
    M["proy"] = pd.to_numeric(M["proy"], errors="coerce").fillna(0)
    M["pMO"] = pd.to_numeric(M["pMO"], errors="coerce").fillna(0).clip(lower=0)
    I = C.merge(M.drop(columns="frS"), on="cod", how="left").fillna({"proy": 0, "pMO": 0})
    I["pct_fact"] = np.where(I["venta"] > 0, I["facturado"] / I["venta"], 0)

    venta = I["venta"].sum()
    proy_total = M["proy"].sum()
    PV, EV = top.loc[sem, "PV"], top.loc[sem, "EV"]
    fact = I["facturado"].sum()

    st.subheader("Presupuesto vs. ejecutado")
    txt_corte = f"corte {ult}" + (f" ({_f(FEC[ult], True)})" if ult in FEC else "") if ult else "sin cortes"
    st.caption(f"Programación al {_f(sem, True)} · facturado hasta el {txt_corte} · costo proyectado y MO del APU de SINCO "
               "(Parametros_Alertas) · MO real de la app. Valores en costo directo, sin AIU ni IVA.")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Contrato (costo directo)", f"$ {venta/1e6:,.0f} M")
    k1.caption(f"Utilidad prevista: $ {(venta - proy_total)/1e6:,.0f} M ({(venta - proy_total)/venta:.1%})")
    k2.metric(f"Programado al {_f(sem)}", f"$ {PV/1e6:,.0f} M")
    k2.caption(f"{PV/venta:.1%} de la obra")
    k3.metric(f"Ejecutado al {_f(sem)}", f"$ {EV/1e6:,.0f} M")
    k3.caption(f"{'Atraso' if EV < PV else 'Adelanto'} de $ {abs(EV-PV)/1e6:,.0f} M · índice {EV/PV:.2f} (1,00 = al día)")
    k4.metric("Facturado (cortes)", f"$ {fact/1e6:,.0f} M")
    k4.caption(f"{fact/venta:.1%} del contrato · ejecutado sin facturar $ {(EV-fact)/1e6:,.0f} M")
    if not np.isnan(total_archivo) and abs(venta - total_archivo) > 1000:
        st.warning(f"La suma de ítems de la hoja de cortes ($ {venta:,.0f}) no cuadra con su fila total ($ {total_archivo:,.0f}). Revisa filas título con valor.")

    # ---- curva S
    cs = top.reset_index()[["semana", "PV", "EV"]].copy()
    cs.loc[cs["semana"] > sem, "EV"] = np.nan
    cs = cs.melt(id_vars="semana", var_name="serie", value_name="valor")
    cs["serie"] = cs["serie"].map({"PV": "Programado", "EV": "Ejecutado"})
    fac = []
    acum = 0.0
    for k in cortes_con_valor:
        acum += C[f"v{k}"].sum()
        if k in FEC:
            fac.append({"semana": FEC[k], "serie": "Facturado (cortes)", "valor": acum})
    cs = pd.concat([cs, pd.DataFrame(fac)], ignore_index=True)
    cs["valor_M"] = cs["valor"] / 1e6
    fig = px.line(cs.dropna(subset=["valor"]), x="semana", y="valor_M", color="serie", markers=False,
                  labels={"semana": "", "valor_M": "Millones de pesos (acumulado)", "serie": ""},
                  color_discrete_map={"Programado": "#94A3B8", "Ejecutado": "#1F3A5F", "Facturado (cortes)": "#D97706"})
    fig.update_traces(selector=dict(name="Facturado (cortes)"), mode="lines+markers", line_shape="hv")
    fig.add_vline(x=sem, line_dash="dot", line_color="#64748B")
    fig.update_layout(height=380, legend=dict(orientation="h", y=1.08), margin=dict(l=10, r=10, t=30, b=10))
    st.plotly_chart(fig, width='stretch')
    if not FEC:
        st.caption("⚠️ No encontré la pestaña Cortes en Parametros_Alertas: la línea de facturado no se puede ubicar en el tiempo.")

    # ---- datos comunes: MO de la app por código y frente, residentes, umbrales
    _, _, U, rs = cargar_parametros()
    RES = dict(zip(rs["Cód. frente"], rs["Residente(s)"].fillna("—")))
    MO_AM, MO_RO = float(U.get("mo_amarillo", 0.8)), float(U.get("mo_rojo", 1.0))
    mo_app = mo_base.copy()
    mo_app["cod"] = mo_app["codact"].astype(str).str.strip()
    np_sup = mo_app["cod"].str.extract(r"^(?:NP|SUP)-(\d\d)")[0]
    fr_cod = mo_app["cod"].where(mo_app["cod"].str.match(r"^\d\d\."))
    fr_app = pd.to_numeric(mo_app.get("codfrente"), errors="coerce").astype("Int64").astype(str).str.zfill(2)
    mo_app["fr"] = fr_cod.str[:2].fillna(np_sup).fillna(fr_app)
    mo_fr = mo_app.groupby("fr")["costo_calculado"].sum()

    g = I.groupby("fr").agg(venta=("venta", "sum"), facturado=("facturado", "sum"))
    g = g.join(M.groupby("frS")[["proy", "pMO"]].sum(), how="left").fillna({"proy": 0, "pMO": 0})
    F = g.join(fr_sem[["PV", "EV"]], how="left").fillna({"PV": 0, "EV": 0})
    F = F[F.index.isin(NOMBRE_FRENTE)].copy()
    F["avance"] = np.where(F["PV"] > 0, F["EV"] / F["PV"], np.nan)
    F["mo_ganada"] = np.where(F["venta"] > 0, F["EV"] / F["venta"] * F["pMO"], 0)
    F["mo_app"] = mo_fr.reindex(F.index).fillna(0)
    F["uso_mo"] = np.where(F["mo_ganada"] > 0, F["mo_app"] / F["mo_ganada"], np.nan)

    # ================= 1. RESUMEN POR FRENTE (compacto) =================
    st.markdown("#### Resumen por frente")

    def s_avance(v):
        if np.isnan(v):
            return "—"
        return f"{'🔴' if v < 0.8 else '🟠' if v < 0.95 else '🟢'} {v:.0%}"

    def s_mo(r):
        if r.name in FRENTES_GUIAS:
            return "no aplica"
        if r["pMO"] <= 0 and r["mo_app"] > 0:
            return "🔴 sin MO en APU"
        if np.isnan(r["uso_mo"]):
            return "—"
        return f"{'🔴' if r['uso_mo'] >= MO_RO else '🟠' if r['uso_mo'] >= MO_AM else '🟢'} {r['uso_mo']:.0%}"

    def alerta(r):
        a = []
        if r.name not in FRENTES_GUIAS:
            if r["pMO"] <= 0 and r["mo_app"] > 0:
                a.append("MO propia en frente subcontratado")
            elif not np.isnan(r["uso_mo"]) and r["uso_mo"] >= MO_RO:
                a.append("MO ya pasó lo que paga el APU")
        if not np.isnan(r["avance"]) and r["avance"] < 0.8:
            a.append("atraso fuerte")
        if r["proy"] > r["venta"]:
            a.append("costo proyectado > venta")
        if r["facturado"] - r["EV"] > 10e6:
            a.append("facturado > ejecutado")
        return " · ".join(a) if a else "—"

    R1 = pd.DataFrame({
        "Frente": [NOMBRE_FRENTE[i] for i in F.index],
        "Residente": [RES.get(i, "—") for i in F.index],
        "Avance vs. programado": F["avance"].map(s_avance).values,
        "Atraso ($ M)": ((F["EV"] - F["PV"]) / 1e6).round(0).values,
        "Ejecutado sin facturar ($ M)": ((F["EV"] - F["facturado"]) / 1e6).round(0).values,
        "MO gastada ($ M)": (F["mo_app"] / 1e6).round(1).values,
        "MO que paga el APU ($ M)": (F["mo_ganada"] / 1e6).round(1).values,
        "MO usada": F.apply(s_mo, axis=1).values,
        "Alerta": F.apply(alerta, axis=1).values,
    })
    R1 = R1[(F["EV"].values > 0) | (F["mo_app"].values > 0)]
    st.dataframe(R1.style.format({"Atraso ($ M)": "{:,.0f}", "Ejecutado sin facturar ($ M)": "{:,.0f}",
                                  "MO gastada ($ M)": "{:,.1f}", "MO que paga el APU ($ M)": "{:,.1f}"}),
                 hide_index=True, width='stretch')
    st.caption(f"**Avance** = ejecutado ÷ programado (🟢 ≥ 95 %, 🟠 80–95 %, 🔴 < 80 %). "
               f"**MO usada** = MO gastada según la app ÷ MO que el APU paga por lo ejecutado del frente "
               f"(🟠 desde {MO_AM:.0%}, 🔴 desde {MO_RO:.0%}). La app tiene registros solo desde el 17-ago: "
               "un 🔴 es sobrecosto seguro; un 🟢 puede esconder MO de mayo a agosto que no se registró.")
    with st.expander("Ver tabla completa (venta, costo proyectado, programado, ejecutado, facturado)"):
        T = F.assign(Frente=[NOMBRE_FRENTE[i] for i in F.index])[["Frente", "venta", "proy", "PV", "EV", "facturado", "pMO", "mo_ganada", "mo_app"]]
        T = T.rename(columns={"venta": "Venta CD", "proy": "Costo proyectado", "PV": "Programado", "EV": "Ejecutado",
                              "facturado": "Facturado", "pMO": "MO propia APU (total)", "mo_ganada": "MO que paga el APU", "mo_app": "MO app"})
        st.dataframe(T.style.format({c: "{:,.0f}" for c in T.columns if c != "Frente"}), hide_index=True, width='stretch')

    # ================= 2. ALERTAS DE SOBRECOSTO DE MO POR PERIODO DE CORTE =================
    st.markdown("#### 🚨 Alertas de sobrecosto de mano de obra — por periodo de corte")
    cortes_fechados = [k for k in cortes_con_valor if k in FEC and (k - 1) in FEC]
    if not cortes_fechados:
        st.info("Para estas alertas hacen falta las fechas de corte en la pestaña Cortes de Parametros_Alertas.")
    else:
        opciones = {f"Corte {k}: {_f(FEC[k-1] + pd.Timedelta(days=1))} → {_f(FEC[k])}": [k] for k in cortes_fechados}
        for a, b in zip(cortes_fechados, cortes_fechados[1:]):
            opciones[f"Cortes {a} y {b}: {_f(FEC[a-1] + pd.Timedelta(days=1))} → {_f(FEC[b])}"] = [a, b]
        claves = list(opciones)
        mo_min = mo_app["fecha"].min()
        lista = [c for c in claves if FEC[opciones[c][0] - 1] + pd.Timedelta(days=1) >= mo_min - pd.Timedelta(days=7)] or claves
        cubiertas = [i for i, c in enumerate(lista) if FEC[opciones[c][-1]] <= mo_app["fecha"].max() + pd.Timedelta(days=1)]
        sel = st.selectbox("Periodo", lista, index=(cubiertas[-1] if cubiertas else len(lista) - 1),
                           help="Rango entre fechas de corte. Solo tiene sentido desde que la app tiene registros (17-ago).")
        ks = opciones[sel]
        d_ini, d_fin = FEC[ks[0] - 1] + pd.Timedelta(days=1), FEC[ks[-1]]
        hasta = [k for k in cortes_con_valor if k <= ks[-1]]
        antes = [k for k in cortes_con_valor if k < ks[0]]
        J = I.copy()
        J["fact_per"] = J[[f"v{k}" for k in ks]].sum(axis=1)
        J["pct_antes"] = np.where(J["venta"] > 0, J[[f"v{k}" for k in antes]].sum(axis=1) / J["venta"], 0) if antes else 0.0
        J["pct_hasta"] = np.where(J["venta"] > 0, J[[f"v{k}" for k in hasta]].sum(axis=1) / J["venta"], 0)
        mper = mo_app[(mo_app["fecha"] >= d_ini) & (mo_app["fecha"] <= d_fin)]
        macu = mo_app[mo_app["fecha"] <= d_fin]
        J = J.merge(mper.groupby("cod")["costo_calculado"].sum().rename("mo_per"), on="cod", how="outer")
        J = J.merge(macu.groupby("cod")["costo_calculado"].sum().rename("mo_acu"), on="cod", how="left")
        J = J.fillna({c: 0 for c in ["fact_per", "mo_per", "mo_acu", "pMO", "pct_antes", "pct_hasta", "venta"]})
        J["fr"] = J["fr"].fillna(J["cod"].str.extract(r"^(?:NP|SUP)-(\d\d)")[0]).fillna(J["cod"].str[:2])
        J["Frente"] = J["fr"].map(NOMBRE_FRENTE).fillna(J["fr"])
        J["Residente"] = J["fr"].map(RES).fillna("—")
        desc_app = mo_app.drop_duplicates("cod").set_index("cod")["actDesc"] if "actDesc" in mo_app.columns else pd.Series(dtype=str)
        J["Actividad"] = J["desc"].fillna(J["cod"].map(desc_app)).astype(str).str.slice(0, 70)
        J = J[~J["fr"].isin(FRENTES_GUIAS)]
        J["mo_paga_acu"] = J["pct_hasta"] * J["pMO"]

        a1 = J[(J["mo_per"] > 0) & (J["pMO"] > 0) & (J["pct_antes"] >= 0.999)].copy()
        a2 = J[(J["mo_per"] > 0) & (J["pMO"] > 0) & (J["pct_antes"] < 0.999) & (J["pct_hasta"] > 0) & (J["mo_acu"] > J["mo_paga_acu"])].copy()
        a2["Exceso"] = a2["mo_acu"] - a2["mo_paga_acu"]
        a3 = J[(J["mo_per"] > 0) & (J["pMO"] <= 0)].copy()
        a3["Tipo"] = np.select([a3["cod"].str.startswith("SUP"), a3["cod"].str.startswith("NP"), a3["venta"] > 0],
                               ["Supervisión de maestro (va a la A)", "No presupuestada", "Subcontratada en el APU"], "Sin venta en el contrato")
        a4 = J[(J["mo_per"] > 0) & (J["pMO"] > 0) & (J["pct_antes"] < 0.999) & (J["fact_per"] <= 0) & ~J.index.isin(a2.index)].copy()

        st.caption(f"MO registrada en la app del {_f(d_ini)} al {_f(d_fin)} (sin guías): **$ {J['mo_per'].sum()/1e6:,.1f} M**. "
                   f"Facturado en {'el corte' if len(ks) == 1 else 'los cortes'} {' y '.join(map(str, ks))}. Último registro de la app: {_f(mo_app['fecha'].max())}.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🔴 Ítems ya cobrados que siguen gastando MO", f"$ {a1['mo_per'].sum()/1e6:,.1f} M")
        m1.caption(f"{len(a1)} ítems · costo sin ingreso")
        m2.metric("🔴 MO que ya pasó lo que paga el APU", f"$ {a2['Exceso'].sum()/1e6:,.1f} M")
        m2.caption(f"{len(a2)} ítems · exceso acumulado")
        m3.metric("🟠 MO en actividades sin MO en el APU", f"$ {a3['mo_per'].sum()/1e6:,.1f} M")
        m3.caption(f"{a3['cod'].nunique()} códigos · nadie la paga hoy")
        m4.metric("⚪ MO en obra aún sin facturar", f"$ {a4['mo_per'].sum()/1e6:,.1f} M")
        m4.caption(f"{len(a4)} ítems · debe entrar en el próximo corte")

        fmt_m = {"MO del periodo": "{:,.0f}", "MO gastada (acum.)": "{:,.0f}", "MO que paga el APU": "{:,.0f}", "Exceso": "{:,.0f}", "% facturado": "{:.0%}"}
        with st.expander(f"🔴 1. Ya cobrados al 100 % y siguen gastando MO ({len(a1)})", expanded=len(a1) > 0):
            st.markdown("**Decisión:** el residente confirma qué pasó. Si es **repaso o garantía** → es costo sin ingreso, va a plan de mejora. "
                        "Si es **código equivocado** en la app → se corrige el código del registro.")
            t = a1.sort_values("mo_per", ascending=False)[["Frente", "Residente", "cod", "Actividad", "mo_per"]]
            st.dataframe(t.rename(columns={"cod": "Código", "mo_per": "MO del periodo"}).style.format(fmt_m), hide_index=True, width='stretch')
        with st.expander(f"🔴 2. La MO gastada ya superó lo que paga el APU por lo facturado ({len(a2)})", expanded=len(a2) > 0):
            st.markdown("**Decisión:** si la cantidad real es mayor que la facturada → **cobrarla en el próximo corte**. "
                        "Si no → revisar **rendimiento o tamaño de la cuadrilla** con el residente.")
            t = a2.sort_values("Exceso", ascending=False)[["Frente", "Residente", "cod", "Actividad", "pct_hasta", "mo_paga_acu", "mo_acu", "Exceso"]]
            st.dataframe(t.rename(columns={"cod": "Código", "pct_hasta": "% facturado", "mo_paga_acu": "MO que paga el APU",
                                           "mo_acu": "MO gastada (acum.)"}).style.format(fmt_m), hide_index=True, width='stretch')
            st.caption("Acumulado hasta el fin del periodo. La app solo tiene registros desde el 17-ago, así que el exceso real es igual o mayor.")
        with st.expander(f"🟠 3. MO propia en actividades que el APU no paga con MO propia ({a3['cod'].nunique()})"):
            st.markdown("**Decisión según el tipo:** *Subcontratada* → ¿por qué trabaja personal propio? descontar al subcontratista o justificar. "
                        "*No presupuestada* → tramitar como adicional para poder cobrarla. *Supervisión* → se carga a la A.")
            t = a3.sort_values("mo_per", ascending=False)[["Frente", "Residente", "Tipo", "cod", "Actividad", "mo_per"]]
            st.dataframe(t.rename(columns={"cod": "Código", "mo_per": "MO del periodo"}).style.format(fmt_m), hide_index=True, width='stretch')
        with st.expander(f"⚪ 4. MO en obra que todavía no se ha facturado ({len(a4)})"):
            st.markdown("**Decisión:** confirmar que estas actividades **entren en el próximo corte**. Si siguen sin facturarse dos cortes seguidos, revisar.")
            t = a4.sort_values("mo_per", ascending=False)[["Frente", "Residente", "cod", "Actividad", "pct_hasta", "mo_per"]]
            st.dataframe(t.rename(columns={"cod": "Código", "pct_hasta": "% facturado", "mo_per": "MO del periodo"}).style.format(fmt_m), hide_index=True, width='stretch')

        import io as _io
        buf = _io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            for nom, d, cols in [("1 Cobrados que gastan MO", a1, ["Frente", "Residente", "cod", "Actividad", "mo_per"]),
                                 ("2 MO supera el APU", a2, ["Frente", "Residente", "cod", "Actividad", "pct_hasta", "mo_paga_acu", "mo_acu", "Exceso"]),
                                 ("3 MO sin MO en APU", a3, ["Frente", "Residente", "Tipo", "cod", "Actividad", "mo_per"]),
                                 ("4 Obra sin facturar", a4, ["Frente", "Residente", "cod", "Actividad", "pct_hasta", "mo_per"])]:
                d[cols].rename(columns={"cod": "Código", "mo_per": "MO del periodo", "pct_hasta": "% facturado",
                                        "mo_paga_acu": "MO que paga el APU", "mo_acu": "MO gastada (acum.)"}).to_excel(xw, sheet_name=nom, index=False)
        st.download_button("📥 Descargar estas alertas (Excel, para los residentes)", buf.getvalue(),
                           file_name=f"Alertas_MO_cortes_{'_'.join(map(str, ks))}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ================= 3. ATRASO PARA GERENCIA =================
    st.markdown(f"#### ⏱️ Atraso por frente al {_f(sem)} — qué estaba programado y no se ha hecho")
    ws_ = set(P["wbs"].unique())
    Ps = P[P["semana"] == sem]
    nom = dict(Ps[["wbs", "nombre"]].values)
    Ls = Ps[(Ps["wbs"] != "0") & ~Ps["wbs"].apply(lambda w: any(o.startswith(w + ".") for o in ws_))].copy()
    Ls["gap"] = Ls["EV"].fillna(0) - Ls["PV"].fillna(0)
    Ls["l1"] = Ls["wbs"].str.split(".").str[0]
    l1fr = dict(P[(P["nivel"] == 1) & P["fr"].notna()][["wbs", "fr"]].drop_duplicates().values)
    Ls["fr"] = Ls["l1"].map(l1fr)
    Ls["zona"] = Ls["wbs"].apply(lambda w: nom.get(".".join(w.split(".")[:2]), "") if w.count(".") >= 2 else "")
    filas = []
    for fr_, d in Ls[Ls["fr"].notna()].groupby("fr"):
        atr = d[d["gap"] < -1e6].sort_values("gap")
        if atr.empty:
            continue
        top3 = " · ".join(f"{r.nombre} {('(' + r.zona.strip() + ')') if r.zona else ''} {abs(r.gap)/1e6:,.0f} M" for r in atr.head(3).itertuples())
        sin_iniciar = int(((d["PV"] > 0) & (d["EV"].fillna(0) <= 0)).sum())
        filas.append({"Frente": NOMBRE_FRENTE[fr_], "Residente": RES.get(fr_, "—"),
                      "Avance vs. programado": s_avance(F.loc[fr_, "avance"]) if fr_ in F.index else "—",
                      "Atraso neto ($ M)": round((d["gap"].sum()) / 1e6),
                      "Actividades sin iniciar que ya debían ir": sin_iniciar,
                      "Lo más atrasado": top3})
    if filas:
        A3 = pd.DataFrame(filas).sort_values("Atraso neto ($ M)")
        st.dataframe(A3, hide_index=True, width='stretch')
        st.caption("Atraso neto = ejecutado − programado del frente (lo adelantado compensa lo atrasado). "
                   "Úsalo para pedir a cada residente un **plan de recuperación** de las actividades listadas. "
                   "Ojo: si lo atrasado es un **suministro** (equipos, muebles, cubierta) la causa suele ser compras, no la obra.")
    with st.expander("Ver las 20 actividades con mayor atraso del proyecto"):
        A_ = Ls.sort_values("gap").head(20).assign(Frente=lambda d: d["fr"].map(NOMBRE_FRENTE))
        A_ = A_[["Frente", "zona", "nombre", "PV", "EV", "gap"]].rename(columns={"zona": "Zona / grupo", "nombre": "Actividad", "PV": "Programado",
                                                                               "EV": "Ejecutado", "gap": "Atraso"})
        st.dataframe(A_.style.format({c: "{:,.0f}" for c in ["Programado", "Ejecutado", "Atraso"]}), hide_index=True, width='stretch')

    # ---- calidad de los archivos
    avisos = []
    sobre = I[(I["cant"] > 0) & (I[[f"q{k}" for k in cortes_con_valor]].sum(axis=1) > I["cant"] * 1.0001)]
    for _, r in sobre.iterrows():
        avisos.append(f"Ítem {r['cod']}: facturada {r[[f'q{k}' for k in cortes_con_valor]].sum():,.2f} {r['und']} contra {r['cant']:,.2f} del contrato.")
    sin_fecha = [k for k in cortes_con_valor if k not in FEC]
    if sin_fecha:
        avisos.append(f"Cortes sin fecha en la pestaña Cortes: {', '.join(map(str, sin_fecha))}.")
    fut = top[(top.index > hoy) & (top["EV"] > 0)]
    if not fut.empty:
        avisos.append(f"La programación tiene 'ejecutado' en semanas futuras ({', '.join(f'{_f(d)}' for d in fut.index)}); no se tienen en cuenta.")
    if avisos:
        with st.expander(f"⚠️ Datos para revisar en los archivos ({len(avisos)})"):
            for a in avisos:
                st.markdown(f"- {a}")


# Vistas (7-oct-2026): GERENCIA = Resumen general + Alertas de la semana + Presupuesto vs. ejecutado.
# CONTROL INTERNO (control de costos y residentes) = Cobertura, Códigos por persona, Productividad, Cargo vs. tarea.
if VISTA_GERENCIA:
    tab0, tabA, tabP = st.tabs(["📋 Resumen general", "🚨 Alertas de la semana", "💰 Presupuesto vs. ejecutado"])
else:
    tab8, tab5, tab3, tab7 = st.tabs(["✅ Cobertura de registro", "🔎 Códigos por Persona", "📈 Productividad", "⚠️ Cargo vs. tarea"])

if VISTA_GERENCIA:
    with tabA:
        mostrar_alertas(mo[mo["frente"].isin(frente_sel)] if frente_sel else mo, tarifas)
    with tabP:
        mostrar_presupuesto(mo)

# ============ TAB RESUMEN EJECUTIVO SEMANAL ============
if VISTA_GERENCIA:
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
                    "dias_faltantes": ", ".join(dia_es(d) for d in faltantes),
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
                        f"📌 **{dia_es(dia_top, largo=True)}** es el día que más se repite sin "
                        f"registro entre quienes sí trabajaron el resto de la semana ({veces} personas)."
                    )

            with st.expander(f"🔴 Sin ningún registro ({n_total})"):
                st.dataframe(
                    asist.loc[asist["estado"] == "Total", ["nombre", "cargo"]].rename(columns=COLS_ASIST),
                    width='stretch', hide_index=True,
                )
            with st.expander(f"🟡 Registro parcial ({n_parcial})"):
                st.dataframe(
                    asist.loc[asist["estado"] == "Parcial", ["nombre", "cargo", "dias_con_registro", "dias_totales", "dias_faltantes"]].rename(columns=COLS_ASIST),
                    width='stretch', hide_index=True,
                )
            with st.expander(f"✅ Registro completo ({n_completo})"):
                st.dataframe(
                    asist.loc[asist["estado"] == "Completo", ["nombre", "cargo"]].rename(columns=COLS_ASIST),
                    width='stretch', hide_index=True,
                )

# ============ (antes TAB MANO DE OBRA) -> continúa en Resumen semanal ============
if VISTA_GERENCIA:
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
                      "tarifa_hora", "costo_calculado", "semana_etiqueta"]].sort_values("fecha", ascending=False)
                .rename(columns={"fecha": "Fecha", "nombre": "Nombre", "cargo": "Cargo", "frente": "Frente",
                                 "horas": "Horas", "tipo_hora": "Tipo de hora", "tarifa_hora": "Tarifa hora",
                                 "costo_calculado": "Costo", "semana_etiqueta": "Semana"}),
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
if not VISTA_GERENCIA:
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
if not VISTA_GERENCIA:
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

if not VISTA_GERENCIA:
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
        st.dataframe(sin_tarifa.rename(columns={"nombre": "Nombre", "cargo": "Cargo", "frente": "Frente", "horas": "Horas"}), width='stretch')
        st.info(
            "Estos nombres no cruzaron con la tabla maestra de personal. "
            "Revisa si hay errores de digitación o personal nuevo sin tarifa asignada."
        )
    else:
        st.success("Todos los registros tienen tarifa confirmada.")

    st.subheader("Resumen de cobertura")
    st.metric("Total de registros en vivo", len(mo))

# ============ TAB CÓDIGOS POR PERSONA ============
if not VISTA_GERENCIA:
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

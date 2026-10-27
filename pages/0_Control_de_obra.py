"""Control de obra · AIRPLAN T1 JMC

Dos vistas en una sola página:
  - Control semanal (administrador): semana de lunes a domingo desde abril, día a día, maestros, alertas, actividades.
  - Gerencia: alertas de sobrecosto en miles de pesos (cobro, subcontratos, mano de obra, actividades, equipos, horas extra).

Los datos se calculan en control_obra_datos.py con las mismas reglas de app.py y pages/1_Gerencia.py.
El dibujo está en componentes/control_obra/index.html.
"""
import json
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import gerencia_datos as gd
import control_obra_datos as co
from control_obra_componente import control_obra

st.set_page_config(page_title="Control de obra · AIRPLAN T1", page_icon="🏗️", layout="wide")
ZONA = ZoneInfo("America/Bogota")
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ============ FUENTES (las mismas de pages/1_Gerencia.py) ============
URL_BD_NOMINA = "https://drive.google.com/uc?export=download&id=1OB3w9yq_u8opkAgMEkRhaafjh1Tp2mvT"      # BD_Nomina_MO
URL_PARAMS = "https://drive.google.com/uc?export=download&id=1c31oKbuKnXO2XV2AC3n4VOqhQmxm6FTr"         # Parametros_Alertas
URL_PROGRAMACION = "https://drive.google.com/uc?export=download&id=1VZr9GvsMjxG5rmw58Tb9GGIpNBHytGnx"   # Curva S
URL_TARIFAS = "https://drive.google.com/uc?export=download&id=1gUGj51qvXvmwZvyhhgy-05CY5INWHSST"        # Maestro de Personal
URL_HITOS = "https://drive.google.com/uc?export=download&id=1KmOde_OuHgE7aCE8p1B_wqkCffp5u-74"          # cuadro.resumen.pagos.hitos
URL_BASE = "https://drive.google.com/uc?export=download&id=1mCDMXj8M-mq9WlD218YRxPzC7ZB09ja2"           # Base_Tablero_Gerencia
URL_APP = "https://docs.google.com/spreadsheets/d/1Sbt-r9yR_pcyluP_JLX9nA2nwxpW3cU2ATVPMwvNP-I/export?format=csv&gid=0"  # registros app
URL_SINCO_CONTRATOS = "https://drive.google.com/uc?export=download&id=1GFY0KewZ1uyKkU5I9IK64dC06r1c21uo"   # SINCO > Informe contratos detallado
URL_SINCO_ACTAS = "https://drive.google.com/uc?export=download&id=1MosjVLGeqceP2IIZIOMtv42S0kZT0MKp"       # SINCO > Detallado de actas
# SINCO > 101 Informe relación de pagos (anticipos). Mientras esté vacío se usa anticipos_contratistas.csv del repositorio.
URL_RELACION_PAGOS = ""

st.markdown("""<style>
.block-container { padding-top: 3rem; padding-bottom: 0.5rem; padding-left: 0.6rem; padding-right: 0.6rem; max-width: 100%; }
header[data-testid="stHeader"] { height: 2.2rem; }
iframe[title="control_obra.control_obra"] { border-radius: 14px; }
</style>""", unsafe_allow_html=True)


@st.cache_data(ttl=600, show_spinner=False)
def _bytes(url, nombre, html_ok=False):
    return gd.descargar(url, nombre, html_ok=html_ok)


@st.cache_data(ttl=60, show_spinner=False)
def _app():
    ruta = os.environ.get("GERENCIA_LOCAL_APP")
    return pd.read_csv(ruta or URL_APP)


def _csv(nombre, **kw):
    return pd.read_csv(os.path.join(RAIZ, nombre), **kw)


@st.cache_data(ttl=600, show_spinner=False)
def _fuentes():
    """Nómina, parámetros, curva S, hitos, base, maestro y SINCO: se leen una vez cada 10 minutos."""
    fuentes = dict(nomina=_bytes(URL_BD_NOMINA, "nomina"), params=_bytes(URL_PARAMS, "params"), curva=_bytes(URL_PROGRAMACION, "curva"),
                   hitos=_bytes(URL_HITOS, "hitos"), base=_bytes(URL_BASE, "base"), maestro=_bytes(URL_TARIFAS, "maestro"))
    for k, u in [("contratos", URL_SINCO_CONTRATOS), ("actas", URL_SINCO_ACTAS), ("relacion_pagos", URL_RELACION_PAGOS)]:
        if u or os.environ.get("GERENCIA_LOCAL_" + k.upper()):
            try:
                fuentes[k] = _bytes(u, k, html_ok=True)
            except Exception:
                pass
    return co.leer_fuentes(fuentes)


@st.cache_data(ttl=600, show_spinner=False, max_entries=4)
def construir(_leidas, app_hash, _app_df):
    datos = co.construir(None, _app_df, hoy=datetime.now(ZONA).replace(tzinfo=None), leidas=_leidas,
                         recargos=_csv("recargos.csv"),
                         eq_fact=_csv("equipos_facturado.csv", dtype={"frente": str}),
                         eq_apu=_csv("equipos_apu.csv", dtype={"frente": str}),
                         anticipos=_csv("anticipos_contratistas.csv"))
    return json.dumps(datos, ensure_ascii=False, allow_nan=False)


with st.sidebar:
    if st.button("Actualizar datos ahora", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.caption("La app se relee cada minuto; la nómina, la curva S, los hitos y SINCO cada 10 minutos.")

try:
    with st.spinner("Leyendo la app, la nómina, la curva S y SINCO..."):
        leidas = _fuentes()
        app = _app()
        app_hash = str(pd.util.hash_pandas_object(app.astype(str), index=False).sum()) + str(len(app))
        datos_json = construir(leidas, app_hash, app)
except Exception as e:
    st.error(f"No se pudieron armar los datos: {e}")
    st.stop()

control_obra(data_json=datos_json, key="control_obra", default=None)

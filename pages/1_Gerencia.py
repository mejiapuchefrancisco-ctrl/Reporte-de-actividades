"""Tablero de gerencia · Mano de obra AIRPLAN T1 JMC.

Lee en vivo la nómina, los parámetros, la curva S, el cuadro de hitos, la base del tablero
y los registros de la app; los cálculos están en gerencia_datos.py (raíz del repositorio).
Para cambiar un enlace, edite la sección ENLACES de este archivo.
"""
from datetime import datetime
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import importlib

import gerencia_datos as gd

gd = importlib.reload(gd)  # siempre la versión más reciente subida a GitHub

# ============ ENLACES (Google Drive, compartidos como "Cualquier persona con el enlace") ============
URL_BD_NOMINA = "https://drive.google.com/uc?export=download&id=1OB3w9yq_u8opkAgMEkRhaafjh1Tp2mvT"      # BD_Nomina_MO
URL_PARAMS = "https://drive.google.com/uc?export=download&id=1c31oKbuKnXO2XV2AC3n4VOqhQmxm6FTr"         # Parametros_Alertas
URL_PROGRAMACION = "https://drive.google.com/uc?export=download&id=1VZr9GvsMjxG5rmw58Tb9GGIpNBHytGnx"   # Curva S
URL_TARIFAS = "https://drive.google.com/uc?export=download&id=1gUGj51qvXvmwZvyhhgy-05CY5INWHSST"        # Maestro de Personal
URL_HITOS = "https://drive.google.com/uc?export=download&id=1KmOde_OuHgE7aCE8p1B_wqkCffp5u-74"                                                  # cuadro.resumen.pagos.hitos
URL_BASE = "https://drive.google.com/uc?export=download&id=1mCDMXj8M-mq9WlD218YRxPzC7ZB09ja2"                                              # Base_Tablero_Gerencia
URL_APP = "https://docs.google.com/spreadsheets/d/1Sbt-r9yR_pcyluP_JLX9nA2nwxpW3cU2ATVPMwvNP-I/export?format=csv&gid=0"  # registros app
# Exportes de SINCO tal como salen (.xls). Se reemplazan cada semana en Drive con "Administrar versiones" (el enlace no cambia).
URL_SINCO_CONTRATOS = "https://drive.google.com/uc?export=download&id=1GFY0KewZ1uyKkU5I9IK64dC06r1c21uo"   # SINCO > Informe contratos detallado
URL_SINCO_ACTAS = "https://drive.google.com/uc?export=download&id=1MosjVLGeqceP2IIZIOMtv42S0kZT0MKp"       # SINCO > Detallado de actas
RUTA_RECARGOS = "recargos.csv"
RUTA_LOGO = "aia.png"
ZONA = ZoneInfo("America/Bogota")

# ============ ESTILO ============
C = dict(verde="#1A5632", verde_osc="#0F331D", verde_cl="#4C8A64", dorado="#C9A227", gris="#7A8B82",
         rojo="#B5442E", ambar="#C98A12", ok="#2E7D4F", fondo="#F3F5F2", tinta="#16211A", tinta2="#5B6A60", linea="#DDE3DC")
SERIES = [C["verde"], C["dorado"], "#3C6E8F", C["verde_cl"], "#8C5E3C", "#7C5A9C", "#4C9A9A", "#A3ADA6"]

st.set_page_config(page_title="Gerencia · Mano de obra", layout="wide", page_icon=RUTA_LOGO)
st.markdown(f"""
<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;700&display=swap" rel="stylesheet">
<style>
  .block-container {{ padding-top: 1.6rem; max-width: 1400px; }}
  h1, h2, h3, .g-num {{ font-family: 'Archivo', 'Source Sans Pro', sans-serif; }}
  .g-head {{ display:flex; justify-content:space-between; align-items:flex-end; flex-wrap:wrap; gap:.6rem;
            border-bottom: 3px solid {C['verde']}; padding-bottom:.6rem; margin-bottom:.8rem; }}
  .g-head h1 {{ font-size: 1.9rem; margin:0; color:{C['verde_osc']}; letter-spacing:-.01em; }}
  .g-chips {{ display:flex; gap:.4rem; flex-wrap:wrap; }}
  .g-chip {{ background:{C['fondo']}; border:1px solid {C['linea']}; color:{C['tinta2']}; border-radius:999px; padding:2px 10px; font-size:.78rem; }}
  .g-chip b {{ color:{C['tinta']}; font-weight:600; }}
  .g-lead {{ font-size:1.05rem; color:{C['tinta']}; max-width: 980px; margin:.2rem 0 1rem; line-height:1.45; }}
  .g-kpis {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(150px,1fr)); gap:.7rem; margin-bottom:1.1rem; }}
  .g-kpi {{ background:white; border:1px solid {C['linea']}; border-radius:10px; padding:.75rem .9rem; border-top:4px solid {C['verde']}; }}
  .g-kpi.rojo {{ border-top-color:{C['rojo']}; }} .g-kpi.ambar {{ border-top-color:{C['ambar']}; }} .g-kpi.ok {{ border-top-color:{C['ok']}; }}
  .g-kpi .l {{ font-size:.78rem; color:{C['tinta2']}; line-height:1.25; min-height:2em; }}
  .g-kpi .v {{ font-family:'Archivo',sans-serif; font-size:1.75rem; font-weight:700; color:{C['tinta']}; line-height:1.15; margin:.15rem 0; font-variant-numeric: tabular-nums; }}
  .g-kpi.rojo .v {{ color:{C['rojo']}; }} .g-kpi.ok .v {{ color:{C['ok']}; }}
  .g-kpi .s {{ font-size:.75rem; color:{C['tinta2']}; }}
  .g-sec {{ font-family:'Archivo',sans-serif; font-size:1.15rem; font-weight:600; color:{C['verde_osc']}; margin: .6rem 0 .5rem; }}
  .g-cards {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(300px,1fr)); gap:.7rem; margin-bottom:1rem; }}
  .g-card {{ background:white; border:1px solid {C['linea']}; border-left:6px solid {C['ambar']}; border-radius:10px; padding:.75rem .9rem; }}
  .g-card.Rojo {{ border-left-color:{C['rojo']}; }}
  .g-card .t {{ display:flex; justify-content:space-between; gap:.5rem; align-items:baseline; }}
  .g-card .t b {{ font-family:'Archivo',sans-serif; font-size:1.02rem; color:{C['tinta']}; }}
  .g-card .r {{ font-size:.75rem; color:{C['tinta2']}; }}
  .g-card .n {{ display:grid; grid-template-columns: repeat(3,1fr); gap:.3rem; margin:.5rem 0; }}
  .g-card .n div {{ background:{C['fondo']}; border-radius:6px; padding:.3rem .45rem; }}
  .g-card .n span {{ display:block; font-size:.68rem; color:{C['tinta2']}; line-height:1.15; }}
  .g-card .n strong {{ font-family:'Archivo',sans-serif; font-size:.98rem; color:{C['tinta']}; font-variant-numeric: tabular-nums; }}
  .g-card .d {{ font-size:.85rem; color:{C['tinta']}; line-height:1.4; }}
  .pill {{ display:inline-block; padding:1px 9px; border-radius:999px; font-size:.74rem; font-weight:600; border:1px solid currentColor; white-space:nowrap; }}
  .pill.Rojo {{ color:{C['rojo']}; }} .pill.Amarillo {{ color:{C['ambar']}; }} .pill.Verde {{ color:{C['ok']}; }}
  table.g-t {{ width:100%; border-collapse:collapse; font-size:.86rem; }}
  table.g-t th {{ text-align:left; color:{C['tinta2']}; font-weight:600; border-bottom:2px solid {C['linea']}; padding:.35rem .5rem; vertical-align:bottom; }}
  table.g-t td {{ border-bottom:1px solid {C['linea']}; padding:.4rem .5rem; vertical-align:top; }}
  table.g-t td.n {{ text-align:right; white-space:nowrap; font-variant-numeric: tabular-nums; }}
  .g-note {{ font-size:.78rem; color:{C['tinta2']}; }}
  .stButton button[kind="primary"] {{ background:{C['verde']}; border-color:{C['verde']}; }}
  .stButton button[kind="primary"]:hover {{ background:{C['verde_osc']}; border-color:{C['verde_osc']}; }}
  .stTabs [aria-selected="true"] {{ color: {C['verde']} !important; }}
  .stTabs [data-baseweb="tab-highlight"] {{ background-color: {C['verde']} !important; }}
</style>""", unsafe_allow_html=True)


# ============ CARGA (con caché; el botón Actualizar la borra) ============
@st.cache_data(ttl=600, show_spinner=False)
def _bytes(url, nombre):
    return gd.descargar(url, nombre)


@st.cache_data(ttl=600, show_spinner=False)
def cargar_fuentes():
    n, aj = gd.leer_nomina(_bytes(URL_BD_NOMINA, "nomina"))
    pa, mop = gd.leer_parametros(_bytes(URL_PARAMS, "params"))
    P = gd.leer_curva(_bytes(URL_PROGRAMACION, "curva"))
    H, I, V = gd.leer_hitos(_bytes(URL_HITOS, "hitos"))
    base = gd.leer_base(_bytes(URL_BASE, "base"))
    try:
        ma = gd.leer_maestro(_bytes(URL_TARIFAS, "maestro"))
    except Exception:
        ma = None
    return n, aj, pa, mop, P, H, I, V, base, ma


@st.cache_data(ttl=600, show_spinner=False)
def cargar_sinco():
    """Contratos y actas de SINCO para la pestaña Subcontratistas. Devuelve None si aún no están configurados."""
    import os
    if not (URL_SINCO_CONTRATOS or os.environ.get("GERENCIA_LOCAL_CONTRATOS")):
        return None
    Cs = gd.leer_contratos_sinco(gd.descargar(URL_SINCO_CONTRATOS, "contratos", html_ok=True))
    As = gd.leer_actas_sinco(gd.descargar(URL_SINCO_ACTAS, "actas", html_ok=True), Cs)
    return Cs, As


@st.cache_data(ttl=120, show_spinner=False)
def cargar_app(_ma_hash, ma):
    """Registros de la app para estimar los meses sin nómina (horas × tarifa × recargo)."""
    try:
        import os
        ruta = os.environ.get("GERENCIA_LOCAL_APP")
        mo = pd.read_csv(ruta or URL_APP)
    except Exception:
        return None
    mo = mo.dropna(subset=["nombre"]).copy()
    if "codact" not in mo.columns and "codcat" in mo.columns:
        mo = mo.rename(columns={"codcat": "codact"})
    mo["fecha"] = pd.to_datetime(mo["fecha"], errors="coerce")
    mo = gd.quitar_excluidos(mo, gd.leer_excluidos(_bytes(URL_PARAMS, "params")))
    mo["horas"] = pd.to_numeric(mo["horas"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    mo["nombre_norm"] = mo["nombre"].map(gd.normalizar_nombre)
    if ma is None:
        return None
    t = ma.copy()
    t["nombre_norm"] = t["NombreNormalizado"].map(gd.normalizar_nombre)
    t["TarifaHoraReal"] = pd.to_numeric(t["TarifaHoraReal"], errors="coerce")
    rec = pd.read_csv(RUTA_RECARGOS)
    mo = mo.merge(t[["nombre_norm", "TarifaHoraReal"]].drop_duplicates("nombre_norm"), on="nombre_norm", how="left")
    mo = mo.merge(rec[["TipoHora", "Factor"]], left_on="tipoHora", right_on="TipoHora", how="left")
    mo["costo_calculado"] = mo["horas"] * mo["TarifaHoraReal"] * mo["Factor"]
    return mo


def M(v, d=1, signo=False):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    s = f"{abs(v) / 1e6:,.{d}f}".replace(",", "#").replace(".", ",").replace("#", ".")
    pre = "−" if v < 0 else ("+" if signo and v > 0 else "")
    return f"{pre}{s} M"


def PCT(v):
    return "—" if v is None or (isinstance(v, float) and np.isnan(v)) else f"{v:.0%}".replace("%", " %")


def fecha_txt(t):
    meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
    return f"{t.day}-{meses[t.month - 1]}"


def layout(fig, h=340, leyenda=True):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor="white", paper_bgcolor="white",
                      font=dict(family="Source Sans Pro, sans-serif", size=13, color=C["tinta"]),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title=None) if leyenda else None, showlegend=leyenda,
                      hoverlabel=dict(bgcolor="white", font_size=13), separators=",.")
    fig.update_yaxes(gridcolor="#EEF1ED", zerolinecolor=C["linea"], tickformat=",.0f")
    fig.update_xaxes(showgrid=False)
    return fig


# ============ BARRA LATERAL ============
with st.sidebar:
    st.image(RUTA_LOGO, width=120)
    if st.button("Actualizar datos ahora", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.caption("La página se actualiza sola cada 10 minutos. La nómina cambia cada mes, la curva S cada semana, el cobro con cada acta y la app todos los días.")

# Vuelve a leer todo cada 10 minutos aunque nadie toque la página (pantalla de obra o reunión).
if hasattr(st, "fragment"):
    import time

    @st.fragment(run_every="10m")
    def _auto_refresco():
        t0 = st.session_state.setdefault("_t0", time.time())
        if time.time() - t0 > 590:
            st.session_state["_t0"] = time.time()
            st.cache_data.clear()
            st.rerun()

    _auto_refresco()

# ============ DATOS ============
try:
    with st.spinner("Leyendo nómina, curva S, cuadro de hitos y parámetros..."):
        n, aj, pa, mop, P, H, I, V, base, ma = cargar_fuentes()
except Exception as e:
    st.error(f"No se pudieron leer las fuentes: {e}")
    st.stop()

R0, meses_nom = gd.base_nomina(n, aj)
app = cargar_app(len(ma) if ma is not None else 0, ma)
est = gd.estimar_con_app(app, meses_nom)
R, meses_nom = gd.base_nomina(n, aj, est)
hoy = datetime.now(ZONA).replace(tzinfo=None)
o = gd.calcular(R, meses_nom, pa, mop, P, H, I, V, base, ma, hoy=hoy)
F, EA, ct = o["frentes"], o["maestros"], o["cortes"]
par = base["par"]

# ============ ENCABEZADO ============
ult_nom = ct["nomina"]
chips = [f"Nómina SINCO hasta <b>{gd.MES_CORTO[ult_nom.month].lower()} {ult_nom.year}</b>",
         f"Curva S <b>{fecha_txt(ct['curva'])}</b>", f"Cobro hasta el <b>acta {ct['acta']}</b>",
         f"Actualizado <b>{fecha_txt(hoy)} {hoy:%H:%M}</b>"]
if ct["estimado"] and ct.get("estimado_hasta") is not None:
    chips.insert(1, f"App al <b>{fecha_txt(pd.Timestamp(ct['estimado_hasta']))}</b>")
st.markdown(f"""<div class="g-head"><h1>Mano de obra · AIRPLAN T1 JMC</h1>
<div class="g-chips">{''.join(f'<span class="g-chip">{c}</span>' for c in chips)}</div></div>""", unsafe_allow_html=True)

he_m = o["he_mes"].tail(3)
pct_he = he_m.pagado.sum() / he_m.proyectado.sum() if len(he_m) else np.nan
sobre_ppto = EA["cierre_plan"] - EA["ppto"]
sin_cobrar = F.sin_cobrar.sum() + o["sin_cobrar_fuera"]
contr_sobre = F.contratado_sobre.sum()
rojos = F[F.semaforo == "Rojo"]
T = o["tendencia"]
peor = F[(F.ganada_ult_mes > 1e6)].assign(r=lambda d: d.nomina_ult_mes / d.ganada_ult_mes).sort_values("r", ascending=False)
frase_mo = ""
if len(peor) and peor.r.iloc[0] > 1.5:
    frase_mo = f" En {peor.frente.iloc[0]} la nómina del último mes fue {str(round(peor.r.iloc[0], 1)).replace('.', ',')} veces lo que ganó el avance."
lead = (f"La nómina va por debajo de lo proyectado: maestros al {EA['real'] / EA['proy_hasta']:.0%} y horas extra al {pct_he:.0%} de lo previsto. "
        + (f"Pero si se llenan las vacantes, los maestros cierran {M(sobre_ppto)} por encima del presupuesto. " if sobre_ppto > 0 else "")
        + f"Hay {M(contr_sobre)} contratados por encima de SINCO y {M(sin_cobrar)} pagados a subcontratistas por obra sin cobrar." + frase_mo).replace("%", " %")
st.markdown(f'<p class="g-lead">{lead}</p>', unsafe_allow_html=True)

kpis = [
    ("Nómina propia acumulada", M(R[R.fuente == "Nómina SINCO"].costo.sum()), "Nómina SINCO" + (f"; mes en curso estimado con la app: {M(o['estimado_mes'])}" if ct["estimado"] else ""), ""),
    ("Maestros al cierre, si se llenan las vacantes", M(sobre_ppto, signo=True), f"Frente al presupuesto de {M(EA['ppto'], 0)}. Tope: {M(EA['tope_mes'])}/mes", "rojo" if sobre_ppto > 0 else "ok"),
    ("Horas extra: usado de lo proyectado", PCT(pct_he), "Últimos 3 meses, tope 40 % sobre el básico", "ok" if pct_he <= 1 else "rojo"),
    ("Contratado por encima de SINCO", M(contr_sobre), "Comprometido aunque no se haya pagado", "rojo" if contr_sobre >= par["rojo_valor"] else "ambar"),
    ("Pagado a subcontratistas sin cobrar", M(sin_cobrar), "Obra pagada que AIRPLAN aún no paga", "rojo" if sin_cobrar >= par["rojo_pagado_sin_cobrar"] else "ambar"),
    ("Frentes en rojo", str(len(rojos)), f"{(F.semaforo == 'Amarillo').sum()} más en amarillo", "rojo" if len(rojos) else "ok"),
]
st.markdown('<div class="g-kpis">' + "".join(f'<div class="g-kpi {c}"><div class="l">{l}</div><div class="v">{v}</div><div class="s">{s}</div></div>' for l, v, s, c in kpis) + "</div>", unsafe_allow_html=True)

# ============ FRENTES QUE REQUIEREN DECISIÓN ============
st.markdown('<div class="g-sec">Frentes que requieren decisión</div>', unsafe_allow_html=True)
alerta = F[F.semaforo.isin(["Rojo", "Amarillo"])].copy()
alerta["peso"] = alerta.contratado_sobre + alerta.sin_cobrar + alerta.valor_sobre_apu + np.where(alerta.pct_gastado > 1, (alerta.real - alerta.meta_ejecutado), 0)
alerta = alerta.sort_values(["semaforo", "peso"], ascending=[False, False])
alerta = pd.concat([alerta[alerta.semaforo == "Rojo"], alerta[alerta.semaforo == "Amarillo"]])
cards = []
for k, r in alerta.iterrows():
    cards.append(f"""<div class="g-card {r.semaforo}"><div class="t"><b>{r.frente}</b><span class="pill {r.semaforo}">{r.semaforo}</span></div>
<div class="r">{r.residente} · ejecutado {PCT(r.pct_ejecutado)} · cobrado {PCT(r.pct_cobrado)}</div>
<div class="n"><div><span>Nómina propia / presupuesto</span><strong>{PCT(r.pct_gastado)}</strong></div>
<div><span>Contratado sobre SINCO</span><strong>{M(r.contratado_sobre)}</strong></div>
<div><span>Pagado sin cobrar</span><strong>{M(r.sin_cobrar)}</strong></div></div>
<div class="d">{r.decidir or 'Vigilar.'}</div></div>""")
st.markdown('<div class="g-cards">' + "".join(cards) + "</div>", unsafe_allow_html=True)

# ============ PESTAÑAS ============
t1, t2, t3, t4, t5, t6 = st.tabs(["Proyecciones", "Gasto y avance", "Cobro", "Horas extra", "Actividades", "Subcontratistas"])

with t1:
    oa = o["oa"]
    rows = [
        ("Maestros: lo corrido del proyecto", EA["proy_hasta"], EA["real"], EA["real"] <= EA["proy_hasta"], f"{EA['real'] / EA['proy_hasta']:.0%} de lo proyectado."),
        ("Maestros al cierre, si se llenan las vacantes del plan", EA["ppto"], EA["cierre_plan"], EA["cierre_plan"] <= EA["ppto"],
         f"Decidir qué vacantes se llenan. Para cerrar dentro del presupuesto, máximo {M(EA['tope_mes'])} al mes."),
        ("Maestros al cierre, al ritmo del último mes", EA["ppto"], EA["cierre_ritmo"], EA["cierre_ritmo"] <= EA["ppto"], f"Ritmo actual {M(EA['ritmo'])} al mes."),
        ("Oficiales y ayudantes proyectados en la A", oa["proy"], oa["real_a"], None, "Trabajan en actividades: sacarlos de la proyección de la A para no contarlos dos veces."),
        ("Horas extra (últimos 3 meses)", he_m.proyectado.sum(), he_m.pagado.sum(), pct_he <= 1,
         f"Bajar el supuesto a 30–35 % libera {M(EA['cierre_plan'] - EA['cierre_he_real'])} en maestros al cierre."),
        ("MO propia en actividades (acumulado)", T.ganada_acum.iloc[-1], T.nomina_acum.iloc[-1], T.nomina_acum.iloc[-1] <= T.ganada_acum.iloc[-1],
         f"Nómina ÷ MO ganada: {T.nomina.iloc[0] / T.ganada.iloc[0]:.0%} el primer mes, {T.nomina.iloc[-1] / T.ganada.iloc[-1]:.0%} el último."),
        ("Cobro programado en hitos", F.programado.sum(), F.cobrado.sum(), F.cobrado.sum() >= 0.95 * F.programado.sum(),
         f"{M(F.ejecutado.sum() - F.cobrado.sum(), 0)} ejecutados sin cobrar. Acta {ct['acta'] + 1}: {M(F.prog_siguiente.sum(), 0)} programados."),
    ]
    html = '<table class="g-t"><tr><th>Proyección</th><th>Proyectado</th><th>Real o esperado</th><th>¿Se cumple?</th><th>Qué hacer</th></tr>'
    for a, p, rv, okk, q in rows:
        pill = '<span class="pill Amarillo">Revisar</span>' if okk is None else ('<span class="pill Verde">Sí</span>' if okk else '<span class="pill Rojo">No</span>')
        html += f'<tr><td>{a}</td><td class="n">{M(p)}</td><td class="n">{M(rv)}</td><td>{pill}</td><td>{q.replace("%", " %")}</td></tr>'
    st.markdown(html + "</table>", unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Maestros: proyectado, real y tope mensual** · millones de pesos")
        adm = o["adm_mes"].copy()
        x = [f"{gd.MES_CORTO[p_.month]} {str(p_.year)[2:]}" for p_ in adm.index]
        fig = go.Figure()
        fig.add_bar(x=x, y=adm["Maestros"] / 1e6, name="Proyectado", marker_color="#C9D3CB")
        if "real_Maestros" in adm:
            fig.add_bar(x=x, y=adm["real_Maestros"] / 1e6, name="Real", marker_color=C["verde"])
        fut = adm["real_Maestros"].isna() if "real_Maestros" in adm else pd.Series(True, index=adm.index)
        fig.add_scatter(x=[xi for xi, f in zip(x, fut) if f], y=[EA["tope_mes"] / 1e6] * int(fut.sum()), name="Tope para cerrar en presupuesto",
                        mode="lines", line=dict(color=C["rojo"], dash="dash", width=2))
        fig.update_layout(barmode="group", bargap=.25)
        st.plotly_chart(layout(fig), use_container_width=True)
    with c2:
        st.markdown("**MO propia: nómina contra lo que ganó el avance, por mes** · frentes de obra")
        xm = [("Abr–May" if i == 0 else gd.MES_CORTO[p_.month]) for i, p_ in enumerate(T.index)]
        fig = go.Figure()
        fig.add_bar(x=xm, y=T.ganada / 1e6, name="MO ganada por el avance", marker_color="#C9D3CB")
        fig.add_bar(x=xm, y=T.nomina / 1e6, name="Nómina pagada", marker_color=C["verde"])
        fig.add_scatter(x=xm, y=T.nomina / T.ganada * 100, name="Nómina ÷ ganada (%)", yaxis="y2", mode="lines+markers+text",
                        text=[f"{v:.0%}".replace("%", " %") for v in T.nomina / T.ganada], textposition="top center", line=dict(color=C["dorado"], width=2))
        fig.update_layout(barmode="group", yaxis2=dict(overlaying="y", side="right", range=[0, 160], showgrid=False, ticksuffix=" %"))
        st.plotly_chart(layout(fig), use_container_width=True)
    st.markdown('<p class="g-note">MO ganada = mano de obra propia que SINCO presupuestó para el frente × avance de la curva S, medido unos días después del cierre del mes. Debajo de 100 % se gasta menos de lo que el avance permite.</p>', unsafe_allow_html=True)

with t2:
    c1, c2 = st.columns([1.15, 1])
    with c1:
        st.markdown("**Nómina propia por mes** · millones de pesos")
        g = o["gasto_mes"].copy()
        g["serie2"] = np.where(g.serie.isin(["Administración: maestros", "Administración: oficiales y ayudantes"]), g.serie, None)
        top = g[g.serie2.isna()].groupby("serie").costo.sum().sort_values(ascending=False).index[:5]
        g["serie2"] = g.serie2.fillna(g.serie.where(g.serie.isin(top), "Otros frentes y guías"))
        piv = g.pivot_table(index="mes", columns="serie2", values="costo", aggfunc="sum").fillna(0)
        orden = ["Administración: maestros", "Administración: oficiales y ayudantes"] + list(top) + ["Otros frentes y guías"]
        fig = go.Figure()
        xm = [f"{gd.MES_CORTO[p_.month]}" for p_ in piv.index]
        for i, s in enumerate([s for s in orden if s in piv.columns]):
            fig.add_bar(x=xm, y=piv[s] / 1e6, name=s, marker_color=SERIES[i % len(SERIES)])
        tot = piv.sum(axis=1) / 1e6
        fig.add_scatter(x=xm, y=tot, mode="text", text=[f"{v:,.0f}".replace(",", ".") for v in tot], textposition="top center", showlegend=False)
        fig.update_layout(barmode="stack", bargap=.3, yaxis_range=[0, tot.max() * 1.15])
        st.plotly_chart(layout(fig, 380), use_container_width=True)
        if ct["estimado"]:
            st.markdown('<p class="g-note">El último mes es estimado con las horas de la app (horas × tarifa × recargo) hasta que llegue la nómina de SINCO. Las comparaciones contra presupuesto y proyección usan solo la nómina SINCO.</p>', unsafe_allow_html=True)
    with c2:
        st.markdown("**Mano de obra por frente: propia y subcontratada** · millones")
        d = F[(F.real + F.sub_pagado) > 0].copy()
        d = d.assign(t=d.real + d.sub_pagado).sort_values("t")
        fig = go.Figure()
        fig.add_bar(y=d.frente, x=d.real / 1e6, orientation="h", name="Nómina propia", marker_color=C["verde"])
        fig.add_bar(y=d.frente, x=d.sub_pagado / 1e6, orientation="h", name="Subcontratistas de MO (pagado)", marker_color=C["dorado"])
        fig.update_layout(barmode="stack", bargap=.3)
        f2 = layout(fig, 380)
        f2.update_xaxes(gridcolor="#EEF1ED", tickformat=",.0f")
        st.plotly_chart(f2, use_container_width=True)
    st.markdown("**Todos los frentes**")
    tb = F[(F.real > 0) | (F.contratado_sobre > 0) | (F.cobrado > 0)].copy()
    html = ('<table class="g-t"><tr><th>Frente</th><th>Semáforo</th><th>% ejecutado</th><th>% cobrado</th><th>Nómina propia</th><th>Presupuesto por lo ejecutado</th>'
            '<th>% gastado</th><th>Al cierre (ritmo actual)</th><th>Presupuesto total</th><th>Subcontratistas pagado</th>'
            + ('<th>Mes en curso (app, estimado)</th>' if ct["estimado"] else '') + '</tr>')
    for _, r in tb.iterrows():
        sem = f'<span class="pill {r.semaforo}">{r.semaforo}</span>' if r.semaforo else ""
        html += (f'<tr><td>{r.frente}<br><span class="g-note">{r.residente}</span></td><td>{sem}</td><td class="n">{PCT(r.pct_ejecutado)}</td><td class="n">{PCT(r.pct_cobrado)}</td>'
                 f'<td class="n">{M(r.real)}</td><td class="n">{M(r.meta_ejecutado)}</td><td class="n">{PCT(r.pct_gastado)}</td><td class="n">{M(r.cierre)}</td>'
                 f'<td class="n">{M(r.mo_meta)}</td><td class="n">{M(r.sub_pagado)}</td>'
                 + (f'<td class="n">{M(r.estimado_mes)}</td>' if ct["estimado"] else '') + '</tr>')
    st.markdown(html + "</table>", unsafe_allow_html=True)
    st.markdown(f'<p class="g-note">Nómina propia = oficiales y ayudantes cargados a las actividades del frente. Presupuesto por lo ejecutado = MO propia de SINCO × avance de la curva S al {fecha_txt(ct["curva_comparacion"])} (la medición más cercana al corte de la nómina). Al cierre = nómina ÷ % ejecutado.</p>', unsafe_allow_html=True)

with t3:
    d = F[F.venta > 0].copy()
    d["sin"] = d.ejecutado - d.cobrado
    if True:
        st.markdown(f"**Ejecutado sin cobrar por frente** · millones (curva S {fecha_txt(ct['curva'])} − cobrado al acta {ct['acta']})")
        dd = d[d.sin.abs() > 20e6].sort_values("sin")
        fig = go.Figure(go.Bar(y=dd.frente, x=dd.sin / 1e6, orientation="h", marker_color=[C["rojo"] if v > 0 else C["gris"] for v in dd.sin],
                               text=[M(v, 0) for v in dd.sin], textposition="outside", cliponaxis=False))
        f3 = layout(fig, 60 + 30 * len(dd), leyenda=False)
        lo, hi = min(dd.sin.min() / 1e6, 0), max(dd.sin.max() / 1e6, 0)
        f3.update_xaxes(gridcolor="#EEF1ED", range=[lo - 0.25 * (hi - lo), hi + 0.2 * (hi - lo)])
        st.plotly_chart(f3, use_container_width=True)
        st.markdown('<p class="g-note">Gris = cobrado por encima de lo medido en la curva S: verificar la medición.</p>', unsafe_allow_html=True)
        st.markdown("**Programado, cobrado y ejecutado** · millones, costo directo")
        html = f'<table class="g-t"><tr><th>Frente</th><th>Contrato</th><th>Programado al acta {ct["acta"]}</th><th>Cobrado</th><th>Ejecutado</th><th>% cobrado</th><th>Pagado a subcontratistas sin cobrar</th></tr>'
        for _, r in d[(d.cobrado + d.ejecutado + d.programado) > 0].sort_values("sin", ascending=False).iterrows():
            html += (f'<tr><td>{r.frente}</td><td class="n">{M(r.venta, 0)}</td><td class="n">{M(r.programado, 0)}</td><td class="n">{M(r.cobrado, 0)}</td>'
                     f'<td class="n">{M(r.ejecutado, 0)}</td><td class="n">{PCT(r.pct_cobrado)}</td><td class="n">{M(r.sin_cobrar)}</td></tr>')
        html += (f'<tr><td><b>Total</b></td><td class="n"><b>{M(d.venta.sum(), 0)}</b></td><td class="n"><b>{M(d.programado.sum(), 0)}</b></td><td class="n"><b>{M(d.cobrado.sum(), 0)}</b></td>'
                 f'<td class="n"><b>{M(d.ejecutado.sum(), 0)}</b></td><td class="n"><b>{PCT(d.cobrado.sum() / d.venta.sum())}</b></td><td class="n"><b>{M(sin_cobrar)}</b></td></tr>')
        st.markdown(html + "</table>", unsafe_allow_html=True)
        st.markdown(f'<p class="g-note">Pagado a subcontratistas sin cobrar: en los ítems con contrato de MO, lo pagado menos (contrato × % cobrado del ítem). Incluye {M(o["sin_cobrar_fuera"])} en ítems no presupuestados.</p>', unsafe_allow_html=True)

with t4:
    hm = o["he_mes"]
    c1, c2 = st.columns([1.1, 1])
    with c1:
        st.markdown("**Pagado en horas extra y recargos contra lo proyectado (40 %)** · millones")
        xm = [gd.MES_CORTO[p_.month] for p_ in hm.index]
        fig = go.Figure()
        fig.add_bar(x=xm, y=hm.proyectado / 1e6, name="Proyectado (40 % del básico)", marker_color="#C9D3CB")
        fig.add_bar(x=xm, y=hm.pagado / 1e6, name="Pagado", marker_color=C["dorado"], text=[f"{v:.0%}".replace("%", " %") for v in hm.pagado / hm.proyectado], textposition="outside")
        fig.update_layout(barmode="group", bargap=.3)
        st.plotly_chart(layout(fig), use_container_width=True)
    with c2:
        st.markdown("**% pagado sobre el básico, por grupo**")
        hg = o["he_grupo"]
        html = '<table class="g-t"><tr><th>Grupo</th>' + "".join(f"<th>{gd.MES_CORTO[p_.month]}</th>" for p_ in hg.columns) + "</tr>"
        for gname, row in hg.iterrows():
            celdas = ""
            for v in row:
                col = C["rojo"] if v > 0.40 else (C["ambar"] if v >= 0.30 else C["tinta"])
                celdas += f'<td class="n" style="color:{col};font-weight:{700 if v >= 0.30 else 400}">{PCT(v)}</td>'
            html += f"<tr><td>{gname}</td>{celdas}</tr>"
        st.markdown(html + "</table>", unsafe_allow_html=True)
        st.markdown('<p class="g-note">Rojo: sobre el 40 % proyectado. Ámbar: entre 30 % y 40 %. Solo personas con el mes completo.</p>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Dónde se pagó por encima del 40 %** · por frente")
        hf = o["he_frente"].sort_values(ascending=True)
        etq = [("Administración (A)" if k == "A" else F.frente.get(k, k)) for k in hf.index]
        hf = hf[hf > 50000]
        etq = [("Administración (A)" if k == "A" else F.frente.get(k, k)) for k in hf.index]
        fig = go.Figure(go.Bar(y=etq, x=hf / 1e6, orientation="h", marker_color=C["dorado"], text=[M(v) for v in hf], textposition="outside", cliponaxis=False))
        f4 = layout(fig, 60 + 32 * len(hf), leyenda=False)
        f4.update_xaxes(gridcolor="#EEF1ED", range=[0, hf.max() / 1e6 * 1.25])
        st.plotly_chart(f4, use_container_width=True)
    with c2:
        st.markdown("**Personas que pasaron el 40 %**")
        hp = o["he_personas"]
        html = '<table class="g-t"><tr><th>Persona</th><th>Grupo</th><th>Meses</th><th>Pagado sobre el 40 %</th><th>Máximo</th></tr>'
        for _, r in hp.iterrows():
            html += f'<tr><td>{str(r.nombre).title()}</td><td>{r.grupo}</td><td>{r.meses}</td><td class="n">{M(r.exceso)}</td><td class="n">{PCT(r.maximo)}</td></tr>'
        st.markdown(html + "</table>", unsafe_allow_html=True)
    st.markdown("**Acciones:** tope de 40 % por persona y por mes; toda hora extra autorizada por el residente y ligada a una actividad; "
                "para trasiegos, escombros y demoliciones, evaluar volqueta, minicargador o subcontrato a precio fijo; rotar personal antes de pagar extras a las mismas personas.")

with t5:
    A = o["actividades"].copy()
    A = A[(A.tipo == "Nómina por encima del APU") | ((A.tipo == "Sin MO en el APU") & (A.nomina >= 1e6))]
    Cc = o["contratos"]
    Cc = Cc[Cc.sobre > 0]
    filas = [dict(cod=k, fr=r.fr, act=r.desc, tipo=r.tipo, pc=r.pc, nomina=r.nomina, ref=r.apu, valor=r.valor) for k, r in A.iterrows()]
    filas += [dict(cod=r["Código"], fr=r.fr, act=r["Descripción"], tipo="Contratado por encima de SINCO", pc=r.pc, nomina=r.nomina,
                   ref=r["Presupuesto SINCO comparable"], valor=r.sobre) for _, r in Cc.iterrows()]
    T5 = pd.DataFrame(filas)
    T5["frente"] = T5.fr.map(lambda k: F.frente.get(k, "No presupuestadas" if k == "19" else k))
    c1, c2 = st.columns([1, 1])
    with c1:
        fsel = st.multiselect("Frente", sorted(T5.frente.unique()), placeholder="Todos los frentes")
    with c2:
        tsel = st.multiselect("Tipo de alerta", sorted(T5.tipo.unique()), placeholder="Todas las alertas")
    if fsel:
        T5 = T5[T5.frente.isin(fsel)]
    if tsel:
        T5 = T5[T5.tipo.isin(tsel)]
    T5 = T5.sort_values("valor", ascending=False)
    st.markdown(f"**{len(T5)} actividades · {M(T5.valor.sum())} por encima**")
    T5["act"] = T5.act.where(T5.act.astype(str).str.strip() != "", "Obra no presupuestada")
    vista = pd.DataFrame({"Por encima (M)": T5.valor / 1e6, "Frente": T5.frente, "Alerta": T5.tipo, "% cobrado": T5.pc,
                          "Nómina propia (M)": T5.nomina / 1e6, "Referencia: APU o SINCO (M)": T5.ref / 1e6, "Código": T5.cod,
                          "Actividad": T5.act.astype(str).str[:110].str.capitalize()})
    st.dataframe(vista, hide_index=True, use_container_width=True, height=520,
                 column_config={"Por encima (M)": st.column_config.NumberColumn(format="%.1f", width="small"),
                                "Frente": st.column_config.TextColumn(width="medium"), "Alerta": st.column_config.TextColumn(width="medium"),
                                "% cobrado": st.column_config.NumberColumn(format="percent", width="small"),
                                "Nómina propia (M)": st.column_config.NumberColumn(format="%.1f", width="small"),
                                "Referencia: APU o SINCO (M)": st.column_config.NumberColumn(format="%.1f", width="small"),
                                "Código": st.column_config.TextColumn(width="small"), "Actividad": st.column_config.TextColumn(width="large")})
    st.markdown('<p class="g-note">Nómina por encima del APU: la nómina propia ya superó toda la mano de obra de oficiales y ayudantes del APU. Contratado por encima de SINCO: contratos de MO + otros contratos del ítem + nómina propia, menos lo que SINCO proyectó.</p>', unsafe_allow_html=True)


with t6:
    try:
        sinco = cargar_sinco()
    except Exception as e:
        sinco = None
        st.error(f"No se pudieron leer los exportes de SINCO: {e}")
    if sinco is None:
        st.info("Para ver esta pestaña, suba a Drive los dos exportes de SINCO (Informe contratos detallado y Detallado de actas) "
                "compartidos con 'Cualquier persona con el enlace', y ponga sus enlaces en URL_SINCO_CONTRATOS y URL_SINCO_ACTAS.")
    else:
        Cs, As = sinco
        sb = gd.subcontratistas(Cs, As, mop, o["items"], P, hoy=hoy)
        k = sb["kpi"]
        nom_fr = lambda c: F.frente.get(c, {"18": "18 - Transportes internos", "19": "19 - No presupuestadas", "TODOS": "Transversales"}.get(c, c))
        st.markdown(f'<p class="g-lead">Hay {M(k["contratado"], 0)} contratados con {k["n_contratos"]} contratos en ítems de obra y se han pagado {M(k["pagado"], 0)} '
                    f'({PCT(k["pagado"] / k["contratado"] if k["contratado"] else np.nan)}). {k["n_sobre"]} ítems ya están contratados por encima de su proyección SINCO, '
                    f'por {M(k["sobre"], 0)}: es costo comprometido que aún se puede negociar antes de pagar. '
                    f'Además, {M(k["sin_cobrar"], 0)} pagados superan lo que ya se cobró de esos ítems.</p>'.replace("%", " %"), unsafe_allow_html=True)
        kp = [("Contratado en ítems de obra", M(k["contratado"], 0), f"Proyección SINCO de esos ítems: {M(k['proy_contratado'], 0)}", ""),
              ("Pagado en actas", M(k["pagado"], 0), f"Última acta: {fecha_txt(k['ult_acta'])}" if pd.notna(k["ult_acta"]) else "", ""),
              ("Pendiente por pagar", M(k["pendiente"], 0), "Contratado menos pagado", ""),
              ("Contratado sobre la proyección", M(k["sobre"], 0), f"{k['n_sobre']} ítems", "rojo" if k["sobre"] >= par["rojo_valor"] else "ambar"),
              ("Pagado por encima de lo cobrado", M(k["sin_cobrar"], 0), "Contrato × % cobrado del ítem", "rojo" if k["sin_cobrar"] >= par["rojo_pagado_sin_cobrar"] else "ambar"),
              ("Contratos indirectos", M(k["ind_contratado"], 0), f"Pagado {M(k['ind_pagado'], 0)} (gastos generales y otros)", "")]
        st.markdown('<div class="g-kpis">' + "".join(f'<div class="g-kpi {c}"><div class="l">{l}</div><div class="v">{v}</div><div class="s">{s_}</div></div>' for l, v, s_, c in kp) + "</div>", unsafe_allow_html=True)

        c1, c2 = st.columns([1.15, 1])
        with c1:
            tm = sb["mensual"]
            tm = tm[(tm.pagado > 0) | (tm.ejecutado_acum > 0)]
            st.markdown("**Pagos a contratistas por mes y lo que falta por pagar de lo ya ejecutado** · millones")
            xm = [gd.MES_CORTO[p_.month] for p_ in tm.index]
            fig = go.Figure()
            fig.add_bar(x=xm, y=tm.pagado / 1e6, name="Pagado en el mes", marker_color=C["verde"])
            fig.add_scatter(x=xm, y=tm.pagado_acum / 1e6, name="Pagado acumulado", mode="lines+markers", line=dict(color=C["verde_osc"], width=2))
            fig.add_scatter(x=xm, y=tm.ejecutado_acum / 1e6, name="Contratado ya ejecutado (curva S)", mode="lines+markers", line=dict(color=C["dorado"], width=2, dash="dash"))
            st.plotly_chart(layout(fig), use_container_width=True)
            st.markdown(f'<p class="g-note">Contratado ya ejecutado = contratado de cada ítem × avance de su frente en la curva S. La distancia entre las dos líneas '
                        f'({M(tm.pendiente_estimado.iloc[-1], 0)} hoy) es obra hecha que los contratistas aún no han cobrado: va a llegar en sus próximos cortes.</p>', unsafe_allow_html=True)
        with c2:
            fr_ = sb["frentes"].copy()
            fr_ = fr_[fr_.contratado > 0].sort_values("contratado", ascending=False)
            st.markdown("**Por frente** · millones")
            html = '<table class="g-t"><tr><th>Frente</th><th>Contratado</th><th>Pagado</th><th>Sobre proyección</th><th>Pagado sin cobrar</th></tr>'
            for kf, r in fr_.iterrows():
                col = C["rojo"] if r.sobre >= par["rojo_valor"] else (C["ambar"] if r.sobre >= par["amarillo_valor"] else C["tinta"])
                html += (f'<tr><td>{nom_fr(kf)}</td><td class="n">{M(r.contratado, 0)}</td><td class="n">{M(r.pagado, 0)}</td>'
                         f'<td class="n" style="color:{col};font-weight:700">{M(r.sobre)}</td><td class="n">{M(r.sin_cobrar)}</td></tr>')
            st.markdown(html + "</table>", unsafe_allow_html=True)

        st.markdown("**Por contratista** · millones")
        cc = sb["contratistas"]
        cc = cc[cc.contratado >= 5e6]
        vista = pd.DataFrame({"Contratista": cc.index.str[:45], "Contratos": cc.contratos, "Contratado (M)": cc.contratado / 1e6, "Pagado (M)": cc.pagado / 1e6,
                              "% pagado": cc.pct_pagado * 100, "Sobre proyección (M)": cc.sobre / 1e6, "Pagado sin cobrar (M)": cc.sin_cobrar / 1e6,
                              "Última acta": cc.ult_acta.dt.strftime("%d-%m-%Y").fillna("—")})
        st.dataframe(vista, hide_index=True, use_container_width=True, height=360,
                     column_config={"Contratado (M)": st.column_config.NumberColumn(format="%.1f"), "Pagado (M)": st.column_config.NumberColumn(format="%.1f"),
                                    "% pagado": st.column_config.ProgressColumn(format="%.0f %%", min_value=0, max_value=100),
                                    "Sobre proyección (M)": st.column_config.NumberColumn(format="%.1f"), "Pagado sin cobrar (M)": st.column_config.NumberColumn(format="%.1f")})

        st.markdown("**Ítems contratados por encima de su proyección SINCO** · ordenados por valor")
        itx = sb["items"]
        itx = itx[itx.sobre > 0].sort_values("sobre", ascending=False)
        lect = np.where(itx.pagado <= 0, "Nada pagado: renegociar o tramitar mayor valor antes del primer pago",
                        np.where(itx.pc >= 0.99, "Ya cobrado: confirmar si hay mayor cantidad para cobrar", "Revisar cantidades antes de pagar el saldo"))
        vista = pd.DataFrame({"Sobre (M)": itx.sobre / 1e6, "Frente": itx.fr.map(nom_fr), "Proyectado SINCO (M)": itx.proyectado / 1e6,
                              "Contratado (M)": itx.contratado / 1e6, "Pagado (M)": itx.pagado / 1e6, "% cobrado": itx.pc, "Contratista": itx.contratistas,
                              "Qué hacer": lect, "Código": itx.index, "Actividad": itx.desc.astype(str).str[:100].str.capitalize()})
        st.dataframe(vista, hide_index=True, use_container_width=True, height=480,
                     column_config={"Sobre (M)": st.column_config.NumberColumn(format="%.1f", width="small"),
                                    "Proyectado SINCO (M)": st.column_config.NumberColumn(format="%.1f", width="small"),
                                    "Contratado (M)": st.column_config.NumberColumn(format="%.1f", width="small"),
                                    "Pagado (M)": st.column_config.NumberColumn(format="%.1f", width="small"),
                                    "% cobrado": st.column_config.NumberColumn(format="percent", width="small"),
                                    "Qué hacer": st.column_config.TextColumn(width="large"), "Actividad": st.column_config.TextColumn(width="large")})
        st.markdown('<p class="g-note">Contratado y pagado antes de IVA, con los exportes de SINCO (informe de contratos y detallado de actas). '
                    'Proyección SINCO por ítem: hoja MO_proyectada de Parametros_Alertas. Sobre la proyección = contratado del ítem menos todo su costo proyectado; '
                    'puede ser mayor cantidad (cobrable a AIRPLAN) o mayor precio (a negociar). Pagado sin cobrar = pagado menos (contratado × % cobrado del ítem).</p>', unsafe_allow_html=True)

"""Datos de la página "Control de obra" (pages/0_Control_de_obra.py).

Vista del administrador (control semanal y diario desde abril) y vista de gerencia (alertas de sobrecosto).
Las funciones marcadas "copiada de app.py" son idénticas a las del tablero actual, para que las cifras coincidan.
"""
import io
import re
import json
import unicodedata

import numpy as np
import openpyxl
import pandas as pd

import gerencia_datos as gd


def normalizar_nombre(nombre):  # copiada de app.py
    if not isinstance(nombre, str):
        return ""
    n = nombre.strip().upper()
    n = unicodedata.normalize("NFKD", n).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", n)


# ---------------- copiadas de app.py (mismas reglas de nómina, frentes e histórico semanal) ----------------
INTERV_A_FRENTE = {"3": "01", "3A": "02", "3B": "03", "4": "04", "5": "05", "5A": "06", "5B": "07", "6": "08",
                   "7": "09", "8": "10", "9": "11", "10": "12", "11": "13", "13": "14", "14": "15", "15": "16", "16": "17"}

NOMBRE_FRENTE = {"01": "3 - Check in", "02": "3A - Equipaje saliendo BHS", "03": "3B - Filtro seg. internacional",
                 "04": "4 - Filtro seg. nacional", "05": "5 - Sala embarque remota", "06": "5A - Equipaje llegando",
                 "07": "5B - Sala de espera actual", "08": "6 - Plataforma fase 1", "09": "7 - Nuevo centro conexiones",
                 "10": "8 - Inmigración", "11": "9 - Emigración", "12": "10 - Plataforma fase 2",
                 "13": "11 y 12 - Segundo filtro", "14": "13 - Dobles lectoras", "15": "14 - Guía socioambiental",
                 "16": "15 - Guía SST-SMS", "17": "16 - PMT"}

FRENTES_GUIAS = {"15", "16", "17"}

def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan

MES_CORTO = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

INICIO_OBRA = pd.Period("2026-04", freq="M")

def _grupo_cargo(c):
    c = str(c).lower()
    if "maestro" in c:
        return "Maestros"
    if "oficial" in c:
        return "Oficiales"
    if "ayudante" in c and "sst" not in c:
        return "Ayudantes"
    if "sin confirmar" in c:
        return "Ayudantes"
    return None

def _clave_persona(n):
    """Mismo nombre aunque venga 'APELLIDOS NOMBRES' (SINCO) o 'NOMBRES APELLIDOS' (app)."""
    return " ".join(sorted(re.sub(r"[^A-Z ]", "", normalizar_nombre(n)).split()))

def _base_operativa(mo_base, n, aj, U):
    """Una fila por persona × mes × código con costo, grupo, destino e imputación."""
    avisos = []
    n = n.dropna(subset=["Cedula"]).copy()
    for c in ["Valor", "Porcentaje", "Anio", "Mes"]:
        n[c] = pd.to_numeric(n[c], errors="coerce")
    malos = n[n[["Valor", "Porcentaje", "Anio", "Mes"]].isna().any(axis=1)]
    if len(malos):
        avisos.append(f"{len(malos)} filas de la nómina sin Valor, Porcentaje, Año o Mes válidos (no se tienen en cuenta).")
    n = n.drop(malos.index)
    n["mes"] = pd.PeriodIndex([pd.Period(year=int(a), month=int(m), freq="M") for a, m in zip(n["Anio"], n["Mes"])])
    cod = n["No Item Ppto"].fillna("").astype(str).str.replace(r"\s+", "", regex=True)
    raros = cod[(cod != "") & ~cod.str.match(r"^(\d\d(\.\d\d)+(GG)?|NP-.*|SUP-.*)$", case=False)]
    if len(raros):
        avisos.append(f"Códigos con formato extraño en la nómina (revisar en SINCO, o si se dañaron al pegar): {', '.join(sorted(raros.unique())[:8])}.")
    NO = pd.DataFrame({"mes": n["mes"], "fuente": "Nómina SINCO", "nombre": n["Nombre Empleado"].astype(str).str.strip(),
                       "cargo": n["Cargo"], "cod": cod, "costo": n["Valor"] * n["Porcentaje"], "desc": ""})
    meses_nom = sorted(NO["mes"].unique())
    dup = n.groupby("mes")["Nomina"].nunique()
    for m_, k in dup[dup > 1].items():
        avisos.append(f"{MES_CORTO[m_.month]}: hay {k} nóminas distintas pegadas (columna Nomina). Si una es provisional, bórrala.")

    ap = mo_base.dropna(subset=["fecha"]).copy()
    ap["mes"] = ap["fecha"].dt.to_period("M")
    ult_nom = max(meses_nom) if meses_nom else INICIO_OBRA - 1
    ap = ap[(ap["mes"] > ult_nom) & (ap["mes"] >= INICIO_OBRA)]
    AP = pd.DataFrame({"mes": ap["mes"], "fuente": "App (estimado)", "nombre": ap["nombre"].astype(str), "cargo": ap["cargo"],
                       "cod": ap["codact"].astype(str).str.strip(), "costo": ap["costo_calculado"],
                       "desc": ap["actDesc"].fillna("").astype(str) if "actDesc" in ap.columns else ""})
    if ap["costo_calculado"].isna().any():
        avisos.append(f"{ap.loc[ap['costo_calculado'].isna(), 'nombre'].nunique()} personas de la app sin tarifa: su costo no se cuenta en los meses estimados.")
    R = pd.concat([NO, AP], ignore_index=True)
    R["clave"] = R["nombre"].map(_clave_persona)
    R["grupo"] = R["cargo"].map(_grupo_cargo)
    if len(aj) and {"Nombre", "Grupo"}.issubset(aj.columns):
        mapa = {_clave_persona(a): str(g).strip() for a, g in zip(aj["Nombre"], aj["Grupo"]) if isinstance(a, str) and isinstance(g, str)}
        R["grupo"] = R["clave"].map(mapa).fillna(R["grupo"])
    R = R[R["grupo"].isin(["Maestros", "Oficiales", "Ayudantes"])].copy()

    cu = R["cod"].str.upper()
    R["fr"] = R["cod"].str.extract(r"^(\d\d)\.")[0].fillna(R["cod"].str.extract(r"^(?:NP|SUP)-(\d\d)", flags=re.I)[0])
    desc = R["desc"].str.upper()
    camp = (desc.str.contains(str(U.get("palabras_campamento", "CAMPAMENT")), regex=True)
            & ~desc.str.contains(str(U.get("palabras_no_campamento", "TRASLAD|RETIR|DESMONT|TRASIEG|ACARRE")), regex=True))
    R["destino"] = np.select(
        [R["cod"] == "", cu.str.endswith("GG"), R["fr"].isin(FRENTES_GUIAS), cu.str.startswith("SUP"), camp],
        ["Sin código", "Código GG", "Guías (15, 16, 17)", "Supervisión", "Campamento"], "Ítem de obra")
    R["imputa"] = np.where(R["grupo"] == "Maestros", "A",
                  np.where(R["destino"].isin(["Código GG", "Campamento", "Supervisión"]), "A",
                  np.where(R["destino"] == "Guías (15, 16, 17)", "Guías",
                  np.where(R["destino"] == "Sin código", "Sin código", "APU"))))
    sin = R[(R["destino"] == "Sin código")]
    if len(sin):
        avisos.append("Sin código de ítem: " + ", ".join(f"{r.nombre} ({MES_CORTO[r.mes.month]}, $ {r.costo/1e6:,.1f} M)"
                                                         for r in sin.groupby(["nombre", "mes"]).costo.sum().reset_index().itertuples()) + ".")
    return R, meses_nom, avisos

FESTIVOS_CO = pd.to_datetime([  # festivos de Colombia (ley Emiliani) 2026-2027
    "2026-01-01", "2026-01-12", "2026-03-23", "2026-04-02", "2026-04-03", "2026-05-01", "2026-05-18", "2026-06-08",
    "2026-06-15", "2026-06-29", "2026-07-20", "2026-08-07", "2026-08-17", "2026-10-12", "2026-11-02", "2026-11-16",
    "2026-12-08", "2026-12-25", "2027-01-01", "2027-01-11", "2027-03-22", "2027-03-25", "2027-03-26", "2027-05-01",
    "2027-05-10", "2027-05-31", "2027-06-07", "2027-07-05", "2027-07-20", "2027-08-07", "2027-08-16", "2027-10-18",
    "2027-11-01", "2027-11-15", "2027-12-08", "2027-12-25"])

def _fr_app(df):
    """Frente de cada registro de la app: código de ítem, NP-/SUP-xx o, si no, el frente elegido en la app."""
    cod = df["codact"].astype(str).str.strip()
    fr_cod = cod.where(cod.str.match(r"^\d\d\.")).str[:2]
    np_sup = cod.str.extract(r"^(?:NP|SUP)-(\d\d)", flags=re.I)[0]
    fr_app = pd.to_numeric(df.get("codfrente"), errors="coerce").astype("Int64").astype(str).str.zfill(2)
    return fr_cod.fillna(np_sup).fillna(fr_app.where(fr_app != "<NA>"))

def _historico_semanal(R, mo_base, aj, P, U_):
    """Costo semanal (lunes a sábado) por frente desde abril.
    - Desde el primer día con registros en la app: la app tal cual (no se toca).
    - Antes: nómina SINCO del mes repartida por días hábiles lun-sáb sin festivos.
      Frente: el del código de ítem; los códigos GG (cargos de gastos generales, no dicen frente) se reparten
      (1) según los ítems de esa misma persona en ese mes, (2) según dónde trabajó en la app, solo en frentes
      con avance ese mes en la programación, o (3) quedan 'Sin frente'."""
    # ---- app: solo maestros, oficiales y ayudantes
    ap = mo_base.dropna(subset=["fecha"]).copy()
    ap["clave"] = ap["nombre"].map(_clave_persona)
    ap["grupo"] = ap["cargo"].map(_grupo_cargo)
    if len(aj) and {"Nombre", "Grupo"}.issubset(aj.columns):
        mapa = {_clave_persona(a): str(g).strip() for a, g in zip(aj["Nombre"], aj["Grupo"]) if isinstance(a, str) and isinstance(g, str)}
        ap["grupo"] = ap["clave"].map(mapa).fillna(ap["grupo"])
    ap = ap[ap["grupo"].isin(["Maestros", "Oficiales", "Ayudantes"]) & ap["costo_calculado"].notna()].copy()
    ap["fr"] = _fr_app(ap).fillna("SF")
    # la app se usa desde que el registro quedó completo (17-ago); los registros sueltos anteriores no se suman
    # porque esos días ya los cubre la nómina. Se puede cambiar en Umbrales con el parámetro inicio_app (AAAA-MM-DD).
    ini_app = pd.to_datetime(str(U_.get("inicio_app", "2026-08-17"))[:10], errors="coerce")
    if pd.isna(ini_app):
        ini_app = pd.Timestamp("2026-08-17")
    ap = ap[ap["fecha"] >= ini_app]
    reparto_app = ap.groupby(["clave", "fr"])["costo_calculado"].sum()

    # ---- frentes con avance a cada fin de mes (programación)
    L1 = P[(P["nivel"] == 1) & P["fr"].notna()]
    def activos(m):
        """Frentes con avance DENTRO del mes: ejecutado acumulado al cierre del mes > al cierre del mes anterior."""
        fin = L1[L1["semana"] <= m.end_time].groupby("fr")["EV"].max()
        ini = L1[L1["semana"] <= (m - 1).end_time].groupby("fr")["EV"].max().reindex(fin.index).fillna(0)
        return set(fin[(fin.fillna(0) - ini) > 0].index)

    # ---- nómina: asignar frente
    N = R[R["fuente"] == "Nómina SINCO"].copy()
    directo = N[(N["destino"] != "Código GG") & N["fr"].notna()]
    filas = [(r.mes, r.fr, "Nómina · código del ítem", r.costo) for r in directo.itertuples()]
    propio = directo.groupby(["clave", "mes", "fr"])["costo"].sum()
    for r in N.drop(directo.index).itertuples():
        w, origen = None, None
        if (r.clave, r.mes) in propio.index.droplevel(2):
            w, origen = propio.loc[(r.clave, r.mes)], "Nómina · repartido (ítems de la persona en el mes)"
        elif r.clave in reparto_app.index.get_level_values(0):
            w = reparto_app.loc[r.clave]
            act = activos(r.mes)
            w2 = w[w.index.isin(act)]
            w, origen = (w2 if w2.sum() > 0 else w), "Nómina · repartido (frentes de la persona en la app)"
        if w is None or w.sum() <= 0:
            filas.append((r.mes, "SF", "Nómina · sin frente", r.costo))
        else:
            for fr_, v in (w / w.sum()).items():
                filas.append((r.mes, fr_, origen, r.costo * v))
    NM = pd.DataFrame(filas, columns=["mes", "fr", "origen", "costo"])

    # ---- nómina mensual -> semanas, solo días antes del inicio de la app
    out = []
    for m, d in NM.groupby("mes"):
        dias = pd.date_range(m.start_time, m.end_time.normalize())
        dias = dias[(dias.dayofweek < 6) & ~dias.isin(FESTIVOS_CO)]
        if len(dias) == 0:
            continue
        usa = dias[dias < ini_app]
        if len(usa) == 0:
            continue
        lun = pd.Series(usa - pd.to_timedelta(usa.dayofweek, unit="D")).value_counts() / len(dias)
        for r in d.groupby(["fr", "origen"])["costo"].sum().reset_index().itertuples():
            for l, w in lun.items():
                out.append((l, r.fr, r.origen, r.costo * w))
    H = pd.DataFrame(out, columns=["lunes", "fr", "origen", "costo"])
    ap["lunes"] = (ap["fecha"] - pd.to_timedelta(ap["fecha"].dt.dayofweek, unit="D")).dt.normalize()
    A = ap.groupby(["lunes", "fr"])["costo_calculado"].sum().reset_index().rename(columns={"costo_calculado": "costo"})
    A["origen"] = "App (registro diario)"
    H = pd.concat([H, A], ignore_index=True)
    H["Frente"] = H["fr"].map(NOMBRE_FRENTE).fillna(H["fr"].map(lambda x: "Sin frente" if x == "SF" else "19 - Transversal (no presup.)" if x == "19" else f"Frente {x}"))
    return H, ini_app

# =====================================================================================
# Conjuntos de datos para la página "Control de obra" (vista del administrador y de gerencia).
# Todo se calcula con las mismas reglas de app.py y gerencia_datos.py; aquí solo se arman
# las tablas que dibuja el componente HTML (componentes/control_obra/index.html).
# =====================================================================================

def _leer_programacion_bytes(b):
    """Igual que cargar_programacion() de app.py, pero desde los bytes del archivo de la curva S."""
    wb = openpyxl.load_workbook(io.BytesIO(b), read_only=True, data_only=True)
    rows = list(wb["AVANCE SEMANA A SEMANA"].iter_rows(values_only=True))
    h2 = rows[1]
    semanas = []
    c = 10
    while c < len(h2) and h2[c] is not None:
        semanas.append((pd.Timestamp(h2[c]).normalize(), c))
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


def leer_umbrales(b_params):
    um = pd.read_excel(io.BytesIO(b_params), sheet_name="Umbrales")
    U = dict(zip(um["Parámetro"].astype(str), um["Valor"]))
    for k, v in list(U.items()):
        try:
            U[k] = float(v)
        except (TypeError, ValueError):
            U[k] = str(v)
    return U


def leer_cortes(b_params):
    try:
        c = pd.read_excel(io.BytesIO(b_params), sheet_name="Cortes")
        return {int(a): pd.Timestamp(f) for a, f in zip(c["Corte"], c["Fecha"]) if pd.notna(a) and pd.notna(f)}
    except Exception:
        return {}


def preparar_app(mo_raw, maestro, recargos, excluidos):
    """Registros de la app como los deja cargar_datos() de app.py (horas, tarifa, recargo, costo)."""
    mo = mo_raw.dropna(subset=["nombre"]).copy()
    ultima = mo.columns[-1]
    if ultima != "loteId":
        mo = mo.rename(columns={ultima: "loteId"})
    if "codact" not in mo.columns and "codcat" in mo.columns:
        mo = mo.rename(columns={"codcat": "codact"})
    mo["fecha"] = pd.to_datetime(mo["fecha"], errors="coerce")
    mo["horas"] = pd.to_numeric(mo["horas"].astype(str).str.replace(",", ".", regex=False), errors="coerce").fillna(0)
    mo = gd.quitar_excluidos(mo, excluidos)
    mo["nombre_norm"] = mo["nombre"].map(gd.normalizar_nombre)
    t = maestro.copy()
    t["nombre_norm"] = t["NombreNormalizado"].map(gd.normalizar_nombre)
    t["TarifaHoraReal"] = pd.to_numeric(t["TarifaHoraReal"], errors="coerce")
    mo = mo.merge(t[["nombre_norm", "TarifaHoraReal"]].drop_duplicates("nombre_norm"), on="nombre_norm", how="left")
    mo = mo.merge(recargos[["TipoHora", "Factor"]], left_on="tipoHora", right_on="TipoHora", how="left")
    mo["costo_calculado"] = mo["horas"] * mo["TarifaHoraReal"] * mo["Factor"]
    return mo


def _r(v, d=3):
    return None if v is None or (isinstance(v, float) and np.isnan(v)) else round(float(v), d)


def _M(v):
    """Pesos a millones con 3 decimales (la página los muestra en miles para gerencia)."""
    return None if v is None or (isinstance(v, float) and np.isnan(v)) else round(float(v) / 1e6, 3)


GRUPO_CORTO = lambda c: ("Maestro" if "maestro" in str(c).lower() and "auxiliar" not in str(c).lower()
                         else "Ayudante" if "ayudante" in str(c).lower() else "Oficial")


def datos_administrador(mo, maestro, excluidos, H, inicio_app, hoy):
    """Registro diario por persona, personal, costo por día y frente, semanas e histórico por frente."""
    NORM = lambda x: gd.normalizar_nombre(x).replace("JARRISON", "JARINSON")
    r = mo.dropna(subset=["fecha"]).copy()
    r["nn"] = r["nombre_norm"]
    r["ad"] = r["actDesc"].astype(str)
    lote = r["loteId"].astype(str) if "loteId" in r.columns else pd.Series("", index=r.index)
    r["es_cant"] = r["ad"].str.startswith("[CANTIDAD SEMANAL]") | lote.str.startswith("CANT_")
    r["es_nov"] = r["ad"].str.startswith("[NOVEDAD]") | r["codact"].astype(str).str.startswith("NOV")
    D1 = min(pd.Timestamp(hoy).normalize(), r["fecha"].max())
    w = r[(r["fecha"] >= inicio_app) & (r["fecha"] <= D1)].copy()
    w["costo"] = np.where(w["es_cant"] | w["es_nov"], 0, w["costo_calculado"].fillna(0))
    w["conc"] = w["registradoPor"].astype(str).str.upper().str.startswith(("CONCILIACI", "AJUSTE"))
    w["rpn"] = w["registradoPor"].astype(str).map(NORM)
    TP = lambda tp: "dom" if str(tp).startswith("Dominical") else "noc" if "Nocturn" in str(tp) else "ext" if str(tp).startswith("Extra") else "ord"
    w["b"] = w["tipoHora"].map(TP)
    base = w[~w["conc"] & ~w["es_cant"] & (w["fecha"] >= D1 - pd.Timedelta(days=15))]
    resp = base.groupby("nn")["rpn"].agg(lambda s: s.value_counts().index[0])
    resp_all = w[~w["conc"] & ~w["es_cant"]].groupby("nn")["rpn"].agg(lambda s: s.value_counts().index[0])
    t = maestro.copy()
    t["nn"] = t["NombreNormalizado"].map(gd.normalizar_nombre)
    NM = {gd.normalizar_nombre(x): " ".join(str(x).title().split()) for x in t["NombreCompleto"].dropna()}
    titulo = lambda s: NM.get(NORM(s), " ".join(str(s).title().split()))
    ro = t[t["CargoReal"].astype(str).str.contains("Maestro|Oficial|Ayudante", case=False)
           & ~t["CargoReal"].astype(str).str.contains("Almacen|Servicios Generales", case=False)].copy()
    claves_ex = {k for k, d in excluidos}
    ro = ro[~ro["NombreNormalizado"].map(lambda s: " ".join(sorted(gd.normalizar_nombre(s).split()))).isin(claves_ex)]
    ultimo = w[~w["es_cant"]].groupby("nn")["fecha"].max()
    pers = []
    for _, p in ro.drop_duplicates("nn").iterrows():
        m = resp.get(p.nn, resp_all.get(p.nn, ""))
        pers.append(dict(id=p.nn, nombre=" ".join(titulo(p.NombreCompleto).split()), cargo=p.CargoReal, grupo=GRUPO_CORTO(p.CargoReal),
                         maestro=titulo(m) if m else "Sin maestro asignado",
                         ultimo_registro=str(ultimo.get(p.nn).date()) if p.nn in ultimo.index else None, en_maestro_personal=True))
    ids = {p["id"] for p in pers}
    otros = w[~w["es_cant"] & ~w["nn"].isin(ids) & w["cargo"].astype(str).str.contains("Maestro|Oficial|Ayudante|Sin confirmar", case=False)]
    for nn, y in otros.sort_values("fecha").groupby("nn"):
        m = resp_all.get(nn, "")
        pers.append(dict(id=nn, nombre=" ".join(titulo(y["nombre"].iloc[-1]).split()), cargo=y["cargo"].iloc[-1], grupo=GRUPO_CORTO(y["cargo"].iloc[-1]),
                         maestro=titulo(m) if m else "Sin maestro asignado", ultimo_registro=str(y["fecha"].max().date()), en_maestro_personal=False))
    P = {p["id"]: p for p in pers}

    x = w[~w["es_cant"] & w["nn"].isin(P.keys())].copy()
    # columnas por fila (vectorizado) y luego suma por persona y día
    tp = x["tipoHora"].astype(str)
    cod = x["codact"].astype(str)
    hv = np.where(x["es_nov"], 0.0, x["horas"].astype(float))
    x["h_tot"] = hv
    x["h_ext"] = np.where(tp.str.contains("Extra"), hv, 0.0)
    x["h_noc"] = np.where(tp.str.contains("Nocturn"), hv, 0.0)
    x["h_dom"] = np.where(tp.str.startswith("Dominical"), hv, 0.0)
    x["h_np"] = np.where(cod.str.startswith("NP"), hv, 0.0)
    x["h_sup"] = np.where(cod.str.startswith("SUP"), hv, 0.0)
    x["fr2"] = x["codfrente"].astype(str).str.split(".").str[0].str.strip().where(lambda s: ~s.isin(["", "nan", "None"]), "").str.zfill(2).replace("00", "")
    x["rp_t"] = x["registradoPor"].astype(str).map(titulo)
    x["nov_t"] = np.where(x["es_nov"], x["ad"].str.replace("[NOVEDAD] ", "", regex=False).str.strip(), "")
    x["act"] = [None if nv else {"cod": c_, "ds": d_[:110], "hr": round(float(h_), 2), "tp": b_, "fr": f_}
                for nv, c_, d_, h_, b_, f_ in zip(x["es_nov"], cod, x["ad"], x["horas"], x["b"], x["fr2"])]
    x["dstr"] = x["fecha"].dt.strftime("%Y-%m-%d")
    g = x.groupby(["nn", "dstr"], sort=True)
    sums = g[["h_tot", "h_ext", "h_noc", "h_dom", "h_np", "h_sup"]].sum()
    otros_ = g.agg(nov=("nov_t", lambda s: "; ".join(sorted({v for v in s if v}))), rp=("rp_t", lambda s: " / ".join(sorted(set(s)))),
                   env=("creado", "min"), conc=("conc", "all"), acts=("act", lambda s: [a for a in s if a]))
    dias = []
    for (nn, d), s_, o_ in zip(sums.index, sums.itertuples(index=False), otros_.itertuples(index=False)):
        acts = o_.acts
        dias.append(dict(clave=f"{nn}|{d}", id=nn, fecha=d, horas=round(s_.h_tot, 2), ordinarias=round(s_.h_tot - s_.h_ext, 2), extras=round(s_.h_ext, 2),
                         nocturnas=round(s_.h_noc, 2), dominicales=round(s_.h_dom, 2), no_presupuestadas=round(s_.h_np, 2), supervision=round(s_.h_sup, 2),
                         frentes=",".join(sorted({a["fr"] for a in acts if a["fr"]})), novedad=o_.nov, registrado_por=o_.rp,
                         enviado=str(o_.env)[:16], conciliado=bool(o_.conc), actividades=json.dumps(acts, ensure_ascii=False)))

    FN = {**NOMBRE_FRENTE}
    x2 = x[~x["es_nov"]].copy()
    x2["f"] = x2["codfrente"].astype(str).str.split(".").str[0].str.zfill(2)
    x2["grupo"] = x2["nn"].map(lambda k: P[k]["grupo"])
    cd = x2.groupby([x2["fecha"].dt.strftime("%Y-%m-%d"), "f", "grupo"]).agg(horas=("horas", "sum"), costo=("costo", "sum")).reset_index()
    costo_dia = [dict(clave=f"{c.fecha}|{c.f}|{c.grupo}", fecha=c.fecha, frente=c.f, frente_nombre=FN.get(c.f, "Frente " + c.f), grupo=c.grupo,
                      horas=round(c.horas, 2), costo=round(c.costo)) for c in cd.itertuples()]

    c = w[w["es_cant"]].copy()
    c["lunes"] = (c["fecha"] - pd.to_timedelta(c["fecha"].dt.weekday, "D")).dt.strftime("%Y-%m-%d")
    cierres = []
    for (l, m), y in c.groupby(["lunes", c["registradoPor"].astype(str).map(titulo)]):
        cierres.append(dict(clave=f"{l}|{m}", lunes=l, maestro=m, lineas=int(len(y)),
                            frentes=",".join(sorted(set(y["codfrente"].dropna().astype(str).str.split(".").str[0].str.zfill(2))))))

    # semanas: costo del histórico (nómina repartida antes de la app, app después) + estadísticas de la app
    Hh = H.copy()
    Hh["fuente"] = np.where(Hh["origen"].str.startswith("App"), "App", "Nómina SINCO repartida")
    z = x[~x["es_nov"]].copy()
    z["lunes"] = (z["fecha"] - pd.to_timedelta(z["fecha"].dt.weekday, "D")).dt.normalize()
    z["gr"] = z["nn"].map(lambda k: P[k]["grupo"])
    semanas = []
    for l, y in Hh.groupby("lunes"):
        fu = "App" if (y["fuente"] == "App").any() else "Nómina SINCO repartida"
        o = dict(lunes=str(pd.Timestamp(l).date()), fuente=fu, costo=round(float(y["costo"].sum())), horas=None, personas=None, extra=None, np=None, dias=None, costo_maestros=None)
        za = z[z["lunes"] == l]
        if fu == "App" and len(za) and za["horas"].sum() > 0:
            hs = za["horas"].sum()
            o.update(horas=round(float(hs), 1), personas=int(za["nn"].nunique()), dias=int(za["fecha"].nunique()),
                     extra=round(float(za[za["tipoHora"].astype(str).str.contains("Extra")]["horas"].sum() / hs), 3),
                     np=round(float(za[za["codact"].astype(str).str.startswith("NP")]["horas"].sum() / hs), 3),
                     costo_maestros=round(float(za[za["gr"] == "Maestro"]["costo"].sum())))
        semanas.append(o)
    hf = Hh.groupby(["lunes", "fr", "fuente"])["costo"].sum().reset_index()
    historico = [dict(clave=f"{pd.Timestamp(h.lunes).date()}|{h.fr}", lunes=str(pd.Timestamp(h.lunes).date()), frente=h.fr,
                      frente_nombre=FN.get(h.fr, "Sin frente" if h.fr == "SF" else "Frente " + str(h.fr)), fuente=h.fuente, costo=round(h.costo))
                 for h in hf.itertuples() if abs(h.costo) > 0.5]
    return dict(personal=pers, registro_diario=dias, costo_dia_frente=costo_dia, semanas=semanas, historico_frente=historico, cierres_cantidades=cierres)


def _personas_nomina(R, semanas):
    """Para semanas anteriores a la app: personas y costo de maestros de la nómina del mes."""
    nom = R[R["fuente"] == "Nómina SINCO"]
    pm = nom.groupby("mes")["clave"].nunique().to_dict()
    pmg = nom.groupby(["mes", "grupo"])["costo"].sum()
    for s in semanas:
        if s["fuente"] == "App":
            continue
        mes = pd.Period(pd.Timestamp(s["lunes"]) + pd.Timedelta(days=2), "M")
        s["personas"] = int(pm.get(mes, 0))
        if mes in pmg.index.get_level_values(0):
            tot = float(pmg.loc[mes].sum())
            s["costo_maestros"] = round(s["costo"] * float(pmg.get((mes, "Maestros"), 0)) / tot) if tot else None
    return semanas


def leer_relacion_pagos(b):
    """SINCO > 101 Informe relación de pagos (xls en HTML): anticipo y amortización por contrato."""
    p = pd.read_html(io.StringIO(b.decode("latin-1", errors="ignore")))[0]
    p.columns = ["Fecha", "Contrato", "Beneficiario", "NIT", "Acta", "Factura", "Neto", "IVA", "Total", "Anticipo", "Garantias", "FIC", "Amort",
                 "OtrosDesc", "ValorPago", "SemanaPrg", "CausTipo", "CausNo", "CausFecha", "CausValor", "CruceNo", "CruceFecha", "FormaPago",
                 "FechaRealPago", "ProvNo", "ProvFecha", "ProvValor", "NoPago", "DifADPRO", "DifCxP", "x1", "x2"][: len(p.columns)]
    p = p[pd.to_numeric(p["Contrato"], errors="coerce").notna()].copy()
    for c in ["Anticipo", "Amort"]:
        p[c] = pd.to_numeric(p[c], errors="coerce").fillna(0)
    g = p.groupby("Contrato").agg(contratista=("Beneficiario", "first"), anticipo=("Anticipo", "sum"), amortizado=("Amort", "sum")).reset_index()
    return g.rename(columns={"Contrato": "contrato"})


def datos_gerencia(o, S, H, P_curva, cortes, eq_fact, eq_apu, anticipos, hoy):
    """Indicadores, frentes, actividades, MO, subcontratos, equipos y cobro para la vista de gerencia (millones con 3 decimales)."""
    F = o["frentes"].copy()
    F["cod"] = F.index.astype(str).str.zfill(2)
    ejec, cob = float(F["ejecutado"].sum()), float(F["cobrado"].sum())
    EA = o["maestros"]
    k = S["kpi"] if S else {}
    # equipos: facturado (relaciones contra factura) frente a la parte de equipo del APU según el ejecutado
    ef = eq_fact.copy(); ef["frente"] = ef["frente"].astype(str).str.zfill(2)
    ea = eq_apu.copy(); ea["frente"] = ea["frente"].astype(str).str.zfill(2)
    ppto = ea.set_index("frente")["ppto_equipo"]
    piv = ef.pivot_table(index="frente", columns="mes", values="valor", aggfunc="sum").fillna(0)
    meses_eq = sorted(piv.columns)
    equipos = []
    for _, r in F.iterrows():
        f = r.cod
        real = float(piv.loc[f].sum()) if f in piv.index else 0.0
        pp = float(ppto.get(f, 0))
        meta = pp * (r.pct_ejecutado if pd.notna(r.pct_ejecutado) else 0)
        if real or pp:
            equipos.append(dict(f=f, frente=r.frente, ppto=_M(pp), meta=_M(meta), real=_M(real), rel=round(real / meta, 2) if meta > 0 else None,
                                **{f"m{i + 1}": _M(piv.loc[f, m]) if f in piv.index else 0 for i, m in enumerate(meses_eq)}))
    eq_total = float(ef["valor"].sum())
    eq_meta = sum((e["meta"] or 0) for e in equipos) * 1e6
    act_all = o["actividades"].reset_index()
    act = act_all[act_all["tipo"] == "Nómina por encima del APU"].sort_values("valor", ascending=False)
    sin_mo = act_all[act_all["tipo"] == "Sin MO en el APU"]
    ant = anticipos.copy()
    saldo_ant = float((ant["anticipo"] - ant["amortizado"]).sum()) if len(ant) else None
    he = o["he_mes"]
    ind = [
        ("venta", _M(F["venta"].sum())), ("ejecutado", _M(ejec)), ("cobrado", _M(cob)), ("ejecutado_sin_cobrar", _M(ejec - cob)),
        ("sub_contratado", _M(k.get("contratado"))), ("sub_pagado", _M(k.get("pagado"))), ("sub_sobre_sinco", _M(k.get("sobre"))),
        ("sub_lineas_sobre", int(k.get("n_sobre", 0) or 0)), ("sub_sin_cobrar", _M(k.get("sin_cobrar"))),
        ("anticipo_girado", _M(ant["anticipo"].sum()) if len(ant) else None), ("anticipo_amortizado", _M(ant["amortizado"].sum()) if len(ant) else None),
        ("anticipo_por_amortizar", _M(saldo_ant)),
        ("maestros_ppto", _M(EA["ppto"])), ("maestros_cierre_plan", _M(EA["cierre_plan"])), ("maestros_cierre_ritmo", _M(EA["cierre_ritmo"])),
        ("maestros_real", _M(EA["real"])), ("maestros_proy_hasta", _M(EA["proy_hasta"])), ("maestros_tope_mes", _M(EA["tope_mes"])),
        ("he_real", _r(EA["he_real"])), ("he_tope", _r(o["he_tope"])), ("mo_estimado_mes", _M(o["estimado_mes"])),
        ("equipos_facturado", _M(eq_total)), ("equipos_meta", _M(eq_meta)), ("equipos_ppto", _M(ppto.sum())),
        ("act_sobre_apu", _M(act["valor"].sum())), ("act_n_sobre_apu", int(len(act))), ("act_sin_mo", _M(sin_mo["valor"].sum())), ("act_n_sin_mo", int(len(sin_mo))),
    ]
    frentes = []
    for _, r in F.iterrows():
        frentes.append(dict(f=r.cod, frente=r.frente, residente=r.residente, sem=r.semaforo or "Sin dato", venta=_M(r.venta), ejecutado=_M(r.ejecutado), cobrado=_M(r.cobrado),
                            pct_ej=_r(r.pct_ejecutado), pct_cob=_r(r.pct_cobrado), mo_meta=_M(r.meta_ejecutado), mo_real=_M(r.real), pct_gastado=_r(r.pct_gastado, 2),
                            mo_cierre=_M(r.cierre), mo_ppto=_M(r.mo_meta), sobre_apu=_M(r.valor_sobre_apu), n_sobre_apu=int(r.n_sobre_apu or 0),
                            contr_sobre=_M(r.contratado_sobre), sub_sin_cobrar=_M(r.sin_cobrar), ejec_sin_cobrar=_M((r.ejecutado or 0) - (r.cobrado or 0)),
                            decidir=a_miles(r.decidir or "")))
    actividades = [dict(cod=a.cod, f=a.fr, desc=str(a.desc)[:120], nomina=_M(a.nomina), apu=_M(a.apu), exceso=_M(a.valor), personas=int(a.personas)) for a in act.itertuples()]
    maestros_mes = [dict(mes=str(i), proy=_M(rw["Maestros"]), real=_M(rw.get("real_Maestros"))) for i, rw in o["adm_mes"].iterrows()]
    he_mes = [dict(mes=str(i), pagado=_M(rw.pagado), proyectado=_M(rw.proyectado), base=_M(rw.base), pct=_r(rw.pagado / rw.base) if rw.base else None) for i, rw in he.iterrows()]
    T = o["tendencia"]
    mo_mes = [dict(mes=str(i), nomina=_M(rw.nomina), ganada=_M(rw.ganada)) for i, rw in T.iterrows()]
    sub = []
    if S:
        C = S["contratistas"].reset_index()
        top = C.assign(peso=C["sobre"].fillna(0) + C["sin_cobrar"].fillna(0)).sort_values("peso", ascending=False).head(12)
        sub = [dict(c=x.ctta, n=int(x.contratos), contratado=_M(x.contratado), pagado=_M(x.pagado), pct=_r(x.pct_pagado, 2), sobre=_M(x.sobre), sin_cobrar=_M(x.sin_cobrar))
               for x in top.itertuples()]
    # cobro: curva S (programado y ejecutado acumulados por semana) y cobrado acumulado por corte
    tot = P_curva[P_curva["fr"] == "TOTAL"].sort_values("semana")
    cob_acta, fechas = _cobro_por_acta(H, cortes)
    cobro = []
    hoy = pd.Timestamp(hoy).normalize()
    for s in tot.itertuples():
        if s.semana > hoy + pd.Timedelta(days=70):
            break
        c_ = sum(v for a, v in cob_acta.items() if a in fechas and fechas[a] <= s.semana)
        cobro.append(dict(semana=str(s.semana.date()), programado=_M(s.PV), ejecutado=_M(s.EV) if s.EV else None, cobrado=_M(c_) if s.semana <= hoy + pd.Timedelta(days=6) else None))
    return dict(indicadores_gerencia=[dict(indicador=a, valor=b) for a, b in ind], frentes_gerencia=frentes, actividades_sobre_apu=actividades,
                maestros_mes=maestros_mes, horas_extra_mes=he_mes, mo_nomina_vs_ganada=mo_mes, subcontratistas=sub, equipos_frente=equipos,
                equipos_meses=meses_eq, cobro_semana=cobro)


def _cobro_por_acta(H, cortes, mascara=None):
    """Cobrado por acta (suma de los hitos; `mascara` filtra los hitos de un frente) y fecha de cada acta.
    Las actas sin fecha en la pestaña Cortes toman el siguiente día 3 o 24."""
    Hx = H if mascara is None else H[mascara]
    cob_acta = {}
    for col in [c for c in H.columns if str(c).startswith("c") and str(c)[1:].isdigit()]:
        cob_acta[int(str(col)[1:])] = float(pd.to_numeric(Hx[col], errors="coerce").fillna(0).sum())
    fechas = dict(cortes)
    todas = {int(str(c)[1:]): float(pd.to_numeric(H[c], errors="coerce").fillna(0).sum())
             for c in H.columns if str(c).startswith("c") and str(c)[1:].isdigit()}
    for a in sorted(todas):
        if a not in fechas and todas[a] > 0 and fechas:
            f0 = fechas[max(x for x in fechas if x < a)]
            fechas[a] = f0.replace(day=24) if f0.day == 3 else (f0 + pd.offsets.MonthBegin(1)).replace(day=3)
    return cob_acta, fechas


def datos_por_frente(o, S, Cs, As, H, P_curva, cortes, mo, hoy):
    """Las gráficas de gerencia filtradas por frente (al tocar la tarjeta de un frente).
    - cobro_frente: programado y ejecutado (curva S del frente) y cobrado (hitos del frente) acumulados por semana.
    - mo_frente: nómina propia de oficiales y ayudantes contra la MO que paga el avance, por mes.
    - maestros_frente: costo de maestros (nómina SINCO, administración) repartido por las horas de los maestros en el frente según la app;
      en los meses sin nómina, el estimado de la app.
    - subcontratistas_frente: por contratista, solo los ítems del frente."""
    hoy = pd.Timestamp(hoy).normalize()
    out = dict(cobro_frente=[], mo_frente=[], maestros_frente=[], subcontratistas_frente=[])
    # ---- cobro por frente
    Pf = P_curva[(P_curva["fr"] != "TOTAL") & P_curva["fr"].notna()]
    for f, g in Pf.groupby("fr"):
        if f not in H["fr"].astype(str).values and g["PV"].max() <= 0:
            continue
        cob_acta, fechas = _cobro_por_acta(H, cortes, H["fr"].astype(str) == f)
        sem = g.groupby("semana")[["PV", "EV"]].sum().sort_index()
        for s, r in sem.iterrows():
            if s > hoy + pd.Timedelta(days=70):
                break
            c_ = sum(v for a, v in cob_acta.items() if a in fechas and fechas[a] <= s)
            out["cobro_frente"].append(dict(f=f, semana=str(s.date()), programado=_M(r.PV), ejecutado=_M(r.EV) if r.EV else None,
                                            cobrado=_M(c_) if s <= hoy + pd.Timedelta(days=6) else None))
    # ---- MO propia por frente
    rm, dmeta = o["tendencia_fr"]
    for f in rm.columns:
        for mes in rm.index:
            out["mo_frente"].append(dict(f=f, mes=str(mes), nomina=_M(rm.loc[mes, f]), ganada=_M(dmeta.loc[mes, f])))
    # ---- maestros por frente
    ap = mo.dropna(subset=["fecha"]).copy()
    cargo = ap["cargo"].astype(str).str.lower()
    ap = ap[cargo.str.contains("maestro") & ~cargo.str.contains("auxiliar") & ap["costo_calculado"].notna()]
    ap = ap[~ap["actDesc"].astype(str).str.startswith(("[CANTIDAD SEMANAL]", "[NOVEDAD]"))]
    ap["fr"] = _fr_app(ap)
    ap["mes"] = ap["fecha"].dt.to_period("M")
    ap = ap[ap["mes"] >= pd.Period("2026-08", "M")]  # la app registra completo desde el 17-ago
    pm = ap.pivot_table(index="mes", columns="fr", values="costo_calculado", aggfunc="sum").fillna(0)
    adm = o["adm_mes"]
    real_m = adm["real_Maestros"] if "real_Maestros" in adm else pd.Series(dtype=float)
    for mes in pm.index:
        fila = pm.loc[mes]; tot = fila.sum()
        real = real_m.get(mes, np.nan) if mes in real_m.index else np.nan
        for f, v in fila.items():
            if v <= 0:
                continue
            if pd.notna(real) and tot > 0:
                out["maestros_frente"].append(dict(f=f, mes=str(mes), real=_M(real * v / tot), estimado=None, pct=_r(v / tot)))
            else:
                out["maestros_frente"].append(dict(f=f, mes=str(mes), real=None, estimado=_M(v), pct=_r(v / tot) if tot else None))
    # ---- subcontratistas por frente (ítems directos del frente)
    if S is not None and Cs is not None and As is not None:
        it = S["items"]
        frc = it["fr"]
        C2 = Cs.assign(fr=Cs["cod"].map(frc))
        A2 = As.merge(Cs.drop_duplicates("contrato")[["contrato", "ctta"]], on="contrato", how="left").assign(fr=lambda d: d["cod"].map(frc))
        con = C2.dropna(subset=["fr"]).groupby(["fr", "ctta"])["base"].sum()
        pag = A2.dropna(subset=["fr"]).groupby(["fr", "ctta"])["base"].sum()
        ci = C2[C2["cod"].isin(it.index[it["directo"]])].groupby(["ctta", "cod"])["base"].sum().rename("c").reset_index()
        ci["peso"] = ci["c"] / ci["cod"].map(ci.groupby("cod")["c"].sum())
        ci["fr"] = ci["cod"].map(frc)
        ci["sobre"] = ci["peso"] * ci["cod"].map(it["sobre"])
        ci["sin_cobrar"] = ci["peso"] * ci["cod"].map(it["sin_cobrar"])
        ss = ci.groupby(["fr", "ctta"])[["sobre", "sin_cobrar"]].sum()
        T = pd.DataFrame({"contratado": con, "pagado": pag}).join(ss, how="outer").fillna(0).reset_index()
        T = T[(T["fr"] != "—") & ((T["contratado"] > 0) | (T["pagado"] > 0))]
        T["peso"] = T["sobre"] + T["sin_cobrar"]
        for f, g in T.groupby("fr"):
            for x in g.sort_values(["peso", "contratado"], ascending=False).head(12).itertuples():
                out["subcontratistas_frente"].append(dict(f=f, c=x.ctta, contratado=_M(x.contratado), pagado=_M(x.pagado),
                                                          pct=_r(x.pagado / x.contratado, 2) if x.contratado > 0 else None,
                                                          sobre=_M(x.sobre), sin_cobrar=_M(x.sin_cobrar)))
    return out


def a_miles(t):
    """'143,2 M' -> '$ 143.200' dentro de los textos de decisión del tablero gerencial."""
    def f(mo):
        v = int(mo.group(1).replace(".", "")) * 1000 + int(mo.group(2)) * 100
        return "$ " + f"{v:,}".replace(",", ".")
    return re.sub(r"(\d{1,3}(?:\.\d{3})*),(\d) M", f, t or "")


def leer_fuentes(fuentes):
    """Lee una sola vez los archivos (lo más lento: curva S e hitos). El resultado se puede guardar en caché."""
    n, aj = gd.leer_nomina(fuentes["nomina"])
    pa, mop = gd.leer_parametros(fuentes["params"])
    L = dict(n=n, aj=aj, pa=pa, mop=mop, U=leer_umbrales(fuentes["params"]), excluidos=gd.leer_excluidos(fuentes["params"]),
             cortes=leer_cortes(fuentes["params"]), Pc=gd.leer_curva(fuentes["curva"]), Pp=_leer_programacion_bytes(fuentes["curva"]),
             base=gd.leer_base(fuentes["base"]), ma=gd.leer_maestro(fuentes["maestro"]), Cs=None, As=None, rel_pagos=None)
    L["H"], L["I"], L["V"] = gd.leer_hitos(fuentes["hitos"])
    if fuentes.get("contratos") and fuentes.get("actas"):
        L["Cs"] = gd.leer_contratos_sinco(fuentes["contratos"])
        L["As"] = gd.leer_actas_sinco(fuentes["actas"], L["Cs"])
    if fuentes.get("relacion_pagos"):
        try:
            L["rel_pagos"] = leer_relacion_pagos(fuentes["relacion_pagos"])
        except Exception:
            pass
    return L


def construir(fuentes, mo_raw, hoy=None, recargos=None, eq_fact=None, eq_apu=None, anticipos=None, leidas=None):
    """fuentes: bytes de nomina, params, curva, hitos, base, maestro, contratos, actas (y opcional relacion_pagos);
    o bien `leidas` = leer_fuentes(fuentes). mo_raw: la hoja Registros de la app. Devuelve {nombre: filas} para el componente."""
    hoy = pd.Timestamp(hoy or pd.Timestamp.now(tz="America/Bogota").tz_localize(None))
    L = leidas or leer_fuentes(fuentes)
    n, aj, pa, mop, U, excluidos, Pc, base, ma = L["n"], L["aj"], L["pa"], L["mop"], L["U"], L["excluidos"], L["Pc"], L["base"], L["ma"]
    H, I, V = L["H"], L["I"], L["V"]
    mo = preparar_app(mo_raw, ma, recargos, excluidos)
    # gerencia: el mismo cálculo de pages/1_Gerencia.py
    mo_g = mo[~mo["actDesc"].astype(str).str.startswith(("[CANTIDAD SEMANAL]", "[NOVEDAD]"))].copy()
    R0, meses_nom = gd.base_nomina(n, aj)
    est = gd.estimar_con_app(mo_g, meses_nom)
    R, meses_nom = gd.base_nomina(n, aj, est)
    o = gd.calcular(R, meses_nom, pa, mop, Pc, H, I, V, base, ma, hoy=hoy)
    S = gd.subcontratistas(L["Cs"], L["As"], mop, o["items"], Pc, hoy=hoy) if L["Cs"] is not None else None
    if L["rel_pagos"] is not None:
        anticipos = L["rel_pagos"]
    # administrador: histórico semanal con las reglas de app.py
    Rb, _, _ = _base_operativa(mo, n, aj, U)
    Pp = L["Pp"]
    Hs, ini_app = _historico_semanal(Rb, mo, aj, Pp, U)
    adm = datos_administrador(mo, ma, excluidos, Hs, ini_app, hoy)
    adm["semanas"] = _personas_nomina(Rb, adm["semanas"])
    ger = datos_gerencia(o, S, H, Pc, L["cortes"], eq_fact, eq_apu, anticipos, hoy)
    ger.update(datos_por_frente(o, S, L["Cs"], L["As"], H, Pc, L["cortes"], mo, hoy))
    ct = o["cortes"]
    meta = dict(generado=str(hoy)[:16], nomina_hasta=str(ct["nomina"]), curva=str(pd.Timestamp(ct["curva"]).date()), acta=int(ct["acta"]),
                app_hasta=str(mo["fecha"].max().date()), inicio_app=str(ini_app.date()),
                sinco=str(S["kpi"]["ult_acta"].date()) if S and pd.notna(S["kpi"].get("ult_acta")) else None)
    return _limpio({**adm, **ger, "meta": meta})


def _limpio(o):
    """Convierte tipos de numpy y NaN a valores JSON simples."""
    if isinstance(o, dict):
        return {k: _limpio(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_limpio(v) for v in o]
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating, float)):
        return None if np.isnan(o) else float(o)
    if isinstance(o, (np.bool_,)):
        return bool(o)
    return o

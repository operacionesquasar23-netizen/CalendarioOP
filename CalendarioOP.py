"""
CalendarioOP.py - Calendario de Visitas de Operadores (Quasar BTL)

Ejecutar:   streamlit run CalendarioOP.py
Necesita:   motor_calendario.py en la misma carpeta.

Flujo mensual:
  1. Subir el Q del mes, HorarioTiendas.xlsx y el calendario del mes anterior.
  2. "Generar calendario".
  3. Revisar / mover tiendas de día en la pestaña Calendario.
  4. Descargar el Excel (ese archivo es el "calendario anterior" del mes siguiente).
"""
import datetime as dt

import pandas as pd
import pydeck as pdk
import streamlit as st

import motor_calendario as mc

st.set_page_config(page_title="Calendario de Visitas - Operadores", layout="wide")

# Paleta: azul marino #1F3A5F, azul medio #4D7EA8, azul claro #A1C6EA, beige #F2EFEA, arena #C9B29B.
# El tema base está en .streamlit/config.toml; aquí se afinan los detalles.
st.markdown("""
<style>
[data-testid="stMetric"]{background:#fff;border:1px solid #C9B29B;border-radius:10px;padding:14px 18px}
[data-baseweb="tab-highlight"]{background-color:#4D7EA8 !important}
[data-testid="stSidebar"] button[kind="primary"],
[data-testid="stSidebar"] [data-testid="stBaseButton-primary"]{background:#C9B29B;color:#1F3A5F;border:0;font-weight:700}
[data-testid="stSidebar"] button[kind="primary"] p,
[data-testid="stSidebar"] [data-testid="stBaseButton-primary"] p{color:#1F3A5F}
[data-testid="stDownloadButton"] button{background:#1F3A5F;color:#fff;border:0;font-weight:600}
[data-testid="stDownloadButton"] button p{color:#fff}
</style>
""", unsafe_allow_html=True)

DIA_CORTO = {"LUNES": "Lun", "MARTES": "Mar", "MIERCOLES": "Mié",
             "JUEVES": "Jue", "VIERNES": "Vie", "SABADO": "Sáb"}


# Colores del mapa: la paleta + 3 tonos extra para poder distinguir hasta 8 grupos
COLORES = ["#1F3A5F", "#4D7EA8", "#C9B29B", "#B5651D", "#3A8F85", "#7B5E8D", "#A3A847", "#C0504D"]


def _rgb(hex_: str, alfa: int = 230):
    h = hex_.lstrip("#")
    return [int(h[i:i + 2], 16) for i in (0, 2, 4)] + [alfa]


def _casco(puntos):
    """Envolvente convexa (cadena monótona) de una lista de (lon, lat)."""
    pts = sorted(set(puntos))
    if len(pts) < 3:
        return pts

    def cruz(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    bajo, alto = [], []
    for pt in pts:
        while len(bajo) >= 2 and cruz(bajo[-2], bajo[-1], pt) <= 0:
            bajo.pop()
        bajo.append(pt)
    for pt in reversed(pts):
        while len(alto) >= 2 and cruz(alto[-2], alto[-1], pt) <= 0:
            alto.pop()
        alto.append(pt)
    return bajo[:-1] + alto[:-1]


# ---------------------------------------------------------------- lectura (con caché)
@st.cache_data(show_spinner=False)
def _leer_q(contenido: bytes, nombre: str):
    return mc.leer_q(contenido, nombre)


@st.cache_data(show_spinner=False)
def _leer_horario(contenido: bytes):
    return mc.leer_horario(contenido)


@st.cache_data(show_spinner=False)
def _leer_previo(contenido: bytes):
    return mc.leer_calendario_previo(contenido)


def _reiniciar(op: str):
    st.session_state["reset"][op] = st.session_state["reset"].get(op, 0) + 1


if "reset" not in st.session_state:
    st.session_state["reset"] = {}

# ---------------------------------------------------------------- barra lateral
st.sidebar.title("Calendario de visitas")
st.sidebar.header("1. Archivos")
f_q = st.sidebar.file_uploader("Q del mes (Excel o CSV)", type=["xlsx", "xls", "csv"])
f_h = st.sidebar.file_uploader("HorarioTiendas.xlsx", type=["xlsx"])
f_p = st.sidebar.file_uploader("Calendario del mes anterior (opcional)", type=["xlsx"])
if not f_p:
    st.sidebar.caption("Sin calendario anterior se programa todo desde cero. Si tu Q no trae las tiendas "
                       "con Q = 0, agrégalas en 'Tiendas sin Q'.")

q_df = hor = prev = None
info_q = info_h = {}
if f_q and f_h:
    try:
        q_df, info_q = _leer_q(f_q.getvalue(), f_q.name)
        hor, info_h = _leer_horario(f_h.getvalue())
        prev = _leer_previo(f_p.getvalue()) if f_p else {}
    except Exception as e:  # noqa: BLE001 - se muestra el motivo al usuario
        st.sidebar.error(f"No pude leer los archivos: {e}")
        q_df = hor = prev = None

generar = False
if hor is not None:
    por_defecto = mc.Params()
    st.sidebar.header("2. Parámetros")
    with st.sidebar.expander("Reglas de FVS", expanded=False):
        n3 = st.number_input("Tiendas con FVS 3 (primeras por Q)", 0, 100, por_defecto.n_fvs3)
        n2 = st.number_input("Tiendas con FVS 2 (siguientes)", 0, 100, por_defecto.n_fvs2)
        todas = hor["tienda"].tolist()
        claves_hor = {mc.norm(t): t for t in todas}
        def_exc = [claves_hor[mc.norm(t)] for t in por_defecto.exc_tiendas if mc.norm(t) in claves_hor]
        exc_t = st.multiselect("Excepciones FVS 1 - tiendas puntuales", todas, default=def_exc)
        exc_c = st.text_input("Excepciones FVS 1 - cadena (palabras clave)", ", ".join(por_defecto.exc_cadenas))
        exc_z = st.text_input("Excepciones FVS 1 - zona (prefijo)", ", ".join(por_defecto.exc_zonas))
    with st.sidebar.expander("Topes y días", expanded=False):
        cap_dia = st.number_input("Tope de tiendas por operador, lunes a viernes", 1, 20, por_defecto.cap_dia)
        cap_sab = st.number_input("Tope de tiendas por operador, sábado", 1, 20, por_defecto.cap_sab)
        detectadas = mc.detectar_solo_lv(hor)
        solo_lv = st.multiselect("Solo lunes a viernes (sin sábado)", todas, default=detectadas)
        fvs3_sab = st.checkbox("Las tiendas FVS 3 sí pueden ir sábado", por_defecto.fvs3_sabado)
        conservar = st.checkbox("Conservar los días del calendario anterior (tiendas sin cambios)",
                                value=bool(prev), disabled=not prev,
                                help="Si lo desmarcas, se reprograma todo desde cero con el criterio de "
                                     "cercanía. Recomendado la primera vez que uses el mapa y las distancias.")
        min_dia = st.number_input("Mínimo ideal de tiendas por día (L-V)", 0, 20, por_defecto.min_dia,
                                  help="Preferencia: el programador intenta completar cada día con tiendas "
                                       "cercanas hasta llegar a este número. No es obligatorio.")
        dist_km = st.number_input("Avisar si dos tiendas del mismo día están a más de (km)", 1.0, 100.0,
                                  float(por_defecto.dist_aviso_km), step=1.0)
    with st.sidebar.expander("Tiendas sin Q", expanded=False):
        _, info0 = mc.armar_universo(q_df, hor, prev)
        extra = st.multiselect("Agregar al programa (Q = 0)", info0["candidatas_sin_q"], default=[])
    with st.sidebar.expander("Avanzado", expanded=False):
        semilla = st.number_input("Semilla del programador", 0, 99999, por_defecto.seed)
        rondas = st.number_input("Esfuerzo de búsqueda", 10, 500, por_defecto.rondas)
    generar = st.sidebar.button("Generar calendario", type="primary")

if generar:
    p = mc.Params(
        n_fvs3=int(n3), n_fvs2=int(n2), cap_dia=int(cap_dia), cap_sab=int(cap_sab),
        min_dia=int(min_dia), dist_aviso_km=float(dist_km),
        exc_tiendas=tuple(exc_t),
        exc_cadenas=tuple(x.strip() for x in exc_c.split(",") if x.strip()),
        exc_zonas=tuple(x.strip() for x in exc_z.split(",") if x.strip()),
        solo_lv=tuple(solo_lv), fvs3_sabado=bool(fvs3_sab), seed=int(semilla), rondas=int(rondas))
    with st.spinner("Calculando FVS y programando días..."):
        uni, info = mc.armar_universo(q_df, hor, prev, extra)
        u = mc.asignar_fvs(uni, p, prev)
        sched = mc.programar(u, p, prev if conservar else {})
    # ¿tiendas que antes eran Periferia y ahora ya no son excepción?
    antes_per = [prev[k]["tienda"] for k in u.loc[~u["excepcion"], "key"]
                 if k in prev and mc.norm(prev[k].get("zona", "")).startswith("periferia")]
    st.session_state["res"] = {
        "sched": sched, "params": p, "prev": prev, "uni": u, "info": info,
        "info_q": info_q, "info_h": info_h, "antes_periferia": antes_per,
        "run": st.session_state.get("res", {}).get("run", 0) + 1}
    st.session_state["reset"] = {}

res = st.session_state.get("res")
if not res:
    st.title("Calendario de visitas de operadores")
    st.info("Sube en la barra lateral el Q del mes y HorarioTiendas (el calendario del mes anterior es "
            "opcional) y pulsa **Generar calendario**.")
    st.stop()

sched, p, prev, run = res["sched"], res["params"], res["prev"], res["run"]
info = res["info"]
operadores = sorted(sched["Operador"].unique(), key=str.lower)

st.title("Calendario de visitas de operadores")
top = st.container()

# ---------------------------------------------------------------- revisión de datos
avisos_datos = []
if res["antes_periferia"]:
    avisos_datos.append(("warning", "Estas tiendas estaban en Periferia en el calendario anterior y ahora "
                         "no son excepción de FVS 1 (¿cambió su zona?): " + ", ".join(res["antes_periferia"])))
if info["q_sin_horario"]:
    avisos_datos.append(("warning", "En el Q pero no en Horario (Lima), no se programan: " +
                         ", ".join(info["q_sin_horario"])))
if info["sin_operador"]:
    avisos_datos.append(("warning", "Sin Operador2 en Horario, no se programan: " + ", ".join(info["sin_operador"])))
if info["salen_del_programa"]:
    avisos_datos.append(("info", "Estaban en el calendario anterior y ya no están en Horario (Lima): " +
                         ", ".join(info["salen_del_programa"])))
if info["q_cero_por_calendario"]:
    avisos_datos.append(("info", "Sin Q este mes, se mantienen en el programa con Q = 0: " +
                         ", ".join(info["q_cero_por_calendario"])))
if info["agregadas_manual"]:
    avisos_datos.append(("info", "Agregadas manualmente (Q = 0): " + ", ".join(info["agregadas_manual"])))
if res["info_q"].get("q_vacios"):
    avisos_datos.append(("warning", f"{res['info_q']['q_vacios']} tienda(s) con Q vacío en el archivo; "
                         "se tomaron como 0."))
if not res.get("info_h", {}).get("tiene_distrito", True):
    avisos_datos.append(("info", "HorarioTiendas no trae la columna 'Distrito': se agrupa por zona y por distancia."))
if "Lat" in sched.columns:
    sin_coord = sched[sched["Lat"].isna() | sched["Lon"].isna()]["Tienda"].tolist()
    if sin_coord:
        avisos_datos.append(("warning", "Sin latitud/longitud en Horario (no se considera la distancia ni "
                             "salen en el mapa): " + ", ".join(sin_coord)))
if avisos_datos:
    with st.expander(f"Revisión de datos ({len(avisos_datos)})", expanded=True):
        for tipo, msg in avisos_datos:
            (st.warning if tipo == "warning" else st.info)(msg)

def _tabla_carga(df: pd.DataFrame):
    """Tabla de cargas con las celdas que superan el tope resaltadas."""
    topes = {DIA_CORTO[d]: (p.cap_sab if d == "SABADO" else p.cap_dia) for d in mc.DIAS}
    df = df.rename(columns=DIA_CORTO)

    def resaltar(x):
        est = pd.DataFrame("", index=x.index, columns=x.columns)
        for dia, tope in topes.items():
            est[dia] = x[dia].map(lambda v, t=tope: "background-color:#C9B29B; font-weight:700" if v > t else "")
        return est

    return df.style.apply(resaltar, axis=None)

# ---------------------------------------------------------------- pestañas
t_cal, t_hor, t_car, t_cam, t_rank, t_avisos = st.tabs(
    ["Calendario", "Horario semanal", "Carga", "Cambios", "Ranking y FVS", "Avisos"])

editados = {}
with t_cal:
    st.caption("Marca o desmarca los días de cada tienda para moverla. Los avisos (tope por día, días "
               "consecutivos, FVS) no bloquean el cambio. Los cambios viven en esta sesión: "
               "**descarga el Excel antes de cerrar**.")
    pestanas = st.tabs(operadores)
    for op, pest in zip(operadores, pestanas):
        with pest:
            base = (sched[sched["Operador"] == op]
                    .sort_values(["Zona", "FVS", "Q", "Tienda"], ascending=[True, False, False, True])
                    .set_index("key"))
            con_distrito = base["Distrito"].astype(str).str.strip().ne("").any() if "Distrito" in base else False
            cols_info = ["Tienda"] + (["Distrito"] if con_distrito else []) + ["Zona", "Q", "FVS"]
            vista = base[cols_info + mc.DIAS].copy()
            cfg = {d: st.column_config.CheckboxColumn(DIA_CORTO[d], width="small") for d in mc.DIAS}
            cfg["Q"] = st.column_config.NumberColumn("Q", width="small", format="%d")
            cfg["FVS"] = st.column_config.NumberColumn("FVS", width="small", format="%d")
            clave = f"ed_{run}_{st.session_state['reset'].get(op, 0)}_{op}"
            ed = st.data_editor(vista, key=clave, hide_index=True,
                                disabled=cols_info, column_config=cfg)
            editados[op] = (base, ed)
            st.button("Restablecer propuesta de este operador", key=f"rst_{run}_{op}",
                      on_click=_reiniciar, args=(op,))

# calendario final = propuesta + ajustes hechos en pantalla
final = sched.set_index("key").copy()
ajustadas = set()
for op, (base, ed) in editados.items():
    nuevo = ed[mc.DIAS].astype(bool)
    final.loc[nuevo.index, mc.DIAS] = nuevo
    cambiadas = (base.loc[nuevo.index, mc.DIAS].astype(bool) != nuevo).any(axis=1)
    ajustadas.update(cambiadas[cambiadas].index)
final[mc.DIAS] = final[mc.DIAS].astype(bool)
final = final.reset_index()
validaciones = mc.validar(final, p)
opcionales = mc.calcular_opcionales(final, p)

# resumen por operador debajo de cada editor (se muestra dentro de la pestaña Calendario)
with t_cal:
    st.divider()
    st.subheader("Mapa de visitas")
    if "Lat" not in final.columns or final["Lat"].notna().sum() == 0:
        st.info("HorarioTiendas no trae Latitud/Longitud, así que no se puede dibujar el mapa.")
    else:
        m1, m2, m3, m4 = st.columns([2, 2, 2, 2])
        op_map = m1.selectbox("Operador", ["Todos"] + operadores, key="map_op")
        dia_map = m2.selectbox("Día", ["Toda la semana"] + [DIA_CORTO[d] for d in mc.DIAS], key="map_dia")
        color_por = m3.radio("Colorear por", ["Operador", "Zona"], horizontal=True, key="map_color")
        area = m4.checkbox("Mostrar área cubierta", value=True, key="map_area")
        pts = final[final["Lat"].notna() & final["Lon"].notna()].copy()
        if op_map != "Todos":
            pts = pts[pts["Operador"] == op_map]
        if dia_map != "Toda la semana":
            col_dia = [d for d in mc.DIAS if DIA_CORTO[d] == dia_map][0]
            pts = pts[pts[col_dia]]
        if pts.empty:
            st.info("No hay visitas para ese operador y día.")
        else:
            campo = "Operador" if color_por == "Operador" else "Zona"
            grupos = sorted(pts[campo].unique(), key=str.lower)
            tono = {g: COLORES[i % len(COLORES)] for i, g in enumerate(grupos)}
            pts["color"] = pts[campo].map(lambda g: _rgb(tono[g]))
            pts["Días"] = pts.apply(lambda r: mc.abreviar_dias(mc.dias_de(r)), axis=1)
            datos = pts[["Tienda", "Distrito", "Zona", "Operador", "FVS", "Días", "Lat", "Lon", "color"]]
            capas = []
            if area:
                poligonos = []
                for g, sub in pts.groupby(campo):
                    hull = _casco(list(zip(sub["Lon"], sub["Lat"])))
                    if len(hull) >= 3:
                        poligonos.append({"nombre": g, "poligono": [list(p_) for p_ in hull],
                                          "relleno": _rgb(tono[g], 45), "borde": _rgb(tono[g], 200)})
                if poligonos:
                    capas.append(pdk.Layer("PolygonLayer", data=poligonos, get_polygon="poligono",
                                           get_fill_color="relleno", get_line_color="borde",
                                           line_width_min_pixels=2, pickable=False))
            capas.append(pdk.Layer("ScatterplotLayer", data=datos, get_position="[Lon, Lat]",
                                   get_fill_color="color", get_line_color=[255, 255, 255],
                                   get_radius=250, radius_min_pixels=6, stroked=True,
                                   line_width_min_pixels=1, pickable=True))
            vista_mapa = pdk.ViewState(latitude=float(pts["Lat"].mean()), longitude=float(pts["Lon"].mean()),
                                       zoom=10.2 if op_map == "Todos" else 10.8)
            tooltip = {"html": "<b>{Tienda}</b><br/>{Distrito} · {Zona}<br/>Operador: {Operador}<br/>"
                               "FVS {FVS} · Días: {Días}",
                       "style": {"backgroundColor": "#1F3A5F", "color": "white"}}
            st.pydeck_chart(pdk.Deck(layers=capas, initial_view_state=vista_mapa, tooltip=tooltip,
                                     map_style="light"))
            st.markdown(" &nbsp; ".join(f"<span style='color:{tono[g]};font-size:18px'>●</span> {g}"
                                        for g in grupos), unsafe_allow_html=True)
    st.divider()
    st.subheader("Carga por día")
    st.dataframe(_tabla_carga(mc.cargas(final).reindex(operadores)))
    if ajustadas:
        nombres = final[final["key"].isin(ajustadas)]["Tienda"].tolist()
        st.info(f"Tiendas ajustadas a mano ({len(nombres)}): " + ", ".join(nombres))

with t_hor:
    for op in operadores:
        with st.expander(op, expanded=True):
            g = final[final["Operador"] == op].sort_values(["Zona", "Tienda"])
            cols = st.columns(6)
            for d, (nombre, col) in enumerate(zip(mc.DIAS, cols)):
                x = g[g[nombre]]
                col.markdown(f"**{DIA_CORTO[nombre]}** ({len(x)})")
                for _, r in x.iterrows():
                    marca = " ⚠️ opcional" if (r["key"], d) in opcionales else ""
                    col.markdown(f"- {r['Tienda']} · _{mc.abreviar_zona(r['Zona'])}_{marca}")

with t_car:
    tabla = mc.cargas(final).reindex(operadores)
    tabla["Capacidad"] = 5 * p.cap_dia + p.cap_sab
    tabla["Exceso"] = (tabla["TOTAL"] - tabla["Capacidad"]).clip(lower=0)
    st.dataframe(_tabla_carga(tabla))
    st.caption(f"Tope: {p.cap_dia} tiendas por día de lunes a viernes y {p.cap_sab} el sábado (medio día).")

with t_cam:
    if not prev:
        st.info("No subiste calendario anterior, así que no hay nada con qué comparar. "
                "Desde el próximo mes esta pestaña mostrará los cambios.")
    else:
        cam = mc.comparar(final, prev)
        resumen = cam["Estado"].value_counts()
        st.write(" · ".join(f"**{k}**: {v}" for k, v in resumen.items()))
        sel = st.multiselect("Filtrar por estado", sorted(cam["Estado"].unique()),
                             default=[e for e in sorted(cam["Estado"].unique()) if e != "Sin cambios"])
        st.dataframe(cam[cam["Estado"].isin(sel)] if sel else cam, hide_index=True)

with t_rank:
    r = res["uni"].copy()
    r["Excepción"] = r["excepcion"].map({True: "Sí", False: ""})
    r["FVS anterior"] = r["fvs_previo"].replace(0, pd.NA)
    r = r.sort_values(["posicion", "q"], ascending=[True, False], na_position="last")
    r = r.rename(columns={"posicion": "Posición", "tienda": "Tienda", "zona": "Zona",
                          "operador": "Operador", "q": "Q", "fvs": "FVS nuevo"})
    st.caption("Las excepciones no entran al ranking: quedan fijas en FVS 1.")
    st.dataframe(r[["Posición", "Tienda", "Zona", "Operador", "Q", "Excepción", "FVS nuevo", "FVS anterior"]],
                 hide_index=True)

with t_avisos:
    if validaciones.empty:
        st.success("Sin avisos: todas las tiendas cumplen FVS, días no consecutivos y topes por día.")
    else:
        st.dataframe(validaciones, hide_index=True)

# ---------------------------------------------------------------- cabecera: métricas y descarga
with top:
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Tiendas", len(final))
    m2.metric("Operadores", len(operadores))
    m3.metric("Visitas por semana", int(final[mc.DIAS].sum().sum()))
    m4.metric("Avisos", len(validaciones))
    m5.metric("Ajustadas a mano", len(ajustadas))
    excel = mc.exportar_excel(final, p, prev or None, ajustadas)
    st.download_button("Descargar calendario (Excel)", data=excel,
                       file_name=f"Calendario_Visitas_Operadores_{dt.date.today():%Y-%m}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

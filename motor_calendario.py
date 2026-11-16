"""
motor_calendario.py
===================
Motor del Calendario de Visitas de Operadores (Quasar BTL).

Qué hace
--------
1. Lee los 3 archivos mensuales:
     - Q del mes (Excel/CSV con columnas "Tienda" y "Q Doméstico")
     - HorarioTiendas.xlsx (hoja "Retail": Tienda, Operador2, Zona, Cadena, Condición)
     - Calendario del mes anterior (el mismo Excel que descarga la app)
2. Asigna el FVS (Frecuencia de Visita Semanal) por tienda:
     - Excepciones fijas en FVS 1 (Periferia, Makro, y tiendas puntuales).
     - Resto de Lima: se ordena por Q de mayor a menor; las N3 primeras FVS 3,
       las N2 siguientes FVS 2 y las demás FVS 1.
3. Programa los días de visita por operador (Operador2 de Horario):
     - Sin días consecutivos.
     - Máximo `cap_dia` tiendas por día (L-V) y `cap_sab` el sábado.
     - Agrupa por zona dentro de cada día.
     - Las tiendas cuyo operador y FVS no cambiaron conservan sus días.
     - Si un operador no tiene cupo, el exceso recae en la tienda de menor
       prioridad (menor FVS / menor Q) y se marca como "opcional".
4. Valida, compara contra el mes anterior y exporta a Excel.

Este módulo no depende de Streamlit.
"""
from __future__ import annotations

import io
import itertools
import random
import re
import unicodedata
import zlib
import datetime as dt
from dataclasses import dataclass, field, asdict

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DIAS = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"]
ABREV = ["L", "M", "Mi", "J", "V", "S"]
SAB = 5


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def norm(s) -> str:
    """Minúsculas, sin tildes, sin espacios repetidos. Sirve de llave de cruce."""
    if s is None or (not isinstance(s, str) and pd.isna(s)):
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def _txt(x) -> str:
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return ""
    return str(x).strip()


def _abrir(origen):
    """Acepta bytes, ruta o archivo (p. ej. el UploadedFile de Streamlit)."""
    if isinstance(origen, (bytes, bytearray)):
        return io.BytesIO(origen)
    if hasattr(origen, "seek"):
        origen.seek(0)
    return origen


def abreviar_dias(dias) -> str:
    return "-".join(ABREV[d] for d in sorted(dias)) if dias else "—"


def abreviar_zona(z: str) -> str:
    z = _txt(z)
    if z.lower().startswith("zona "):
        return "Z." + z[5:]
    if z.lower().startswith("periferia "):
        return "Per. " + z[10:]
    return z


def inferir_cadena(tienda: str) -> str:
    t = norm(tienda)
    if t.startswith("makro"):
        return "Makro"
    if t.startswith("viv"):
        return "Vivanda"
    if t.startswith(("pvea", "svea", "xvea", "pv ")):
        return "Plaza Vea"
    return ""


# ---------------------------------------------------------------------------
# Parámetros
# ---------------------------------------------------------------------------
@dataclass
class Params:
    n_fvs3: int = 20                    # primeras N tiendas por Q -> FVS 3
    n_fvs2: int = 20                    # siguientes N tiendas por Q -> FVS 2
    cap_dia: int = 6                    # tope de tiendas por operador, L-V
    cap_sab: int = 3                    # tope de tiendas por operador, sábado (medio día)
    min_dia: int = 4                    # mínimo ideal de tiendas por día L-V (preferencia, no obligación)
    dist_aviso_km: float = 10.0         # aviso si dos tiendas del mismo día quedan a más de estos km
    exc_tiendas: tuple = ("PVEA Ventanilla", "PVEA Puente Piedra")
    exc_cadenas: tuple = ("makro",)     # palabra clave en cadena / inicio del nombre
    exc_zonas: tuple = ("periferia",)   # prefijo de zona
    solo_lv: tuple = ()                 # tiendas que no se visitan en sábado
    fvs3_sabado: bool = True            # las FVS 3 sí pueden ir sábado
    seed: int = 42                      # mismo insumo + misma semilla = mismo calendario
    rondas: int = 60                    # esfuerzo de búsqueda del programador


# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------
def leer_q(origen, nombre: str = ""):
    """
    Lee el Q mensual. Busca (en cualquier hoja, primeras 30 filas) un encabezado
    con "Tienda" y una columna "Q Doméstico" (o "Q").
    Devuelve (DataFrame[tienda, key, q], info).
    """
    nombre = (nombre or getattr(origen, "name", "") or "").lower()
    if nombre.endswith((".csv", ".txt")):
        crudos = [("csv", pd.read_csv(_abrir(origen), header=None, dtype=object,
                                      sep=None, engine="python"))]
    else:
        xl = pd.ExcelFile(_abrir(origen))
        crudos = [(h, xl.parse(h, header=None, dtype=object)) for h in xl.sheet_names]

    for hoja, raw in crudos:
        for i in range(min(len(raw), 30)):
            fila = [norm(x) for x in raw.iloc[i].tolist()]
            if "tienda" not in fila:
                continue
            ct = fila.index("tienda")
            cq = [j for j, x in enumerate(fila) if x in ("q domestico", "q dom", "q")]
            if not cq:
                cq = [j for j, x in enumerate(fila) if x.startswith("q ")]
            if not cq:
                continue
            cq = cq[0]
            datos = raw.iloc[i + 1:, [ct, cq]].copy()
            datos.columns = ["tienda", "q"]
            datos["tienda"] = datos["tienda"].map(_txt)
            datos = datos[(datos["tienda"] != "") & (datos["tienda"].map(norm) != "total")]
            q_num = pd.to_numeric(datos["q"], errors="coerce")
            info = {"hoja": hoja, "columna_q": _txt(raw.iloc[i, cq]),
                    "q_vacios": int(q_num.isna().sum()), "duplicadas": 0}
            datos["q"] = q_num.fillna(0)
            datos["key"] = datos["tienda"].map(norm)
            info["duplicadas"] = int(datos["key"].duplicated().sum())
            datos = (datos.groupby("key", as_index=False, sort=False)
                     .agg(tienda=("tienda", "first"), q=("q", "sum")))
            return datos[["tienda", "key", "q"]], info
    raise ValueError(
        "No encontré en el archivo del Q una columna 'Tienda' junto a una columna "
        "'Q Doméstico' (o 'Q'). Revisa los encabezados.")


def leer_horario(origen):
    """
    Lee HorarioTiendas.xlsx (hoja "Retail"). Usa la columna Operador2 como operador.
    Solo Región = Lima. Devuelve (DataFrame, info).
    """
    xl = pd.ExcelFile(_abrir(origen))
    hojas = (["Retail"] + [h for h in xl.sheet_names if h != "Retail"]
             if "Retail" in xl.sheet_names else xl.sheet_names)
    df, cols = None, {}
    for h in hojas:
        tmp = xl.parse(h, dtype=object)
        c = {norm(x): x for x in tmp.columns}
        if "tienda" in c and ("operador2" in c or "operador" in c):
            df, cols = tmp, c
            break
    if df is None:
        raise ValueError("No encontré en Horario una hoja con columnas 'Tienda' y 'Operador2'.")

    col_op = cols.get("operador2") or cols["operador"]
    out = pd.DataFrame({
        "cadena": df[cols["cadena"]] if "cadena" in cols else "",
        "tienda": df[cols["tienda"]],
        "zona": df[cols["zona"]] if "zona" in cols else "",
        "operador": df[col_op],
        "condicion": df[cols["condicion"]] if "condicion" in cols else "",
        "region": df[cols["region"]] if "region" in cols else "Lima",
        "distrito": df[cols["distrito"]] if "distrito" in cols else "",
        "lat": pd.to_numeric(df[cols["latitud"]], errors="coerce") if "latitud" in cols
               else pd.Series(float("nan"), index=df.index),
        "lon": pd.to_numeric(df[cols["longitud"]], errors="coerce") if "longitud" in cols
               else pd.Series(float("nan"), index=df.index),
    })
    for c in ["cadena", "tienda", "zona", "operador", "condicion", "region", "distrito"]:
        out[c] = out[c].map(_txt)
    out = out[out["tienda"] != ""]
    out = out[out["region"].map(norm) == "lima"]
    out["operador"] = out["operador"].where(~out["operador"].isin(["-", "nan", "NaN"]), "")
    out["key"] = out["tienda"].map(norm)
    duplicadas = out[out["key"].duplicated(keep="first")]["tienda"].tolist()
    out = out.drop_duplicates("key", keep="first").reset_index(drop=True)
    for i, r in out.iterrows():
        if not r["cadena"]:
            out.at[i, "cadena"] = inferir_cadena(r["tienda"])
    info = {"usa_operador2": "operador2" in cols, "duplicadas": duplicadas,
            "tiene_distrito": "distrito" in cols,
            "con_coordenadas": int((out["lat"].notna() & out["lon"].notna()).sum())}
    return out, info


def leer_calendario_previo(origen):
    """
    Lee el calendario del mes anterior (hoja "Calendario": Operador, Cadena, Tienda,
    Zona, Q Doméstico, FVS, LUNES..SABADO con "x").
    Devuelve dict key -> {tienda, operador, fvs, dias(set de índices 0..5)}.
    """
    wb = load_workbook(_abrir(origen), data_only=True)
    ws = wb["Calendario"] if "Calendario" in wb.sheetnames else wb.worksheets[0]
    cab, fila_cab = {}, None
    for r in range(1, 21):
        vals = [norm(c) for c in next(ws.iter_rows(min_row=r, max_row=r, values_only=True))]
        if "tienda" in vals and "lunes" in vals:
            cab = {v: i for i, v in enumerate(vals) if v}
            fila_cab = r
            break
    if fila_cab is None:
        raise ValueError("No encontré el encabezado (Tienda, LUNES..SABADO) en el calendario anterior.")
    prev = {}
    for fila in ws.iter_rows(min_row=fila_cab + 1, values_only=True):
        t = _txt(fila[cab["tienda"]])
        if not t:
            continue
        dias = {i for i, d in enumerate(DIAS) if norm(fila[cab[norm(d)]]).startswith("x")}
        fvs = len(dias)
        if "fvs" in cab:
            try:
                fvs = int(fila[cab["fvs"]])
            except (TypeError, ValueError):
                pass
        prev[norm(t)] = {"tienda": t, "operador": _txt(fila[cab["operador"]]) if "operador" in cab else "",
                         "zona": _txt(fila[cab["zona"]]) if "zona" in cab else "",
                         "fvs": fvs, "dias": dias}
    return prev


def detectar_solo_lv(hor: pd.DataFrame) -> list:
    """Tiendas cuya condición en Horario no incluye sábado (p. ej. 'Lunes a viernes - FDS previa coordinación')."""
    res = []
    for _, r in hor.iterrows():
        c = norm(r["condicion"])
        if c and "sabado" not in c.split(" - ")[0]:
            res.append(r["tienda"])
    return res


# ---------------------------------------------------------------------------
# Universo de tiendas y FVS
# ---------------------------------------------------------------------------
def armar_universo(q_df: pd.DataFrame, hor: pd.DataFrame, prev: dict, extra=()):
    """
    Tiendas a programar:
      - las del Q que existen en Horario (Lima),
      - las del calendario anterior que siguen en Horario pero no vienen en el Q (Q = 0),
      - las que el usuario agregue manualmente (Q = 0 si no vienen en el Q).
    """
    hidx = hor.set_index("key")
    qmap = dict(zip(q_df["key"], q_df["q"]))
    qnom = dict(zip(q_df["key"], q_df["tienda"]))
    claves, info = [], {"q_sin_horario": [], "q_cero_por_calendario": [], "salen_del_programa": [],
                        "agregadas_manual": [], "sin_operador": []}

    for k in q_df["key"]:
        if k in hidx.index:
            claves.append(k)
        else:
            info["q_sin_horario"].append(qnom[k])
    for k, v in prev.items():
        if k in claves:
            continue
        if k in hidx.index and k not in qmap:
            claves.append(k)
            info["q_cero_por_calendario"].append(hidx.loc[k, "tienda"])
        elif k not in hidx.index:
            info["salen_del_programa"].append(v["tienda"])
    for t in extra:
        k = norm(t)
        if k in hidx.index and k not in claves:
            claves.append(k)
            info["agregadas_manual"].append(hidx.loc[k, "tienda"])

    uni = hidx.loc[claves].reset_index()
    uni["q"] = uni["key"].map(lambda k: float(qmap.get(k, 0)))
    sin_op = uni[uni["operador"] == ""]
    info["sin_operador"] = sin_op["tienda"].tolist()
    uni = uni[uni["operador"] != ""].reset_index(drop=True)
    info["candidatas_sin_q"] = [r["tienda"] for _, r in hor.iterrows() if r["key"] not in claves]
    return uni, info


def es_excepcion(row, p: Params) -> bool:
    k, cad, t, z = row["key"], norm(row["cadena"]), norm(row["tienda"]), norm(row["zona"])
    if k in {norm(x) for x in p.exc_tiendas}:
        return True
    for kw in p.exc_cadenas:
        kw = norm(kw)
        if kw and (kw in cad or t.startswith(kw)):
            return True
    for kw in p.exc_zonas:
        kw = norm(kw)
        if kw and z.startswith(kw):
            return True
    return False


def asignar_fvs(uni: pd.DataFrame, p: Params, prev: dict) -> pd.DataFrame:
    """
    Excepciones -> FVS 1. El resto se ordena por Q (mayor a menor):
    primeras n_fvs3 -> 3, siguientes n_fvs2 -> 2, demás -> 1.
    Desempate: FVS del mes anterior (el mayor primero) y luego alfabético.
    """
    u = uni.copy()
    u["excepcion"] = u.apply(lambda r: es_excepcion(r, p), axis=1)
    u["fvs_previo"] = u["key"].map(lambda k: prev.get(k, {}).get("fvs", 0))
    resto = u[~u["excepcion"]].sort_values(["q", "fvs_previo", "tienda"],
                                           ascending=[False, False, True])
    pos = {k: i + 1 for i, k in enumerate(resto["key"])}
    u["posicion"] = u["key"].map(pos)

    def fvs(r):
        if r["excepcion"]:
            return 1
        return 3 if r["posicion"] <= p.n_fvs3 else (2 if r["posicion"] <= p.n_fvs3 + p.n_fvs2 else 1)

    u["fvs"] = u.apply(fvs, axis=1).astype(int)
    return u


# ---------------------------------------------------------------------------
# Programador de días
# ---------------------------------------------------------------------------
_ZW, _BW, _HARD = 4.0, 1.5, 1_000_000.0
_DW = 1.5     # por cada distrito extra en un mismo día
_GW = 0.6     # por km de dispersión (suma de distancias de cada tienda al centro del día)
_MW = 2.0     # por cada tienda que falta para llegar al mínimo ideal del día


def _a_km(lat, lon):
    """Coordenadas en km (plano local, suficiente para Lima). None si falta el dato."""
    if lat is None or lon is None or pd.isna(lat) or pd.isna(lon):
        return None
    return ((float(lon) + 77.04) * 111.32 * 0.9779, (float(lat) + 12.05) * 110.57)


def distancia_km(a, b) -> float:
    return ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5


def _patrones(f: int, permitidos) -> list:
    """Combinaciones de f días sin días consecutivos (si no existen, se relaja)."""
    perm = sorted(permitidos)
    if f <= 0:
        return [()]
    if f > len(perm):
        return [tuple(perm)]
    todos = list(itertools.combinations(perm, f))
    ok = [c for c in todos if all(b - a > 1 for a, b in zip(c, c[1:]))]
    return ok or todos


def _costo_patron(f: int, pat: tuple) -> float:
    """Preferencia por visitas bien espaciadas y por no usar el sábado (medio día)."""
    c = 0.0
    if f == 2 and len(pat) == 2:
        c += {3: 0.0, 2: 1.0, 4: 1.0}.get(pat[1] - pat[0], 2.0)
    elif f == 3:
        c += 0.0 if pat in ((0, 2, 4), (1, 3, 5)) else 1.0
    if SAB in pat and f < 3:
        c += 0.3
    return c


def _resolver(stores: list, p: Params, rng: random.Random):
    """
    Programa los días de UN operador. Búsqueda local con reinicios.
    Costo = exceso de capacidad (cae en las tiendas de menor prioridad)
          + zonas distintas por día + desbalance + preferencias de patrón.
    El sábado nunca se excede (penalización dura).
    """
    n = len(stores)
    prio = [s["prio"] for s in stores]
    zon = [s["zona"] for s in stores]
    dis = [s["distrito"] for s in stores]
    xy = [s["xy"] for s in stores]
    pats = [[s["fija"]] if s["fija"] else s["pats"] for s in stores]
    pcost = []
    for s, pl in zip(stores, pats):
        fila = []
        for pt in pl:
            c = 0.0 if s["fija"] else _costo_patron(s["fvs"], pt)
            if not s["fija"] and s["prev_dias"]:
                c -= 1.5 * len(s["prev_dias"] & set(pt))     # continuidad si cambió el FVS
            fila.append(c)
        pcost.append(fila)
    W = [1000 + 100 * (n - 1 - prio[i]) for i in range(n)]    # exceso barato = baja prioridad
    miembros = [set() for _ in range(6)]
    asig = [0] * n

    def costo_dia(d, m):
        L = len(m)
        if not L:
            return 0.0
        cap = p.cap_sab if d == SAB else p.cap_dia
        c = _ZW * (len({zon[i] for i in m}) - 1) + _BW * (L / cap) ** 2
        ds = {dis[i] for i in m if dis[i]}
        if len(ds) > 1:
            c += _DW * (len(ds) - 1)
        pts = [xy[i] for i in m if xy[i] is not None]
        if len(pts) > 1:
            cx = sum(a for a, _ in pts) / len(pts)
            cy = sum(b for _, b in pts) / len(pts)
            c += _GW * sum(((a - cx) ** 2 + (b - cy) ** 2) ** 0.5 for a, b in pts)
        if d != SAB and L < p.min_dia:
            c += _MW * (p.min_dia - L)
        if L > cap:
            ex = L - cap
            if d == SAB:
                c += _HARD * ex
            else:
                c += sum(W[i] for i in sorted(m, key=lambda i: -prio[i])[:ex])
        return c

    def evaluar(i):
        """Costo incremental de cada patrón de i (con i fuera de `miembros`)."""
        base = [costo_dia(d, miembros[d]) for d in range(6)]
        res = []
        for pi, pt in enumerate(pats[i]):
            delta = pcost[i][pi]
            for d in pt:
                miembros[d].add(i)
                delta += costo_dia(d, miembros[d]) - base[d]
                miembros[d].discard(i)
            res.append(delta)
        return res

    def poner(i, pi):
        asig[i] = pi
        for d in pats[i][pi]:
            miembros[d].add(i)

    def quitar(i):
        for d in pats[i][asig[i]]:
            miembros[d].discard(i)

    def total():
        return (sum(costo_dia(d, miembros[d]) for d in range(6))
                + sum(pcost[i][asig[i]] for i in range(n)))

    # Construcción inicial: fijas primero, luego por prioridad
    orden = sorted(range(n), key=lambda i: (0 if stores[i]["fija"] else 1, prio[i]))
    for i in orden:
        if stores[i]["fija"]:
            poner(i, 0)
        else:
            res = evaluar(i)
            poner(i, res.index(min(res)))

    libres = [i for i in range(n) if not stores[i]["fija"]]

    def barrido():
        mejoro = False
        ord_ = libres[:]
        rng.shuffle(ord_)
        for i in ord_:
            cur = asig[i]
            quitar(i)
            res = evaluar(i)
            m = min(res)
            if m < res[cur] - 1e-9:
                nuevo = rng.choice([k for k, v in enumerate(res) if v <= m + 1e-9])
                mejoro = True
            elif len(res) > 1 and rng.random() < 0.25:
                c = [k for k, v in enumerate(res) if abs(v - m) <= 1e-9 and k != cur and v <= res[cur] + 1e-9]
                nuevo = rng.choice(c) if c else cur
            else:
                nuevo = cur
            poner(i, nuevo)
        return mejoro

    def descenso():
        for _ in range(30):
            if not barrido():
                break

    def reconstruir(sol):
        for d in range(6):
            miembros[d].clear()
        for i, pi in enumerate(sol):
            poner(i, pi)

    mejor, mejor_costo = asig[:], total()
    if libres:
        descenso()
        c = total()
        if c < mejor_costo - 1e-9:
            mejor, mejor_costo = asig[:], c
        sin_mejora = 0
        for _ in range(p.rondas):
            reconstruir(mejor)
            for i in rng.sample(libres, min(len(libres), rng.randint(2, 5))):
                quitar(i)
                poner(i, rng.randrange(len(pats[i])))
            descenso()
            c = total()
            if c < mejor_costo - 1e-9:
                mejor, mejor_costo, sin_mejora = asig[:], c, 0
            else:
                sin_mejora += 1
            if sin_mejora >= 40:
                break
        reconstruir(mejor)
    return [pats[i][mejor[i]] for i in range(n)]


def programar(uni: pd.DataFrame, p: Params, prev: dict) -> pd.DataFrame:
    """
    Devuelve el calendario: una fila por tienda con booleanos LUNES..SABADO.
    `uni` debe traer la columna fvs (ver asignar_fvs).
    """
    solo = {norm(t) for t in p.solo_lv}
    filas = []
    for op, g in uni.groupby("operador", sort=True):
        g = g.sort_values(["fvs", "q", "tienda"], ascending=[False, False, True]).reset_index(drop=True)
        stores = []
        for r, row in g.iterrows():
            permit = set(range(6))
            restringida = row["key"] in solo and not (p.fvs3_sabado and row["fvs"] >= 3)
            if restringida:
                permit.discard(SAB)
            pats = _patrones(int(row["fvs"]), permit)
            pv = prev.get(row["key"])
            mismo_op = bool(pv) and norm(pv["operador"]) == norm(op)
            fija = None
            if mismo_op and len(pv["dias"]) == int(row["fvs"]):
                t = tuple(sorted(pv["dias"]))
                if t in set(pats):
                    fija = t
            stores.append({"prio": r, "zona": row["zona"], "distrito": row.get("distrito", ""),
                           "xy": _a_km(row.get("lat"), row.get("lon")), "pats": pats, "fija": fija,
                           "fvs": int(row["fvs"]), "key": row["key"], "restringida": restringida,
                           "prev_dias": set(pv["dias"]) if mismo_op else set()})
        # las fijas que desbordan el tope de un día se liberan (la de menor prioridad primero)
        for d in range(6):
            cap = p.cap_sab if d == SAB else p.cap_dia
            en_dia = [s for s in stores if s["fija"] and d in s["fija"]]
            while len(en_dia) > cap:
                peor = max(en_dia, key=lambda s: s["prio"])
                peor["fija"] = None
                en_dia.remove(peor)

        # Se resuelve; si queda un exceso evitable por culpa de tiendas fijas, se liberan
        # (de menor a mayor prioridad) las fijas de los días desbordados y se vuelve a resolver.
        while True:
            rng = random.Random(p.seed ^ zlib.crc32(op.encode("utf-8")))
            soluciones = _resolver(stores, p, rng)
            cargas = [sum(1 for so in soluciones if d in so) for d in range(6)]
            exceso = sum(max(0, cargas[d] - p.cap_dia) for d in range(5)) + max(0, cargas[SAB] - p.cap_sab)
            inevitable = max(0, sum(cargas) - (5 * p.cap_dia + p.cap_sab))
            if exceso <= inevitable:
                break
            cand = [s for s, so in zip(stores, soluciones) if s["fija"]
                    and any(cargas[d] > (p.cap_sab if d == SAB else p.cap_dia) for d in so)]
            if not cand:
                break
            max(cand, key=lambda s: s["prio"])["fija"] = None
        for s, sol, (_, row) in zip(stores, soluciones, g.iterrows()):
            fila = {"Operador": op, "Cadena": row["cadena"], "Tienda": row["tienda"],
                    "Distrito": row.get("distrito", ""), "Zona": row["zona"],
                    "Lat": row.get("lat"), "Lon": row.get("lon"),
                    "Q": row["q"], "FVS": int(row["fvs"]),
                    "Excepcion": bool(row["excepcion"]), "SoloLV": s["restringida"],
                    "Fijada": s["fija"] is not None, "key": row["key"]}
            for d, nombre in enumerate(DIAS):
                fila[nombre] = d in sol
            filas.append(fila)
    return pd.DataFrame(filas)


# ---------------------------------------------------------------------------
# Validaciones, cargas, opcionales y cambios
# ---------------------------------------------------------------------------
def dias_de(fila) -> list:
    return [d for d, n in enumerate(DIAS) if bool(fila[n])]


def cargas(sched: pd.DataFrame) -> pd.DataFrame:
    """Visitas por operador y día (+ total)."""
    c = sched.groupby("Operador")[DIAS].sum().astype(int)
    c["TOTAL"] = c.sum(axis=1)
    return c


def calcular_opcionales(sched: pd.DataFrame, p: Params) -> set:
    """
    {(key, día)} de las visitas que exceden el tope L-V de su operador.
    El exceso recae en las tiendas de menor prioridad (menor FVS, luego menor Q).
    """
    res = set()
    for _, g in sched.groupby("Operador"):
        orden = g.sort_values(["FVS", "Q", "Tienda"], ascending=[False, False, True])["key"].tolist()
        rango = {k: i for i, k in enumerate(orden)}
        for d in range(5):
            miembros = g[g[DIAS[d]]]["key"].tolist()
            if len(miembros) > p.cap_dia:
                peores = sorted(miembros, key=lambda k: -rango[k])[: len(miembros) - p.cap_dia]
                res.update((k, d) for k in peores)
    return res


def validar(sched: pd.DataFrame, p: Params) -> pd.DataFrame:
    """Avisos (no bloquean): tope por día, días consecutivos, FVS vs visitas, sábado restringido."""
    avisos = []
    for op, g in sched.groupby("Operador"):
        for d in range(6):
            cap = p.cap_sab if d == SAB else p.cap_dia
            n = int(g[DIAS[d]].sum())
            if n > cap:
                avisos.append(("Tope por día", op, "", f"{DIAS[d].capitalize()}: {n} tiendas (tope {cap})"))
    if "Lat" in sched.columns:
        for op, g in sched.groupby("Operador"):
            for d in range(6):
                pts = [(r["Tienda"], _a_km(r["Lat"], r["Lon"])) for _, r in g[g[DIAS[d]]].iterrows()]
                pts = [(t, xy) for t, xy in pts if xy is not None]
                lejos = max(((distancia_km(a[1], b[1]), a[0], b[0]) for i, a in enumerate(pts) for b in pts[i + 1:]),
                            default=None)
                if lejos and lejos[0] > p.dist_aviso_km:
                    avisos.append(("Tiendas lejanas", op, "",
                                   f"{DIAS[d].capitalize()}: {lejos[1]} y {lejos[2]} a {lejos[0]:.1f} km "
                                   f"(aviso desde {p.dist_aviso_km:g} km)"))
    for _, r in sched.iterrows():
        dias = dias_de(r)
        if len(dias) != r["FVS"]:
            avisos.append(("FVS", r["Operador"], r["Tienda"], f"{len(dias)} visita(s) programada(s), FVS {r['FVS']}"))
        if any(b - a == 1 for a, b in zip(dias, dias[1:])):
            avisos.append(("Días consecutivos", r["Operador"], r["Tienda"], abreviar_dias(dias)))
        if r["SoloLV"] and r[DIAS[SAB]]:
            avisos.append(("Sábado no permitido", r["Operador"], r["Tienda"], "Solo se visita de lunes a viernes"))
    return pd.DataFrame(avisos, columns=["Tipo", "Operador", "Tienda", "Detalle"])


def comparar(sched: pd.DataFrame, prev: dict) -> pd.DataFrame:
    """Compara el calendario nuevo contra el del mes anterior."""
    filas = []
    for _, r in sched.iterrows():
        pv = prev.get(r["key"])
        dias = set(dias_de(r))
        if not pv:
            estado = "Nueva en el programa"
        elif norm(pv["operador"]) != norm(r["Operador"]):
            estado = "Cambió de operador"
        elif pv["fvs"] != r["FVS"]:
            estado = f"Cambió FVS ({pv['fvs']} → {r['FVS']})"
        elif pv["dias"] != dias:
            estado = "Reprogramada (mismo FVS)"
        else:
            estado = "Sin cambios"
        filas.append({"Tienda": r["Tienda"],
                      "Operador anterior": pv["operador"] if pv else "",
                      "Operador nuevo": r["Operador"],
                      "FVS anterior": pv["fvs"] if pv else None,
                      "FVS nuevo": r["FVS"],
                      "Días anteriores": abreviar_dias(pv["dias"]) if pv else "",
                      "Días nuevos": abreviar_dias(dias),
                      "Estado": estado})
    vigentes = set(sched["key"])
    for k, pv in prev.items():
        if k not in vigentes:
            filas.append({"Tienda": pv["tienda"], "Operador anterior": pv["operador"], "Operador nuevo": "",
                          "FVS anterior": pv["fvs"], "FVS nuevo": None,
                          "Días anteriores": abreviar_dias(pv["dias"]), "Días nuevos": "",
                          "Estado": "Sale del programa"})
    return pd.DataFrame(filas)


# ---------------------------------------------------------------------------
# Exportación a Excel (misma estructura que el calendario original)
# ---------------------------------------------------------------------------
_F = "Arial"
_AZUL = PatternFill("solid", fgColor="1F3A5F")      # azul marino
_CLARO = PatternFill("solid", fgColor="A1C6EA")      # azul claro: día de visita
_NARANJA = PatternFill("solid", fgColor="C9B29B")    # arena: visita opcional
_ROJO = PatternFill("solid", fgColor="FFC7CE")
_VERDE = PatternFill("solid", fgColor="A1C6EA")      # control OK
_BEIGE = PatternFill("solid", fgColor="F2EFEA")
_borde = Border(*(Side(style="thin", color="BFBFBF"),) * 4)


def _enc(ws, fila, ncols, ini=1):
    for c in range(ini, ini + ncols):
        cel = ws.cell(row=fila, column=c)
        cel.font = Font(name=_F, bold=True, color="FFFFFF", size=10)
        cel.fill = _AZUL
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = _borde


def _nombre_hoja(nombre: str, usados: set) -> str:
    base = re.sub(r"[\[\]\*\?/\\:]", "", nombre)[:28] or "Operador"
    n, i = base, 2
    while n in usados:
        n, i = f"{base[:25]}_{i}", i + 1
    usados.add(n)
    return n


def exportar_excel(sched: pd.DataFrame, p: Params, prev: dict | None = None,
                   ajustadas: set | None = None) -> bytes:
    """
    Genera el Excel con las hojas: Calendario, Horario semanal, Resumen, una hoja
    por operador, Cambios y Parámetros. La hoja "Calendario" mantiene el formato
    que lee la app el mes siguiente.
    """
    ajustadas = ajustadas or set()
    opcionales = calcular_opcionales(sched, p)
    s = sched.copy()
    s["_o"] = s["Operador"].str.lower()
    s = s.sort_values(["_o", "Zona", "FVS", "Q", "Tienda"], ascending=[True, True, False, False, True])
    operadores = sorted(s["Operador"].unique(), key=str.lower)

    wb = Workbook()
    # ------------------------------------------------------------ Calendario
    ws = wb.active
    ws.title = "Calendario"
    ws["A1"] = "Calendario semanal de visitas por tienda y operador"
    ws["A1"].font = Font(name=_F, bold=True, size=14, color="1F3A5F")
    ws["A2"] = ("Fuente: Q del mes + HorarioTiendas (Operador2, Zona). \"x\" = día de visita. Sin visitas en días "
                f"consecutivos. Tope: {p.cap_dia} tiendas por día (L-V) y {p.cap_sab} el sábado. "
                f"Generado el {dt.date.today():%d/%m/%Y}.")
    ws["A2"].font = Font(name=_F, italic=True, size=9)
    cab = ["Operador", "Cadena", "Tienda", "Distrito", "Zona", "Q Doméstico", "FVS"] + DIAS + \
          ["Visitas programadas", "Control", "Observación"]
    for i, h in enumerate(cab, 1):
        ws.cell(row=4, column=i, value=h)
    _enc(ws, 4, len(cab))
    r0 = 5
    for i, (_, r) in enumerate(s.iterrows()):
        fila = r0 + i
        obs = []
        if r["Excepcion"]:
            obs.append("Excepción FVS 1")
        if r["SoloLV"]:
            obs.append("Solo L-V")
        if r["key"] in ajustadas:
            obs.append("Ajustada a mano")
        vals = [r["Operador"], r["Cadena"], r["Tienda"], r.get("Distrito", ""), r["Zona"], r["Q"], int(r["FVS"])]
        for c, v in enumerate(vals, 1):
            cel = ws.cell(row=fila, column=c, value=v)
            cel.font = Font(name=_F, size=10)
            cel.border = _borde
        for d, nombre in enumerate(DIAS):
            cel = ws.cell(row=fila, column=8 + d)
            cel.border = _borde
            cel.alignment = Alignment(horizontal="center")
            cel.font = Font(name=_F, size=10, bold=True)
            if r[nombre]:
                cel.value = "x"
                if (r["key"], d) in opcionales:
                    cel.fill = _NARANJA
                    obs.append(f"Visita opcional ({DIAS[d].capitalize()}): excede el tope del operador")
                else:
                    cel.fill = _CLARO
        ws.cell(row=fila, column=14, value=f'=COUNTIF(H{fila}:M{fila},"x")')
        ws.cell(row=fila, column=15, value=f'=IF(N{fila}=G{fila},"OK","REVISAR")')
        ws.cell(row=fila, column=16, value="; ".join(obs))
        for c in (14, 15, 16):
            ws.cell(row=fila, column=c).font = Font(name=_F, size=10)
            ws.cell(row=fila, column=c).border = _borde
        ws.cell(row=fila, column=14).alignment = Alignment(horizontal="center")
        ws.cell(row=fila, column=15).alignment = Alignment(horizontal="center")
    ult = r0 + len(s) - 1
    rango_ctrl = f"O{r0}:O{ult}"
    ws.conditional_formatting.add(rango_ctrl, CellIsRule(operator="equal", formula=['"REVISAR"'], fill=_ROJO))
    ws.conditional_formatting.add(rango_ctrl, CellIsRule(operator="equal", formula=['"OK"'], fill=_VERDE))
    for col, w in zip("ABCDEFGHIJKLMNOP", [18, 11, 28, 18, 15, 11, 6, 9, 9, 12, 9, 9, 9, 13, 10, 52]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[4].height = 32
    ws.freeze_panes = "D5"
    ws.page_setup.orientation = "landscape"          # al imprimir, todo el ancho en una página
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "4:4"
    ws.auto_filter.ref = f"A4:P{ult}"
    ws["A" + str(ult + 2)] = ("Celdas en naranja = visita opcional: el operador no tiene cupo ese día y se asignó a su "
                              "tienda de menor prioridad.")
    ws["A" + str(ult + 2)].font = Font(name=_F, italic=True, size=9)

    # -------------------------------------------------------- Horario semanal
    wh = wb.create_sheet("Horario semanal")
    wh["A1"] = "Horario semanal: tiendas a visitar por operador y día"
    wh["A1"].font = Font(name=_F, bold=True, size=14, color="1F3A5F")
    for i, h in enumerate(["Operador"] + DIAS, 1):
        wh.cell(row=3, column=i, value=h)
    _enc(wh, 3, 7)
    for i, op in enumerate(operadores):
        fila = 4 + i
        g = s[s["Operador"] == op].sort_values(["Zona", "Tienda"])
        wh.cell(row=fila, column=1, value=op).font = Font(name=_F, bold=True, size=10)
        maxl = 1
        for d, nombre in enumerate(DIAS):
            items = []
            for _, r in g[g[nombre]].iterrows():
                t = f"{r['Tienda']} ({r.get('Distrito') or abreviar_zona(r['Zona'])})"
                items.append(t + (" [opcional]" if (r["key"], d) in opcionales else ""))
            maxl = max(maxl, len(items))
            cel = wh.cell(row=fila, column=2 + d, value="\n".join(items))
            cel.alignment = Alignment(wrap_text=True, vertical="top")
            cel.font = Font(name=_F, size=9)
            cel.border = _borde
        wh.cell(row=fila, column=1).alignment = Alignment(vertical="top")
        wh.row_dimensions[fila].height = max(30, 12.5 * maxl + 6)
    wh.column_dimensions["A"].width = 20
    wh.page_setup.orientation = "landscape"
    wh.page_setup.fitToWidth = 1
    wh.page_setup.fitToHeight = 0
    wh.sheet_properties.pageSetUpPr.fitToPage = True
    for c in "BCDEFG":
        wh.column_dimensions[c].width = 33

    # ---------------------------------------------------------------- Resumen
    wr = wb.create_sheet("Resumen")
    wr["A1"] = "Carga de visitas por operador y día"
    wr["A1"].font = Font(name=_F, bold=True, size=14, color="1F3A5F")
    cabr = ["Operador", "Tiendas"] + DIAS + ["Total visitas", "Controles OK", "Capacidad semanal", "Exceso"]
    for i, h in enumerate(cabr, 1):
        wr.cell(row=3, column=i, value=h)
    _enc(wr, 3, len(cabr))
    rng_op = f"Calendario!$A${r0}:$A${ult}"
    for i, op in enumerate(operadores):
        f = 4 + i
        wr.cell(row=f, column=1, value=op)
        wr.cell(row=f, column=2, value=f"=COUNTIF({rng_op},A{f})")
        for d in range(6):
            col = get_column_letter(8 + d)
            wr.cell(row=f, column=3 + d,
                    value=f'=COUNTIFS({rng_op},$A{f},Calendario!${col}${r0}:${col}${ult},"x")')
        wr.cell(row=f, column=9, value=f"=SUM(C{f}:H{f})")
        wr.cell(row=f, column=10, value=f'=COUNTIFS({rng_op},$A{f},Calendario!$O${r0}:$O${ult},"OK")')
        wr.cell(row=f, column=11, value=f"=5*{p.cap_dia}+{p.cap_sab}")
        wr.cell(row=f, column=12, value=f"=MAX(0,I{f}-K{f})")
    ft = 4 + len(operadores)
    wr.cell(row=ft, column=1, value="TOTAL")
    for c in range(2, 13):
        L = get_column_letter(c)
        wr.cell(row=ft, column=c, value=f"=SUM({L}4:{L}{ft - 1})")
    for f in range(4, ft + 1):
        for c in range(1, 13):
            cel = wr.cell(row=f, column=c)
            cel.font = Font(name=_F, size=10, bold=(f == ft), color="1F3A5F")
            cel.border = _borde
            if f == ft:
                cel.fill = _BEIGE
            if c > 1:
                cel.alignment = Alignment(horizontal="center")
    wr.cell(row=ft + 2, column=1,
            value="Control: Total visitas debe ser igual a la suma de FVS de todas las tiendas. "
                  f"Capacidad semanal = 5 × {p.cap_dia} (L-V) + {p.cap_sab} (sábado).").font = Font(name=_F, italic=True, size=9)
    wr.column_dimensions["A"].width = 20
    for c in range(2, 13):
        wr.column_dimensions[get_column_letter(c)].width = 12

    # ------------------------------------------------- Hoja por operador
    usados = {"Calendario", "Horario semanal", "Resumen", "Cambios", "Parámetros"}
    for op in operadores:
        wo = wb.create_sheet(_nombre_hoja(op, usados))
        for i, h in enumerate([op.upper(), "DISTRITO", "ZONA"] + DIAS, 1):
            wo.cell(row=1, column=i, value=h)
        _enc(wo, 1, 9)
        g = s[s["Operador"] == op].sort_values(["Zona", "FVS", "Tienda"], ascending=[True, False, True])
        for j, (_, r) in enumerate(g.iterrows()):
            f = 2 + j
            wo.cell(row=f, column=1, value=r["Tienda"])
            wo.cell(row=f, column=2, value=r.get("Distrito", ""))
            wo.cell(row=f, column=3, value=r["Zona"])
            for d, nombre in enumerate(DIAS):
                cel = wo.cell(row=f, column=4 + d)
                cel.alignment = Alignment(horizontal="center")
                if r[nombre]:
                    cel.value = "x"
                    cel.fill = _NARANJA if (r["key"], d) in opcionales else _CLARO
            for c in range(1, 10):
                wo.cell(row=f, column=c).font = Font(name=_F, size=10)
                wo.cell(row=f, column=c).border = _borde
        wo.column_dimensions["A"].width = 30
        wo.column_dimensions["B"].width = 18
        wo.column_dimensions["C"].width = 16
        wo.freeze_panes = "A2"

    # ---------------------------------------------------------------- Cambios
    if prev is not None:
        wc = wb.create_sheet("Cambios")
        cam = comparar(sched, prev)
        for i, h in enumerate(cam.columns, 1):
            wc.cell(row=1, column=i, value=h)
        _enc(wc, 1, len(cam.columns))
        for j, (_, r) in enumerate(cam.iterrows()):
            for i, v in enumerate(r.tolist(), 1):
                if isinstance(v, float) and pd.isna(v):
                    v = None
                cel = wc.cell(row=2 + j, column=i, value=v)
                cel.font = Font(name=_F, size=10)
                cel.border = _borde
        for col, w in zip("ABCDEFGH", [30, 18, 18, 12, 10, 16, 16, 28]):
            wc.column_dimensions[col].width = w
        wc.freeze_panes = "A2"
        wc.auto_filter.ref = f"A1:H{len(cam) + 1}"

    # ------------------------------------------------------------- Parámetros
    wp = wb.create_sheet("Parámetros")
    wp["A1"] = "Parámetros usados para generar este calendario"
    wp["A1"].font = Font(name=_F, bold=True, size=14, color="1F3A5F")
    filas = [
        ("Fecha de generación", f"{dt.date.today():%d/%m/%Y}"),
        ("Tiendas con FVS 3 (primeras por Q, sin excepciones)", p.n_fvs3),
        ("Tiendas con FVS 2 (siguientes por Q)", p.n_fvs2),
        ("Tope de tiendas por operador y día (L-V)", p.cap_dia),
        ("Tope de tiendas por operador el sábado (medio día)", p.cap_sab),
        ("Mínimo ideal de tiendas por día (L-V)", p.min_dia),
        ("Aviso de distancia entre tiendas del mismo día (km)", p.dist_aviso_km),
        ("Excepciones FVS 1 - tiendas", ", ".join(p.exc_tiendas)),
        ("Excepciones FVS 1 - cadena (palabra clave)", ", ".join(p.exc_cadenas)),
        ("Excepciones FVS 1 - zona (prefijo)", ", ".join(p.exc_zonas)),
        ("Tiendas solo lunes a viernes", ", ".join(p.solo_lv)),
        ("Las FVS 3 pueden ir sábado", "Sí" if p.fvs3_sabado else "No"),
        ("Semilla del programador", p.seed),
    ]
    for i, (k, v) in enumerate(filas, 3):
        wp.cell(row=i, column=1, value=k).font = Font(name=_F, bold=True, size=10)
        wp.cell(row=i, column=2, value=v).font = Font(name=_F, size=10)
    wp.column_dimensions["A"].width = 52
    wp.column_dimensions["B"].width = 70

    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

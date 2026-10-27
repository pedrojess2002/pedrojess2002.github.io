# -*- coding: utf-8 -*-
"""Base Maestra de Equivalencias: cada número de la mutual (con su DNI) -> nuestro Id de Plenario.

Lee TODAS las rendiciones (originales y copias trabajadas) que encontró inventario_rendiciones.py, junta cada
par DNI + número de la mutual, y lo resuelve con la evidencia disponible, de mayor a menor confianza:
  1 Confirmado: override confirmado a mano (ENTIDAD1), o ID Cooperativa cargado a mano en una rendición trabajada
  2 Crosswalk existente (Maestro TIPO NUMERO Entidad 5, crosswalk ENTIDAD1 / ENTIDAD3 / ENTIDAD2)
  2 Coincide con Plenario (número = Código de Referencia con el mismo DNI; o referencia = Id)
  3 Sugerido: el DNI tiene un solo crédito en esa mutual
  4 Sin asignar: para completar a mano
Si dos fuentes no coinciden, el caso va a la hoja Conflictos (no se elige en silencio).
Uso: py construir_maestro.py      Salida: 07_Base_Maestra\\Base_Maestra_Equivalencias.xlsx
"""
import collections
import csv
import re
import sys
from pathlib import Path

import openpyxl

AQUI = Path(__file__).resolve().parent
sys.path[:0] = [str(AQUI), str(AQUI.parent)]
sys.stdout.reconfigure(encoding="utf-8")
import comun as C  # noqa: E402
import inventario_rendiciones as INV  # noqa: E402
from construir_modelo import mutual_de  # noqa: E402

SALIDA = C.RAIZ / "07_Base_Maestra"
def plenario_mas_reciente():
    """El 'data plenario DD-MM-AAAA.xls' completo más nuevo (por la fecha del nombre) de Descargas o de CarteraDW.
    Se descartan los exports filtrados (los que traen menos de 8 mutuales)."""
    import datetime as dt
    import xlrd
    cands = []
    for carpeta in (Path(r"<ruta>"), Path(r"procesados"), Path(r"entrada")):
        if not carpeta.exists():
            continue
        for f in carpeta.glob("*.xls"):
            m = re.search(r"plenario\D*(\d{1,2})-(\d{1,2})-(\d{2,4})", f.name, re.I)
            if m:
                a = int(m.group(3)) + (2000 if len(m.group(3)) == 2 else 0)
                cands.append((dt.date(a, int(m.group(2)), int(m.group(1))), f))
    for fecha, f in sorted(cands, reverse=True):
        try:
            sh = xlrd.open_workbook(str(f), on_demand=True).sheet_by_index(0)
            hdr = [str(x).strip() for x in sh.row_values(0)]
            ie = hdr.index("Entidad")
            if len({sh.cell_value(r, ie) for r in range(1, sh.nrows)}) >= 8:
                return f
        except Exception:
            continue
    raise SystemExit("No encontré un data plenario completo")


PLENARIO = plenario_mas_reciente()
NOMBRE_MUTUAL = {"entidad5": "Entidad 5", "entidad3": "Entidad 3", "entidad1": "ENTIDAD1", "entidad6": "ENTIDAD6", "entidad2": "ENTIDAD2",
                 "tresabril": "Entidad 8"}
TIPO_NUMERO = {"entidad5": "TIPO NUMERO / Nro. Cred.", "entidad3": "OP", "entidad1": "Docum", "entidad6": "PADRON",
               "entidad2": "Comprob.", "tresabril": "Referencia externa (= Id)"}

# columnas por mutual (nombres normalizados, en orden de preferencia)
COLS = {
    "entidad5": {"dni": ["tipo nro documento", "nro,doc,", "dni", "documento"],
                "num": ["tipo numero", "nro,cred,", "numero cred ellos", "nro referencia"],
                "nom": ["apellido nombre", "nombre", "beneficiario", "apenom"],
                "imp": ["imputado", "imputado (pre)", "importe pagado", "importe proveedor", "pagado", "importe"],
                "per": ["periodo"], "id": ["id cooperativa", "id_cooperativa"]},
    "entidad3": {"dni": ["dni"], "num": ["op", "op entidad3"], "nom": ["nombre"],
                 "imp": ["importe", "monto pagado", "importe cuota"], "per": [],
                 "id": ["id cooperativa", "id_encontrado", "id"]},
    "entidad1": {"dni": ["cuil"], "num": ["docum", "orden"], "nom": ["nombre_y_apellido", "nombre apellido", "nombre y apellido"],
            "imp": ["monto", "valor cuota", "cuota"], "per": [],
            "id": ["id cooperativa", "id cooperativa correcto", "id identificado (a2)"]},
    "entidad6": {"dni": ["d.n.i."], "num": ["padron"], "nom": ["apellido y nombre"], "imp": ["desc."], "per": [], "id": []},
    "entidad2": {"dni": ["cuil"], "num": ["comprob."], "nom": ["apellido/s y nombre/s"], "imp": ["importe cobrado"],
            "per": ["periodo"], "id": ["i.d. cooperativa"]},
    "tresabril": {"dni": ["socnrodocumento"], "num": ["conreferenciaexterna"], "nom": ["socapellido"],
                  "imp": ["ctactotal"], "per": [], "id": []},
}


def dni_de(v):
    """DNI desde 'DNI - 00000000', 00000000 o un CUIL '20-00000000-0' (11 dígitos -> el del medio)."""
    d = re.sub(r"\D", "", str(v if not isinstance(v, float) else int(v)) if v not in (None, "") else "")
    if len(d) == 11:
        d = d[2:10]
    return d.lstrip("0")


def num_de(v):
    if v in (None, ""):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return C.norm_num(v)


def idx(hdr, nombres):
    norm = [C.sin_acentos(h) for h in hdr]
    for n in nombres:
        if n in norm:
            return norm.index(n)
    return None


def periodos(v, mes_archivo):
    ps = re.findall(r"(20\d\d)(0[1-9]|1[0-2])", str(v or ""))
    return {f"{a}-{m}" for a, m in ps} or ({mes_archivo} if mes_archivo else set())


# ---------------------------------------------------------------- 1. observaciones de las rendiciones
def leer_observaciones():
    wb = openpyxl.load_workbook(SALIDA / "Inventario_Rendiciones.xlsx", read_only=True)
    rows = list(wb["Archivos"].iter_rows(values_only=True))
    wb.close()
    h = rows[0]
    hojas_ok = [dict(zip(h, r)) for r in rows[1:]]
    hojas_ok = [d for d in hojas_ok if d.get("Origen") in ("Original de la mutual", "Copia trabajada") and d.get("Ruta")]
    obs = {}
    leidas = collections.Counter()
    por_ruta = collections.defaultdict(list)
    for d in hojas_ok:
        por_ruta[d["Ruta"]].append(d)
    for ruta, ds in por_ruta.items():
        try:
            hs = dict(INV.hojas(Path(ruta)))
        except Exception as e:
            print("  no se pudo leer", ruta, e)
            continue
        for d in ds:
            m, filas = d["Mutual"], hs.get(d["Hoja"], [])
            i, hdr = INV.encabezado(filas)
            if i is None:
                continue
            cc = {k: idx(hdr, v) for k, v in COLS[m].items()}
            if m == "tresabril" and cc["num"] is not None and len(hdr) > 24 and C.sin_acentos(hdr[0]) == "socapellido":
                # el encabezado viene corrido (campos duplicados): posiciones fijas confirmadas
                cc = {"dni": 3, "num": 23, "nom": 0, "imp": 24, "per": None, "id": None}
            icta2 = None
            if m == "entidad6":
                # dos columnas 'CTA': la 1ª es el número de cuota, la 2ª el monto teórico (identifica el crédito)
                ctas = [j for j, x in enumerate(hdr) if C.sin_acentos(x) == "cta"]
                icta2 = ctas[1] if len(ctas) > 1 else None
            if cc["dni"] is None:
                continue
            n = 0
            for f in filas[i + 1:]:
                if len(f) <= cc["dni"]:
                    continue
                dni = dni_de(f[cc["dni"]])
                if not dni or len(dni) < 6:
                    continue
                num = num_de(f[cc["num"]]) if cc["num"] is not None and len(f) > cc["num"] else ""
                if m == "entidad6" and icta2 is not None and len(f) > icta2 and C.num(f[icta2]):
                    num = f"{num}|{C.num(f[icta2]):.2f}"  # PADRON (persona) + cuota teórica (crédito)
                k = (m, dni, num)
                o = obs.setdefault(k, {"nombre": "", "per": set(), "arch": set(), "ids": collections.Counter(),
                                       "imp": [], "veces": 0, "tipo_txt": ""})
                o["veces"] += 1
                if cc["nom"] is not None and len(f) > cc["nom"] and f[cc["nom"]] and not o["nombre"]:
                    o["nombre"] = str(f[cc["nom"]]).strip()
                if cc["num"] is not None and len(f) > cc["num"] and not o["tipo_txt"]:
                    o["tipo_txt"] = str(f[cc["num"]]).strip()
                per_v = f[cc["per"]] if cc["per"] is not None and len(f) > cc["per"] else None
                o["per"] |= periodos(per_v, d.get("Mes") or "")
                o["arch"].add(f'{d["Archivo"]} [{d["Hoja"]}]')
                if cc["imp"] is not None and len(f) > cc["imp"]:
                    v = C.num(f[cc["imp"]])
                    if v:
                        o["imp"].append(v)
                if cc["id"] is not None and len(f) > cc["id"]:
                    idv = C.norm_id(f[cc["id"]]) if not isinstance(f[cc["id"]], str) or f[cc["id"]].strip().isdigit() else None
                    if idv:
                        o["ids"][(idv, d["Origen"])] += 1
                n += 1
            leidas[(m, d["Origen"])] += 1
    return obs, leidas


# ---------------------------------------------------------------- 2. crosswalks existentes
def leer_crosswalks():
    cw = collections.defaultdict(dict)  # mutual -> clave -> (Id, fuente, confianza)
    # Entidad 5: Maestro TIPO NUMERO (DNI + Nº)
    wb = openpyxl.load_workbook(C.RAIZ / "04_Reglas" / "Maestro_TIPO_NUMERO_Entidad 5.xlsx", read_only=True, data_only=True)
    rs = list(wb["Buscar por TIPO NUMERO"].iter_rows(values_only=True))
    hh = [str(x or "") for x in rs[0]]
    for r in rs[1:]:
        d = dict(zip(hh, r))
        if d.get("Id Plenario"):
            cw["entidad5"][(dni_de(d["DNI"]), num_de(d["Nº"]))] = (C.norm_id(d["Id Plenario"]), "Maestro TIPO NUMERO Entidad 5", 2)
    wb.close()
    # ENTIDAD1: overrides confirmados > PASA A > crosswalk ABR26 (por Docum)
    wb = openpyxl.load_workbook(Path(r"<ruta>"), read_only=True, data_only=True)
    for r in list(wb["Crosswalk_Base_ABR26"].iter_rows(min_row=3, values_only=True)):
        if r[0] and r[1]:
            cw["entidad1"][("*", num_de(r[0]))] = (C.norm_id(r[1]), "Crosswalk ENTIDAD1 ABR26 (errores conocidos)", 2)
    for r in list(wb["Overrides_Confirmados"].iter_rows(min_row=2, values_only=True)):
        if r[0] and r[1]:
            cw["entidad1"][("*", num_de(r[0]))] = (C.norm_id(r[1]), f"Override confirmado ENTIDAD1 ({r[4] or ''})", 0)
    wb.close()
    # ENTIDAD3: OP -> Id
    with open(Path(r"C:\ENTIDAD3\Entidad 3_crosswalk_combinado.csv"), encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=";"):
            if r.get("Id_Cooperativa"):
                cw["entidad3"][("*", num_de(r["Orden Entidad 3"]))] = (C.norm_id(r["Id_Cooperativa"]), "Crosswalk ENTIDAD3 combinado", 2)
    # Crosswalk_ENTIDAD3.xlsx de Descargas (el más nuevo de el analista): pisa al combinado
    wb = openpyxl.load_workbook(Path(r"<ruta>"), read_only=True, data_only=True)
    rs = list(wb["Crosswalk"].iter_rows(values_only=True))
    hh = [str(x or "") for x in rs[0]]
    for r in rs[1:]:
        d = dict(zip(hh, r))
        if d.get("OP") and d.get("Id"):
            fuente = str(d.get("Fuente") or "") + (f" / {d['Cómo se cruzó']}" if d.get("Cómo se cruzó") else "")
            conf = 1 if any(k in fuente.lower() for k in ("verificado", "manual confirmado")) else 2
            cw["entidad3"][("*", num_de(d["OP"]))] = (C.norm_id(d["Id"]), f"Crosswalk_ENTIDAD3.xlsx ({fuente})", conf)
    wb.close()
    # ENTIDAD2: Comprob + DNI -> Id
    wb = openpyxl.load_workbook(Path(r"<ruta>"), read_only=True, data_only=True)
    rs = list(wb["Crosswalk Plenario-ENTIDAD2"].iter_rows(values_only=True))
    hh = [str(x or "") for x in rs[0]]
    for r in rs[1:]:
        d = dict(zip(hh, r))
        if d.get("Id_Plenario") and d.get("Comprob_SUE") and str(d.get("Match_Status", "")).startswith("Confirmado"):
            cw["entidad2"][(dni_de(d["DNI"]), num_de(d["Comprob_SUE"]))] = (C.norm_id(d["Id_Plenario"]), "Crosswalk Plenario-ENTIDAD2", 2)
    wb.close()
    return cw


def leer_reportes_cerrados():
    """Hojas Matcheados/Cancelados de los Reporte_APL_* / Reporte_ENTIDAD3_* (imputaciones ya validadas).
    De cada mes se toma solo la ÚLTIMA versión (por fecha de modificación)."""
    raices = [Path(r"<ruta>"), Path(r"C:\ENTIDAD1"), Path(r"C:\ENTIDAD3"), INV.COBRANZA / "ENTIDAD1",
              INV.COBRANZA / "ENTIDAD3"]
    cand = collections.defaultdict(list)
    for r in raices:
        if not r.exists():
            continue
        for f in r.rglob("Reporte_*.xlsx"):
            n = f.name
            mm = re.search(r"Reporte_(ENTIDAD1|ENTIDAD3)_([A-Za-z]+)(20\d\d)", n)
            if not mm or any(x in n for x in ("Imputacion", "Diferencias", "SANCHEZ")) or f.name.startswith("~$"):
                continue
            cand[(mm.group(1), mm.group(2).lower(), mm.group(3))].append(f)
    ev = collections.defaultdict(dict)  # mutual -> ("*", num) -> (Id, fuente)
    ev_obs = []  # filas del reporte como observaciones propias (créditos imputados por fuera del número habitual)
    usados = []
    for (mut, mes, anio), fs in sorted(cand.items()):
        f = max(fs, key=lambda x: x.stat().st_mtime)
        m = "entidad1" if mut == "ENTIDAD1" else "entidad3"
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        except Exception:
            continue
        n = 0
        for hoja in ("Matcheados", "Cancelados"):
            if hoja not in wb.sheetnames:
                continue
            rs = list(wb[hoja].iter_rows(values_only=True))
            hh = [C.sin_acentos(x or "") for x in rs[0]]
            inum = next((i for i, x in enumerate(hh) if "docum" in x or x.startswith("op/") or x == "op"), None)
            iid = next((i for i, x in enumerate(hh) if x in ("id_credito", "id", "id_credito_cancelado")), None)
            idni = next((i for i, x in enumerate(hh) if x == "dni"), None)
            if inum is None or iid is None:
                continue
            for r in rs[1:]:
                idv = C.norm_id(r[iid])
                # un reporte puede traer varios Docum juntos ("13215, 13215")
                for parte in re.split(r"[,;/]", str(r[inum] or "")):
                    num = num_de(parte.strip())
                    if num and idv:
                        fuente = f"Reporte cerrado {mes} {anio} ({f.name}, hoja {hoja})"
                        ev[m][("*", num)] = (idv, fuente)
                        dni_r = dni_de(r[idni]) if idni is not None else ""
                        ev_obs.append((m, dni_r, num, idv, fuente, f"{anio}-{MES_NUM.get(mes, 0):02d}"))
                        n += 1
        wb.close()
        usados.append([mut, f"{mes} {anio}", f.name, str(f.parent), n])
    return ev, usados, ev_obs


MES_NUM = {m: i for i, m in enumerate(["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
                                       "septiembre", "octubre", "noviembre", "diciembre"], 1)}


# ---------------------------------------------------------------- 3. resolver
CORRECCIONES = SALIDA / "Correcciones_Manuales.csv"
ENCABEZADO_CORR = "Mutual;DNI;Numero;Id;Motivo;Fecha\n"


def leer_correcciones():
    """Lo único que se edita a mano. Una fila por caso; manda sobre cualquier otra fuente y nunca se pisa.
    Id = el Id de Plenario correcto, o EXCLUIR para sacar ese número del maestro. DNI vacío = vale para cualquier DNI."""
    if not CORRECCIONES.exists():
        CORRECCIONES.write_text(ENCABEZADO_CORR, encoding="utf-8-sig")
    inv = {v: v for v in NOMBRE_MUTUAL.values()} | {k: v for k, v in NOMBRE_MUTUAL.items()}
    out, errores = {}, []
    with open(CORRECCIONES, encoding="utf-8-sig", newline="") as f:
        for n, r in enumerate(csv.DictReader(f, delimiter=";"), 2):
            if not any((v or "").strip() for v in r.values()):
                continue
            mu = inv.get((r.get("Mutual") or "").strip()) or mutual_de(r.get("Mutual") or "")
            num = num_de((r.get("Numero") or "").strip())
            idv = (r.get("Id") or "").strip()
            if mu not in NOMBRE_MUTUAL.values() or not num or not (idv.upper() == "EXCLUIR" or C.norm_id(idv)):
                errores.append(f"Correcciones_Manuales.csv fila {n}: falta Mutual / Numero / Id válido -> {dict(r)}")
                continue
            dni = dni_de(r.get("DNI")) or "*"
            out[(mu, dni, num)] = ("EXCLUIR" if idv.upper() == "EXCLUIR" else C.norm_id(idv),
                                   (r.get("Motivo") or "").strip(), (r.get("Fecha") or "").strip())
    for e in errores:
        print("AVISO", e)
    return out


def main():
    corr = leer_correcciones()
    print(f"Correcciones manuales: {len(corr)}")
    print(f"Leyendo Plenario: {PLENARIO.name}")
    plen = [p for p in C.leer_plenario(PLENARIO)]
    por_id = {p["Id"]: p for p in plen}
    por_dni = collections.defaultdict(list)
    por_dni_ref = collections.defaultdict(list)
    for p in plen:
        mu = mutual_de(p["Entidad"])
        por_dni[(mu, p["DNI"])].append(p)
        por_dni_ref[(mu, p["DNI"], C.norm_num(p["CodRef"]))].append(p)
    print("Leyendo rendiciones...")
    obs, leidas = leer_observaciones()
    print("Leyendo crosswalks...")
    cw = leer_crosswalks()
    rep, rep_usados, rep_obs = leer_reportes_cerrados()
    # los pares del reporte que no aparecen en ninguna rendición leída también entran al maestro
    nums_vistos = {(k[0], k[2]) for k in obs if k[2]}
    for m, dni_r, num, idv, fuente, per in rep_obs:
        dni_r = dni_r or (por_id[idv]["DNI"] if idv in por_id else "")
        if not dni_r or (m, num) in nums_vistos:
            continue  # ese número ya está en una rendición: no se duplica con el DNI del titular
        o = obs.setdefault((m, dni_r, num), {"nombre": por_id[idv]["Cliente"] if idv in por_id else "", "per": set(),
                                             "arch": set(), "ids": collections.Counter(), "imp": [], "veces": 0,
                                             "tipo_txt": num})
        o["per"].add(per)
        o["arch"].add(fuente)
        o["veces"] += 1

    # pares que el analista cargó a mano y no aparecen en ninguna rendición leída
    clave_m = {v: k for k, v in NOMBRE_MUTUAL.items()}
    for (mu, dni, num), (idv, mot, fecha) in corr.items():
        if dni != "*" and (clave_m[mu], dni, num) not in obs:
            obs[(clave_m[mu], dni, num)] = {"nombre": "", "per": set(), "arch": {"Correcciones_Manuales.csv"},
                                            "ids": collections.Counter(), "imp": [], "veces": 0, "tipo_txt": num}
    usadas = set()
    filas, conflictos = [], []
    for (m, dni, num), o in sorted(obs.items()):
        mu = NOMBRE_MUTUAL[m]
        cand = []  # (confianza, Id, fuente)
        for (idv, origen), n in o["ids"].items():
            cand.append((1 if origen == "Copia trabajada" else 2, idv,
                         f"ID cargado en la rendición ({'copia trabajada' if origen == 'Copia trabajada' else 'original'}, {n} filas)"))
        for clave in [(dni, num), ("*", num)]:
            if num and clave in cw[m]:
                idv, fuente, conf = cw[m][clave]
                cand.append((conf, idv, fuente))
            if num and clave in rep[m]:
                cand.append((1, rep[m][clave][0], rep[m][clave][1]))
        if m == "entidad6" and "|" in num:
            padron, cuota = num.split("|")
            lst = [p_ for p_ in por_dni.get((mu, dni), []) if abs(p_["ValorCuota"] - float(cuota)) <= 1]
            mismo_padron = [p_ for p_ in lst if C.norm_num(p_["CodRef"]) == padron]
            if len(mismo_padron) == 1:
                cand.append((2, mismo_padron[0]["Id"], "PADRON = Código de Referencia + DNI + cuota"))
            elif len(lst) == 1:
                cand.append((2, lst[0]["Id"], "DNI + cuota (el PADRON de Plenario es distinto)"))
            elif len(mismo_padron) > 1 or len(lst) > 1:
                activos_ = [p_ for p_ in (mismo_padron or lst) if p_["Estado"] in ("Vigente", "Mora")]
                if len(activos_) == 1:
                    cand.append((3, activos_[0]["Id"], "Sugerido: único crédito activo con ese DNI y cuota"))
        elif num:
            if m == "tresabril":
                if C.norm_id(num) in por_id:
                    cand.append((2, C.norm_id(num), "Referencia externa = Id de Plenario"))
            lst = por_dni_ref.get((mu, dni, num), [])
            if len(lst) == 1:
                cand.append((2, lst[0]["Id"], "Número = Código de Referencia de Plenario (mismo DNI)"))
        if not cand:
            lst = por_dni.get((mu, dni), [])
            if len(lst) == 1:
                cand.append((3, lst[0]["Id"], "Sugerido: único crédito del DNI en la mutual"))
            elif lst and o["imp"]:
                # DNI + valor de cuota: el importe más frecuente de la rendición contra la cuota de Plenario
                cuota = collections.Counter(round(v, 2) for v in o["imp"]).most_common(1)[0][0]
                x = [p for p in lst if abs(p["ValorCuota"] - cuota) <= 1]
                if len(x) == 1:
                    cand.append((3, x[0]["Id"], f"Sugerido: DNI + cuota {cuota:,.2f}"))
        # un Id solo vale si en Plenario es del MISMO DNI y de la MISMA mutual; lo confirmado a mano se respeta
        # (ENTIDAD1: hay casos PASA A donde paga otra persona) pero queda marcado
        nota, validos = [], []
        for c in sorted(set(cand)):
            p = por_id.get(c[1])
            mismo_dni = bool(p) and p["DNI"] == dni
            misma_mut = bool(p) and mutual_de(p["Entidad"]) == mu
            if mismo_dni and misma_mut:
                validos.append(c)
            elif c[0] == 1 and misma_mut:
                validos.append(c)
                nota.append(f"Confirmado a mano pero el Id {c[1]} es de otro DNI en Plenario ({p['DNI']}): ¿paga otra persona?")
            else:
                motivo = "no existe en Plenario" if not p else (f"es de otra mutual ({p['Entidad']})" if not misma_mut
                                                                 else f"es de otro DNI ({p['DNI']})")
                nota.append(f"Descartado {c[1]} de '{c[2]}': {motivo}")
        ids = sorted({c[1] for c in validos})
        resuelto_por_regla = ""
        if len(ids) > 1:
            # 1) un override confirmado manda sobre todo lo demás
            top = [c for c in validos if c[0] == 0]
            # 2) si no: el Id que sostienen 2+ fuentes distintas, o el de 75%+ de las filas de las rendiciones
            if not top:
                apoyo = collections.Counter()
                filas_id = collections.Counter()
                for c in validos:
                    apoyo[c[1]] += 1
                    mm = re.search(r"\((?:copia trabajada|original), (\d+) filas\)", c[2])
                    if mm:
                        filas_id[c[1]] += int(mm.group(1))
                tot = sum(filas_id.values())
                fuerte = [i for i, n in apoyo.items() if n >= 2 and n > max(v for k, v in apoyo.items() if k != i)]
                mayor = [i for i, n in filas_id.items() if tot and n / tot >= 0.75]
                gana = (fuerte or mayor or [None])[0]
                if gana:
                    top = [c for c in validos if c[1] == gana]
                    resuelto_por_regla = "mayoría"
            else:
                resuelto_por_regla = "override"
            if top:
                otros = sorted({c[1] for c in validos if c[1] != top[0][1]})
                nota.append(f"Ids alternativos descartados por {resuelto_por_regla}: {', '.join(map(str, otros))}")
                validos, ids = top, [top[0][1]]
        if validos:
            mejor = min(validos, key=lambda c: c[0])
            idf, conf, fuente = mejor[1], min(mejor[0], 2) if mejor[0] == 0 else mejor[0], mejor[2]
            conf = 1 if mejor[0] == 0 else conf
        else:
            idf, conf, fuente = None, 4, ""
        if len(ids) > 1:
            conflictos.append([mu, dni, o["tipo_txt"] or num, o["nombre"],
                               " | ".join(f"{c[1]} ← {c[2]}" for c in validos)])
        p = por_id.get(idf) if idf else None
        etiqueta = {1: "1 - Confirmado", 2: "2 - Crosswalk / Plenario", 3: "3 - Sugerido (revisar)", 4: "4 - Sin asignar"}[conf]
        if len(ids) > 1:
            etiqueta = "Revisar (fuentes con Ids distintos)"
        elif conf == 1 and p and p["DNI"] != dni:
            # si sale de una imputación ya validada (reporte cerrado u override), es un caso de "paga otra persona"
            # lo cruzado a mano por el analista (reporte cerrado, override o ID cargado en copia trabajada) se respeta:
            # son los casos "paga otra persona" o filas con un CUIL genérico (ej. 00000000) en la rendición
            etiqueta = ("1 - Confirmado (paga otra persona)" if any(k in fuente for k in ("Reporte cerrado", "Override", "copia trabajada"))
                        else "Revisar (Id de otro DNI)")
        elif conf == 4 and not num:
            etiqueta = "Sin número en la rendición"
        # la corrección manual de el analista manda sobre todo lo anterior
        c_ = corr.get((mu, dni, num)) or corr.get((mu, "*", num))
        if c_:
            usadas.add((mu, dni, num) if (mu, dni, num) in corr else (mu, "*", num))
            if c_[0] == "EXCLUIR":
                idf, p, etiqueta, fuente = None, None, "Excluido (corrección manual)", f"Corrección manual: {c_[1]}"
            else:
                if idf and idf != c_[0]:
                    nota.append(f"Antes de la corrección manual resolvía a {idf}")
                idf = c_[0]
                p = por_id.get(idf)
                etiqueta, fuente = "1 - Confirmado (corrección manual)", f"Corrección manual {c_[2]}: {c_[1]}".strip()
                if not p:
                    nota.append(f"OJO: el Id {idf} de la corrección no existe en Plenario")
        per = sorted(p_ for p_ in o["per"] if p_)
        filas.append([mu, dni, TIPO_NUMERO[m], o["tipo_txt"] or num, num, o["nombre"], idf,
                      p["Cliente"] if p else None, p["Estado"] if p else None, p["ValorCuota"] if p else None,
                      p["CodRef"] if p else None, etiqueta, fuente, "; ".join(nota),
                      per[0] if per else None, per[-1] if per else None, len(per), o["veces"],
                      " ; ".join(sorted(o["arch"]))[:1500]])

    cols = ["Mutual", "DNI", "Tipo de número", "Número de la mutual (como viene)", "Número normalizado", "Nombre en la rendición",
            "Id Plenario", "Cliente Plenario", "Estado Plenario", "Valor Cuota Plenario", "Código Referencia Plenario",
            "Confianza", "Fuente de la equivalencia", "Observación", "Primer período visto", "Último período visto",
            "Períodos", "Filas en rendiciones", "Archivos"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen"
    res = collections.Counter((f[0], f[11]) for f in filas)
    C.encabezado(ws, 1, ["Mutual", "Confianza", "Pares DNI + número"])
    for i, ((mu, cf), n) in enumerate(sorted(res.items()), 2):
        C.escribir_fila(ws, i, [mu, cf, n])
    C.hoja_tabla(wb, "Maestro", cols, filas, resaltar=lambda r: r[11].startswith(("Revisar", "4", "3")))
    C.hoja_tabla(wb, "Sin asignar", cols, [f for f in filas if f[11] == "4 - Sin asignar"])
    C.hoja_tabla(wb, "Revisar", cols, [f for f in filas if f[11].startswith(("Revisar", "3"))])
    C.hoja_tabla(wb, "Sin número (formato viejo)", cols, [f for f in filas if f[11] == "Sin número en la rendición"])
    C.hoja_tabla(wb, "Conflictos", ["Mutual", "DNI", "Número", "Nombre", "Ids propuestos y de dónde salen"], conflictos)
    # lo que hay que resolver ANTES de la próxima imputación
    import datetime as dt
    firme = lambda f: f[11][:1] in ("1", "2")
    ult = {}
    for f in filas:
        solo_reporte = all(a.strip().startswith("Reporte cerrado") for a in (f[18] or "").split(" ; ") if a.strip())
        if f[15] and not solo_reporte:  # la "última rendición" se mide con rendiciones, no con reportes
            ult[f[0]] = max(ult.get(f[0], ""), f[15])
    con_equiv = {(f[0], f[6]) for f in filas if f[6] and firme(f)}
    inicio_mes = dt.date.today().replace(day=1)
    pend = [["En la última rendición sin resolver", ult.get(f[0])] + f for f in filas
            if f[15] == ult.get(f[0]) and not firme(f)]
    for p_ in plen:
        mu = mutual_de(p_["Entidad"])
        if (mu in NOMBRE_MUTUAL.values() and p_["Estado"] in ("Vigente", "Mora") and (mu, p_["Id"]) not in con_equiv
                and p_["F1Vto"] and p_["F1Vto"].date() < inicio_mes):
            pend.append(["Crédito activo (ya venció su 1ª cuota) sin número de la mutual en el maestro", ult.get(mu),
                         mu, p_["DNI"], TIPO_NUMERO[[k for k, v in NOMBRE_MUTUAL.items() if v == mu][0]], None, None, None,
                         p_["Id"], p_["Cliente"], p_["Estado"], p_["ValorCuota"], p_["CodRef"], "", "",
                         f"1ª cuota {p_['F1Vto']:%d/%m/%Y}", None, None, 0, 0, ""])
    C.hoja_tabla(wb, "Pendientes antes de imputar", ["Motivo", "Última rendición de la mutual"] + cols,
                 sorted(pend, key=lambda r: (r[2], r[0], str(r[3]))))
    wb.move_sheet("Pendientes antes de imputar", offset=-(len(wb.sheetnames) - 2))
    print("Pendientes antes de imputar: " + ", ".join(
        f"{k}: {v}" for k, v in sorted(collections.Counter((r[2], r[0][:12]) for r in pend).items())))

    # control: un Id nuestro no puede quedar asignado a dos personas distintas
    dnis_por_id = collections.defaultdict(set)
    for f in filas:
        if f[6]:
            dnis_por_id[(f[0], f[6])].add(f[1])
    ctrl = [f for f in filas if f[6] and len(dnis_por_id[(f[0], f[6])]) > 1]
    C.hoja_tabla(wb, "Control Id con 2 DNIs", cols, sorted(ctrl, key=lambda f: (f[0], f[6])))
    C.hoja_tabla(wb, "Reportes cerrados usados", ["Mutual", "Mes", "Archivo (última versión)", "Carpeta", "Equivalencias"], rep_usados)
    C.hoja_tabla(wb, "Hojas leídas", ["Mutual", "Origen", "Hojas"], [[NOMBRE_MUTUAL[m], o, n] for (m, o), n in sorted(leidas.items())])
    sin_uso = [[k[0], k[1], k[2], v[0], v[1]] for k, v in corr.items() if k not in usadas]
    if sin_uso:
        C.hoja_tabla(wb, "Correcciones sin uso", ["Mutual", "DNI", "Número", "Id", "Motivo"], sin_uso)
        print(f"AVISO {len(sin_uso)} correcciones manuales no coinciden con ningún número de las rendiciones (ver hoja)")
    ws = wb.create_sheet("LEEME", 0)
    for i, t in enumerate([
            "BASE MAESTRA DE EQUIVALENCIAS — número de la mutual + DNI -> Id de Plenario",
            "",
            "Este archivo lo arma el script construir_maestro.py: NO se edita a mano (se regenera entero en cada corrida).",
            "Para corregir un caso: agregar una fila en 07_Base_Maestra\\Correcciones_Manuales.csv (se abre con el Bloc de notas o Excel):",
            "   Mutual;DNI;Numero;Id;Motivo;Fecha",
            "   ej.: ENTIDAD1;00000000;123456;1234;Confirmado con la entidad;19/10/2026",
            "   Id = Id de Plenario correcto, o EXCLUIR para sacar ese número. DNI vacío = vale para cualquier DNI con ese número.",
            "   La corrección manual manda sobre cualquier otra fuente y el script nunca la pisa.",
            "Los scripts de cada mutual leen Base_Maestra.csv (misma información que la hoja Maestro).",
            "Confianza: 1 Confirmado (corrección manual / override / reporte cerrado / ID cargado a mano) · 2 Crosswalk o Plenario · "
            "3 Sugerido (revisar) · 4 Sin asignar.",
            f"Plenario usado: {PLENARIO.name}"], 1):
        ws.cell(i, 1, t)
    C.autofit(wb)
    with open(SALIDA / "Base_Maestra.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(cols[:19])
        for fila in filas:
            w.writerow(["" if v is None else v for v in fila[:19]])
    p = SALIDA / "Base_Maestra_Equivalencias.xlsx"
    try:
        wb.save(p)
    except PermissionError:
        p = SALIDA / "Base_Maestra_Equivalencias_NUEVO.xlsx"
        wb.save(p)
        print("AVISO Base_Maestra_Equivalencias.xlsx estaba abierto en Excel: se guardó como _NUEVO.xlsx")
    print("\n=== RESUMEN ===")
    for mu in NOMBRE_MUTUAL.values():
        t = {cf: n for (m_, cf), n in res.items() if m_ == mu}
        if t:
            print(f"  {mu:24s} total {sum(t.values()):6,} | " + " | ".join(f"{k}: {v:,}" for k, v in sorted(t.items())))
    print(f"Control: {len({(f[0], f[6]) for f in ctrl})} Ids asignados a 2+ DNIs "
          f"({', '.join(f'{k}: {v}' for k, v in collections.Counter(f[0] for f in {(f[0], f[6]): f for f in ctrl}.values()).items())})")
    print(f"\nGuardado: {p}")


if __name__ == "__main__":
    main()

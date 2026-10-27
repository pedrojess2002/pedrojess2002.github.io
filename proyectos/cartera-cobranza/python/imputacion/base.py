# -*- coding: utf-8 -*-
"""Piezas comunes de la imputación (todas las mutuales).

Regla de el analista: la rendición manda. La imputación toma el VtoTeórico del mes (armado por vto_teorico, Cobrado vacío),
le carga lo que rindió la mutual crédito por crédito y agrega la columna "cobrado" al LP crudo de Plenario.
Si algo no cierra al centavo, no se entrega nada: se frena y se muestra qué falta.

Entradas:  cartera\\01_Entrada\\AAAA-MM\\Rendiciones\\  (rendición / mail de cada mutual)
           cartera\\01_Entrada\\AAAA-MM\\Plenario\\     (LP crudo .csv de cada mutual, bajado de Plenario)
           cartera\\01_Entrada\\AAAA-MM\\VtoTeorico\\   (lo arma Generar_VtoTeorico.bat)
Salida:    Cobranza\\Imputaciones\\<Mes Año>\\<Mutual>\\
"""
import collections
import re
import sys
from pathlib import Path

import openpyxl
import pandas as pd

AQUI = Path(__file__).resolve().parent
sys.path[:0] = [str(AQUI.parent), str(AQUI.parent / "vto_teorico")]
import build_vto_teorico as B  # noqa: E402
import comun as C  # noqa: E402
from comun import FrenarProceso  # noqa: E402

COBRANZA = Path(r"<ruta> Solutions\el área de riesgos - Riesgos - Operaciones\Cobranza")
IMPUTACIONES = COBRANZA / "Imputaciones"  # salida: Imputaciones\<Mes Año>\<Mutual>\ (el analista después lo mueve a su carpeta)
SALIDA_PRUEBAS = None  # imputar.py --out: en pruebas, todo se escribe ahí y no se mira Cobranza
# carpeta de cada mutual en Cobranza; adentro, una carpeta por mes con la rendición, el LP crudo y el VtoTeórico
CARPETA_COBRANZA = {"Entidad 5": "Entidad 5", "ENTIDAD1": "ENTIDAD1", "ENTIDAD3": "ENTIDAD3", "ENTIDAD6": "ENTIDAD6",
                    "Entidad 8": "Entidad 8", "Entidad 4": "Entidad 4",
                    "ASOC Entidad 7": "Entidad 7", "ENTIDAD2": "ENTIDAD 2"}
# meses de diferencia entre la carpeta de la rendición y el VtoTeórico donde se imputa (ENTIDAD2: la carpeta "Septiembre"
# trae la rendición del 09/09 al 09/10, que se imputa en octubre)
DESFASE_CARPETA = {"ENTIDAD2": -1}
MESES_TXT = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
             "noviembre", "diciembre"]


def pesos(x):
    return "$" + f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def primer_plenario_del_mes(anio, nmes):
    """El 'data plenario DD-MM-AAAA.xls' más temprano DEL mes (Descargas o CarteraDW), o None."""
    import datetime as dt
    cands = []
    for carpeta in (Path(r"<ruta>"), Path(r"procesados"), Path(r"entrada")):
        if carpeta.exists():
            for f in carpeta.glob("*.xls"):
                m = re.search(r"plenario\D*(\d{1,2})-(\d{1,2})-(\d{2,4})", f.name, re.I)
                if m:
                    a = int(m.group(3)) + (2000 if len(m.group(3)) == 2 else 0)
                    try:
                        d = dt.date(a, int(m.group(2)), int(m.group(1)))
                    except ValueError:
                        continue
                    if (d.year, d.month) == (anio, nmes):
                        cands.append((d, f))
    return sorted(cands)[0][1] if cands else None


def armar_vto_teorico(mes, nombre):
    """La primera vez que se imputa un mes, arma su VtoTeórico con el primer data plenario del mes (queda congelado)."""
    import subprocess
    anio, nmes = map(int, mes.split("-"))
    f = primer_plenario_del_mes(anio, nmes)
    if f is None:
        raise FrenarProceso(f"No está el VtoTeórico de {nombre} y no encontré un data plenario de {nombre} en Descargas: "
                            f"bajá el data plenario de los primeros días del mes (o corré Generar_VtoTeorico.bat con {mes})")
    print(f"No estaba el VtoTeórico de {nombre}: lo armo ahora con {f.name} (queda fijo para todo el mes)...")
    cmd = [sys.executable, "-I", str(AQUI.parent / "vto_teorico" / "generar_vto_teorico.py"), mes, str(f)]
    if C.RAIZ != AQUI.parent.parent:  # (pruebas sobre una copia: no tocar la carpeta real ni Agenda_Vtos)
        cmd += ["--out", str(C.RAIZ / "01_Entrada" / mes / "VtoTeorico")]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if p.returncode not in (0, 3) or not (C.RAIZ / "01_Entrada" / mes / "VtoTeorico" / "_universo.pkl").exists():
        raise FrenarProceso(f"No se pudo armar el VtoTeórico de {nombre}:\n{p.stdout[-1500:]}{p.stderr[-800:]}")
    print("  " + "\n  ".join(l for l in p.stdout.splitlines() if l.strip() and not l.startswith("Salida"))[-1200:])


class Mes:
    """El mes que se imputa: su VtoTeórico (universo congelado) y el Plenario con el que se armó."""

    def __init__(self, mes):
        if not re.fullmatch(r"20\d\d-(0[1-9]|1[0-2])", mes or ""):
            raise FrenarProceso(f"El mes tiene que ser AAAA-MM (ej. 2026-10). Llegó: '{mes}'")
        self.mes = mes
        self.anio, self.nmes = map(int, mes.split("-"))
        self.tag = f"{B.MESES[self.nmes]}{self.anio}"
        self.nombre = f"{B.MESES[self.nmes]} {self.anio}"
        self.entrada = C.RAIZ / "01_Entrada" / mes
        vt = self.entrada / "VtoTeorico"
        if not (vt / "_universo.pkl").exists():
            armar_vto_teorico(mes, self.nombre)
        self.plenario = next(iter(C.archivos(vt / "_plenario_congelado", ".xls", ".xlsx")), None)
        if self.plenario is None:
            raise FrenarProceso(f"El VtoTeórico de {self.nombre} no está armado con un Plenario del mes: "
                                f"volvé a correr Generar_VtoTeorico.bat con {mes} y el primer data plenario del mes")
        self.u = pd.read_pickle(vt / "_universo.pkl")
        self._df = None

    @property
    def df(self):
        """Todos los créditos del Plenario congelado del mes (todas las mutuales y estados), más los créditos dados de
        alta después (los toma del data plenario más reciente: una rendición puede cobrar un crédito nuevo)."""
        if self._df is None:
            df = B.build_universe(str(self.plenario), self.nmes, self.anio)[0]
            df["Alta posterior"] = False
            sys.path.insert(0, str(AQUI.parent / "maestra"))
            from construir_maestro import plenario_mas_reciente
            try:
                nuevo = plenario_mas_reciente()
            except SystemExit:  # no hay un data plenario más nuevo: se trabaja solo con el congelado
                nuevo = None
            if nuevo and Path(nuevo).name != self.plenario.name:
                dn = B.build_universe(str(nuevo), self.nmes, self.anio)[0]
                dn = dn[~dn["Id"].isin(df["Id"])].copy()
                dn["Alta posterior"] = True
                df = pd.concat([df, dn], ignore_index=True)
            df["Categoria"] = df["Días Mora"].fillna(0).apply(B.categoria)
            self._df = df
        return self._df

    def universo(self, ent):
        return self.u[self.u["EntidadMap"] == ent]

    def ids_mutual(self, ent):
        return set(self.df.loc[self.df["EntidadMap"] == ent, "Id"].astype(int))

    @property
    def prueba(self):
        """Corriendo sobre una copia (pruebas): solo se miran las carpetas de la copia, nunca Cobranza."""
        return C.RAIZ != AQUI.parent.parent or SALIDA_PRUEBAS is not None

    def salida(self, ent):
        d = Path(SALIDA_PRUEBAS if SALIDA_PRUEBAS is not None else IMPUTACIONES) / self.nombre / ent
        d.mkdir(parents=True, exist_ok=True)
        # antes de escribir nada: si una salida anterior está abierta en Excel, frenar ahora (y no dejar la mitad escrita)
        abiertos = []
        for f in d.iterdir():
            if f.is_file() and not f.name.startswith("~$"):
                try:
                    with open(f, "r+b"):
                        pass
                except PermissionError:
                    abiertos.append(f.name)
        if abiertos:
            raise FrenarProceso(f"Cerrá en Excel {', '.join(abiertos)} ({d}) y volvé a correr: no escribí nada.")
        return d

    def carpetas_cobranza(self, ent):
        """La carpeta del mes de la mutual en Cobranza (ej. Cobranza\\ENTIDAD1\\Octubre 2026, Cobranza\\Mutual Entidad 5
        Cba\\2026\\Vto Septiembre 2026, Cobranza\\Entidad 4\\2026\\Septiembre): ahí el analista pone la rendición y el LP."""
        raiz = COBRANZA / CARPETA_COBRANZA.get(ent, "")
        if not raiz.exists():
            return []
        a, m = self.anio, self.nmes + DESFASE_CARPETA.get(ent, 0)
        if m < 1:
            a, m = a - 1, m + 12
        mes_txt, anio_txt = MESES_TXT[m], str(a)
        out = []
        for d in raiz.rglob("*"):
            if not d.is_dir() or len(d.relative_to(raiz).parts) > 3:
                continue
            nombre = C.sin_acentos(d.name)
            nombre = re.sub(r"^(vto|vencimiento|entidad6|entidad3|entidad5|entidad1)[\s_-]+", "", nombre).strip()
            nombre = re.sub(r"^(entidad5|ENTIDAD 2", "", nombre).strip()
            if nombre == f"{mes_txt} {anio_txt}" or (nombre == mes_txt and anio_txt in str(d.relative_to(raiz))):
                out.append(d)
        return sorted(out, key=lambda d: len(d.parts))

    def dirs_rendicion(self, ent):
        if self.prueba:
            return [self.entrada / "Rendiciones"]
        return self.carpetas_cobranza(ent) + [self.entrada / "Rendiciones"]

    def dirs_lp(self, ent):
        if self.prueba:
            return [self.entrada / "Plenario"]
        return self.carpetas_cobranza(ent) + [self.entrada / "Plenario"]

    def donde_poner(self, ent):
        """Texto para los avisos: la carpeta del mes de la mutual (o cómo debería llamarse si todavía no existe)."""
        d = self.carpetas_cobranza(ent)
        return str(d[0]) if d else f"Cobranza\\{CARPETA_COBRANZA.get(ent, ent)}\\<carpeta del mes {self.nombre}>"


class LP:
    """LP crudo de Plenario (.csv). Se le agrega la columna 'cobrado' al final sin tocar nada más
    (mismo encoding, separador y fin de línea; mismas filas y en el mismo orden)."""

    def __init__(self, path):
        self.path = Path(path)
        _, self.lineas, self.sep, self.eol, self.enc = C.leer_csv_lp(path)
        self.hdr = self.lineas[0].split(self.sep)
        need = ["conceptonumero", "importecobrado", "fechacobro", "cuota"]
        if any(h not in self.hdr for h in need):
            raise FrenarProceso(f"{self.path.name}: no es un LP de Plenario (faltan columnas {need})")
        ix = {h: self.hdr.index(h) for h in need + ["nombre", "documentonumero"] if h in self.hdr}
        self.filas = []
        for i, l in enumerate(self.lineas[1:], 1):
            if not l.strip():
                continue
            v = l.split(self.sep)
            if len(v) < len(self.hdr):
                raise FrenarProceso(f"{self.path.name}: la fila {i + 1} tiene {len(v)} columnas y el encabezado {len(self.hdr)} "
                                    f"(¿el archivo se cortó o se abrió y guardó en Excel?)")
            self.filas.append({"linea": i, "Id": C.norm_id(v[ix["conceptonumero"]]), "Importe": C.num(v[ix["importecobrado"]]),
                               "FechaTxt": v[ix["fechacobro"]], "Fecha": C.fecha_xls(v[ix["fechacobro"]]),
                               "Cuota": v[ix["cuota"]], "Nombre": v[ix["nombre"]] if "nombre" in ix else ""})

    @property
    def ya_procesado(self):
        """Ya tiene la columna de cobrado agregada (cobrado / M / Cobrado_APL_Agosto2026...): es una salida, no un LP crudo."""
        ult = self.hdr[-1].strip().lower()
        return ult.startswith("cobrado") or ult == "m"

    def escribir(self, valores, destino, vacio=None):
        """valores = {número de línea: monto}; las líneas sin valor quedan vacías (o en `vacio`, ej. 0 en ENTIDAD3)."""
        out = [self.lineas[0] + self.sep + "cobrado"]
        for i, l in enumerate(self.lineas[1:], 1):
            if l == "" and i == len(self.lineas) - 1:
                out.append("")
                continue
            m = valores.get(i, vacio if l.strip() else None)
            out.append(l + self.sep + (f"{m:.2f}" if m is not None else ""))
        Path(destino).write_bytes(self.eol.join(out).encode(self.enc))


def encontrar(mes, ent, detectar, que, *ext):
    """Busca la rendición de la mutual: 1) la carpeta del mes de la mutual en Cobranza (ej. Cobranza\\ENTIDAD1\\Octubre 2026)
    2) 01_Entrada\\AAAA-MM\\Rendiciones\\. Usa la primera carpeta donde haya UN archivo que `detectar`
    reconozca (detectar(f) devuelve lo leído o None). Si en una carpeta hay más de uno, frena y los lista."""
    vistas = []
    for d in mes.dirs_rendicion(ent):
        if not d.exists():
            continue
        vistas.append(str(d))
        cands, abiertos = [], []
        for f in C.archivos(d, *ext):
            n = f.name.lower()
            if n.startswith(("vtoteorico", "reporte_", "~$")) or "con_columna" in n or "cobrado" in n:
                continue  # salidas del sistema (o de el analista), no rendiciones
            try:
                x = detectar(f)
            except PermissionError:
                abiertos.append(f)
                x = None
            except Exception:
                x = None
            if x is not None:
                cands.append((f, x))
        if len(cands) > 1:  # copias / comparaciones / pre-liquidaciones que el analista guarda junto a la original
            orig = [c for c in cands if not re.search(r"^copia|resincronizado|comparacion|no_?imputado|pre liq|explicacion",
                                                      C.sin_acentos(c[0].name))]
            if len(orig) == 1:
                cands = orig
        if len(cands) > 1:
            raise FrenarProceso(f"Hay {len(cands)} archivos que parecen {que} en {d}: " + ", ".join(f.name for f, _ in cands)
                                + ". Dejá uno solo.")
        if cands:
            if not mes.prueba:
                print(f"  Rendición tomada: {d}\\{cands[0][0].name}")
            return cands[0]
        if abiertos:  # puede ser justo la rendición: no seguir buscando en otro lado sin avisar
            raise FrenarProceso(f"Hay archivos abiertos en Excel que no pude leer en {d}: "
                                + ", ".join(f.name for f in abiertos) + ". Cerralos y volvé a correr.")
    raise FrenarProceso(f"No encontré {que}. Ponela en {mes.donde_poner(ent)}"
                        + (f" (busqué en: {' | '.join(vistas)})" if vistas else ""))


def detectar_lp(mes, ent):
    """El LP crudo de la mutual: 1) la carpeta del mes de la mutual en Cobranza  2) 01_Entrada\\AAAA-MM\\Plenario\\.
    Es el .csv cuyos créditos son (casi) todos de esa mutual y que todavía no trae la columna de cobrado."""
    ids_ent = mes.ids_mutual(ent)
    ids_todos = set(mes.df["Id"].astype(int))
    procesados, rotos = [], []
    for d in mes.dirs_lp(ent):
        if not d.exists():
            continue
        cands, abiertos = [], []
        for f in C.archivos(d, ".csv"):
            try:
                lp = LP(f)
            except PermissionError:
                abiertos.append(f.name)
                continue
            except FrenarProceso as e:
                rotos.append(str(e))
                continue
            # los créditos dados de alta después del Plenario congelado no cuentan ni a favor ni en contra
            ids = {x["Id"] for x in lp.filas if x["Id"] is not None} & ids_todos
            if not ids or len(ids & ids_ent) / len(ids) < 0.9:
                continue
            (procesados if lp.ya_procesado else cands).append(lp)
        if not cands and abiertos:  # puede ser justo el LP de la mutual: no seguir buscando en otro lado sin avisar
            raise FrenarProceso(f"Hay LP abiertos en Excel que no pude leer en {d}: {', '.join(abiertos)}. "
                                f"Cerralos y volvé a correr.")
        if not cands:
            continue
        if len(cands) > 1:
            raise FrenarProceso(f"Hay {len(cands)} LP de {ent} en {d}: " + ", ".join(x.path.name for x in cands) + ". Dejá uno solo.")
        if not mes.prueba:
            print(f"  LP tomado: {d}\\{cands[0].path.name}")
        return cands[0]
    extra = ""
    if procesados:
        extra = (" — " + ", ".join(x.path.name for x in procesados) + " ya trae la columna de cobrado (es una salida, "
                 "no el LP crudo bajado de Plenario)")
    elif rotos:
        extra = " — " + " | ".join(r for r in rotos if "fila" in r)
    donde = mes.entrada / "Plenario" if mes.prueba else mes.donde_poner(ent)
    raise FrenarProceso(f"No encontré el LP crudo de {ent} (bajalo de Plenario y ponelo en {donde}){extra}")


def escribir_vto_final(mes, ent, cobrado, ids_agregados, destino, pie=None):
    """VtoTeórico del mes con el Cobrado cargado (0 si no rindió) y los cobros no esperados abajo, aparte."""
    rows = mes.universo(ent)
    agreg = mes.df[mes.df["Id"].isin(ids_agregados)].sort_values("Id")
    try:
        B.entity_file(ent, rows, destino, cobrado=cobrado, agregados=agreg, pie=pie)
    except PermissionError:
        raise FrenarProceso(f"{Path(destino).name} está abierto en Excel: cerralo y volvé a correr")
    # control: releer lo escrito
    det = list(openpyxl.load_workbook(destino, read_only=True)["Detalle"].iter_rows(min_row=2, values_only=True))
    ids_u = set(rows["Id"].astype(int))
    cob_orig = sum(r[11] or 0 for r in det if C.norm_id(r[0]) in ids_u)
    cob_agr = sum(r[11] or 0 for r in det if C.norm_id(r[0]) is not None and C.norm_id(r[0]) not in ids_u)
    esp_orig = sum(v for k, v in cobrado.items() if k in ids_u)
    esp_agr = sum(v for k, v in cobrado.items() if k in set(ids_agregados))
    if not (C.cierra(cob_orig, esp_orig) and C.cierra(cob_agr, esp_agr)):
        raise FrenarProceso(f"El VtoTeórico escrito no coincide con lo imputado ({cob_orig:.2f}/{esp_orig:.2f}, "
                            f"{cob_agr:.2f}/{esp_agr:.2f})")
    return C.r2(cob_orig), C.r2(cob_agr)


MAESTRA = C.RAIZ / "07_Base_Maestra"


class _Equivalencias(dict):
    """(DNI, número) -> (Id, fuente). Una corrección manual sin DNI vale para cualquier DNI (clave ('*', número))."""

    def get(self, k, default=None):
        return super().get(k) or super().get(("*", k[1]), default)


_ALIAS = {"entidad3": ("amsur", "recolect"), "entidad1": ("entidad1",), "entidad2": ("entidad2", "union", "electrica"),
          "entidad6": ("entidad6", "docente"), "Entidad 5": ("polic",), "Entidad 8": ("abril", "munici")}


def _es_mutual(texto, mutual):
    t = C.sin_acentos(texto or "")
    return any(k in t for k in _ALIAS.get(mutual.lower(), (C.sin_acentos(mutual),)))


def maestro(mutual):
    """Base Maestra de equivalencias (número de la mutual -> Id Plenario) de una mutual ('Entidad 3', 'ENTIDAD1', ...).
    Solo lo confirmado (1) o de crosswalk/Plenario (2); lo 'sugerido' no se usa solo. Las Correcciones_Manuales
    de el analista mandan sobre todo (Id = EXCLUIR saca el número; DNI vacío = cualquier DNI). Clave: (DNI, número normalizado)."""
    import csv
    out = _Equivalencias()
    f = MAESTRA / "Base_Maestra.csv"
    if not f.exists():
        raise FrenarProceso(f"No está {f}: corré construir_maestro.py")
    with open(f, encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh, delimiter=";"):
            if r["Mutual"] != mutual or not r["Confianza"][:1] in ("1", "2") or not r["Id Plenario"]:
                continue
            out[(C.norm_dni(r["DNI"]), r["Número normalizado"])] = (C.norm_id(r["Id Plenario"]), r["Fuente de la equivalencia"])
    fc = MAESTRA / "Correcciones_Manuales.csv"
    if fc.exists():
        with open(fc, encoding="utf-8-sig", newline="") as fh:
            for n, r in enumerate(csv.DictReader(fh, delimiter=";"), 2):
                if not _es_mutual(r.get("Mutual"), mutual) or not (r.get("Numero") or "").strip():
                    continue
                v = (r.get("Id") or "").strip()
                if not (v.upper() == "EXCLUIR" or C.norm_id(v)):
                    raise FrenarProceso(f"Correcciones_Manuales.csv fila {n}: el Id '{v}' no es un número ni EXCLUIR")
                out[(C.norm_dni(r.get("DNI")) or "*", C.norm_num(r["Numero"]))] = (
                    "EXCLUIR" if v.upper() == "EXCLUIR" else C.norm_id(v), f"Corrección manual: {r.get('Motivo', '')}")
    return out


def escribir_reporte(path, hojas):
    """Reporte en Excel: hojas = [(nombre, encabezados, filas)]. Encabezado azul 2E5395, montos $ y fechas dd/mm/aaaa."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nombre, hdr, filas in hojas:
        C.hoja_tabla(wb, nombre, hdr, filas)
    C.autofit(wb)
    try:
        wb.save(path)
    except PermissionError:
        raise FrenarProceso(f"{Path(path).name} está abierto en Excel: cerralo y volvé a correr")


def agregar_hoja(path, nombre, hdr, filas):
    """Agrega una hoja al final de un Excel ya escrito (ej. 'No imputado' cuando quedan casos sin resolver)."""
    wb = openpyxl.load_workbook(path)
    C.hoja_tabla(wb, nombre, hdr, filas)
    ws = wb[nombre[:31]]
    for row in ws.iter_rows():
        for c in row:
            c.font = openpyxl.styles.Font(name="Arial", size=10, bold=c.row == 1, color="FFFFFF" if c.row == 1 else None)
            if c.row == 1:
                c.fill = B.FILL  # mismo azul que el resto del VtoTeórico
    for j, h in enumerate(hdr, 1):
        largo = max([len(str(h))] + [len(f"{f[j - 1]:,.2f}" if isinstance(f[j - 1], float) else str(f[j - 1] or "")) for f in filas])
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = min(max(largo + 2, 8), 70)
    wb.save(path)


def asignar_lp(lp, cobrado_por_id, modo, mes=None):
    """Pone lo cobrado de cada crédito en sus filas del LP (de la cuota más vieja a la más nueva).
    modo 'primera': todo en la primera cuota pendiente, las demás vacías (Entidad 8, ENTIDAD1).
    modo 'cascada': hasta el importe de cada cuota y el resto a la siguiente (Entidad 5).
    Devuelve (valores {línea: monto}, sobrante {Id: monto que no entró en el LP})."""
    import datetime as dt
    por_id = collections.defaultdict(list)
    for f in lp.filas:
        if f["Id"] is not None:
            por_id[f["Id"]].append(f)
    valores, sobrante = {}, {}
    for idc, total in cobrado_por_id.items():
        total = C.r2(total)
        if total <= 0:
            continue
        cuotas = sorted(por_id.get(idc, []), key=lambda f: (f["Fecha"] or dt.datetime.max, C.norm_id(f["Cuota"]) or 0, f["linea"]))
        if not cuotas:
            sobrante[idc] = total
            continue
        if modo == "primera":
            valores[cuotas[0]["linea"]] = total
            continue
        if modo == "del_mes":  # la cuota que vence en el mes de la imputación; si no hay, la más vieja (Entidad 8)
            del_mes = [f for f in cuotas if f["Fecha"] is not None and (f["Fecha"].year, f["Fecha"].month) == (mes.anio, mes.nmes)]
            valores[(del_mes or cuotas)[0]["linea"]] = total
            continue
        resto = total
        for f in cuotas:
            if resto <= 0:
                break
            a = C.r2(min(resto, f["Importe"]))
            if a > 0:
                valores[f["linea"]] = a
                resto = C.r2(resto - a)
        if 0 < resto < 1:  # ruido de redondeo: a la última fila del crédito
            valores[cuotas[-1]["linea"]] = C.r2(valores.get(cuotas[-1]["linea"], 0) + resto)
            resto = 0.0
        if resto > 0:
            sobrante[idc] = resto
    if modo == "cascada" and valores:
        # LP "crudo" que en realidad ya fue imputado (lo cobrado escrito encima de importecobrado, Entidad 5 oct-26):
        # entonces lo imputado coincide con el importe de TODAS las filas. Un LP crudo siempre trae cuotas sin cobrar.
        con_importe = [f for f in lp.filas if f["Id"] is not None and (f["Importe"] or 0) > 0]
        if con_importe and all(abs(valores.get(f["linea"], 0) - f["Importe"]) < 0.005 for f in con_importe):
            print(f"  AVISO: {lp.path.name} parece un LP YA IMPUTADO (lo cobrado coincide con importecobrado en todas "
                  f"las filas). Si es así, bajá el LP crudo de nuevo de Plenario: con este, los topes de cada cuota "
                  f"salen de tu propia imputación.")
    return valores, sobrante


def leer_texto_o_excel(path):
    """Filas de una rendición que puede venir como texto con tabulaciones (aunque diga .XLS) o como Excel
    con el texto pegado en pocas celdas. Devuelve listas de campos (str, sin espacios de relleno)."""
    raw = Path(path).read_bytes()
    if raw[:2] == b"PK" or raw[:4] == b"\xd0\xcf\x11\xe0":  # Excel de verdad
        filas = []
        if raw[:2] == b"PK":
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            it = (r for ws in wb.worksheets for r in ws.iter_rows(values_only=True))
        else:
            import xlrd
            bk = xlrd.open_workbook(str(path))
            it = (sh.row_values(i) for sh in bk.sheets() for i in range(sh.nrows))
        for r in it:
            partes = [str(c) for c in r if c not in (None, "")]
            filas.append([c.strip() for c in "\t".join(partes).split("\t")])
        return filas
    try:
        txt = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        txt = raw.decode("cp1252", errors="replace")
    return [[c.strip() for c in l.split("\t")] for l in txt.splitlines() if l.strip()]


def sumar_por_id(lp, valores):
    s = collections.defaultdict(float)
    for f in lp.filas:
        if f["linea"] in valores and f["Id"] is not None:
            s[f["Id"]] += valores[f["linea"]]
    return {k: C.r2(v) for k, v in s.items()}

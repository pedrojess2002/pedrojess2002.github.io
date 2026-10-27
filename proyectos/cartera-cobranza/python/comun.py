# -*- coding: utf-8 -*-
"""Funciones comunes a todos los adaptadores: lectura de archivos, normalización,
formato Excel y controles. Nada de acá conoce reglas de una mutual en particular."""
import csv
import datetime as dt
import re
import unicodedata
from pathlib import Path

import openpyxl
import xlrd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
TOL = 0.005  # tolerancia para "cierra al centavo"

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


class FrenarProceso(Exception):
    """Algo no cierra o falta un archivo: se frena y se informa, no se entrega nada."""


# ---------------------------------------------------------------- normalización
def sin_acentos(s):
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in s if not unicodedata.combining(c)).lower().strip()


def norm_dni(x):
    """'DNI - 00000000' / 'DNI-00000000' / 'DNI 00000000' / 00000000 -> '00000000'."""
    if x is None:
        return ""
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    d = re.sub(r"\D", "", str(x))
    return d.lstrip("0")


def norm_num(x):
    """Código de Referencia / número de crédito: sin espacios ni ceros a la izquierda."""
    if x is None:
        return ""
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    s = str(x).strip()
    m = re.search(r"#\s*0*(\d+)", s)
    if m:
        return m.group(1)
    s = re.sub(r"\D", "", s)
    return s.lstrip("0")


def norm_id(x):
    if x is None or x == "":
        return None
    try:
        return int(float(x))
    except (TypeError, ValueError):
        return None


def num(x):
    if x is None or x == "":
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("$", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def r2(x):
    return round(x + 0.0, 2)


def cierra(a, b):
    return abs(a - b) < TOL


def fecha_xls(v, datemode=0):
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, (int, float)):
        return xlrd.xldate.xldate_as_datetime(v, datemode)
    for f in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(str(v).strip()[:10], f)
        except ValueError:
            pass
    return None


def categoria_mora(dias):
    """Misma regla que el armado del VtoTeórico: <30 Al día, 30-60, >60."""
    d = dias or 0
    if d < 30:
        return "Al día"
    if d <= 60:
        return "Mora 30 a 60 días"
    return "Mora mayor a 60 días"


# ---------------------------------------------------------------- lectura
def leer_tabla_xlsx(path, hoja, encabezado_contiene):
    """Devuelve (encabezados, filas) buscando la fila de encabezado por contenido."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = next((w for w in wb.worksheets if w.title.strip().upper() == hoja.strip().upper()), None)
    if ws is None:
        raise FrenarProceso(f"{Path(path).name}: no tiene la hoja '{hoja}'")
    filas = list(ws.iter_rows(values_only=True))
    wb.close()
    for i, f in enumerate(filas[:15]):
        celdas = [str(c).strip() if c is not None else "" for c in f]
        if all(any(sin_acentos(c) == sin_acentos(e) for c in celdas) for e in encabezado_contiene):
            hdr = celdas
            return hdr, filas[i + 1:]
    raise FrenarProceso(f"{Path(path).name} / {hoja}: no encuentro el encabezado {encabezado_contiene}")


def col(hdr, *nombres, obligatoria=True):
    for n in nombres:
        for i, h in enumerate(hdr):
            if sin_acentos(h) == sin_acentos(n):
                return i
    if obligatoria:
        raise FrenarProceso(f"Falta la columna {nombres[0]} (encabezados: {hdr})")
    return None


def leer_plenario(path):
    """Extracto completo de Plenario (.xls crudo, nunca re-guardado en Excel).
    Devuelve lista de dicts con las columnas normalizadas que usan los adaptadores."""
    bk = xlrd.open_workbook(str(path))
    sh = bk.sheet_by_index(0)
    hdr = [str(x).strip() for x in sh.row_values(0)]
    c = {n: col(hdr, n) for n in ["Id", "Codigo Referencia", "Estado", "Línea", "Cliente",
                                   "Tipo y Nº Doc.", "Valor Cuota", "F. Operación", "Capital",
                                   "Plazo", "Entidad", "F. 1º Vto.", "F. Próximo Vto.",
                                   "F. Último Vto.", "Días Mora"]}
    c["Marca"] = col(hdr, "Marca", obligatoria=False)  # algunos exports viejos no la traen
    out = []
    for r in range(1, sh.nrows):
        v = sh.row_values(r) + [""]
        idc = norm_id(v[c["Id"]])
        if idc is None:
            continue
        out.append({
            "Id": idc,
            "CodRef": str(v[c["Codigo Referencia"]]).strip(),
            "Estado": str(v[c["Estado"]]).strip(),
            "Linea": str(v[c["Línea"]]).strip(),
            "Cliente": str(v[c["Cliente"]]).strip(),
            "Doc": str(v[c["Tipo y Nº Doc."]]).strip(),
            "DNI": norm_dni(v[c["Tipo y Nº Doc."]]),
            "ValorCuota": num(v[c["Valor Cuota"]]),
            "FOperacion": fecha_xls(v[c["F. Operación"]], bk.datemode),
            "Capital": num(v[c["Capital"]]),
            "Plazo": norm_id(v[c["Plazo"]]),
            "Entidad": str(v[c["Entidad"]]).strip(),
            "F1Vto": fecha_xls(v[c["F. 1º Vto."]], bk.datemode),
            "FProxVto": fecha_xls(v[c["F. Próximo Vto."]], bk.datemode),
            "FUltVto": fecha_xls(v[c["F. Último Vto."]], bk.datemode),
            "DiasMora": norm_id(v[c["Días Mora"]]) or 0,
            "Marca": str(v[c["Marca"]]).strip() if c["Marca"] is not None else None,  # None = el export no trae Marca
        })
    return out


def leer_csv_lp(path):
    """LP crudo de Plenario. Devuelve (texto_original, lineas, separador, fin_de_linea, encoding)."""
    raw = Path(path).read_bytes()
    enc = "utf-8-sig" if raw.startswith(b"\xef\xbb\xbf") else "utf-8"
    try:
        txt = raw.decode(enc)
    except UnicodeDecodeError:
        enc = "cp1252"
        txt = raw.decode(enc)
    eol = "\r\n" if "\r\n" in txt else "\n"
    sep = ";" if txt.split(eol, 1)[0].count(";") >= txt.split(eol, 1)[0].count(",") else ","
    lineas = txt.split(eol)
    return txt, lineas, sep, eol, enc


# ---------------------------------------------------------------- detección de archivos
def archivos(carpeta, *ext):
    p = Path(carpeta)
    if not p.exists():
        return []
    return sorted(f for f in p.iterdir() if f.is_file() and f.suffix.lower() in ext
                  and not f.name.startswith("~$"))


def uno_solo(candidatos, que):
    if len(candidatos) == 1:
        return candidatos[0]
    if not candidatos:
        raise FrenarProceso(f"No encontré {que}")
    raise FrenarProceso(f"Hay {len(candidatos)} archivos que parecen {que}: "
                        + ", ".join(c.name for c in candidatos) + ". Dejá uno solo.")


# ---------------------------------------------------------------- formato Excel
AZUL = "2E5395"
_borde = Side(style="thin", color="BFBFBF")
BORDE = Border(left=_borde, right=_borde, top=_borde, bottom=_borde)
FMT_PESOS = "$#,##0.00"
FMT_FECHA = "DD/MM/YYYY"
FMT_PCT = "0.0%"
RESALTADO = PatternFill("solid", fgColor="FFF2CC")


def encabezado(ws, fila, valores):
    for j, v in enumerate(valores, 1):
        c = ws.cell(fila, j, v)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.border = BORDE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def escribir_fila(ws, fila, valores, formatos=None, relleno=None, negrita=False):
    for j, v in enumerate(valores, 1):
        c = ws.cell(fila, j, v)
        c.border = BORDE
        if formatos and formatos.get(j - 1):
            c.number_format = formatos[j - 1]
        elif isinstance(v, dt.datetime):
            c.number_format = FMT_FECHA
        if relleno:
            c.fill = relleno
        if negrita:
            c.font = Font(bold=True)


def formatos_por_nombre(hdr):
    """Asigna formato por nombre de columna (montos, fechas, %)."""
    f = {}
    for i, h in enumerate(hdr):
        n = sin_acentos(h or "")
        if n.startswith("%") or "eficiencia" in n or "cumplimiento" in n or n in ("performance", "part."):
            f[i] = FMT_PCT
        elif n.startswith("f.") or "fecha" in n:
            f[i] = FMT_FECHA
        elif any(k in n for k in ("cuota", "cobrado", "monto", "importe", "diferencia", "capital",
                                  "saldo", "vto teorico", "al dia", "mora 30", "mora mayor",
                                  "esperado", "liquidado", "imputado", "total", "rendido")) \
                and "cant" not in n and "plazo" not in n and n != "cuota (liquidacion)":
            f[i] = FMT_PESOS
    return f


def hoja_tabla(wb, nombre, hdr, filas, resaltar=None):
    """Hoja simple: encabezado azul + filas con formato por nombre de columna."""
    ws = wb.create_sheet(nombre[:31])
    encabezado(ws, 1, hdr)
    fm = formatos_por_nombre(hdr)
    for i, f in enumerate(filas, 2):
        escribir_fila(ws, i, f, fm, RESALTADO if resaltar and resaltar(f) else None)
    ws.freeze_panes = "A2"
    return ws


def autofit(wb):
    for ws in wb.worksheets:
        anchos = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if isinstance(c.value, dt.datetime):
                    largo = 10
                elif isinstance(c.value, float):
                    largo = len(f"{c.value:,.2f}") + 2
                else:
                    largo = len(str(c.value))
                anchos[c.column] = max(anchos.get(c.column, 0), largo)
        for k, v in anchos.items():
            ws.column_dimensions[get_column_letter(k)].width = min(max(v + 2, 8), 60)


def nombre_mes(aaaamm):
    a, m = aaaamm.split("-")
    return f"{MESES[int(m) - 1]}{a}"

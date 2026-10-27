# -*- coding: utf-8 -*-
"""
Cargador CarteraDW: lee los archivos de entrada\\, detecta que son
por su CONTENIDO (no por el nombre) y los carga a las tablas stg.

Despues de cargar stg, si existe el procedimiento de el analista
(dw.sp_cargar_plenario / dw.sp_cargar_operaciones) lo ejecuta para pasar los
datos al modelo dw. Si todavia no existe, deja la carga en estado STAGING para
que el analista corra el EXEC a mano.

Uso:  py cargar.py            (o doble click en Cargar.bat)
      py cargar.py --db CarteraDW_test
"""
import argparse
import datetime as dt
import hashlib
import math
import re
import shutil
import sys
import unicodedata
from pathlib import Path

import pandas as pd
import pyodbc

SERVIDOR = r"(localdb)\MSSQLLocalDB"
BASE = "CarteraDW"
RAIZ = Path(__file__).resolve().parent

# Un export "completo" no puede tener menos de este % de filas que la ultima
# carga OK del mismo tipo (protege contra exports filtrados a una mutual).
MIN_PCT_VS_ANTERIOR = 0.90
# Si no hay carga anterior, minimo absoluto de filas para aceptar el archivo.
MIN_FILAS_PRIMERA_CARGA = 10000


# ---------------------------------------------------------------- utilidades

def normalizar(texto):
    """'F. 1º Vto.' -> 'f1vto'  (sin acentos, minusculas, solo letras/numeros)."""
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", t.lower())


def hash_archivo(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def fecha_corte_desde_nombre(ruta):
    """Busca DD-MM-AAAA, DD-MM-AA o DD-MM en el nombre del archivo."""
    m = re.search(r"(\d{1,2})-(\d{1,2})(?:-(\d{2,5}))?", ruta.stem)
    if not m:
        return None
    dia, mes, anio = int(m.group(1)), int(m.group(2)), m.group(3)
    if anio is None:
        anio = dt.datetime.fromtimestamp(ruta.stat().st_mtime).year
    elif len(anio) == 2:
        anio = 2000 + int(anio)
    else:
        anio = int(anio[-4:])  # tolera typos tipo '22026'
    try:
        return dt.date(anio, mes, dia)
    except ValueError:
        return None


# ------------------------------------------------- definicion de cada formato
# clave normalizada del encabezado original -> (columna stg, tipo)
# tipos: int, dec2, dec4, dec6, str, date, bit

COLS_OPERACIONES = {
    "idop": ("Id_Op", "int"), "cuil": ("CUIL", "str"), "tipodoc": ("Tipo_Doc", "str"),
    "nrodoc": ("Nro_Doc", "int"), "apellido": ("Apellido", "str"), "nombre": ("Nombre", "str"),
    "domicilio": ("Domicilio", "str"), "provincia": ("Provincia", "str"),
    "localidad": ("Localidad", "str"), "fliq": ("F_Liq", "date"),
    "f1ervto": ("F_1er_Vto", "date"), "fultvto": ("F_Ult_Vto", "date"),
    "tem": ("TEM", "dec6"), "capsolicitado": ("Cap_Solicitado", "dec2"),
    "capotorgado": ("Cap_Otorgado", "dec2"), "monto": ("Monto", "dec2"),
    "cantcuotas": ("Cant_Cuotas", "int"), "valcuota": ("Val_Cuota", "dec2"),
    "cantcuotasimpvencidas": ("Cant_Cuotas_Imp_Vencidas", "int"),
    "impcuotasimpvencidas": ("Imp_Cuotas_Imp_Vencidas", "dec2"),
    "capcuotasimpvencidas": ("Cap_Cuotas_Imp_Vencidas", "dec2"),
    "intcuotasimpvencidas": ("Int_Cuotas_Imp_Vencidas", "dec2"),
    "ivacuotasimpvencidas": ("Iva_Cuotas_Imp_Vencidas", "dec2"),
    "cantcuotasimpavencer": ("Cant_Cuotas_Imp_a_Vencer", "int"),
    "impcuotasimpavencer": ("Imp_Cuotas_Imp_a_Vencer", "dec2"),
    "capcuotasimpavencer": ("Cap_Cuotas_Imp_a_Vencer", "dec2"),
    "intcuotasimpavencer": ("Int_Cuotas_Imp_a_Vencer", "dec2"),
    "ivacuotasimpavencer": ("Iva_Cuotas_Imp_a_Vencer", "dec2"),
    "diasdemora": ("Dias_de_Mora", "int"), "entidad": ("Entidad", "str"),
    "importeultimopago": ("Importe_Ultimo_Pago", "dec2"),
    "fultimos3pagos": ("F_Ultimos_3_Pagos", "str"),
    "tramodemora": ("Tramo_de_Mora", "str"), "tna": ("TNA", "dec4"),
    "marcacedido": ("Marca_Cedido", "str"),
    "cantidadcuotaspagas": ("Cantidad_Cuotas_Pagas", "int"),
    "capitalpagado": ("Capital_Pagado", "dec2"), "interespagado": ("Interes_Pagado", "dec2"),
    "ivapagado": ("IVA_Pagado", "dec2"),
}

COLS_PLENARIO = {
    "id": ("Id", "int"), "idsolicitud": ("Id_Solicitud", "int"),
    "codigoreferencia": ("Codigo_Referencia", "str"),
    "idcredterceros": ("Id_Cred_Terceros", "int"), "estado": ("Estado", "str"),
    "origen": ("Origen", "str"), "linea": ("Linea", "str"),
    "idcliente": ("Id_Cliente", "int"), "cliente": ("Cliente", "str"),
    "tipoynodoc": ("Tipo_Nro_Doc", "str"), "cuitl": ("CUIL", "str"),
    "valorcuota": ("Valor_Cuota", "dec2"), "foperacion": ("F_Operacion", "date"),
    "capital": ("Capital", "dec2"), "tasa": ("Tasa_Pct", "dec4"), "iva": ("IVA_Pct", "dec2"),
    "plazo": ("Plazo", "int"),
    "fcancelacionanticipada": ("F_Cancelacion_Anticipada", "date"),
    "entidadidnombre": ("Entidad_Id", "entidad_id"), "vendido": ("Vendido", "str"),
    "f1ovto": ("F_1er_Vto", "date"), "fproximovto": ("F_Proximo_Vto", "date"),
    "fultimovto": ("F_Ultimo_Vto", "date"),
    "fultimopagoparcial": ("F_Ult_Pago_Parcial", "date"),
    "fultimopagototal": ("F_Ult_Pago_Total", "date"), "diasmora": ("Dias_Mora", "int"),
    "cuotascobradas": ("Cuotas_Cobradas", "int"),
    "capitaladeudado": ("Capital_Adeudado", "dec2"), "totalcobrado": ("Total_Cobrado", "dec2"),
    "nolegajo": ("Nro_Legajo", "str"), "capitalcobrado": ("Capital_Cobrado", "dec2"),
    "capitaladeudadovencido": ("Capital_Adeudado_Vencido", "dec2"),
    "f1ovtoimpago": ("F_1er_Vto_Impago", "date"),
}

TIPOS = {
    "OPERACIONES": {"cols": COLS_OPERACIONES, "stg": "stg.operaciones",
                    "sp": "dw.sp_cargar_operaciones", "clave_mutual": "Entidad"},
    "PLENARIO": {"cols": COLS_PLENARIO, "stg": "stg.plenario",
                 "sp": "dw.sp_cargar_plenario", "clave_mutual": "Entidad_Id"},
}
ORDEN_TIPOS = ["PLENARIO", "OPERACIONES"]  # Plenario primero: trae el Id Cliente


def detectar_tipo(columnas):
    norm = {normalizar(c) for c in columnas}
    if {"idop", "diasdemora", "tramodemora"} <= norm:
        return "OPERACIONES"
    if {"id", "idcliente", "entidadidnombre", "estado"} <= norm:
        return "PLENARIO"
    return None


# --------------------------------------------------------- conversion valores

def convertir(valor, tipo):
    if valor is None or (isinstance(valor, float) and math.isnan(valor)) or valor is pd.NaT:
        return None
    if isinstance(valor, str):
        valor = valor.strip()
        if valor == "":
            return None
    if tipo == "str":
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))  # 100378.0 -> '100378'
        return str(valor)
    if tipo == "int":
        return int(float(valor))
    if tipo == "entidad_id":  # '7 - Asoc. Mutual de ...' -> 7
        m = re.match(r"\s*(\d+)\s*-", str(valor))
        return int(m.group(1)) if m else None
    if tipo == "date":
        if isinstance(valor, (dt.datetime, pd.Timestamp)):
            return valor.date()
        if isinstance(valor, dt.date):
            return valor
        return pd.to_datetime(valor, dayfirst=True).date()
    decimales = {"dec2": 2, "dec4": 4, "dec6": 6}[tipo]
    return round(float(valor), decimales)


def preparar_filas(df, cols_def):
    mapa = {}
    for col in df.columns:
        clave = normalizar(col)
        if clave in cols_def:
            mapa[col] = cols_def[clave]
    faltan = {v[0] for v in cols_def.values()} - {v[0] for v in mapa.values()}
    if faltan:
        raise ValueError("Faltan columnas en el archivo: " + ", ".join(sorted(faltan)))
    origen = list(mapa.keys())
    destino = [mapa[c][0] for c in origen]
    tipos = [mapa[c][1] for c in origen]
    filas = []
    for registro in df[origen].itertuples(index=False, name=None):
        filas.append(tuple(convertir(v, t) for v, t in zip(registro, tipos)))
    return destino, filas


# ------------------------------------------------------------ base de datos

def conectar(base):
    return pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={SERVIDOR};DATABASE={base};Trusted_Connection=yes;",
        autocommit=False,
    )


def existe_sp(cur, nombre):
    return cur.execute("SELECT OBJECT_ID(?, 'P')", nombre).fetchval() is not None


def validar_completo(cur, tipo, filas, columnas):
    """Rechaza exports filtrados (una sola mutual, una fecha, bajas, etc.)."""
    idx = columnas.index(TIPOS[tipo]["clave_mutual"])
    presentes = {f[idx] for f in filas if f[idx] is not None}
    if tipo == "PLENARIO":
        activas = {r[0] for r in cur.execute("SELECT mutual_id FROM dw.dim_mutual WHERE activa = 1")}
    else:
        activas = {r[0] for r in cur.execute("SELECT nombre_plenario FROM dw.dim_mutual WHERE activa = 1")}
    faltan = activas - presentes
    if faltan:
        raise ValueError(
            f"Parece un export FILTRADO: le faltan {len(faltan)} de las {len(activas)} mutuales activas. "
            "Para la foto diaria hace falta el export completo de todas las mutuales.")
    anterior = cur.execute(
        "SELECT TOP 1 filas_leidas FROM ctl.carga WHERE tipo = ? AND estado IN ('OK','STAGING') "
        "ORDER BY fecha_corte DESC", tipo).fetchval()
    minimo = anterior * MIN_PCT_VS_ANTERIOR if anterior else MIN_FILAS_PRIMERA_CARGA
    if len(filas) < minimo:
        raise ValueError(
            f"Tiene {len(filas)} filas y se esperaban al menos {int(minimo)} "
            f"(ultima carga de {tipo}: {anterior or 'ninguna'}). Parece un export parcial.")


def mover(ruta, carpeta, motivo=None):
    destino_dir = RAIZ / carpeta
    destino_dir.mkdir(exist_ok=True)
    destino = destino_dir / ruta.name
    if destino.exists():
        destino = destino_dir / f"{ruta.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{ruta.suffix}"
    shutil.move(str(ruta), str(destino))
    if motivo:
        destino.with_suffix(destino.suffix + ".MOTIVO.txt").write_text(motivo, encoding="utf-8")


# ---------------------------------------------------------------- principal

def procesar(base):
    entrada = RAIZ / "entrada"
    entrada.mkdir(exist_ok=True)
    archivos = sorted(p for p in entrada.iterdir() if p.suffix.lower() in (".xls", ".xlsx"))
    otros = [p for p in entrada.iterdir() if p.is_file() and p not in archivos]
    for p in otros:
        print(f"[IGNORADO] {p.name}: todavia no se cargan archivos de este tipo, queda en entrada.")
    if not archivos:
        print("No hay archivos .xls/.xlsx en", entrada)
        return

    con = conectar(base)
    cur = con.cursor()
    cur.fast_executemany = True

    # 1) Leer y clasificar todo antes de cargar nada
    pendientes = []
    for ruta in archivos:
        print(f"\nLeyendo {ruta.name} ...")
        try:
            df = pd.read_excel(ruta)
            tipo = detectar_tipo(df.columns)
            if tipo is None:
                raise ValueError("No reconozco el formato (no es Operaciones a fecha ni Data plenario).")
            fecha = fecha_corte_desde_nombre(ruta)
            if fecha is None:
                raise ValueError("No encuentro la fecha en el nombre. Renombralo con DD-MM-AAAA "
                                 "(ej: 'Operaciones a fecha 30-09-2026.xls').")
            if fecha > dt.date.today():
                raise ValueError(f"La fecha del nombre ({fecha:%d/%m/%Y}) es futura.")
            pendientes.append((ORDEN_TIPOS.index(tipo), fecha, ruta, tipo, df))
            print(f"  -> {tipo}, fecha de corte {fecha:%d/%m/%Y}, {len(df)} filas")
        except Exception as e:
            print(f"  [ERROR] {e}")
            mover(ruta, "error", str(e))

    # 2) Cargar: Plenario antes que Operaciones, y por fecha
    stg_ocupado = set()
    for _, fecha, ruta, tipo, df in sorted(pendientes, key=lambda x: (x[1], x[0])):
        conf = TIPOS[tipo]
        print(f"\nCargando {ruta.name} ({tipo} {fecha:%d/%m/%Y})")
        if tipo in stg_ocupado:
            print(f"  [ESPERA] {conf['stg']} tiene una carga sin pasar a dw. "
                  f"Corre el EXEC pendiente y volve a ejecutar Cargar.bat.")
            continue
        carga_id = None
        try:
            h = hash_archivo(ruta)
            previa = cur.execute("SELECT carga_id, estado FROM ctl.carga WHERE hash_archivo = ?", h).fetchone()
            if previa and previa.estado != "ERROR":
                print(f"  [DUPLICADO] Este archivo ya se cargo (carga_id {previa.carga_id}, {previa.estado}).")
                mover(ruta, "procesados", f"Duplicado de carga_id {previa.carga_id}")
                continue
            if previa:  # reintento de una carga que fallo
                cur.execute("DELETE FROM ctl.carga WHERE carga_id = ?", previa.carga_id)

            columnas, filas = preparar_filas(df, conf["cols"])
            validar_completo(cur, tipo, filas, columnas)

            carga_id = cur.execute(
                "INSERT INTO ctl.carga (archivo, hash_archivo, tipo, fecha_corte, filas_leidas, estado) "
                "OUTPUT INSERTED.carga_id VALUES (?, ?, ?, ?, ?, 'STAGING')",
                ruta.name, h, tipo, fecha, len(filas)).fetchval()
            cur.execute(f"TRUNCATE TABLE {conf['stg']}")
            marcas = ", ".join("?" * len(columnas))
            cur.executemany(
                f"INSERT INTO {conf['stg']} ({', '.join(columnas)}) VALUES ({marcas})", filas)
            con.commit()
            print(f"  stg OK: {len(filas)} filas en {conf['stg']} (carga_id {carga_id})")

            if existe_sp(cur, conf["sp"]):
                cur.execute(f"EXEC {conf['sp']} @carga_id = ?", carga_id)
                while cur.nextset():
                    pass
                con.commit()
                fila = cur.execute("SELECT estado, filas_cargadas, mensaje FROM ctl.carga WHERE carga_id = ?",
                                   carga_id).fetchone()
                print(f"  dw: {fila.estado} - {fila.filas_cargadas} filas. {fila.mensaje or ''}")
            else:
                stg_ocupado.add(tipo)
                print(f"  [PENDIENTE] Todavia no existe {conf['sp']}.")
                print(f"  Pasalo a dw a mano en SSMS (Guia 2) y despues volve a correr Cargar.bat.")
            mover(ruta, "procesados")
        except Exception as e:
            con.rollback()
            msg = str(e)
            print(f"  [ERROR] {msg}")
            if carga_id:
                cur.execute("UPDATE ctl.carga SET estado = 'ERROR', mensaje = ? WHERE carga_id = ?",
                            msg[:1000], carga_id)
                con.commit()
            mover(ruta, "error", msg)
    con.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default=BASE)
    ap.add_argument("--carpeta", default=None, help="carpeta con entrada/procesados/error (para pruebas)")
    args = ap.parse_args()
    if args.carpeta:
        RAIZ = Path(args.carpeta)
    procesar(args.db)
    print("\nListo.")

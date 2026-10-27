# -*- coding: utf-8 -*-
"""Pruebas de "romper" el generador de VtoTeórico con exports armados a propósito con fallas.
Cada caso dice qué tiene que pasar (frenar con mensaje claro, o seguir y avisar en el control) y se verifica.
Uso: py tests_vto_teorico.py      (escribe solo en una carpeta temporal, no toca 01_Entrada)
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
import pandas as pd

AQUI = Path(__file__).resolve().parent
sys.stdout.reconfigure(encoding="utf-8")
BASE = Path(r"<ruta> plenario 02-10-2026.xls")
TMP = Path(tempfile.gettempdir()) / "tests_vto_teorico"
shutil.rmtree(TMP, ignore_errors=True)
TMP.mkdir()
df0 = pd.read_excel(BASE)


def correr(nombre, df=None, archivo=None, mes="2026-10"):
    f = archivo or TMP / f"{nombre} 02-10-2026.xlsx"  # con fecha en el nombre, como los exports reales
    if df is not None:
        df.to_excel(f, index=False)
    out = TMP / f"out_{nombre}"
    r = subprocess.run([sys.executable, "-I", str(AQUI / "generar_vto_teorico.py"), mes, str(f), "--out", str(out)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return r.returncode, r.stdout + r.stderr, out


def total_detalle(out):
    """Fila TOTAL del Detalle del GENERAL: tiene que ser =SUM de toda la columna Valor Cuota."""
    f = next(out.glob("VtoTeorico_*_GENERAL.xlsx"))
    ws = openpyxl.load_workbook(f, read_only=True)["Detalle"]
    rows = list(ws.iter_rows(values_only=True))
    return rows[-1], len(rows)


def creditos(salida):
    for line in salida.splitlines():
        if line.startswith("  Créditos en el VtoTeórico:"):
            return int(line.split(":")[1])
    return None


resultados = []


def caso(nombre, esperado, ok, detalle=""):
    resultados.append((nombre, esperado, "PASA" if ok else "FALLA", detalle))
    print(f"[{'PASA ' if ok else 'FALLA'}] {nombre} — {esperado}" + (f" | {detalle}" if detalle and not ok else ""))


# 0. export normal (convertido a xlsx): tiene que dar lo mismo que el original
rc, s, out = correr("t0_normal", df0.copy())
caso("Export normal", "5802 créditos, Todo OK, sin archivo de control", rc == 0 and creditos(s) == 5802
     and "Todo OK" in s and not list(out.glob("Control_*")), s[-400:])
tot, nfil = total_detalle(out)
caso("Fila TOTAL en el Detalle", "TOTAL con =SUM(L2:L5803) en Valor Cuota",
     tot[2] == "TOTAL" and tot[0] is None and tot[11] == "=SUM(L2:L5803)" and nfil == 5804, str(tot))

# 1. Id duplicado -> frenar
d = pd.concat([df0, df0.iloc[[100]]], ignore_index=True)
rc, s, _ = correr("t1_id_duplicado", d)
caso("Id duplicado", "frena y dice qué Id", rc == 2 and "duplicados" in s, s[-300:])

# 2. Entidad de Entidad 5 con otro nombre que sigue diciendo 'Entidad 5' -> la reconoce
d = df0.copy()
d.loc[d["Entidad"].str.contains("Entidad 5", na=False), "Entidad"] = "Asociacion Mutual Entidad 5 (nuevo nombre)"
rc, s, _ = correr("t2_entidad5_renombrada", d)
caso("Entidad renombrada (con palabra clave)", "la reconoce, mismos 5802", rc == 0 and creditos(s) == 5802, s[-300:])

# 3. Entidad de Entidad 5 con un nombre irreconocible -> frenar (antes la dejaba afuera sin avisar)
d = df0.copy()
d.loc[d["Entidad"].str.contains("Entidad 5", na=False), "Entidad"] = "AMEPPC"
rc, s, _ = correr("t3_entidad5_irreconocible", d)
caso("Entidad irreconocible", "frena: falta Entidad 5", rc == 2 and "Entidad 5" in s, s[-300:])

# 4. export filtrado (solo ENTIDAD1) -> frenar
rc, s, _ = correr("t4_filtrado", df0[df0["Entidad"] == "ENTIDAD1"].copy())
caso("Export filtrado (solo ENTIDAD1)", "frena: faltan mutuales", rc == 2 and "no trae ninguna fila" in s, s[-300:])

# 5. fechas de vencimiento vacías en 3 créditos activos -> sigue, pero avisa
d = df0.copy()
idx = d[(d["Estado"] != "Cancelado") & d["Entidad"].str.contains("Entidad 5", na=False)].index[:3]
d.loc[idx, "F. 1º Vto."] = pd.NaT
rc, s, out = correr("t5_fecha_vacia", d)
caso("Fecha de vencimiento vacía", "avisa los 3 en pantalla", rc == 0 and s.count("fecha de vencimiento vacía") == 3, s[-400:])

# 6. Días Mora vacío -> avisa
d = df0.copy()
idx = d[(d["Estado"] != "Cancelado") & (d["Entidad"] == "ENTIDAD1")].index[:2]
d.loc[idx, "Días Mora"] = None
rc, s, out = correr("t6_dias_mora_vacio", d)
caso("Días Mora vacío", "avisa los 2 en pantalla", rc == 0 and s.count("Días Mora vacío") == 2, s[-400:])

# 7. fila de marca de agua arriba del encabezado (export DevExpress) -> funciona igual
f7 = TMP / "t7_marca_agua 02-10-2026.xlsx"
with pd.ExcelWriter(f7) as w:
    pd.DataFrame([["Exportado desde Plenario"]]).to_excel(w, index=False, header=False, startrow=0)
    df0.to_excel(w, index=False, startrow=1)
rc, s, _ = correr("t7_marca_agua", archivo=f7)
caso("Marca de agua arriba del encabezado", "funciona igual, 5802", rc == 0 and creditos(s) == 5802, s[-300:])

# 8. archivo de texto con extensión .xls (como los de Entidad 8) -> frenar con mensaje
f8 = TMP / "t8_texto.xls"
f8.write_text("socAPELLIDO\tsocNOMBRE\nPEREZ\tJUAN\n", encoding="utf-8")
rc, s, _ = correr("t8_texto", archivo=f8)
caso("Archivo que no es Excel", "frena: no se pudo leer", rc == 2 and "No se pudo leer" in s, s[-300:])

# 9. falta una columna -> frenar
rc, s, _ = correr("t9_sin_valor_cuota", df0.drop(columns=["Valor Cuota"]))
caso("Falta la columna Valor Cuota", "frena: faltan columnas", rc == 2 and "faltan las columnas" in s, s[-300:])

# 10. mes mal escrito -> frenar
rc, s, _ = correr("t10_mes_malo", df0.head(5), mes="2026-13")
caso("Mes mal escrito (2026-13)", "frena: formato AAAA-MM", rc == 2 and "AAAA-MM" in s, s[-300:])

# 11. fechas como texto dd/mm/aaaa -> las entiende
d = df0.copy()
for c in ("F. 1º Vto.", "F. Último Vto."):
    d[c] = pd.to_datetime(d[c]).dt.strftime("%d/%m/%Y")
rc, s, _ = correr("t11_fechas_texto", d)
caso("Fechas como texto dd/mm/aaaa", "las entiende, 5802", rc == 0 and creditos(s) == 5802, s[-300:])

# 12. armar un mes PASADO con un export posterior -> avisa
rc, s, out = correr("t12_export_posterior", df0.copy(), mes="2026-08")
caso("Mes pasado con export posterior", "avisa en pantalla", rc == 0 and "POSTERIOR" in s, s[-400:])

# 13. el archivo no existe -> frenar
rc, s, _ = correr("t13_no_existe", archivo=TMP / "no_existe.xls")
caso("El archivo no existe", "frena: no existe", rc == 2 and "no existe" in s, s[-300:])

# 14. el mes se congela con el 1er Plenario del mes: octubre con el 02-10; volver a correrlo con el 19-10 usa el 02-10
out14 = TMP / "out_t14_congelada"
r1 = subprocess.run([sys.executable, "-I", str(AQUI / "generar_vto_teorico.py"), "2026-10", str(BASE), "--out", str(out14)],
                    capture_output=True, text=True, encoding="utf-8", errors="replace")
u1 = pd.read_pickle(out14 / "_universo.pkl")
r2 = subprocess.run([sys.executable, "-I", str(AQUI / "generar_vto_teorico.py"), "2026-10",
                     r"<ruta> plenario 19-10-2026.xls", "--out", str(out14)],
                    capture_output=True, text=True, encoding="utf-8", errors="replace")
u2 = pd.read_pickle(out14 / "_universo.pkl")
caso("Mes congelado con el 1er Plenario del mes", "la 2ª corrida (export 19-10) da igual que la del 02-10",
     r1.returncode == 0 and r2.returncode == 0 and "ya está armado con el Plenario" in r2.stdout
     and u1[["Id", "Valor Cuota", "Categoria"]].equals(u2[["Id", "Valor Cuota", "Categoria"]]) and len(u2) == 5802,
     f"1ª {len(u1)} / 2ª {len(u2)}")

# 15. export anterior al mes -> BORRADOR, no congela
rc, s, out = correr("t15_borrador", df0.copy(), mes="2026-11")
caso("Export anterior al mes", "sale BORRADOR y no congela la mora vieja",
     rc == 0 and "BORRADOR" in s and not (out / "_mora_vieja_congelada.csv").exists(), s[-400:])

# 16. las 5 métricas del Resumen cierran entre los 8 archivos y el GENERAL, ENTIDAD9 afuera (si algo falla no dice Todo OK)
caso("Suma de los 8 = GENERAL en las 5 métricas", "Todo OK", r2.returncode == 0 and "Todo OK" in r2.stdout,
     r2.stdout[-400:])

print(f"\n{sum(1 for r in resultados if r[2] == 'PASA')} de {len(resultados)} pruebas pasan")

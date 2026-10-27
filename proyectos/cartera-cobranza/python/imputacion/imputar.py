# -*- coding: utf-8 -*-
"""Imputación del mes.  Uso:  py imputar.py AAAA-MM [mutual ...]
Mutuales: entidad4, entidad7, Entidad 8, entidad6, entidad3, entidad5, entidad1, entidad2   (sin mutual: todas las que estén listas)
Ej.:  py imputar.py 2026-10 entidad4
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.stdout.reconfigure(encoding="utf-8")

import base  # noqa: E402
import sin_rendicion  # noqa: E402
import entidad8  # noqa: E402
import entidad6  # noqa: E402
import entidad3  # noqa: E402
import entidad5  # noqa: E402
import entidad1  # noqa: E402
import entidad2  # noqa: E402

PROCESOS = {"entidad4": lambda m: sin_rendicion.procesar(m, "entidad4"),
            "entidad7": lambda m: sin_rendicion.procesar(m, "entidad7"),
            "Entidad 8": entidad8.procesar,
            "entidad6": entidad6.procesar,
            "entidad3": entidad3.procesar,
            "entidad5": entidad5.procesar,
            "entidad1": entidad1.procesar,
            "entidad2": entidad2.procesar}


def main():
    if "--out" in sys.argv:  # solo para pruebas: escribir en otra carpeta y no mirar Cobranza
        k = sys.argv.index("--out")
        base.SALIDA_PRUEBAS = Path(sys.argv[k + 1])
        del sys.argv[k:k + 2]
    if "--raiz" in sys.argv:  # solo para pruebas: leer 01_Entrada / 04_Reglas / 07_Base_Maestra de una copia
        k = sys.argv.index("--raiz")
        base.C.RAIZ = Path(sys.argv[k + 1])
        base.MAESTRA = base.C.RAIZ / "07_Base_Maestra"
        del sys.argv[k:k + 2]
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    try:
        mes = base.Mes(sys.argv[1].strip())
    except base.FrenarProceso as e:
        print(f"PROCESO FRENADO: {e}")
        return 2
    pedidas = [m.strip().lower() for m in sys.argv[2:]] or list(PROCESOS)
    resultado = {}
    for m in pedidas:
        if m not in PROCESOS:
            print(f"No conozco la mutual '{m}'. Opciones: {', '.join(PROCESOS)}")
            resultado[m] = "ERROR"
            continue
        print()
        try:
            resultado[m] = PROCESOS[m](mes)
        except base.FrenarProceso as e:
            print(f"PROCESO FRENADO: {e}")
            resultado[m] = "FRENADO"
        except PermissionError as e:
            print(f"PROCESO FRENADO: {Path(e.filename).name if e.filename else 'un archivo'} está abierto en Excel. "
                  f"Cerralo y volvé a correr.")
            resultado[m] = "FRENADO"
        except Exception as e:  # algo que el sistema no esperaba: mostrarlo claro (y el detalle, para revisarlo)
            import traceback
            print(f"ERROR INESPERADO en {m}: {type(e).__name__}: {e}\n"
                  f"Revisá los archivos de la carpeta del mes; si no ves el problema, pasale este mensaje a soporte:")
            traceback.print_exc(file=sys.stdout)
            resultado[m] = "ERROR"
    if len(resultado) > 1:
        print("\n" + " | ".join(f"{k}: {v}" for k, v in resultado.items()))
    return 0 if all(v == "OK" for v in resultado.values()) else 2


if __name__ == "__main__":
    sys.exit(main())

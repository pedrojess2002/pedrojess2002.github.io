"""Vencimiento Teórico mensual — 8 mutuales + GENERAL.
Uso: python3 build_vto_teorico.py <export_plenario.xls> --mes N --anio AAAA --out ./salida

Copia del script del proyecto de claude.ai "Vencimientos Teóricos" (build_vto_teorico_oficial.py, sin tocar).
Único cambio, pedido por el analista (19/10/2026): la columna Cobrado sale VACÍA en las 8 mutuales (el oficial ponía
=Valor Cuota salvo en ENTIDAD3). Se completa después, en la imputación, solo con lo que la mutual rindió.
Validado: con data plenario 02-10-2026 reproduce el VtoTeorico_Octubre2026_GENERAL de las 8 mutuales
(créditos y montos exactos; la única diferencia es ENTIDAD4 6038, agregado a mano en el GENERAL).
CAMBIO 2: la Entidad se reconoce por palabras clave y frena si una mutual sale vacía (ver mapear_entidad).
CAMBIO 4: formato Arial 10 (encabezado blanco sobre azul 1F4E78) en los 9 archivos; el oficial solo en ENTIDAD3.
CAMBIO 3: la mora vieja de ENTIDAD3 se puede congelar al 1º del mes (parámetro mora_vieja_fija).
Los cancelados siguen quedando afuera siempre (Estado != Cancelado), igual que el oficial.
"""
import argparse, calendar, os, re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

MESES = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
         'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

ENTITY_MAP = {
    'ENTIDAD1': 'ENTIDAD1',
    'ENTIDAD 7': 'ASOC Entidad 7',
    'ENTIDAD 5': 'Entidad 5',
    'ENTIDAD 8': 'Entidad 8',
    'ENTIDAD 3': 'ENTIDAD3',
    'ENTIDAD 4': 'Entidad 4',
    'ENTIDAD 6': 'ENTIDAD6',
    'ENTIDAD 2': 'ENTIDAD2',
}
def mapear_entidad(e):
    """Entidad de Plenario -> nombre de la mutual, por palabras clave (tolera typos y cambios de nombre)."""
    import unicodedata
    if e in ENTITY_MAP:
        return ENTITY_MAP[e]
    t = ''.join(c for c in unicodedata.normalize('NFKD', str(e)) if not unicodedata.combining(c)).lower().strip()
    for claves, nombre in [(('entidad1',), 'ENTIDAD1'), (('Entidad 7',), 'ASOC Entidad 7'), (('entidad5',), 'Entidad 5'),
                           (('munici',), 'Entidad 8'), (('Entidad 3',), 'ENTIDAD3'),
                           (('justicia',), 'Entidad 4'), (('docentes',), 'ENTIDAD6'), (('ENTIDAD 2',), 'ENTIDAD2')]:
        if any((k == t) if k == 'entidad1' else (k in t) for k in claves):
            return nombre
    return None  # ENTIDAD 9 y cualquier otra entidad: fuera


ORDER = ['ENTIDAD1', 'ENTIDAD3', 'ASOC Entidad 7', 'Entidad 8',
         'Entidad 5', 'ENTIDAD2', 'ENTIDAD6', 'Entidad 4']
DIA_VTO = {'ENTIDAD1': 10, 'ASOC Entidad 7': 10, 'Entidad 5': 15, 'Entidad 8': 15,
           'ENTIDAD3': 10, 'Entidad 4': 10, 'ENTIDAD6': 25, 'ENTIDAD2': 15}
FNAME = {'ENTIDAD1': 'ENTIDAD1', 'ENTIDAD3': 'ENTIDAD3', 'ASOC Entidad 7': 'ENTIDAD7',
         'Entidad 8': 'Entidad 8', 'Entidad 5': 'Entidad 5_Cordoba',
         'ENTIDAD2': 'ENTIDAD2', 'ENTIDAD6': 'ENTIDAD6', 'Entidad 4': 'Entidad 4'}
PV = re.compile(r'\(PV\)|PAGO VOLUNTARIO|SOLUTIONS|voluntari', re.IGNORECASE)

MONEY = '"$ "#,##0.00'
DATE = 'dd/mm/yyyy'
PCT = '0.0%'
FILL = PatternFill(start_color='FF1F4E78', end_color='FF1F4E78', fill_type='solid')
CENTER = Alignment(horizontal='center', vertical='center')
STYLES = {
    'standard': (Font(name='Cambria', bold=True, color='FFFFFFFF'), Font(name='Calibri', size=11)),
    'arial': (Font(name='Arial', size=10, bold=True, color='FFFFFFFF'), Font(name='Arial', size=10)),
}
CAT_SHEETS = [('Al_Dia', 'Al día'), ('Mora_30_a_60', 'Mora 30 a 60 días'),
              ('Mora_mayor_60', 'Mora mayor a 60 días')]
SRC = ['Id', 'Codigo Referencia', 'Cliente', 'Tipo y Nº Doc.', 'Línea', 'Estado', 'F. 1º Vto.',
       'F. Último Vto.', 'F. Operación', 'Días Mora', 'Valor Cuota', 'Capital', 'Plazo', 'Categoria']


def categoria(dm):
    if dm < 30:
        return 'Al día'
    if dm <= 60:
        return 'Mora 30 a 60 días'
    return 'Mora mayor a 60 días'


def build_universe(raw, mes, anio, mora_vieja_fija=None):
    """mora_vieja_fija: Ids de la mora vieja ENTIDAD3 congelada al 1º del mes (CAMBIO 3, pedido de el analista 19/10/2026).
    Si viene, la mora vieja del mes es esa lista (y no depende de los días de mora del día del export)."""
    ini = pd.Timestamp(anio, mes, 1)
    fin = pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1])
    try:
        df = pd.read_excel(raw, header=0)
        if 'Id' not in df.columns:  # export con fila de marca de agua DevExpress
            df = pd.read_excel(raw, header=1)
    except Exception as e:
        raise ValueError(f'No se pudo leer {os.path.basename(str(raw))} como Excel de Plenario ({type(e).__name__}: {e})')
    faltan_cols = [c for c in SRC[:-1] + ['Entidad'] if c not in df.columns]
    if faltan_cols:
        raise ValueError(f'El archivo no es un export completo de Plenario: faltan las columnas {faltan_cols}')
    if not df['Id'].is_unique:
        dup = df.loc[df['Id'].duplicated(keep=False), 'Id'].unique().tolist()[:10]
        raise ValueError(f'Ids duplicados en el export (ej. {dup})')
    # CAMBIO 2 (19/10/2026): la entidad se reconoce por palabras clave, no por el texto exacto. El oficial dejaba
    # afuera a Entidad 8 sin avisar con exports viejos ("Municiapales", con el typo que Plenario corrigió en sep-26).
    df['EntidadMap'] = df['Entidad'].map(mapear_entidad)
    faltan = [e for e in ORDER if not (df['EntidadMap'] == e).any()]
    if faltan:
        raise ValueError(f'El export no trae ninguna fila de: {faltan}. ¿Es un export filtrado, o cambió el nombre '
                         f'de la entidad en Plenario? Entidades que trae: {sorted(df["Entidad"].astype(str).unique())}')
    for c in ('F. 1º Vto.', 'F. Último Vto.', 'F. Operación'):
        df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=True)  # por si una fecha viene como texto
    pv = df['Línea'].astype(str).str.contains(PV) | df['Cliente'].astype(str).str.contains(PV)
    exceptuado = df['Cliente'].astype(str).str.strip().str.upper() == 'CLIENTE EXCEPTUADO'
    base = df['EntidadMap'].notna() & (df['Estado'] != 'Cancelado') & ~(pv & ~exceptuado)
    window = (df['F. 1º Vto.'] <= fin) & (df['F. Último Vto.'] >= ini)
    if mora_vieja_fija is None:
        mora_vieja = (df['EntidadMap'] == 'ENTIDAD3') & (df['Días Mora'] > 0) & ~window
    else:
        mora_vieja = (df['EntidadMap'] == 'ENTIDAD3') & df['Id'].isin(mora_vieja_fija) & ~window
    u = df[base & (window | mora_vieja)].copy()
    u['Categoria'] = u['Días Mora'].apply(categoria)
    u['MoraVieja'] = mora_vieja[u.index]
    return df, u.sort_values('Id').reset_index(drop=True), ini, fin


def header(ws, cols, hf):
    for j, h in enumerate(cols, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font, c.fill, c.alignment = hf, FILL, CENTER
    ws.freeze_panes = 'A2'
    ws.row_dimensions[1].height = 30


def write_detalle(ws, rows, hf, df_, general=False, cobrado=None, agregados=None, pie=None):
    """general=True inserta Entidad tras Tipo y Nº Doc. Cobrado vacío (VtoTeórico del mes) o, en la imputación,
    cobrado = {Id: monto rendido} y agregados = créditos que cobraron sin estar en el VtoTeórico (van abajo, aparte)."""
    cols = ['Id', 'Codigo Referencia', 'Cliente', 'Tipo y Nº Doc.'] + (['Entidad'] if general else []) + \
           ['Línea', 'Estado', 'F. 1º Vto.', 'F. Último Vto.', 'F. Operación', 'Días Mora', 'Valor Cuota',
            'Cobrado', 'Diferencia', '% Cobrado', 'Capital', 'Plazo', 'Categoria']
    header(ws, cols, hf)
    o = 1 if general else 0
    K, L = ('L', 'M') if general else ('K', 'L')
    bold = Font(name=df_.name, size=df_.size, bold=True)

    def bloque(filas, i0):
        for i, r in enumerate(filas[SRC + ['EntidadMap']].itertuples(index=False, name=None), i0):
            vals = list(r[:4]) + ([r[14]] if general else []) + list(r[4:11])
            for j, v in enumerate(vals, 1):
                ws.cell(row=i, column=j, value=None if pd.isna(v) else (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v))
            for j in (7 + o, 8 + o, 9 + o):
                ws.cell(row=i, column=j).number_format = DATE
            ws.cell(row=i, column=11 + o).number_format = MONEY
            # Cobrado vacío en el VtoTeórico del mes; en la imputación, lo que rindió la mutual (0 si no rindió nada)
            cob = None if cobrado is None else round(float(cobrado.get(int(r[0]), 0.0)), 2)
            ws.cell(row=i, column=12 + o, value=cob).number_format = MONEY
            ws.cell(row=i, column=13 + o, value=f'={K}{i}-{L}{i}').number_format = MONEY
            ws.cell(row=i, column=14 + o, value=f'=IFERROR({L}{i}/{K}{i},"")').number_format = PCT
            ws.cell(row=i, column=15 + o, value=r[11]).number_format = MONEY
            ws.cell(row=i, column=16 + o, value=r[12])
            ws.cell(row=i, column=17 + o, value=r[13])
            for j in range(1, len(cols) + 1):
                ws.cell(row=i, column=j).font = df_

    def total(t, desde, hasta, texto):
        ws.cell(row=t, column=3, value=texto).font = bold
        for j in (11 + o, 12 + o, 13 + o):
            L_ = ws.cell(row=1, column=j).column_letter
            c = ws.cell(row=t, column=j, value=f'=SUM({L_}{desde}:{L_}{hasta})')
            c.font, c.number_format = bold, MONEY

    bloque(rows, 2)
    # CAMBIO 5: fila TOTAL abajo (pedido de el analista). Id vacío y "TOTAL" en Cliente, así los scripts que leen el Detalle la saltean
    total(len(rows) + 2, 2, max(len(rows) + 1, 2), 'TOTAL')
    if agregados is not None and len(agregados):
        # cobros no esperados (la mutual pagó un crédito que no estaba en el VtoTeórico del mes): debajo del TOTAL,
        # dejando una fila en blanco, con todas sus columnas y sin título ni subtotal (como lo arma el analista, oct-26).
        # El Resumen los suma al Cobrado (no al Vto Teórico), así Cobrado del VtoTeórico = columna cobrado del LP.
        bloque(agregados, len(rows) + 4)
    if pie:  # renglones finales de la imputación (ej. comisión de la mutual), en la columna de Valor Cuota / Cobrado
        p0 = ws.max_row + 3
        for k, (texto, valor) in enumerate(pie):
            ws.cell(row=p0 + k, column=11 + o, value=texto).font = bold
            c = ws.cell(row=p0 + k, column=12 + o, value=valor)
            c.font, c.number_format = bold, MONEY
    w = [10, 14, 28, 16] + ([22] if general else []) + [26, 10, 12, 12, 12, 10, 14, 14, 13, 10, 13, 8, 18]
    for j, wd in enumerate(w, 1):
        ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = wd


def entity_file(ent, rows, path, cobrado=None, agregados=None, pie=None):
    hf, df_ = STYLES['arial']  # CAMBIO 4: Arial 10 en las 9 (pedido de el analista)
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Resumen')
    write_detalle(wb.create_sheet('Detalle'), rows, hf, df_, cobrado=cobrado, agregados=agregados, pie=pie)
    for sn, cat in CAT_SHEETS:
        write_detalle(wb.create_sheet(sn), rows[rows['Categoria'] == cat], hf, df_, cobrado=cobrado)
    last = max(len(rows) + 1, 2)
    header(ws, ['Entidad', 'Fecha de Vto (día)', 'Cant. Créditos', 'Vto Teórico Total', 'Al día',
                'Mora 30 a 60', 'Mora mayor a 60', 'Saldo a Cobrar (Con Mora)', 'Performance',
                'Esperado', 'Cobrado', 'Eficiencia', 'Saldo'], hf)
    f = [ent, DIA_VTO[ent], f'=COUNTA(Detalle!A2:A{last})', f'=SUM(Detalle!K2:K{last})',
         f'=SUMIF(Detalle!Q2:Q{last},"Al día",Detalle!K2:K{last})',
         f'=SUMIF(Detalle!Q2:Q{last},"Mora 30 a 60 días",Detalle!K2:K{last})',
         f'=SUMIF(Detalle!Q2:Q{last},"Mora mayor a 60 días",Detalle!K2:K{last})',
         '=D2-E2', None, '=IF(I2="","",I2*D2)',
         f'=SUM(Detalle!L2:L{last})' + (f'+SUM(Detalle!L{len(rows) + 4}:L{len(rows) + 3 + len(agregados)})'
                                         if agregados is not None and len(agregados) else ''),
         '=IFERROR(K2/D2,"")', '=IFERROR(K2-D2,"")']
    for j, v in enumerate(f, 1):
        c = ws.cell(row=2, column=j, value=v)
        c.font = df_
        if j in (4, 5, 6, 7, 8, 10, 11, 13):
            c.number_format = MONEY
        if j == 12:
            c.number_format = PCT
    for col, wd in {'A': 22, 'B': 16, 'C': 14, 'D': 18, 'E': 16, 'F': 14, 'G': 16, 'H': 20, 'I': 12,
                    'J': 16, 'K': 16, 'L': 12, 'M': 16}.items():
        ws.column_dimensions[col].width = wd
    wb.save(path)


def general_file(u, path):
    hf, df_ = STYLES['arial']  # CAMBIO 4: Arial 10 en las 9
    bold = Font(name='Arial', size=10, bold=True)
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Resumen')
    rows = u.assign(_o=u['EntidadMap'].map(ORDER.index)).sort_values(['_o', 'Id'])
    write_detalle(wb.create_sheet('Detalle'), rows, hf, df_, general=True)
    for sn, cat in CAT_SHEETS:
        write_detalle(wb.create_sheet(sn), rows[rows['Categoria'] == cat], hf, df_, general=True)
    header(ws, ['Entidad', 'Fecha de Vto (día)', 'Cant. Créditos', 'Vto Teórico Total', 'Al día',
                'Mora 30 a 60', 'Mora mayor a 60', 'Saldo a Cobrar (Con Mora)', 'Part.', 'Performance',
                'Esperado', 'Cobrado', 'Eficiencia', 'Saldo'], hf)
    tot = len(ORDER) + 2
    for r, ent in enumerate(ORDER, 2):
        f = [ent, DIA_VTO[ent], f'=COUNTIF(Detalle!E:E,"{ent}")', f'=SUMIF(Detalle!E:E,"{ent}",Detalle!L:L)']
        f += [f'=SUMIFS(Detalle!L:L,Detalle!E:E,"{ent}",Detalle!R:R,"{c}")' for _, c in CAT_SHEETS]
        f += [f'=D{r}-E{r}', f'=IFERROR(D{r}/D${tot},0)', None, f'=IF(J{r}="","",J{r}*D{r})', None,
              f'=IF(L{r}="","",L{r}/D{r})', f'=IF(L{r}="","",L{r}-D{r})']
        for j, v in enumerate(f, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = df_
            if j in (4, 5, 6, 7, 8, 11, 12, 14):
                c.number_format = MONEY
            if j in (9, 13):
                c.number_format = PCT
    ws.cell(row=tot, column=1, value='Total').font = bold
    for j, L in enumerate('CDEFGH', 3):
        c = ws.cell(row=tot, column=j, value=f'=IF(COUNT({L}2:{L}{tot-1})=0,"",SUM({L}2:{L}{tot-1}))')
        c.font = bold
        if j >= 4:
            c.number_format = MONEY
    c = ws.cell(row=tot, column=9, value=1)
    c.number_format, c.font = PCT, bold
    for j in (11, 12, 14):
        ws.cell(row=tot, column=j).number_format = MONEY
    ws.cell(row=tot, column=13).number_format = PCT
    for col, wd in {'A': 22, 'B': 16, 'C': 14, 'D': 18, 'E': 16, 'F': 14, 'G': 16, 'H': 20, 'I': 10,
                    'J': 12, 'K': 16, 'L': 16, 'M': 12, 'N': 16}.items():
        ws.column_dimensions[col].width = wd
    wb.save(path)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('raw')
    ap.add_argument('--mes', type=int, required=True)
    ap.add_argument('--anio', type=int, required=True)
    ap.add_argument('--out', default='./salida')
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    df, u, ini, fin = build_universe(a.raw, a.mes, a.anio)
    tag = f'{MESES[a.mes]}{a.anio}'
    print(f'Período {ini:%d/%m/%Y} - {fin:%d/%m/%Y} | universo {len(u)} | ${u["Valor Cuota"].sum():,.2f}')
    for ent in ORDER:
        rows = u[u['EntidadMap'] == ent]
        entity_file(ent, rows, os.path.join(a.out, f'VtoTeorico_{tag}_{FNAME[ent]}.xlsx'))
        print(f'{ent:24s} {len(rows):5d}  ${rows["Valor Cuota"].sum():,.2f}')
    general_file(u, os.path.join(a.out, f'VtoTeorico_{tag}_GENERAL.xlsx'))
    u.to_pickle(os.path.join(a.out, '_universo.pkl'))

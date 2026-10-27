-- Tabla "Operaciones" del tablero, leída de CarteraDW (solo vistas rpt, no modifica nada).
-- Última foto de cada mes, sin Cancelados ni ENTIDAD9 (la vista ya filtra mutuales activas).
-- {N} = parámetro "Meses sin pagar" (lo reemplaza Power BI).
-- Identificacion viene vacía: Power BI la completa con la Marca de Plenario (03_Modelo\Marcas.csv).
-- No usar comentarios al final de una línea: el generador junta la consulta en una sola línea.
WITH fotos AS (
    SELECT MAX(fecha_corte) AS fecha_corte
    FROM rpt.fact_cartera
    GROUP BY YEAR(fecha_corte), MONTH(fecha_corte)
), base AS (
    SELECT f.*, c.titular, c.nro_doc, c.f_operacion, c.f_1er_vto, c.plazo, c.valor_cuota, m.mutual_id, m.nombre_plenario
    FROM rpt.fact_cartera f
    JOIN fotos u ON u.fecha_corte = f.fecha_corte
    JOIN rpt.dim_credito c ON c.credito_id = f.credito_id
    JOIN rpt.dim_mutual m ON m.mutual_id = c.mutual_id
    WHERE f.estado <> 'Cancelado'
)
SELECT
    DATEFROMPARTS(YEAR(fecha_corte), MONTH(fecha_corte), 1) AS [Mes],
    fecha_corte AS [Fecha de Corte],
    CASE mutual_id WHEN 1 THEN 'ENTIDAD1' WHEN 4 THEN 'ENTIDAD2' WHEN 5 THEN 'Entidad 3' WHEN 6 THEN 'Entidad 4'
         WHEN 7 THEN 'Entidad 5' WHEN 8 THEN 'ENTIDAD6' WHEN 9 THEN 'Entidad 7'
         WHEN 10 THEN 'Entidad 8' END AS [Mutual],
    credito_id AS [Id_Op],
    nombre_plenario AS [Entidad],
    titular AS [Cliente],
    CAST(nro_doc AS varchar(20)) AS [Nro_Doc],
    estado AS [Estado],
    f_operacion AS [F_Liq],
    f_1er_vto AS [F_1er_Vto],
    plazo AS [Cant_Cuotas],
    valor_cuota AS [Val_Cuota],
    ISNULL(dias_mora, 0) AS [Dias_de_Mora],
    ISNULL(cuotas_vencidas_impagas, 0) AS [Cant_Cuotas_Imp_Vencidas],
    ISNULL(saldo_vencido, 0) AS [Imp_Cuotas_Imp_Vencidas],
    plazo - cuotas_vencidas_total AS [Cant_Cuotas_Imp_a Vencer],
    ISNULL(saldo_a_vencer, 0) AS [Imp_Cuotas_Imp_a Vencer],
    saldo_total AS [Imp_Cuotas_Imp_V_a V],
    ISNULL(cuotas_pagas, 0) AS [Cantidad_Cuotas_Pagas],
    ISNULL(capital_pagado, 0) AS [Capital_Pagado],
    ISNULL(interes_pagado, 0) AS [Interes_Pagado],
    ISNULL(iva_pagado, 0) AS [IVA_Pagado],
    total_pagado AS [Pagado Total],
    f_ultimo_pago AS [Ultimo Pago],
    CASE WHEN dias_mora < 30 THEN NULL WHEN dias_mora <= 60 THEN '30 a 60' WHEN dias_mora <= 90 THEN '61 a 90'
         WHEN dias_mora <= 180 THEN '91 a 180' ELSE 'Más de 180' END AS [Tramo],
    CASE WHEN dias_mora < 30 THEN NULL WHEN dias_mora <= 60 THEN 1 WHEN dias_mora <= 90 THEN 2
         WHEN dias_mora <= 180 THEN 3 ELSE 4 END AS [Orden Tramo],
    CASE WHEN dias_mora <= 0 THEN NULL WHEN dias_mora <= 30 THEN '1 a 30' WHEN dias_mora <= 60 THEN '31 a 60'
         WHEN dias_mora <= 90 THEN '61 a 90' ELSE '91 o más' END AS [Tramo Mora],
    CASE WHEN total_pagado = 0 AND f_ultimo_pago IS NULL THEN 'Nunca pagó'
         WHEN f_ultimo_pago IS NULL THEN 'Pagó sin fecha de último pago'
         WHEN f_ultimo_pago > EOMONTH(fecha_corte) THEN 'Pago con fecha futura'
         WHEN f_ultimo_pago >= DATEFROMPARTS(YEAR(fecha_corte), MONTH(fecha_corte), 1) THEN 'Viene pagando'
         WHEN f_ultimo_pago < DATEADD(MONTH, -{N}, fecha_corte) THEN 'Dejó de pagar'
         ELSE 'No pagó este mes' END AS [Comportamiento],
    ISNULL(saldo_vencido, 0) AS [Monto Vencido],
    total_vencido AS [Total Vencido],
    cobrado_de_lo_vencido AS [Cobrado de lo Vencido],
    CASE WHEN total_pagado = 0 THEN 'No pago nada' ELSE 'Pago algo (parcial o total)' END AS [Situacion de pago],
    CASE WHEN mutual_id IN (5, 6) AND estado = 'Mora' THEN saldo_total ELSE ISNULL(saldo_vencido, 0) END AS [Monto a Reclamar],
    tramo_mora AS [Tramo Cartera],
    orden_tramo AS [Orden Tramo Cartera],
    CAST(NULL AS varchar(50)) AS [Identificacion],
    'CarteraDW rpt.fact_cartera' AS [Fuente]
FROM base

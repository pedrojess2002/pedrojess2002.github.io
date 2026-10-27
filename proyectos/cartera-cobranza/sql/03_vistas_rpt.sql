CREATE VIEW rpt.dim_mutual
AS
SELECT mutual_id, sigla, nombre_plenario, dia_vencimiento
FROM dw.dim_mutual
WHERE activa = 1;
GO

CREATE VIEW rpt.dim_credito
AS
SELECT c.credito_id,
       c.mutual_id,
       cl.cuil,
       cl.nro_doc,
       cl.apellido,
       cl.nombre,
       CONCAT(cl.apellido, ', ', cl.nombre)                        AS titular,
       c.linea,
       c.f_operacion,
       DATEFROMPARTS(YEAR(c.f_operacion), MONTH(c.f_operacion), 1) AS cosecha,
       c.f_1er_vto,
       c.f_ult_vto,
       c.capital,
       c.plazo,
       c.valor_cuota
FROM dw.dim_credito c
JOIN dw.dim_cliente cl ON cl.cliente_id = c.cliente_id
JOIN dw.dim_mutual  m  ON m.mutual_id   = c.mutual_id
WHERE m.activa = 1;
GO

CREATE   VIEW rpt.fact_cartera
AS
SELECT p.fecha_corte,
       p.credito_id,
       p.estado,
       ISNULL(o.dias_mora, p.dias_mora)                     AS dias_mora,
       p.dias_mora                                          AS dias_mora_plenario,
       CASE WHEN p.estado = 'Cancelado' THEN 'Cancelado'
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 0   THEN 'Al día'
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 30  THEN '1-30'
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 60  THEN '31-60'
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 90  THEN '61-90'
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 180 THEN '91-180'
            ELSE '>180'
       END AS tramo_mora,
       CASE WHEN p.estado = 'Cancelado' THEN 7
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 0   THEN 1
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 30  THEN 2
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 60  THEN 3
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 90  THEN 4
            WHEN ISNULL(o.dias_mora, p.dias_mora) <= 180 THEN 5
            ELSE 6
       END AS orden_tramo,
       o.cuotas_pagas,
       o.cuotas_vencidas_impagas,
       c.plazo - ISNULL(o.cuotas_a_vencer, 0)               AS cuotas_vencidas_total,
       (c.plazo - ISNULL(o.cuotas_a_vencer, 0)) * c.valor_cuota AS total_vencido,
       (c.plazo - ISNULL(o.cuotas_a_vencer, 0)) * c.valor_cuota - ISNULL(o.imp_vencido, 0) AS cobrado_de_lo_vencido,
       o.imp_vencido                                        AS saldo_vencido,
       o.imp_a_vencer                                       AS saldo_a_vencer,
       ISNULL(o.imp_vencido, 0) + ISNULL(o.imp_a_vencer, 0) AS saldo_total,
       o.cap_vencido                                        AS capital_vencido,
       o.int_vencido                                        AS interes_vencido,
       ISNULL(o.cap_vencido, 0) + ISNULL(o.cap_a_vencer, 0) AS capital_pendiente,
       ISNULL(o.int_vencido, 0) + ISNULL(o.int_a_vencer, 0) AS interes_pendiente,
       p.capital_adeudado,
       p.capital_adeudado_vencido,
       o.capital_pagado,
       o.interes_pagado,
       o.iva_pagado,
       ISNULL(o.capital_pagado, 0) + ISNULL(o.interes_pagado, 0) + ISNULL(o.iva_pagado, 0) AS total_pagado,
       o.f_pago_1                                           AS f_ultimo_pago,
       o.f_pago_2                                           AS f_penultimo_pago,
       o.f_pago_3                                           AS f_antepenultimo_pago,
       CASE WHEN p.dias_mora <= 0 AND o.dias_mora > 0 THEN 1 ELSE 0 END AS alerta_pago_fecha_futura,
       CASE WHEN p.fecha_corte = (SELECT MAX(fecha_corte) FROM dw.fact_foto_plenario)
            THEN 1 ELSE 0 END                               AS es_ultima_foto
FROM dw.fact_foto_plenario p
LEFT JOIN dw.fact_foto_operaciones o
       ON o.credito_id  = p.credito_id
      AND o.fecha_corte = p.fecha_corte
JOIN dw.dim_credito c ON c.credito_id = p.credito_id
JOIN dw.dim_mutual  m ON m.mutual_id  = c.mutual_id
WHERE m.activa = 1;
GO

CREATE   PROCEDURE dw.sp_cargar_plenario @carga_id INT
AS
BEGIN
    SET NOCOUNT ON;     -- no mostrar "(N filas afectadas)" por cada instrucción
    SET XACT_ABORT ON;  -- ante cualquier error, abortar la transacción entera

    DECLARE @fecha DATE = (SELECT fecha_corte FROM ctl.carga
                           WHERE carga_id = @carga_id AND tipo = 'PLENARIO');
    IF @fecha IS NULL
        THROW 50001, 'El carga_id no existe o no es de tipo PLENARIO.', 1;

    BEGIN TRY
        BEGIN TRAN;

        -- 0) Control: ninguna mutual desconocida
        IF EXISTS (SELECT 1 FROM stg.plenario s
                   WHERE NOT EXISTS (SELECT 1 FROM dw.dim_mutual m WHERE m.mutual_id = s.Entidad_Id))
            THROW 50002, 'Hay créditos de una mutual que no existe en dw.dim_mutual.', 1;

        -- 1) Clientes
        WITH origen AS (
            SELECT Id_Cliente, CUIL,
                   CASE WHEN Tipo_Nro_Doc LIKE '[0-9]%' THEN NULL
                        ELSE LEFT(Tipo_Nro_Doc, CHARINDEX(' ', Tipo_Nro_Doc + ' ') - 1) END AS tipo_doc,
                   TRY_CAST(CASE WHEN Tipo_Nro_Doc LIKE '[0-9]%' THEN Tipo_Nro_Doc
                                 ELSE SUBSTRING(Tipo_Nro_Doc, CHARINDEX(' ', Tipo_Nro_Doc + ' ') + 1, 20) END
                            AS BIGINT) AS nro_doc,
                   LTRIM(RTRIM(LEFT(Cliente, CHARINDEX(',', Cliente + ',') - 1))) AS apellido,
                   NULLIF(LTRIM(RTRIM(SUBSTRING(Cliente, CHARINDEX(',', Cliente + ',') + 1, 80))), '') AS nombre,
                   ROW_NUMBER() OVER (PARTITION BY Id_Cliente ORDER BY F_Operacion DESC) AS rn
            FROM stg.plenario
        )
        MERGE dw.dim_cliente AS d
        USING (SELECT * FROM origen WHERE rn = 1) AS s
           ON d.cliente_id = s.Id_Cliente
        WHEN MATCHED AND (d.cuil <> s.CUIL OR ISNULL(d.nro_doc, 0) <> ISNULL(s.nro_doc, 0)) THEN
            UPDATE SET cuil = s.CUIL, tipo_doc = s.tipo_doc, nro_doc = s.nro_doc,
                       fecha_modif = SYSDATETIME()
        WHEN NOT MATCHED THEN
            INSERT (cliente_id, cuil, tipo_doc, nro_doc, apellido, nombre)
            VALUES (s.Id_Cliente, s.CUIL, s.tipo_doc, s.nro_doc, s.apellido, s.nombre);

        -- 2) Créditos
        MERGE dw.dim_credito AS d
        USING stg.plenario AS s
           ON d.credito_id = s.Id
        WHEN MATCHED AND (   ISNULL(d.linea, '') <> ISNULL(s.Linea, '')
                          OR ISNULL(d.codigo_referencia, '') <> ISNULL(s.Codigo_Referencia, '')
                          OR ISNULL(d.id_cred_terceros, 0) <> ISNULL(s.Id_Cred_Terceros, 0)
                          OR ISNULL(d.f_ult_vto, '19000101') <> ISNULL(s.F_Ultimo_Vto, '19000101')) THEN
            UPDATE SET linea = s.Linea, codigo_referencia = s.Codigo_Referencia,
                       id_cred_terceros = s.Id_Cred_Terceros, f_ult_vto = s.F_Ultimo_Vto,
                       fecha_modif = SYSDATETIME()
        WHEN NOT MATCHED THEN
            INSERT (credito_id, cliente_id, mutual_id, id_solicitud, codigo_referencia, id_cred_terceros,
                    nro_legajo, linea, origen, f_operacion, f_1er_vto, f_ult_vto, capital, plazo,
                    valor_cuota, tasa_pct, iva_pct)
            VALUES (s.Id, s.Id_Cliente, s.Entidad_Id, s.Id_Solicitud, s.Codigo_Referencia, s.Id_Cred_Terceros,
                    s.Nro_Legajo, s.Linea, s.Origen, s.F_Operacion, s.F_1er_Vto, s.F_Ultimo_Vto, s.Capital,
                    s.Plazo, s.Valor_Cuota, s.Tasa_Pct, s.IVA_Pct);

        -- 3) Foto del día
        INSERT INTO dw.fact_foto_plenario
            (fecha_corte, credito_id, estado, dias_mora, cuotas_cobradas, capital_adeudado,
             capital_adeudado_vencido, capital_cobrado, total_cobrado, f_proximo_vto, f_1er_vto_impago,
             f_ult_pago_parcial, f_ult_pago_total, f_cancelacion_anticipada, vendido, carga_id)
        SELECT @fecha, Id, Estado, Dias_Mora, Cuotas_Cobradas, Capital_Adeudado,
               Capital_Adeudado_Vencido, Capital_Cobrado, Total_Cobrado, F_Proximo_Vto, F_1er_Vto_Impago,
               F_Ult_Pago_Parcial, F_Ult_Pago_Total, F_Cancelacion_Anticipada,
               CASE Vendido WHEN 'Checked' THEN 1 ELSE 0 END, @carga_id
        FROM stg.plenario;

        DECLARE @filas INT = @@ROWCOUNT;   -- filas que insertó la foto (tiene que ir JUSTO después del INSERT)

        UPDATE ctl.carga
        SET estado = 'OK', filas_cargadas = @filas, mensaje = NULL
        WHERE carga_id = @carga_id;

        COMMIT;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK;
        UPDATE ctl.carga SET estado = 'ERROR', mensaje = LEFT(ERROR_MESSAGE(), 1000)
        WHERE carga_id = @carga_id;
        THROW;   -- re-lanza el error para que lo veas
    END CATCH
END;
GO

CREATE   PROCEDURE dw.sp_cargar_operaciones @carga_id INT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @fecha DATE = (SELECT fecha_corte FROM ctl.carga
                           WHERE carga_id = @carga_id AND tipo = 'OPERACIONES');
    IF @fecha IS NULL
        THROW 50001, 'El carga_id no existe o no es de tipo OPERACIONES.', 1;

    BEGIN TRY
        BEGIN TRAN;

        -- 1) Completar clientes
        WITH origen AS (
            SELECT CUIL, Apellido, Nombre, Domicilio, Localidad, Provincia,
                   ROW_NUMBER() OVER (PARTITION BY CUIL ORDER BY F_Liq DESC) AS rn
            FROM stg.operaciones
        )
        UPDATE d
        SET apellido = s.Apellido, nombre = s.Nombre, domicilio = s.Domicilio,
            localidad = s.Localidad, provincia = s.Provincia, fecha_modif = SYSDATETIME()
        FROM dw.dim_cliente d
        JOIN origen s ON s.CUIL = d.cuil AND s.rn = 1
        WHERE ISNULL(d.domicilio, '') <> ISNULL(s.Domicilio, '')
           OR ISNULL(d.localidad, '') <> ISNULL(s.Localidad, '')
           OR ISNULL(d.apellido, '') <> ISNULL(s.Apellido, '');

        -- 2) Completar créditos
        UPDATE d
        SET f_liquidacion = s.F_Liq, capital_otorgado = s.Cap_Otorgado, monto_total = s.Monto,
            tem = s.TEM, tna = s.TNA, fecha_modif = SYSDATETIME()
        FROM dw.dim_credito d
        JOIN stg.operaciones s ON s.Id_Op = d.credito_id
        WHERE d.f_liquidacion IS NULL
           OR d.capital_otorgado <> s.Cap_Otorgado
           OR d.monto_total <> s.Monto;

        -- 3) Foto del día (AHORA CON LAS 3 FECHAS DE PAGO)
        INSERT INTO dw.fact_foto_operaciones
            (fecha_corte, credito_id, dias_mora, tramo_mora, cuotas_pagas, cuotas_vencidas_impagas,
             imp_vencido, cap_vencido, int_vencido, iva_vencido, cuotas_a_vencer, imp_a_vencer,
             cap_a_vencer, int_a_vencer, iva_a_vencer, capital_pagado, interes_pagado, iva_pagado,
             importe_ultimo_pago, marca_cedido, f_pago_1, f_pago_2, f_pago_3, carga_id)
        SELECT @fecha, s.Id_Op, s.Dias_de_Mora, s.Tramo_de_Mora, s.Cantidad_Cuotas_Pagas,
               s.Cant_Cuotas_Imp_Vencidas, s.Imp_Cuotas_Imp_Vencidas, s.Cap_Cuotas_Imp_Vencidas,
               s.Int_Cuotas_Imp_Vencidas, s.Iva_Cuotas_Imp_Vencidas, s.Cant_Cuotas_Imp_a_Vencer,
               s.Imp_Cuotas_Imp_a_Vencer, s.Cap_Cuotas_Imp_a_Vencer, s.Int_Cuotas_Imp_a_Vencer,
               s.Iva_Cuotas_Imp_a_Vencer, s.Capital_Pagado, s.Interes_Pagado, s.IVA_Pagado,
               s.Importe_Ultimo_Pago, s.Marca_Cedido,
               TRY_CONVERT(DATE, NULLIF(SUBSTRING(s.F_Ultimos_3_Pagos,  1, 10), ''), 103),
               TRY_CONVERT(DATE, NULLIF(SUBSTRING(s.F_Ultimos_3_Pagos, 14, 10), ''), 103),
               TRY_CONVERT(DATE, NULLIF(SUBSTRING(s.F_Ultimos_3_Pagos, 27, 10), ''), 103),
               @carga_id
        FROM stg.operaciones s
        JOIN dw.dim_credito d ON d.credito_id = s.Id_Op;

        DECLARE @filas INT = @@ROWCOUNT;
        DECLARE @sin_credito INT = (SELECT COUNT(*) FROM stg.operaciones s
                                    WHERE NOT EXISTS (SELECT 1 FROM dw.dim_credito d
                                                      WHERE d.credito_id = s.Id_Op));

        UPDATE ctl.carga
        SET estado = 'OK', filas_cargadas = @filas,
            mensaje = CASE WHEN @sin_credito > 0
                           THEN CONCAT(@sin_credito, ' créditos no están en dim_credito (falta un Plenario más nuevo).')
                      END
        WHERE carga_id = @carga_id;

        COMMIT;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK;
        UPDATE ctl.carga SET estado = 'ERROR', mensaje = LEFT(ERROR_MESSAGE(), 1000)
        WHERE carga_id = @carga_id;
        THROW;
    END CATCH
END;
GO

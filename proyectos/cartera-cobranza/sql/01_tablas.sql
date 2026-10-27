-- Esquemas: stg (staging, el archivo tal cual llega) | dw (modelo) | rpt (vistas para Power BI) | ctl (auditoría de cargas)
CREATE SCHEMA stg;
GO
CREATE SCHEMA dw;
GO
CREATE SCHEMA rpt;
GO
CREATE SCHEMA ctl;
GO

CREATE TABLE ctl.carga (
    [carga_id] INT IDENTITY(1,1) NOT NULL,
    [archivo] NVARCHAR(260) NOT NULL,
    [hash_archivo] CHAR(64) NOT NULL,
    [tipo] VARCHAR(50) NOT NULL,
    [fecha_corte] DATE NOT NULL,
    [filas_leidas] INT NULL,
    [filas_cargadas] INT NULL,
    [estado] VARCHAR(30) NOT NULL DEFAULT ('EN CURSO'),
    [mensaje] NVARCHAR(1000) NULL,
    [fecha_carga] DATETIME2 NOT NULL DEFAULT (sysdatetime()),
    CONSTRAINT [PK__carga__FC0B27E607B42D83] PRIMARY KEY (carga_id),
    CONSTRAINT [UQ_carga_hash] UNIQUE (hash_archivo)
);
GO

CREATE TABLE stg.operaciones (
    [Id_Op] INT NULL,
    [CUIL] VARCHAR(20) NULL,
    [Tipo_Doc] VARCHAR(10) NULL,
    [Nro_Doc] BIGINT NULL,
    [Apellido] NVARCHAR(60) NULL,
    [Nombre] NVARCHAR(60) NULL,
    [Domicilio] NVARCHAR(120) NULL,
    [Provincia] NVARCHAR(40) NULL,
    [Localidad] NVARCHAR(60) NULL,
    [F_Liq] DATE NULL,
    [F_1er_Vto] DATE NULL,
    [F_Ult_Vto] DATE NULL,
    [TEM] DECIMAL(10,6) NULL,
    [Cap_Solicitado] DECIMAL(18,2) NULL,
    [Cap_Otorgado] DECIMAL(18,2) NULL,
    [Monto] DECIMAL(18,2) NULL,
    [Cant_Cuotas] SMALLINT NULL,
    [Val_Cuota] DECIMAL(18,2) NULL,
    [Cant_Cuotas_Imp_Vencidas] SMALLINT NULL,
    [Imp_Cuotas_Imp_Vencidas] DECIMAL(18,2) NULL,
    [Cap_Cuotas_Imp_Vencidas] DECIMAL(18,2) NULL,
    [Int_Cuotas_Imp_Vencidas] DECIMAL(18,2) NULL,
    [Iva_Cuotas_Imp_Vencidas] DECIMAL(18,2) NULL,
    [Cant_Cuotas_Imp_a_Vencer] SMALLINT NULL,
    [Imp_Cuotas_Imp_a_Vencer] DECIMAL(18,2) NULL,
    [Cap_Cuotas_Imp_a_Vencer] DECIMAL(18,2) NULL,
    [Int_Cuotas_Imp_a_Vencer] DECIMAL(18,2) NULL,
    [Iva_Cuotas_Imp_a_Vencer] DECIMAL(18,2) NULL,
    [Dias_de_Mora] INT NULL,
    [Entidad] NVARCHAR(60) NULL,
    [Importe_Ultimo_Pago] DECIMAL(18,2) NULL,
    [Tramo_de_Mora] VARCHAR(10) NULL,
    [TNA] DECIMAL(10,4) NULL,
    [Marca_Cedido] NVARCHAR(40) NULL,
    [Cantidad_Cuotas_Pagas] SMALLINT NULL,
    [Capital_Pagado] DECIMAL(18,2) NULL,
    [Interes_Pagado] DECIMAL(18,2) NULL,
    [IVA_Pagado] DECIMAL(18,2) NULL,
    [F_Ultimos_3_Pagos] NVARCHAR(40) NULL
);
GO

CREATE TABLE stg.plenario (
    [Id] INT NULL,
    [Id_Solicitud] INT NULL,
    [Codigo_Referencia] VARCHAR(20) NULL,
    [Id_Cred_Terceros] INT NULL,
    [Estado] VARCHAR(15) NULL,
    [Origen] VARCHAR(20) NULL,
    [Linea] NVARCHAR(100) NULL,
    [Id_Cliente] INT NULL,
    [Cliente] NVARCHAR(80) NULL,
    [Tipo_Nro_Doc] VARCHAR(20) NULL,
    [CUIL] VARCHAR(20) NULL,
    [Valor_Cuota] DECIMAL(18,2) NULL,
    [F_Operacion] DATE NULL,
    [Capital] DECIMAL(18,2) NULL,
    [Tasa_Pct] DECIMAL(10,4) NULL,
    [IVA_Pct] DECIMAL(5,2) NULL,
    [Plazo] SMALLINT NULL,
    [F_Cancelacion_Anticipada] DATE NULL,
    [Entidad_Id] INT NULL,
    [Vendido] VARCHAR(10) NULL,
    [F_1er_Vto] DATE NULL,
    [F_Proximo_Vto] DATE NULL,
    [F_Ultimo_Vto] DATE NULL,
    [F_Ult_Pago_Parcial] DATE NULL,
    [F_Ult_Pago_Total] DATE NULL,
    [Dias_Mora] INT NULL,
    [Cuotas_Cobradas] SMALLINT NULL,
    [Capital_Adeudado] DECIMAL(18,2) NULL,
    [Total_Cobrado] DECIMAL(18,2) NULL,
    [Nro_Legajo] VARCHAR(20) NULL,
    [Capital_Cobrado] DECIMAL(18,2) NULL,
    [Capital_Adeudado_Vencido] DECIMAL(18,2) NULL,
    [F_1er_Vto_Impago] DATE NULL
);
GO

CREATE TABLE dw.dim_mutual (
    [mutual_id] INT NOT NULL,
    [sigla] VARCHAR(20) NOT NULL,
    [nombre_plenario] NVARCHAR(80) NOT NULL,
    [activa] BIT NOT NULL DEFAULT ((1)),
    [dia_vencimiento] TINYINT NULL,
    CONSTRAINT [PK__dim_mutu__2AD81483698AE4C5] PRIMARY KEY (mutual_id),
    CONSTRAINT [UQ__dim_mutu__3C47D51981C3A564] UNIQUE (sigla),
    CONSTRAINT [UQ__dim_mutu__FC28B75FC469F8D7] UNIQUE (nombre_plenario)
);
GO

CREATE TABLE dw.dim_cliente (
    [cliente_id] INT NOT NULL,
    [cuil] CHAR(13) NOT NULL,
    [tipo_doc] VARCHAR(5) NULL,
    [nro_doc] BIGINT NULL,
    [apellido] NVARCHAR(60) NULL,
    [nombre] NVARCHAR(60) NULL,
    [domicilio] NVARCHAR(120) NULL,
    [localidad] NVARCHAR(60) NULL,
    [provincia] NVARCHAR(40) NULL,
    [fecha_alta] DATETIME2 NOT NULL DEFAULT (sysdatetime()),
    [fecha_modif] DATETIME2 NULL,
    CONSTRAINT [PK__dim_clie__47E34D641E74568C] PRIMARY KEY (cliente_id)
);
GO

CREATE TABLE dw.dim_credito (
    [credito_id] INT NOT NULL,
    [cliente_id] INT NOT NULL,
    [mutual_id] INT NOT NULL,
    [id_solicitud] INT NULL,
    [codigo_referencia] VARCHAR(20) NULL,
    [id_cred_terceros] INT NULL,
    [nro_legajo] VARCHAR(20) NULL,
    [linea] NVARCHAR(100) NULL,
    [origen] VARCHAR(20) NULL,
    [f_operacion] DATE NULL,
    [f_liquidacion] DATE NULL,
    [f_1er_vto] DATE NULL,
    [f_ult_vto] DATE NULL,
    [capital] DECIMAL(18,2) NULL,
    [capital_otorgado] DECIMAL(18,2) NULL,
    [monto_total] DECIMAL(18,2) NULL,
    [plazo] SMALLINT NULL,
    [valor_cuota] DECIMAL(18,2) NULL,
    [tasa_pct] DECIMAL(10,4) NULL,
    [tem] DECIMAL(10,6) NULL,
    [tna] DECIMAL(10,4) NULL,
    [iva_pct] DECIMAL(5,2) NULL,
    [fecha_alta] DATETIME2 NOT NULL DEFAULT (sysdatetime()),
    [fecha_modif] DATETIME2 NULL,
    CONSTRAINT [PK__dim_cred__57F2CD288DD8E67D] PRIMARY KEY (credito_id),
    CONSTRAINT [FK__dim_credi__mutua__47DBAE45] FOREIGN KEY (mutual_id) REFERENCES dw.dim_mutual (mutual_id),
    CONSTRAINT [FK__dim_credi__clien__46E78A0C] FOREIGN KEY (cliente_id) REFERENCES dw.dim_cliente (cliente_id)
);
GO

CREATE TABLE dw.fact_foto_operaciones (
    [fecha_corte] DATE NOT NULL,
    [credito_id] INT NOT NULL,
    [dias_mora] INT NULL,
    [tramo_mora] VARCHAR(10) NULL,
    [cuotas_pagas] SMALLINT NULL,
    [cuotas_vencidas_impagas] SMALLINT NULL,
    [imp_vencido] DECIMAL(18,2) NULL,
    [cap_vencido] DECIMAL(18,2) NULL,
    [int_vencido] DECIMAL(18,2) NULL,
    [iva_vencido] DECIMAL(18,2) NULL,
    [cuotas_a_vencer] SMALLINT NULL,
    [imp_a_vencer] DECIMAL(18,2) NULL,
    [cap_a_vencer] DECIMAL(18,2) NULL,
    [int_a_vencer] DECIMAL(18,2) NULL,
    [iva_a_vencer] DECIMAL(18,2) NULL,
    [capital_pagado] DECIMAL(18,2) NULL,
    [interes_pagado] DECIMAL(18,2) NULL,
    [iva_pagado] DECIMAL(18,2) NULL,
    [importe_ultimo_pago] DECIMAL(18,2) NULL,
    [marca_cedido] NVARCHAR(40) NULL,
    [carga_id] INT NOT NULL,
    [f_pago_1] DATE NULL,
    [f_pago_2] DATE NULL,
    [f_pago_3] DATE NULL,
    CONSTRAINT [PK_foto_operaciones] PRIMARY KEY (fecha_corte, credito_id),
    CONSTRAINT [FK__fact_foto__carga__4CA06362] FOREIGN KEY (carga_id) REFERENCES ctl.carga (carga_id),
    CONSTRAINT [FK__fact_foto__credi__4BAC3F29] FOREIGN KEY (credito_id) REFERENCES dw.dim_credito (credito_id)
);
GO

CREATE TABLE dw.fact_foto_plenario (
    [fecha_corte] DATE NOT NULL,
    [credito_id] INT NOT NULL,
    [estado] VARCHAR(15) NULL,
    [dias_mora] INT NULL,
    [cuotas_cobradas] SMALLINT NULL,
    [capital_adeudado] DECIMAL(18,2) NULL,
    [capital_adeudado_vencido] DECIMAL(18,2) NULL,
    [capital_cobrado] DECIMAL(18,2) NULL,
    [total_cobrado] DECIMAL(18,2) NULL,
    [f_proximo_vto] DATE NULL,
    [f_1er_vto_impago] DATE NULL,
    [f_ult_pago_parcial] DATE NULL,
    [f_ult_pago_total] DATE NULL,
    [f_cancelacion_anticipada] DATE NULL,
    [vendido] BIT NULL,
    [carga_id] INT NOT NULL,
    CONSTRAINT [PK_foto_plenario] PRIMARY KEY (fecha_corte, credito_id),
    CONSTRAINT [FK__fact_foto__carga__5070F446] FOREIGN KEY (carga_id) REFERENCES ctl.carga (carga_id),
    CONSTRAINT [FK__fact_foto__credi__4F7CD00D] FOREIGN KEY (credito_id) REFERENCES dw.dim_credito (credito_id)
);
GO

from __future__ import annotations

import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import pyodbc
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL


# ==========================================================
# CONFIGURACIÓN
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

CARPETA_CURATED = (
    BASE_DIR
    / "data"
    / "curated"
    / "oece"
)

CARPETA_CALIDAD = (
    BASE_DIR
    / "data"
    / "quality"
    / "oece"
)

CARPETA_CALIDAD.mkdir(
    parents=True,
    exist_ok=True
)

ARCHIVO_VALIDACION = (
    CARPETA_CALIDAD
    / "VALIDACION_BD_LOCAL_OECE.xlsx"
)

# Cambia este valor si tu instancia tiene otro nombre.
# Ejemplos:
# r".\SQLEXPRESS"
# r"DESKTOP-ABC123\SQLEXPRESS"
# r"localhost"
SERVIDOR_SQL = r"VIKDIR\SQLEXPRESS"

BASE_DATOS = "OECE_DW_LOCAL"

USAR_AUTENTICACION_WINDOWS = True

USUARIO_SQL = ""
PASSWORD_SQL = ""

TAMANIO_LOTE = 1000

CANTIDAD_MUESTRAS_ALEATORIAS = 20


# ==========================================================
# ARCHIVOS DEL MODELO CURADO
# ==========================================================

ARCHIVOS_CURATED = {
    "FACT_CONTRATO_ACTUAL":
        CARPETA_CURATED
        / "FACT_CONTRATO_ACTUAL.parquet",

    "BRIDGE_CONTRATO_PROVEEDOR":
        CARPETA_CURATED
        / "BRIDGE_CONTRATO_PROVEEDOR.parquet",

    "DIM_ORGANIZACION":
        CARPETA_CURATED
        / "DIM_ORGANIZACION.parquet",

    "DIM_ENTIDAD":
        CARPETA_CURATED
        / "DIM_ENTIDAD.parquet",

    "DIM_PROVEEDOR":
        CARPETA_CURATED
        / "DIM_PROVEEDOR.parquet",

    "DIM_UBICACION":
        CARPETA_CURATED
        / "DIM_UBICACION.parquet",

    "DIM_MONEDA":
        CARPETA_CURATED
        / "DIM_MONEDA.parquet",

    "DIM_CATEGORIA":
        CARPETA_CURATED
        / "DIM_CATEGORIA.parquet",

    "DIM_METODO_CONTRATACION":
        CARPETA_CURATED
        / "DIM_METODO_CONTRATACION.parquet",

    "DIM_FECHA":
        CARPETA_CURATED
        / "DIM_FECHA.parquet",
}


# ==========================================================
# CONEXIÓN
# ==========================================================

def detectar_driver_odbc() -> str:
    drivers = pyodbc.drivers()

    preferidos = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
    ]

    for driver in preferidos:
        if driver in drivers:
            return driver

    instalados = "\n".join(
        f"- {driver}"
        for driver in drivers
    )

    raise RuntimeError(
        "No se encontró ODBC Driver 18 ni 17 "
        "para SQL Server.\n\n"
        "Drivers instalados:\n"
        f"{instalados}"
    )


def validar_nombre_base_datos(nombre: str) -> None:
    if not re.fullmatch(
        r"[A-Za-z0-9_]+",
        nombre
    ):
        raise ValueError(
            "El nombre de la base de datos contiene "
            "caracteres no permitidos."
        )


def construir_cadena_odbc(
    driver: str,
    base_datos: str
) -> str:

    partes = [
        f"DRIVER={{{driver}}}",
        f"SERVER={SERVIDOR_SQL}",
        f"DATABASE={base_datos}",
        "TrustServerCertificate=yes",
    ]

    if driver == "ODBC Driver 18 for SQL Server":
        partes.append("Encrypt=no")
        partes.append("LongAsMax=yes")
    else:
        partes.append("Encrypt=no")

    if USAR_AUTENTICACION_WINDOWS:
        partes.append(
            "Trusted_Connection=yes"
        )
    else:
        partes.append(
            f"UID={USUARIO_SQL}"
        )
        partes.append(
            f"PWD={PASSWORD_SQL}"
        )

    return ";".join(partes) + ";"


def crear_engine_sqlalchemy(
    driver: str,
    base_datos: str
):
    cadena_odbc = construir_cadena_odbc(
        driver,
        base_datos
    )

    url = URL.create(
        "mssql+pyodbc",
        query={
            "odbc_connect": cadena_odbc
        }
    )

    return create_engine(
        url,
        fast_executemany=True,
        pool_pre_ping=True,
        future=True,
    )


# ==========================================================
# CREAR BASE DE DATOS
# ==========================================================

def crear_base_datos(
    driver: str
) -> None:

    validar_nombre_base_datos(
        BASE_DATOS
    )

    conexion = pyodbc.connect(
        construir_cadena_odbc(
            driver,
            "master"
        ),
        autocommit=True
    )

    try:
        cursor = conexion.cursor()

        consulta = f"""
        IF DB_ID(N'{BASE_DATOS}') IS NULL
        BEGIN
            CREATE DATABASE [{BASE_DATOS}];
        END;
        """

        cursor.execute(consulta)

        print(
            f"Base de datos disponible: "
            f"{BASE_DATOS}"
        )

    finally:
        conexion.close()


# ==========================================================
# VALIDAR ARCHIVOS
# ==========================================================

def validar_archivos_curated() -> None:
    faltantes = [
        str(ruta)
        for ruta in ARCHIVOS_CURATED.values()
        if not ruta.exists()
    ]

    if faltantes:
        detalle = "\n".join(
            f"- {ruta}"
            for ruta in faltantes
        )

        raise FileNotFoundError(
            "Faltan archivos del modelo curado:\n"
            f"{detalle}"
        )


# ==========================================================
# CREAR ESQUEMAS
# ==========================================================

def crear_esquemas(engine) -> None:
    instrucciones = [
        """
        IF SCHEMA_ID('curated') IS NULL
            EXEC('CREATE SCHEMA curated');
        """,

        """
        IF SCHEMA_ID('dw') IS NULL
            EXEC('CREATE SCHEMA dw');
        """,

        """
        IF SCHEMA_ID('control') IS NULL
            EXEC('CREATE SCHEMA control');
        """,
    ]

    with engine.begin() as conexion:
        for instruccion in instrucciones:
            conexion.execute(
                text(instruccion)
            )


# ==========================================================
# PREPARAR DATAFRAME PARA SQL SERVER
# ==========================================================

def preparar_dataframe(
    df: pd.DataFrame
) -> pd.DataFrame:

    df = df.copy()

    for columna in df.columns:
        serie = df[columna]

        if pd.api.types.is_datetime64_any_dtype(
            serie
        ):
            # SQL Server no acepta NaT como parámetro.
            df[columna] = serie.astype(object)

            df.loc[
                pd.isna(df[columna]),
                columna
            ] = None

        elif pd.api.types.is_bool_dtype(
            serie
        ):
            df[columna] = (
                serie
                .astype("boolean")
                .astype(object)
            )

            df.loc[
                pd.isna(df[columna]),
                columna
            ] = None

        elif pd.api.types.is_object_dtype(
            serie
        ) or pd.api.types.is_string_dtype(
            serie
        ):
            df[columna] = serie.astype(object)

            df.loc[
                pd.isna(df[columna]),
                columna
            ] = None

        else:
            df[columna] = serie.astype(object)

            df.loc[
                pd.isna(df[columna]),
                columna
            ] = None

    return df


# ==========================================================
# CARGAR CAPA CURATED
# ==========================================================

def cargar_tablas_curated(
    engine
) -> pd.DataFrame:

    resumen = []

    print("\n========================================")
    print("CARGANDO CAPA CURATED")
    print("========================================")

    for tabla, archivo in ARCHIVOS_CURATED.items():
        inicio = datetime.now()

        print(f"\nLeyendo {tabla}...")
        print(archivo)

        df = pd.read_parquet(
            archivo
        )

        filas = len(df)
        columnas = len(df.columns)

        print(
            f"Filas: {filas:,} | "
            f"Columnas: {columnas}"
        )

        df = preparar_dataframe(df)

        print(
            f"Insertando curated.{tabla}..."
        )

        df.to_sql(
            name=tabla,
            con=engine,
            schema="curated",
            if_exists="replace",
            index=False,
            chunksize=TAMANIO_LOTE,
            method=None,
        )

        fin = datetime.now()

        resumen.append({
            "TABLA": tabla,
            "ARCHIVO": archivo.name,
            "FILAS": filas,
            "COLUMNAS": columnas,
            "DURACION_SEGUNDOS": (
                fin - inicio
            ).total_seconds(),
            "ESTADO": "OK",
        })

        print(
            f"Finalizado: {fin - inicio}"
        )

        del df

    return pd.DataFrame(resumen)


# ==========================================================
# CREAR TABLAS DEL MODELO ESTRELLA
# ==========================================================

SQL_CREAR_MODELO = """
/* ========================================================
   LIMPIAR MODELO ANTERIOR
   ======================================================== */

DROP TABLE IF EXISTS dw.BridgeContratoProveedor;
DROP TABLE IF EXISTS dw.FactContrato;
DROP TABLE IF EXISTS dw.DimProveedor;
DROP TABLE IF EXISTS dw.DimEntidad;
DROP TABLE IF EXISTS dw.DimMetodoContratacion;
DROP TABLE IF EXISTS dw.DimCategoria;
DROP TABLE IF EXISTS dw.DimMoneda;
DROP TABLE IF EXISTS dw.DimUbicacion;
DROP TABLE IF EXISTS dw.DimFecha;


/* ========================================================
   DIMENSIÓN FECHA
   ======================================================== */

CREATE TABLE dw.DimFecha (
    FechaKey               INT            NOT NULL,
    Fecha                  DATE           NOT NULL,
    Anio                   INT            NOT NULL,
    Semestre               TINYINT        NOT NULL,
    Trimestre              TINYINT        NOT NULL,
    MesNumero              TINYINT        NOT NULL,
    MesNombre              NVARCHAR(20)   NOT NULL,
    AnioMes                CHAR(7)        NOT NULL,
    Dia                    TINYINT        NOT NULL,
    DiaSemanaNumero        TINYINT        NOT NULL,
    DiaSemanaNombre        NVARCHAR(20)   NOT NULL,
    EsFinSemana            BIT            NOT NULL,

    CONSTRAINT PK_DimFecha
        PRIMARY KEY (FechaKey),

    CONSTRAINT UQ_DimFecha_Fecha
        UNIQUE (Fecha)
);

INSERT INTO dw.DimFecha (
    FechaKey,
    Fecha,
    Anio,
    Semestre,
    Trimestre,
    MesNumero,
    MesNombre,
    AnioMes,
    Dia,
    DiaSemanaNumero,
    DiaSemanaNombre,
    EsFinSemana
)
SELECT DISTINCT
    TRY_CONVERT(INT, FECHA_KEY),
    TRY_CONVERT(DATE, FECHA),
    TRY_CONVERT(INT, ANIO),
    TRY_CONVERT(TINYINT, SEMESTRE),
    TRY_CONVERT(TINYINT, TRIMESTRE),
    TRY_CONVERT(TINYINT, MES_NUMERO),
    CONVERT(NVARCHAR(20), MES_NOMBRE),
    CONVERT(CHAR(7), ANIO_MES),
    TRY_CONVERT(TINYINT, DIA),
    TRY_CONVERT(TINYINT, DIA_SEMANA_NUMERO),
    CONVERT(NVARCHAR(20), DIA_SEMANA_NOMBRE),
    TRY_CONVERT(BIT, ES_FIN_SEMANA)
FROM curated.DIM_FECHA
WHERE
    TRY_CONVERT(INT, FECHA_KEY) IS NOT NULL
    AND TRY_CONVERT(DATE, FECHA) IS NOT NULL;


/* ========================================================
   DIMENSIÓN UBICACIÓN
   ======================================================== */

CREATE TABLE dw.DimUbicacion (
    UbicacionKey           BIGINT IDENTITY(1,1) NOT NULL,
    ClaveUbicacion         NVARCHAR(100)         NOT NULL,
    Pais                   NVARCHAR(200)         NULL,
    Departamento           NVARCHAR(200)         NULL,
    Region                 NVARCHAR(200)         NULL,
    Localidad              NVARCHAR(300)         NULL,

    CONSTRAINT PK_DimUbicacion
        PRIMARY KEY (UbicacionKey),

    CONSTRAINT UQ_DimUbicacion_Clave
        UNIQUE (ClaveUbicacion)
);

INSERT INTO dw.DimUbicacion (
    ClaveUbicacion,
    Pais,
    Departamento,
    Region,
    Localidad
)
SELECT DISTINCT
    CONVERT(NVARCHAR(100), CLAVE_UBICACION),
    CONVERT(NVARCHAR(200), PAIS),
    CONVERT(NVARCHAR(200), DEPARTAMENTO),
    CONVERT(NVARCHAR(200), REGION),
    CONVERT(NVARCHAR(300), LOCALIDAD)
FROM curated.DIM_UBICACION
WHERE CLAVE_UBICACION IS NOT NULL;


/* ========================================================
   DIMENSIÓN ENTIDAD
   ======================================================== */

CREATE TABLE dw.DimEntidad (
    EntidadKey                  BIGINT IDENTITY(1,1) NOT NULL,
    ClaveOrganizacion          NVARCHAR(450)         NOT NULL,
    OrganizacionHash           CHAR(32)              NULL,
    IdOrganizacionOriginal     NVARCHAR(300)         NULL,
    Ruc                        VARCHAR(20)            NULL,
    RucValido                  BIT                    NULL,
    NombreOriginal             NVARCHAR(1000)        NULL,
    NombreEstandar             NVARCHAR(1000)        NULL,
    TipoOrganizacion           NVARCHAR(100)         NULL,
    Direccion                  NVARCHAR(1500)        NULL,
    Localidad                  NVARCHAR(300)         NULL,
    Region                     NVARCHAR(200)         NULL,
    Departamento               NVARCHAR(200)         NULL,
    Pais                       NVARCHAR(200)         NULL,
    UbicacionKey               BIGINT                NULL,
    EsEntidad                  BIT                   NULL,
    EsProveedor                BIT                   NULL,
    PeriodoPrimeraAparicion    INT                   NULL,
    PeriodoUltimaAparicion     INT                   NULL,

    CONSTRAINT PK_DimEntidad
        PRIMARY KEY (EntidadKey),

    CONSTRAINT UQ_DimEntidad_Clave
        UNIQUE (ClaveOrganizacion),

    CONSTRAINT FK_DimEntidad_Ubicacion
        FOREIGN KEY (UbicacionKey)
        REFERENCES dw.DimUbicacion(UbicacionKey)
);

INSERT INTO dw.DimEntidad (
    ClaveOrganizacion,
    OrganizacionHash,
    IdOrganizacionOriginal,
    Ruc,
    RucValido,
    NombreOriginal,
    NombreEstandar,
    TipoOrganizacion,
    Direccion,
    Localidad,
    Region,
    Departamento,
    Pais,
    UbicacionKey,
    EsEntidad,
    EsProveedor,
    PeriodoPrimeraAparicion,
    PeriodoUltimaAparicion
)
SELECT
    CONVERT(NVARCHAR(450), e.CLAVE_ORGANIZACION),
    CONVERT(CHAR(32), e.ORGANIZACION_HASH),
    CONVERT(NVARCHAR(300), e.ID_ORGANIZACION_ORIGINAL),
    CONVERT(VARCHAR(20), e.RUC),
    TRY_CONVERT(BIT, e.RUC_VALIDO),
    CONVERT(NVARCHAR(1000), e.NOMBRE_ORIGINAL),
    CONVERT(NVARCHAR(1000), e.NOMBRE_ESTANDAR),
    CONVERT(NVARCHAR(100), e.TIPO_ORGANIZACION),
    CONVERT(NVARCHAR(1500), e.DIRECCION),
    CONVERT(NVARCHAR(300), e.LOCALIDAD),
    CONVERT(NVARCHAR(200), e.REGION),
    CONVERT(NVARCHAR(200), e.DEPARTAMENTO),
    CONVERT(NVARCHAR(200), e.PAIS),
    u.UbicacionKey,
    TRY_CONVERT(BIT, e.ES_ENTIDAD),
    TRY_CONVERT(BIT, e.ES_PROVEEDOR),
    TRY_CONVERT(INT, e.PERIODO_PRIMERA_APARICION),
    TRY_CONVERT(INT, e.PERIODO_ULTIMA_APARICION)
FROM curated.DIM_ENTIDAD AS e
LEFT JOIN dw.DimUbicacion AS u
    ON u.ClaveUbicacion =
       CONVERT(NVARCHAR(100), e.CLAVE_UBICACION)
WHERE e.CLAVE_ORGANIZACION IS NOT NULL;


/* ========================================================
   DIMENSIÓN PROVEEDOR
   ======================================================== */

CREATE TABLE dw.DimProveedor (
    ProveedorKey               BIGINT IDENTITY(1,1) NOT NULL,
    ClaveOrganizacion          NVARCHAR(450)         NOT NULL,
    OrganizacionHash           CHAR(32)              NULL,
    IdOrganizacionOriginal     NVARCHAR(300)         NULL,
    Ruc                        VARCHAR(20)            NULL,
    RucValido                  BIT                    NULL,
    NombreOriginal             NVARCHAR(1000)        NULL,
    NombreEstandar             NVARCHAR(1000)        NULL,
    TipoOrganizacion           NVARCHAR(100)         NULL,
    Direccion                  NVARCHAR(1500)        NULL,
    Localidad                  NVARCHAR(300)         NULL,
    Region                     NVARCHAR(200)         NULL,
    Departamento               NVARCHAR(200)         NULL,
    Pais                       NVARCHAR(200)         NULL,
    UbicacionKey               BIGINT                NULL,
    EsEntidad                  BIT                   NULL,
    EsProveedor                BIT                   NULL,
    PeriodoPrimeraAparicion    INT                   NULL,
    PeriodoUltimaAparicion     INT                   NULL,

    CONSTRAINT PK_DimProveedor
        PRIMARY KEY (ProveedorKey),

    CONSTRAINT UQ_DimProveedor_Clave
        UNIQUE (ClaveOrganizacion),

    CONSTRAINT FK_DimProveedor_Ubicacion
        FOREIGN KEY (UbicacionKey)
        REFERENCES dw.DimUbicacion(UbicacionKey)
);

INSERT INTO dw.DimProveedor (
    ClaveOrganizacion,
    OrganizacionHash,
    IdOrganizacionOriginal,
    Ruc,
    RucValido,
    NombreOriginal,
    NombreEstandar,
    TipoOrganizacion,
    Direccion,
    Localidad,
    Region,
    Departamento,
    Pais,
    UbicacionKey,
    EsEntidad,
    EsProveedor,
    PeriodoPrimeraAparicion,
    PeriodoUltimaAparicion
)
SELECT
    CONVERT(NVARCHAR(450), p.CLAVE_ORGANIZACION),
    CONVERT(CHAR(32), p.ORGANIZACION_HASH),
    CONVERT(NVARCHAR(300), p.ID_ORGANIZACION_ORIGINAL),
    CONVERT(VARCHAR(20), p.RUC),
    TRY_CONVERT(BIT, p.RUC_VALIDO),
    CONVERT(NVARCHAR(1000), p.NOMBRE_ORIGINAL),
    CONVERT(NVARCHAR(1000), p.NOMBRE_ESTANDAR),
    CONVERT(NVARCHAR(100), p.TIPO_ORGANIZACION),
    CONVERT(NVARCHAR(1500), p.DIRECCION),
    CONVERT(NVARCHAR(300), p.LOCALIDAD),
    CONVERT(NVARCHAR(200), p.REGION),
    CONVERT(NVARCHAR(200), p.DEPARTAMENTO),
    CONVERT(NVARCHAR(200), p.PAIS),
    u.UbicacionKey,
    TRY_CONVERT(BIT, p.ES_ENTIDAD),
    TRY_CONVERT(BIT, p.ES_PROVEEDOR),
    TRY_CONVERT(INT, p.PERIODO_PRIMERA_APARICION),
    TRY_CONVERT(INT, p.PERIODO_ULTIMA_APARICION)
FROM curated.DIM_PROVEEDOR AS p
LEFT JOIN dw.DimUbicacion AS u
    ON u.ClaveUbicacion =
       CONVERT(NVARCHAR(100), p.CLAVE_UBICACION)
WHERE p.CLAVE_ORGANIZACION IS NOT NULL;


/* ========================================================
   DIMENSIÓN MONEDA
   ======================================================== */

CREATE TABLE dw.DimMoneda (
    MonedaKey              INT IDENTITY(1,1) NOT NULL,
    ClaveMoneda            NVARCHAR(100)      NOT NULL,
    CodigoMoneda           NVARCHAR(50)       NULL,
    NombreMoneda           NVARCHAR(200)      NULL,

    CONSTRAINT PK_DimMoneda
        PRIMARY KEY (MonedaKey),

    CONSTRAINT UQ_DimMoneda_Clave
        UNIQUE (ClaveMoneda)
);

INSERT INTO dw.DimMoneda (
    ClaveMoneda,
    CodigoMoneda,
    NombreMoneda
)
SELECT DISTINCT
    CONVERT(NVARCHAR(100), CLAVE_MONEDA),
    CONVERT(NVARCHAR(50), CODIGO_MONEDA),
    CONVERT(NVARCHAR(200), NOMBRE_MONEDA)
FROM curated.DIM_MONEDA
WHERE CLAVE_MONEDA IS NOT NULL;


/* ========================================================
   DIMENSIÓN CATEGORÍA
   ======================================================== */

CREATE TABLE dw.DimCategoria (
    CategoriaKey           INT IDENTITY(1,1) NOT NULL,
    ClaveCategoria         NVARCHAR(100)      NOT NULL,
    CategoriaPrincipal     NVARCHAR(500)      NULL,

    CONSTRAINT PK_DimCategoria
        PRIMARY KEY (CategoriaKey),

    CONSTRAINT UQ_DimCategoria_Clave
        UNIQUE (ClaveCategoria)
);

INSERT INTO dw.DimCategoria (
    ClaveCategoria,
    CategoriaPrincipal
)
SELECT DISTINCT
    CONVERT(NVARCHAR(100), CLAVE_CATEGORIA),
    CONVERT(NVARCHAR(500), CATEGORIA_PRINCIPAL)
FROM curated.DIM_CATEGORIA
WHERE CLAVE_CATEGORIA IS NOT NULL;


/* ========================================================
   DIMENSIÓN MÉTODO
   ======================================================== */

CREATE TABLE dw.DimMetodoContratacion (
    MetodoKey                  INT IDENTITY(1,1) NOT NULL,
    ClaveMetodo                NVARCHAR(100)      NOT NULL,
    MetodoContratacion         NVARCHAR(500)      NULL,
    DetalleMetodoContratacion  NVARCHAR(1000)     NULL,

    CONSTRAINT PK_DimMetodoContratacion
        PRIMARY KEY (MetodoKey),

    CONSTRAINT UQ_DimMetodo_Clave
        UNIQUE (ClaveMetodo)
);

INSERT INTO dw.DimMetodoContratacion (
    ClaveMetodo,
    MetodoContratacion,
    DetalleMetodoContratacion
)
SELECT DISTINCT
    CONVERT(NVARCHAR(100), CLAVE_METODO),
    CONVERT(NVARCHAR(500), METODO_CONTRATACION),
    CONVERT(NVARCHAR(1000), DETALLE_METODO_CONTRATACION)
FROM curated.DIM_METODO_CONTRATACION
WHERE CLAVE_METODO IS NOT NULL;


/* ========================================================
   TABLA DE HECHOS
   ======================================================== */

CREATE TABLE dw.FactContrato (
    ContratoKey                 BIGINT IDENTITY(1,1) NOT NULL,
    ClaveContrato               NVARCHAR(450)         NOT NULL,
    ContratoHash                CHAR(32)              NULL,
    HashFila                    CHAR(32)              NULL,

    Ocid                        NVARCHAR(200)         NULL,
    IdEntrega                   NVARCHAR(300)         NULL,
    IdContrato                  NVARCHAR(300)         NULL,
    IdAdjudicacion              NVARCHAR(300)         NULL,
    IdLicitacion                NVARCHAR(300)         NULL,

    EntidadKey                  BIGINT                NULL,
    UbicacionEntidadKey         BIGINT                NULL,
    MonedaKey                   INT                   NULL,
    CategoriaKey                INT                   NULL,
    MetodoKey                   INT                   NULL,

    FechaFirmaKey               INT                   NULL,
    FechaInicioKey              INT                   NULL,
    FechaFinKey                 INT                   NULL,

    TituloContrato              NVARCHAR(2000)        NULL,
    DescripcionContrato         NVARCHAR(MAX)         NULL,

    FechaFirma                  DATETIME2             NULL,
    FechaInicio                 DATETIME2             NULL,
    FechaFin                    DATETIME2             NULL,
    FechaFinImplementacion      DATETIME2             NULL,

    DuracionDias                DECIMAL(18,2)         NULL,
    MontoContrato               DECIMAL(28,4)         NULL,
    MontoFinal                  DECIMAL(28,4)         NULL,

    MonedaContrato              NVARCHAR(50)          NULL,
    NombreMonedaContrato        NVARCHAR(200)         NULL,
    MonedaFinal                 NVARCHAR(50)          NULL,
    NombreMonedaFinal           NVARCHAR(200)         NULL,

    RucEntidad                  VARCHAR(20)           NULL,
    IdEntidad                   NVARCHAR(300)         NULL,
    NombreEntidadOriginal       NVARCHAR(1000)        NULL,
    NombreEntidadEstandar       NVARCHAR(1000)        NULL,

    TituloLicitacion            NVARCHAR(2000)        NULL,
    DescripcionLicitacion       NVARCHAR(MAX)         NULL,

    MetodoContratacion          NVARCHAR(500)         NULL,
    DetalleMetodoContratacion   NVARCHAR(1000)        NULL,
    CategoriaPrincipal          NVARCHAR(500)         NULL,

    MontoLicitacion             DECIMAL(28,4)         NULL,
    MonedaLicitacion            NVARCHAR(50)          NULL,
    NombreMonedaLicitacion      NVARCHAR(200)         NULL,

    TipoRelacionRegistro        NVARCHAR(50)          NULL,
    AnioArchivo                 INT                   NULL,
    MesArchivo                  INT                   NULL,
    ArchivoOrigen               NVARCHAR(500)         NULL,

    CONSTRAINT PK_FactContrato
        PRIMARY KEY (ContratoKey),

    CONSTRAINT UQ_FactContrato_Clave
        UNIQUE (ClaveContrato),

    CONSTRAINT FK_FactContrato_Entidad
        FOREIGN KEY (EntidadKey)
        REFERENCES dw.DimEntidad(EntidadKey),

    CONSTRAINT FK_FactContrato_Ubicacion
        FOREIGN KEY (UbicacionEntidadKey)
        REFERENCES dw.DimUbicacion(UbicacionKey),

    CONSTRAINT FK_FactContrato_Moneda
        FOREIGN KEY (MonedaKey)
        REFERENCES dw.DimMoneda(MonedaKey),

    CONSTRAINT FK_FactContrato_Categoria
        FOREIGN KEY (CategoriaKey)
        REFERENCES dw.DimCategoria(CategoriaKey),

    CONSTRAINT FK_FactContrato_Metodo
        FOREIGN KEY (MetodoKey)
        REFERENCES dw.DimMetodoContratacion(MetodoKey),

    CONSTRAINT FK_FactContrato_FechaFirma
        FOREIGN KEY (FechaFirmaKey)
        REFERENCES dw.DimFecha(FechaKey),

    CONSTRAINT FK_FactContrato_FechaInicio
        FOREIGN KEY (FechaInicioKey)
        REFERENCES dw.DimFecha(FechaKey),

    CONSTRAINT FK_FactContrato_FechaFin
        FOREIGN KEY (FechaFinKey)
        REFERENCES dw.DimFecha(FechaKey)
);

INSERT INTO dw.FactContrato (
    ClaveContrato,
    ContratoHash,
    HashFila,
    Ocid,
    IdEntrega,
    IdContrato,
    IdAdjudicacion,
    IdLicitacion,
    EntidadKey,
    UbicacionEntidadKey,
    MonedaKey,
    CategoriaKey,
    MetodoKey,
    FechaFirmaKey,
    FechaInicioKey,
    FechaFinKey,
    TituloContrato,
    DescripcionContrato,
    FechaFirma,
    FechaInicio,
    FechaFin,
    FechaFinImplementacion,
    DuracionDias,
    MontoContrato,
    MontoFinal,
    MonedaContrato,
    NombreMonedaContrato,
    MonedaFinal,
    NombreMonedaFinal,
    RucEntidad,
    IdEntidad,
    NombreEntidadOriginal,
    NombreEntidadEstandar,
    TituloLicitacion,
    DescripcionLicitacion,
    MetodoContratacion,
    DetalleMetodoContratacion,
    CategoriaPrincipal,
    MontoLicitacion,
    MonedaLicitacion,
    NombreMonedaLicitacion,
    TipoRelacionRegistro,
    AnioArchivo,
    MesArchivo,
    ArchivoOrigen
)
SELECT
    CONVERT(NVARCHAR(450), c.CLAVE_CONTRATO),
    CONVERT(CHAR(32), c.CONTRATO_HASH),
    CONVERT(CHAR(32), c.HASH_FILA),
    CONVERT(NVARCHAR(200), c.OCID),
    CONVERT(NVARCHAR(300), c.ID_ENTREGA),
    CONVERT(NVARCHAR(300), c.ID_CONTRATO),
    CONVERT(NVARCHAR(300), c.ID_ADJUDICACION),
    CONVERT(NVARCHAR(300), c.ID_LICITACION),

    e.EntidadKey,
    u.UbicacionKey,
    moneda.MonedaKey,
    categoria.CategoriaKey,
    metodo.MetodoKey,

    fechaFirma.FechaKey,
    fechaInicio.FechaKey,
    fechaFin.FechaKey,

    CONVERT(NVARCHAR(2000), c.TITULO_CONTRATO),
    CONVERT(NVARCHAR(MAX), c.DESCRIPCION_CONTRATO),

    TRY_CONVERT(DATETIME2, c.FECHA_FIRMA),
    TRY_CONVERT(DATETIME2, c.FECHA_INICIO),
    TRY_CONVERT(DATETIME2, c.FECHA_FIN),
    TRY_CONVERT(
        DATETIME2,
        c.FECHA_FIN_IMPLEMENTACION
    ),

    TRY_CONVERT(
        DECIMAL(18,2),
        c.DURACION_DIAS
    ),

    TRY_CONVERT(
        DECIMAL(28,4),
        c.MONTO_CONTRATO
    ),

    TRY_CONVERT(
        DECIMAL(28,4),
        c.MONTO_FINAL
    ),

    CONVERT(
        NVARCHAR(50),
        c.MONEDA_CONTRATO
    ),

    CONVERT(
        NVARCHAR(200),
        c.NOMBRE_MONEDA_CONTRATO
    ),

    CONVERT(
        NVARCHAR(50),
        c.MONEDA_FINAL
    ),

    CONVERT(
        NVARCHAR(200),
        c.NOMBRE_MONEDA_FINAL
    ),

    CONVERT(
        VARCHAR(20),
        c.RUC_ENTIDAD
    ),

    CONVERT(
        NVARCHAR(300),
        c.ID_ENTIDAD
    ),

    CONVERT(
        NVARCHAR(1000),
        c.NOMBRE_ENTIDAD_ORIGINAL
    ),

    CONVERT(
        NVARCHAR(1000),
        c.NOMBRE_ENTIDAD_ESTANDAR
    ),

    CONVERT(
        NVARCHAR(2000),
        c.TITULO_LICITACION
    ),

    CONVERT(
        NVARCHAR(MAX),
        c.DESCRIPCION_LICITACION
    ),

    CONVERT(
        NVARCHAR(500),
        c.METODO_CONTRATACION
    ),

    CONVERT(
        NVARCHAR(1000),
        c.DETALLE_METODO_CONTRATACION
    ),

    CONVERT(
        NVARCHAR(500),
        c.CATEGORIA_PRINCIPAL
    ),

    TRY_CONVERT(
        DECIMAL(28,4),
        c.MONTO_LICITACION
    ),

    CONVERT(
        NVARCHAR(50),
        c.MONEDA_LICITACION
    ),

    CONVERT(
        NVARCHAR(200),
        c.NOMBRE_MONEDA_LICITACION
    ),

    CONVERT(
        NVARCHAR(50),
        c.TIPO_RELACION_REGISTRO
    ),

    TRY_CONVERT(
        INT,
        c.ANIO_ARCHIVO
    ),

    TRY_CONVERT(
        INT,
        c.MES_ARCHIVO
    ),

    CONVERT(
        NVARCHAR(500),
        c.ARCHIVO_ORIGEN
    )

FROM curated.FACT_CONTRATO_ACTUAL AS c

LEFT JOIN dw.DimEntidad AS e
    ON e.ClaveOrganizacion =
       CONVERT(
           NVARCHAR(450),
           c.CLAVE_ENTIDAD
       )

LEFT JOIN dw.DimUbicacion AS u
    ON u.ClaveUbicacion =
       CONVERT(
           NVARCHAR(100),
           c.CLAVE_UBICACION_ENTIDAD
       )

LEFT JOIN dw.DimMoneda AS moneda
    ON moneda.ClaveMoneda =
       CONVERT(
           NVARCHAR(100),
           c.CLAVE_MONEDA
       )

LEFT JOIN dw.DimCategoria AS categoria
    ON categoria.ClaveCategoria =
       CONVERT(
           NVARCHAR(100),
           c.CLAVE_CATEGORIA
       )

LEFT JOIN dw.DimMetodoContratacion AS metodo
    ON metodo.ClaveMetodo =
       CONVERT(
           NVARCHAR(100),
           c.CLAVE_METODO
       )

LEFT JOIN dw.DimFecha AS fechaFirma
    ON fechaFirma.FechaKey =
       TRY_CONVERT(
           INT,
           c.FECHA_FIRMA_KEY
       )

LEFT JOIN dw.DimFecha AS fechaInicio
    ON fechaInicio.FechaKey =
       TRY_CONVERT(
           INT,
           c.FECHA_INICIO_KEY
       )

LEFT JOIN dw.DimFecha AS fechaFin
    ON fechaFin.FechaKey =
       TRY_CONVERT(
           INT,
           c.FECHA_FIN_KEY
       )

WHERE c.CLAVE_CONTRATO IS NOT NULL;


/* ========================================================
   PUENTE CONTRATO-PROVEEDOR
   ======================================================== */

CREATE TABLE dw.BridgeContratoProveedor (
    ContratoKey                BIGINT        NOT NULL,
    ProveedorKey               BIGINT        NOT NULL,
    IdProveedorOriginal        NVARCHAR(300) NULL,
    RucProveedor               VARCHAR(20)   NULL,
    NombreProveedorOriginal    NVARCHAR(1000) NULL,
    NombreProveedorEstandar    NVARCHAR(1000) NULL,

    CONSTRAINT PK_BridgeContratoProveedor
        PRIMARY KEY (
            ContratoKey,
            ProveedorKey
        ),

    CONSTRAINT FK_Bridge_Contrato
        FOREIGN KEY (ContratoKey)
        REFERENCES dw.FactContrato(ContratoKey),

    CONSTRAINT FK_Bridge_Proveedor
        FOREIGN KEY (ProveedorKey)
        REFERENCES dw.DimProveedor(ProveedorKey)
);

INSERT INTO dw.BridgeContratoProveedor (
    ContratoKey,
    ProveedorKey,
    IdProveedorOriginal,
    RucProveedor,
    NombreProveedorOriginal,
    NombreProveedorEstandar
)
SELECT DISTINCT
    contrato.ContratoKey,
    proveedor.ProveedorKey,

    CONVERT(
        NVARCHAR(300),
        bridge.ID_PROVEEDOR
    ),

    CONVERT(
        VARCHAR(20),
        bridge.RUC_PROVEEDOR
    ),

    CONVERT(
        NVARCHAR(1000),
        bridge.NOMBRE_PROVEEDOR_ORIGINAL
    ),

    CONVERT(
        NVARCHAR(1000),
        bridge.NOMBRE_PROVEEDOR_ESTANDAR
    )

FROM curated.BRIDGE_CONTRATO_PROVEEDOR AS bridge

INNER JOIN dw.FactContrato AS contrato
    ON contrato.ClaveContrato =
       CONVERT(
           NVARCHAR(450),
           bridge.CLAVE_CONTRATO
       )

INNER JOIN dw.DimProveedor AS proveedor
    ON proveedor.ClaveOrganizacion =
       CONVERT(
           NVARCHAR(450),
           bridge.CLAVE_PROVEEDOR
       );


/* ========================================================
   ÍNDICES PARA CONSULTAS Y POWER BI
   ======================================================== */

CREATE INDEX IX_FactContrato_EntidadKey
ON dw.FactContrato(EntidadKey);

CREATE INDEX IX_FactContrato_FechaFirmaKey
ON dw.FactContrato(FechaFirmaKey);

CREATE INDEX IX_FactContrato_AnioMes
ON dw.FactContrato(
    AnioArchivo,
    MesArchivo
);

CREATE INDEX IX_Bridge_ProveedorKey
ON dw.BridgeContratoProveedor(
    ProveedorKey
);

CREATE INDEX IX_DimEntidad_UbicacionKey
ON dw.DimEntidad(
    UbicacionKey
);

CREATE INDEX IX_DimProveedor_UbicacionKey
ON dw.DimProveedor(
    UbicacionKey
);
"""


def crear_modelo_estrella(
    engine
) -> None:

    print("\n========================================")
    print("CREANDO MODELO ESTRELLA")
    print("========================================")

    with engine.begin() as conexion:
        conexion.exec_driver_sql(
            SQL_CREAR_MODELO
        )

    print(
        "Modelo estrella creado correctamente."
    )


# ==========================================================
# TABLA DE CONTROL
# ==========================================================

def crear_tabla_control(
    engine
) -> None:

    consulta = """
    IF OBJECT_ID(
        'control.EjecucionCarga',
        'U'
    ) IS NULL
    BEGIN
        CREATE TABLE control.EjecucionCarga (
            EjecucionId        BIGINT IDENTITY(1,1)
                               NOT NULL,
            FechaInicio        DATETIME2 NOT NULL,
            FechaFin           DATETIME2 NULL,
            BaseDatos          NVARCHAR(200) NOT NULL,
            Estado             NVARCHAR(50) NOT NULL,
            Mensaje            NVARCHAR(MAX) NULL,

            CONSTRAINT PK_EjecucionCarga
                PRIMARY KEY (EjecucionId)
        );
    END;
    """

    with engine.begin() as conexion:
        conexion.execute(
            text(consulta)
        )


def registrar_ejecucion(
    engine,
    fecha_inicio: datetime,
    fecha_fin: datetime,
    estado: str,
    mensaje: str
) -> None:

    consulta = text(
        """
        INSERT INTO control.EjecucionCarga (
            FechaInicio,
            FechaFin,
            BaseDatos,
            Estado,
            Mensaje
        )
        VALUES (
            :fecha_inicio,
            :fecha_fin,
            :base_datos,
            :estado,
            :mensaje
        );
        """
    )

    with engine.begin() as conexion:
        conexion.execute(
            consulta,
            {
                "fecha_inicio": fecha_inicio,
                "fecha_fin": fecha_fin,
                "base_datos": BASE_DATOS,
                "estado": estado,
                "mensaje": mensaje,
            }
        )


# ==========================================================
# VALIDACIONES
# ==========================================================

def ejecutar_consulta_dataframe(
    engine,
    consulta: str
) -> pd.DataFrame:

    return pd.read_sql_query(
        text(consulta),
        engine
    )


def validar_modelo(
    engine,
    resumen_carga: pd.DataFrame
) -> None:

    print("\n========================================")
    print("VALIDANDO MODELO")
    print("========================================")

    consulta_indicadores = """
    SELECT
        'Contratos curated' AS INDICADOR,
        COUNT_BIG(*) AS VALOR
    FROM curated.FACT_CONTRATO_ACTUAL

    UNION ALL

    SELECT
        'Contratos DW',
        COUNT_BIG(*)
    FROM dw.FactContrato

    UNION ALL

    SELECT
        'Contratos duplicados por clave',
        COUNT_BIG(*)
    FROM (
        SELECT ClaveContrato
        FROM dw.FactContrato
        GROUP BY ClaveContrato
        HAVING COUNT(*) > 1
    ) AS duplicados

    UNION ALL

    SELECT
        'Relaciones proveedor curated',
        COUNT_BIG(*)
    FROM curated.BRIDGE_CONTRATO_PROVEEDOR

    UNION ALL

    SELECT
        'Relaciones proveedor DW',
        COUNT_BIG(*)
    FROM dw.BridgeContratoProveedor

    UNION ALL

    SELECT
        'Entidades',
        COUNT_BIG(*)
    FROM dw.DimEntidad

    UNION ALL

    SELECT
        'Proveedores',
        COUNT_BIG(*)
    FROM dw.DimProveedor

    UNION ALL

    SELECT
        'Ubicaciones',
        COUNT_BIG(*)
    FROM dw.DimUbicacion

    UNION ALL

    SELECT
        'Monedas',
        COUNT_BIG(*)
    FROM dw.DimMoneda

    UNION ALL

    SELECT
        'Categorías',
        COUNT_BIG(*)
    FROM dw.DimCategoria

    UNION ALL

    SELECT
        'Métodos de contratación',
        COUNT_BIG(*)
    FROM dw.DimMetodoContratacion

    UNION ALL

    SELECT
        'Fechas',
        COUNT_BIG(*)
    FROM dw.DimFecha

    UNION ALL

    SELECT
        'Contratos sin entidad',
        COUNT_BIG(*)
    FROM dw.FactContrato
    WHERE EntidadKey IS NULL

    UNION ALL

    SELECT
        'Contratos sin proveedor',
        COUNT_BIG(*)
    FROM dw.FactContrato AS contrato
    LEFT JOIN dw.BridgeContratoProveedor AS puente
        ON puente.ContratoKey =
           contrato.ContratoKey
    WHERE puente.ContratoKey IS NULL

    UNION ALL

    SELECT
        'Entidades sin ubicación',
        COUNT_BIG(*)
    FROM dw.DimEntidad
    WHERE UbicacionKey IS NULL

    UNION ALL

    SELECT
        'Proveedores sin ubicación',
        COUNT_BIG(*)
    FROM dw.DimProveedor
    WHERE UbicacionKey IS NULL;
    """

    consulta_huerfanos = """
    SELECT
        'FACT_ENTIDAD' AS VALIDACION,
        COUNT_BIG(*) AS CANTIDAD
    FROM dw.FactContrato AS f
    LEFT JOIN dw.DimEntidad AS d
        ON d.EntidadKey = f.EntidadKey
    WHERE
        f.EntidadKey IS NOT NULL
        AND d.EntidadKey IS NULL

    UNION ALL

    SELECT
        'FACT_MONEDA',
        COUNT_BIG(*)
    FROM dw.FactContrato AS f
    LEFT JOIN dw.DimMoneda AS d
        ON d.MonedaKey = f.MonedaKey
    WHERE
        f.MonedaKey IS NOT NULL
        AND d.MonedaKey IS NULL

    UNION ALL

    SELECT
        'FACT_CATEGORIA',
        COUNT_BIG(*)
    FROM dw.FactContrato AS f
    LEFT JOIN dw.DimCategoria AS d
        ON d.CategoriaKey = f.CategoriaKey
    WHERE
        f.CategoriaKey IS NOT NULL
        AND d.CategoriaKey IS NULL

    UNION ALL

    SELECT
        'FACT_METODO',
        COUNT_BIG(*)
    FROM dw.FactContrato AS f
    LEFT JOIN dw.DimMetodoContratacion AS d
        ON d.MetodoKey = f.MetodoKey
    WHERE
        f.MetodoKey IS NOT NULL
        AND d.MetodoKey IS NULL

    UNION ALL

    SELECT
        'BRIDGE_CONTRATO',
        COUNT_BIG(*)
    FROM dw.BridgeContratoProveedor AS b
    LEFT JOIN dw.FactContrato AS f
        ON f.ContratoKey = b.ContratoKey
    WHERE f.ContratoKey IS NULL

    UNION ALL

    SELECT
        'BRIDGE_PROVEEDOR',
        COUNT_BIG(*)
    FROM dw.BridgeContratoProveedor AS b
    LEFT JOIN dw.DimProveedor AS p
        ON p.ProveedorKey = b.ProveedorKey
    WHERE p.ProveedorKey IS NULL;
    """

    consulta_muestra_contratos = f"""
    SELECT TOP ({CANTIDAD_MUESTRAS_ALEATORIAS})
        curated.CLAVE_CONTRATO
            AS CLAVE_CURATED,

        fact.ClaveContrato
            AS CLAVE_DW,

        curated.OCID,
        curated.ID_CONTRATO,

        curated.CLAVE_ENTIDAD
            AS CLAVE_ENTIDAD_CURATED,

        entidad.ClaveOrganizacion
            AS CLAVE_ENTIDAD_DW,

        curated.NOMBRE_ENTIDAD_ESTANDAR
            AS ENTIDAD_CURATED,

        entidad.NombreEstandar
            AS ENTIDAD_DW,

        curated.MONTO_CONTRATO
            AS MONTO_CURATED,

        fact.MontoContrato
            AS MONTO_DW,

        moneda.CodigoMoneda,
        categoria.CategoriaPrincipal,
        metodo.MetodoContratacion,

        ubicacion.Departamento,
        ubicacion.Region,
        ubicacion.Localidad,

        CASE
            WHEN fact.ClaveContrato =
                 curated.CLAVE_CONTRATO
            THEN 'OK'
            ELSE 'REVISAR'
        END AS VALIDACION_CLAVE,

        CASE
            WHEN
                entidad.ClaveOrganizacion =
                curated.CLAVE_ENTIDAD

                OR (
                    entidad.ClaveOrganizacion IS NULL
                    AND curated.CLAVE_ENTIDAD IS NULL
                )
            THEN 'OK'
            ELSE 'REVISAR'
        END AS VALIDACION_ENTIDAD,

        CASE
            WHEN
                fact.MontoContrato =
                TRY_CONVERT(
                    DECIMAL(28,4),
                    curated.MONTO_CONTRATO
                )

                OR (
                    fact.MontoContrato IS NULL
                    AND curated.MONTO_CONTRATO IS NULL
                )
            THEN 'OK'
            ELSE 'REVISAR'
        END AS VALIDACION_MONTO

    FROM curated.FACT_CONTRATO_ACTUAL
         AS curated

    INNER JOIN dw.FactContrato AS fact
        ON fact.ClaveContrato =
           curated.CLAVE_CONTRATO

    LEFT JOIN dw.DimEntidad AS entidad
        ON entidad.EntidadKey =
           fact.EntidadKey

    LEFT JOIN dw.DimMoneda AS moneda
        ON moneda.MonedaKey =
           fact.MonedaKey

    LEFT JOIN dw.DimCategoria AS categoria
        ON categoria.CategoriaKey =
           fact.CategoriaKey

    LEFT JOIN dw.DimMetodoContratacion AS metodo
        ON metodo.MetodoKey =
           fact.MetodoKey

    LEFT JOIN dw.DimUbicacion AS ubicacion
        ON ubicacion.UbicacionKey =
           fact.UbicacionEntidadKey

    ORDER BY NEWID();
    """

    consulta_muestra_proveedores = f"""
    SELECT TOP ({CANTIDAD_MUESTRAS_ALEATORIAS})
        bridgeCurated.CLAVE_CONTRATO
            AS CONTRATO_CURATED,

        fact.ClaveContrato
            AS CONTRATO_DW,

        bridgeCurated.CLAVE_PROVEEDOR
            AS PROVEEDOR_CURATED,

        proveedor.ClaveOrganizacion
            AS PROVEEDOR_DW,

        bridgeCurated.RUC_PROVEEDOR,
        proveedor.Ruc,

        bridgeCurated.NOMBRE_PROVEEDOR_ESTANDAR
            AS NOMBRE_CURATED,

        proveedor.NombreEstandar
            AS NOMBRE_DW,

        ubicacion.Departamento,
        ubicacion.Region,
        ubicacion.Localidad,

        CASE
            WHEN
                bridgeCurated.CLAVE_CONTRATO =
                fact.ClaveContrato
            THEN 'OK'
            ELSE 'REVISAR'
        END AS VALIDACION_CONTRATO,

        CASE
            WHEN
                bridgeCurated.CLAVE_PROVEEDOR =
                proveedor.ClaveOrganizacion
            THEN 'OK'
            ELSE 'REVISAR'
        END AS VALIDACION_PROVEEDOR

    FROM curated.BRIDGE_CONTRATO_PROVEEDOR
         AS bridgeCurated

    INNER JOIN dw.FactContrato AS fact
        ON fact.ClaveContrato =
           bridgeCurated.CLAVE_CONTRATO

    INNER JOIN dw.BridgeContratoProveedor AS bridgeDw
        ON bridgeDw.ContratoKey =
           fact.ContratoKey

    INNER JOIN dw.DimProveedor AS proveedor
        ON proveedor.ProveedorKey =
           bridgeDw.ProveedorKey

        AND proveedor.ClaveOrganizacion =
            bridgeCurated.CLAVE_PROVEEDOR

    LEFT JOIN dw.DimUbicacion AS ubicacion
        ON ubicacion.UbicacionKey =
           proveedor.UbicacionKey

    ORDER BY NEWID();
    """

    consulta_muestra_geografia = f"""
    SELECT TOP ({CANTIDAD_MUESTRAS_ALEATORIAS})
        entidad.ClaveOrganizacion,
        entidad.Ruc,
        entidad.NombreEstandar,
        entidad.Direccion,

        entidad.Departamento
            AS DEPARTAMENTO_ORGANIZACION,

        ubicacion.Departamento
            AS DEPARTAMENTO_DIMENSION,

        entidad.Region
            AS REGION_ORGANIZACION,

        ubicacion.Region
            AS REGION_DIMENSION,

        entidad.Localidad
            AS LOCALIDAD_ORGANIZACION,

        ubicacion.Localidad
            AS LOCALIDAD_DIMENSION,

        CASE
            WHEN entidad.UbicacionKey IS NOT NULL
                 AND ubicacion.UbicacionKey IS NOT NULL
            THEN 'OK'
            ELSE 'REVISAR'
        END AS VALIDACION_UBICACION

    FROM dw.DimEntidad AS entidad

    INNER JOIN dw.DimUbicacion AS ubicacion
        ON ubicacion.UbicacionKey =
           entidad.UbicacionKey

    ORDER BY NEWID();
    """

    df_indicadores = ejecutar_consulta_dataframe(
        engine,
        consulta_indicadores
    )

    df_huerfanos = ejecutar_consulta_dataframe(
        engine,
        consulta_huerfanos
    )

    df_muestra_contratos = ejecutar_consulta_dataframe(
        engine,
        consulta_muestra_contratos
    )

    df_muestra_proveedores = ejecutar_consulta_dataframe(
        engine,
        consulta_muestra_proveedores
    )

    df_muestra_geografia = ejecutar_consulta_dataframe(
        engine,
        consulta_muestra_geografia
    )

    validacion_huerfanos_ok = bool(
        df_huerfanos["CANTIDAD"]
        .fillna(0)
        .eq(0)
        .all()
    )

    validacion_contratos_ok = bool(
        df_muestra_contratos[
            [
                "VALIDACION_CLAVE",
                "VALIDACION_ENTIDAD",
                "VALIDACION_MONTO",
            ]
        ]
        .eq("OK")
        .all()
        .all()
    )

    validacion_proveedores_ok = bool(
        df_muestra_proveedores[
            [
                "VALIDACION_CONTRATO",
                "VALIDACION_PROVEEDOR",
            ]
        ]
        .eq("OK")
        .all()
        .all()
    )

    resumen_validacion = pd.DataFrame([
        {
            "PRUEBA": "Integridad referencial",
            "RESULTADO": (
                "OK"
                if validacion_huerfanos_ok
                else "REVISAR"
            ),
        },
        {
            "PRUEBA": "Muestra aleatoria de contratos",
            "RESULTADO": (
                "OK"
                if validacion_contratos_ok
                else "REVISAR"
            ),
        },
        {
            "PRUEBA": "Muestra aleatoria de proveedores",
            "RESULTADO": (
                "OK"
                if validacion_proveedores_ok
                else "REVISAR"
            ),
        },
    ])

    with pd.ExcelWriter(
        ARCHIVO_VALIDACION,
        engine="openpyxl"
    ) as writer:

        resumen_validacion.to_excel(
            writer,
            sheet_name="RESUMEN_VALIDACION",
            index=False
        )

        resumen_carga.to_excel(
            writer,
            sheet_name="CARGA_CURATED",
            index=False
        )

        df_indicadores.to_excel(
            writer,
            sheet_name="INDICADORES",
            index=False
        )

        df_huerfanos.to_excel(
            writer,
            sheet_name="HUERFANOS",
            index=False
        )

        df_muestra_contratos.to_excel(
            writer,
            sheet_name="MUESTRA_CONTRATOS",
            index=False
        )

        df_muestra_proveedores.to_excel(
            writer,
            sheet_name="MUESTRA_PROVEEDORES",
            index=False
        )

        df_muestra_geografia.to_excel(
            writer,
            sheet_name="MUESTRA_GEOGRAFIA",
            index=False
        )

    print("\nResultado de las validaciones:")
    print(
        resumen_validacion.to_string(
            index=False
        )
    )

    print("\nRegistros huérfanos:")
    print(
        df_huerfanos.to_string(
            index=False
        )
    )

    print("\nReporte:")
    print(ARCHIVO_VALIDACION)


# ==========================================================
# MOSTRAR CANTIDADES FINALES
# ==========================================================

def mostrar_cantidades(
    engine
) -> None:

    consulta = """
    SELECT
        'dw.FactContrato' AS TABLA,
        COUNT_BIG(*) AS FILAS
    FROM dw.FactContrato

    UNION ALL

    SELECT
        'dw.BridgeContratoProveedor',
        COUNT_BIG(*)
    FROM dw.BridgeContratoProveedor

    UNION ALL

    SELECT
        'dw.DimEntidad',
        COUNT_BIG(*)
    FROM dw.DimEntidad

    UNION ALL

    SELECT
        'dw.DimProveedor',
        COUNT_BIG(*)
    FROM dw.DimProveedor

    UNION ALL

    SELECT
        'dw.DimUbicacion',
        COUNT_BIG(*)
    FROM dw.DimUbicacion

    UNION ALL

    SELECT
        'dw.DimMoneda',
        COUNT_BIG(*)
    FROM dw.DimMoneda

    UNION ALL

    SELECT
        'dw.DimCategoria',
        COUNT_BIG(*)
    FROM dw.DimCategoria

    UNION ALL

    SELECT
        'dw.DimMetodoContratacion',
        COUNT_BIG(*)
    FROM dw.DimMetodoContratacion

    UNION ALL

    SELECT
        'dw.DimFecha',
        COUNT_BIG(*)
    FROM dw.DimFecha;
    """

    df = ejecutar_consulta_dataframe(
        engine,
        consulta
    )

    print("\n========================================")
    print("CANTIDADES FINALES")
    print("========================================")

    print(
        df.to_string(
            index=False
        )
    )


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def main() -> None:
    fecha_inicio = datetime.now()

    engine = None

    try:
        print("========================================")
        print("CARGA LOCAL OECE A SQL SERVER")
        print("========================================")

        validar_archivos_curated()

        driver = detectar_driver_odbc()

        print(f"Driver detectado: {driver}")
        print(f"Servidor: {SERVIDOR_SQL}")
        print(f"Base de datos: {BASE_DATOS}")

        crear_base_datos(driver)

        engine = crear_engine_sqlalchemy(
            driver,
            BASE_DATOS
        )

        # Verificar la conexión.
        with engine.connect() as conexion:
            resultado = conexion.execute(
                text("SELECT DB_NAME()")
            ).scalar_one()

            print(
                f"Conectado a: {resultado}"
            )

        crear_esquemas(engine)
        crear_tabla_control(engine)

        resumen_carga = cargar_tablas_curated(
            engine
        )

        crear_modelo_estrella(
            engine
        )

        mostrar_cantidades(
            engine
        )

        validar_modelo(
            engine,
            resumen_carga
        )

        fecha_fin = datetime.now()

        registrar_ejecucion(
            engine,
            fecha_inicio,
            fecha_fin,
            "OK",
            "Carga inicial local y validaciones terminadas."
        )

        print("\n========================================")
        print("PROCESO TERMINADO CORRECTAMENTE")
        print("========================================")
        print(f"Duración total: {fecha_fin - fecha_inicio}")
        print(f"Servidor: {SERVIDOR_SQL}")
        print(f"Base de datos: {BASE_DATOS}")

    except Exception as error:
        fecha_fin = datetime.now()

        print("\n========================================")
        print("ERROR EN LA CARGA")
        print("========================================")
        print(str(error))

        traceback.print_exc()

        if engine is not None:
            try:
                registrar_ejecucion(
                    engine,
                    fecha_inicio,
                    fecha_fin,
                    "ERROR",
                    str(error)
                )
            except Exception:
                pass

        sys.exit(1)

    finally:
        if engine is not None:
            engine.dispose()


if __name__ == "__main__":
    main()
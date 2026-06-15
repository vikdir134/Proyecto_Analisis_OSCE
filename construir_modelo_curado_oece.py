from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import duckdb
import pandas as pd


# ==========================================================
# RUTAS
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

CARPETA_STAGING = (
    BASE_DIR
    / "data"
    / "staging"
    / "oece"
)

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

CARPETA_WORK = (
    BASE_DIR
    / "data"
    / "work"
)

CARPETA_CURATED.mkdir(
    parents=True,
    exist_ok=True
)

CARPETA_CALIDAD.mkdir(
    parents=True,
    exist_ok=True
)

CARPETA_WORK.mkdir(
    parents=True,
    exist_ok=True
)

BASE_DUCKDB = (
    CARPETA_WORK
    / "transformacion_oece.duckdb"
)

ARCHIVO_CONTROL = (
    CARPETA_CALIDAD
    / "CONTROL_MODELO_CURADO_OECE.xlsx"
)

ARCHIVO_DICCIONARIO = (
    CARPETA_CURATED
    / "DICCIONARIO_MODELO_OECE.xlsx"
)

ARCHIVO_HANDOFF = (
    CARPETA_CURATED
    / "ENTREGA_PARA_BASE_DE_DATOS.txt"
)


# ==========================================================
# TABLAS QUE SE EXPORTARÁN
# ==========================================================

TABLAS_EXPORTACION = {
    "fact_contrato_actual":
        "FACT_CONTRATO_ACTUAL",

    "bridge_contrato_proveedor":
        "BRIDGE_CONTRATO_PROVEEDOR",

    "dim_organizacion":
        "DIM_ORGANIZACION",

    "dim_entidad":
        "DIM_ENTIDAD",

    "dim_proveedor":
        "DIM_PROVEEDOR",

    "dim_ubicacion":
        "DIM_UBICACION",

    "dim_moneda":
        "DIM_MONEDA",

    "dim_categoria":
        "DIM_CATEGORIA",

    "dim_metodo":
        "DIM_METODO_CONTRATACION",

    "dim_fecha":
        "DIM_FECHA",
}


# ==========================================================
# FUNCIONES DE ARCHIVOS
# ==========================================================

def obtener_archivos(
    carpeta: Path
) -> list[Path]:

    if not carpeta.exists():
        return []

    return sorted(
        archivo
        for archivo in carpeta.rglob("*.parquet")
        if archivo.is_file()
    )


def lista_sql_archivos(
    archivos: list[Path]
) -> str:

    if not archivos:
        raise RuntimeError(
            "La lista de archivos Parquet está vacía."
        )

    rutas = []

    for archivo in archivos:
        ruta = str(
            archivo.resolve()
        ).replace("\\", "/")

        ruta = ruta.replace(
            "'",
            "''"
        )

        rutas.append(
            f"'{ruta}'"
        )

    return "[" + ", ".join(rutas) + "]"


def ruta_sql(
    ruta: Path
) -> str:

    return (
        str(ruta.resolve())
        .replace("\\", "/")
        .replace("'", "''")
    )


# ==========================================================
# CREAR VISTAS STAGING
# ==========================================================

def crear_vistas_staging(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    carpetas = {
        "stg_contratos":
            CARPETA_STAGING / "contratos",

        "stg_registros":
            CARPETA_STAGING / "registros",

        "stg_proveedores":
            CARPETA_STAGING / "proveedores",

        "stg_partes":
            CARPETA_STAGING / "partes",
    }

    for vista, carpeta in carpetas.items():
        archivos = obtener_archivos(
            carpeta
        )

        if not archivos:
            raise RuntimeError(
                f"No existen archivos para {vista}: "
                f"{carpeta}"
            )

        lista_archivos = lista_sql_archivos(
            archivos
        )

        conexion.execute(
            f"""
            CREATE OR REPLACE VIEW {vista} AS
            SELECT *
            FROM read_parquet(
                {lista_archivos},
                union_by_name = true
            );
            """
        )

        cantidad = conexion.execute(
            f"""
            SELECT COUNT(*)
            FROM {vista};
            """
        ).fetchone()[0]

        print(
            f"{vista}: {cantidad:,} filas"
        )


# ==========================================================
# CONTRATOS ACTUALES
# ==========================================================

def crear_contratos_actuales(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE contrato_actual AS

        WITH clasificados AS (
            SELECT
                *,

                ROW_NUMBER() OVER (
                    PARTITION BY CLAVE_CONTRATO

                    ORDER BY
                        ANIO_ARCHIVO DESC,
                        MES_ARCHIVO DESC,

                        COALESCE(
                            FECHA_FIRMA,
                            FECHA_INICIO,
                            FECHA_FIN,
                            DATE '1900-01-01'
                        ) DESC,

                        ID_ENTREGA DESC
                ) AS RN

            FROM stg_contratos

            WHERE CLAVE_CONTRATO IS NOT NULL
        )

        SELECT
            * EXCLUDE (RN)

        FROM clasificados

        WHERE RN = 1;
        """
    )


# ==========================================================
# REGISTROS ACTUALES
# ==========================================================

def crear_registros_actuales(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE registro_actual_clave AS

        WITH evaluados AS (
            SELECT
                *,

                (
                    CASE
                        WHEN ID_ENTIDAD_CONTRATANTE IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN NOMBRE_ENTIDAD_ESTANDAR IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN ID_COMPRADOR IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN ID_LICITACION IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN METODO_CONTRATACION IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN CATEGORIA_PRINCIPAL IS NOT NULL
                        THEN 1 ELSE 0
                    END
                ) AS COMPLETITUD,

                ROW_NUMBER() OVER (
                    PARTITION BY
                        OCID,
                        ID_ENTREGA

                    ORDER BY
                        ANIO_ARCHIVO DESC,
                        MES_ARCHIVO DESC,
                        COMPLETITUD DESC,

                        COALESCE(
                            FECHA_ENTREGA,
                            FECHA_PUBLICACION,
                            DATE '1900-01-01'
                        ) DESC
                ) AS RN

            FROM stg_registros

            WHERE
                OCID IS NOT NULL
                AND ID_ENTREGA IS NOT NULL
        )

        SELECT
            * EXCLUDE (
                COMPLETITUD,
                RN
            )

        FROM evaluados

        WHERE RN = 1;
        """
    )

    # Respaldo por OCID cuando no coincide el ID de entrega.
    conexion.execute(
        """
        CREATE OR REPLACE TABLE registro_actual_ocid AS

        WITH evaluados AS (
            SELECT
                *,

                (
                    CASE
                        WHEN ID_ENTIDAD_CONTRATANTE IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN NOMBRE_ENTIDAD_ESTANDAR IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN ID_COMPRADOR IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN ID_LICITACION IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN METODO_CONTRATACION IS NOT NULL
                        THEN 1 ELSE 0
                    END
                    +
                    CASE
                        WHEN CATEGORIA_PRINCIPAL IS NOT NULL
                        THEN 1 ELSE 0
                    END
                ) AS COMPLETITUD,

                ROW_NUMBER() OVER (
                    PARTITION BY OCID

                    ORDER BY
                        ANIO_ARCHIVO DESC,
                        MES_ARCHIVO DESC,
                        COMPLETITUD DESC,

                        COALESCE(
                            FECHA_ENTREGA,
                            FECHA_PUBLICACION,
                            DATE '1900-01-01'
                        ) DESC
                ) AS RN

            FROM stg_registros

            WHERE OCID IS NOT NULL
        )

        SELECT
            * EXCLUDE (
                COMPLETITUD,
                RN
            )

        FROM evaluados

        WHERE RN = 1;
        """
    )


# ==========================================================
# PROVEEDORES ACTUALES
# ==========================================================

def crear_proveedores_actuales(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE proveedor_actual AS

        WITH clasificados AS (
            SELECT
                *,

                ROW_NUMBER() OVER (
                    PARTITION BY
                        OCID,
                        ID_ADJUDICACION,
                        ID_PROVEEDOR

                    ORDER BY
                        ANIO_ARCHIVO DESC,
                        MES_ARCHIVO DESC,
                        ID_ENTREGA DESC
                ) AS RN

            FROM stg_proveedores

            WHERE
                OCID IS NOT NULL
                AND ID_ADJUDICACION IS NOT NULL
                AND ID_PROVEEDOR IS NOT NULL
        )

        SELECT
            * EXCLUDE (RN),

            CASE
                WHEN RUC_PROVEEDOR IS NOT NULL
                THEN
                    'RUC|' || RUC_PROVEEDOR

                ELSE
                    'ID|' || ID_PROVEEDOR
            END AS CLAVE_PROVEEDOR

        FROM clasificados

        WHERE RN = 1;
        """
    )


# ==========================================================
# CANDIDATOS DE ORGANIZACIONES
# ==========================================================

def crear_candidatos_organizacion(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE organizacion_candidato AS

        -- Organizaciones procedentes de partes involucradas
        SELECT
            CLAVE_ORGANIZACION,

            ID_ORGANIZACION
                AS ID_ORGANIZACION_ORIGINAL,

            RUC_ORGANIZACION
                AS RUC,

            CAST(
                RUC_VALIDO AS BOOLEAN
            ) AS RUC_VALIDO,

            NOMBRE_ORGANIZACION_ORIGINAL
                AS NOMBRE_ORIGINAL,

            NOMBRE_ORGANIZACION_ESTANDAR
                AS NOMBRE_ESTANDAR,

            DIRECCION_ESTANDAR
                AS DIRECCION,

            LOCALIDAD_ESTANDAR
                AS LOCALIDAD,

            REGION_ESTANDAR
                AS REGION,

            DEPARTAMENTO_ESTANDAR
                AS DEPARTAMENTO,

            PAIS_ESTANDAR
                AS PAIS,

            CAST(
                ES_COMPRADOR_O_ENTIDAD AS INTEGER
            ) AS ES_ENTIDAD,

            CAST(
                ES_PROVEEDOR AS INTEGER
            ) AS ES_PROVEEDOR,

            ANIO_ARCHIVO,
            MES_ARCHIVO,

            3 AS PRIORIDAD_FUENTE,

            (
                CASE WHEN RUC_ORGANIZACION IS NOT NULL
                    THEN 1 ELSE 0 END
                +
                CASE WHEN NOMBRE_ORGANIZACION_ESTANDAR IS NOT NULL
                    THEN 1 ELSE 0 END
                +
                CASE WHEN DEPARTAMENTO_ESTANDAR IS NOT NULL
                    THEN 1 ELSE 0 END
                +
                CASE WHEN REGION_ESTANDAR IS NOT NULL
                    THEN 1 ELSE 0 END
                +
                CASE WHEN LOCALIDAD_ESTANDAR IS NOT NULL
                    THEN 1 ELSE 0 END
            ) AS COMPLETITUD

        FROM stg_partes

        WHERE CLAVE_ORGANIZACION IS NOT NULL


        UNION ALL


        -- Entidades contratantes
        SELECT
            CASE
                WHEN RUC_ENTIDAD_CONTRATANTE IS NOT NULL
                THEN
                    'RUC|' || RUC_ENTIDAD_CONTRATANTE

                ELSE
                    'ID|' || ID_ENTIDAD_CONTRATANTE
            END AS CLAVE_ORGANIZACION,

            ID_ENTIDAD_CONTRATANTE
                AS ID_ORGANIZACION_ORIGINAL,

            RUC_ENTIDAD_CONTRATANTE
                AS RUC,

            CAST(
                RUC_ENTIDAD_VALIDO AS BOOLEAN
            ) AS RUC_VALIDO,

            NOMBRE_ENTIDAD_CONTRATANTE_ORIGINAL
                AS NOMBRE_ORIGINAL,

            NOMBRE_ENTIDAD_ESTANDAR
                AS NOMBRE_ESTANDAR,

            CAST(NULL AS VARCHAR)
                AS DIRECCION,

            CAST(NULL AS VARCHAR)
                AS LOCALIDAD,

            CAST(NULL AS VARCHAR)
                AS REGION,

            CAST(NULL AS VARCHAR)
                AS DEPARTAMENTO,

            CAST(NULL AS VARCHAR)
                AS PAIS,

            1 AS ES_ENTIDAD,
            0 AS ES_PROVEEDOR,

            ANIO_ARCHIVO,
            MES_ARCHIVO,

            2 AS PRIORIDAD_FUENTE,

            (
                CASE WHEN RUC_ENTIDAD_CONTRATANTE IS NOT NULL
                    THEN 1 ELSE 0 END
                +
                CASE WHEN NOMBRE_ENTIDAD_ESTANDAR IS NOT NULL
                    THEN 1 ELSE 0 END
            ) AS COMPLETITUD

        FROM stg_registros

        WHERE ID_ENTIDAD_CONTRATANTE IS NOT NULL


        UNION ALL


        -- Compradores usados como respaldo
        SELECT
            CASE
                WHEN RUC_COMPRADOR IS NOT NULL
                THEN
                    'RUC|' || RUC_COMPRADOR

                ELSE
                    'ID|' || ID_COMPRADOR
            END AS CLAVE_ORGANIZACION,

            ID_COMPRADOR
                AS ID_ORGANIZACION_ORIGINAL,

            RUC_COMPRADOR
                AS RUC,

            CAST(
                RUC_COMPRADOR_VALIDO AS BOOLEAN
            ) AS RUC_VALIDO,

            NOMBRE_COMPRADOR_ORIGINAL
                AS NOMBRE_ORIGINAL,

            NOMBRE_COMPRADOR_ESTANDAR
                AS NOMBRE_ESTANDAR,

            CAST(NULL AS VARCHAR)
                AS DIRECCION,

            CAST(NULL AS VARCHAR)
                AS LOCALIDAD,

            CAST(NULL AS VARCHAR)
                AS REGION,

            CAST(NULL AS VARCHAR)
                AS DEPARTAMENTO,

            CAST(NULL AS VARCHAR)
                AS PAIS,

            1 AS ES_ENTIDAD,
            0 AS ES_PROVEEDOR,

            ANIO_ARCHIVO,
            MES_ARCHIVO,

            2 AS PRIORIDAD_FUENTE,

            (
                CASE WHEN RUC_COMPRADOR IS NOT NULL
                    THEN 1 ELSE 0 END
                +
                CASE WHEN NOMBRE_COMPRADOR_ESTANDAR IS NOT NULL
                    THEN 1 ELSE 0 END
            ) AS COMPLETITUD

        FROM stg_registros

        WHERE ID_COMPRADOR IS NOT NULL


        UNION ALL


        -- Proveedores procedentes de adjudicaciones
        SELECT
            CASE
                WHEN RUC_PROVEEDOR IS NOT NULL
                THEN
                    'RUC|' || RUC_PROVEEDOR

                ELSE
                    'ID|' || ID_PROVEEDOR
            END AS CLAVE_ORGANIZACION,

            ID_PROVEEDOR
                AS ID_ORGANIZACION_ORIGINAL,

            RUC_PROVEEDOR
                AS RUC,

            CAST(
                RUC_PROVEEDOR_VALIDO AS BOOLEAN
            ) AS RUC_VALIDO,

            NOMBRE_PROVEEDOR_ORIGINAL
                AS NOMBRE_ORIGINAL,

            NOMBRE_PROVEEDOR_ESTANDAR
                AS NOMBRE_ESTANDAR,

            CAST(NULL AS VARCHAR)
                AS DIRECCION,

            CAST(NULL AS VARCHAR)
                AS LOCALIDAD,

            CAST(NULL AS VARCHAR)
                AS REGION,

            CAST(NULL AS VARCHAR)
                AS DEPARTAMENTO,

            CAST(NULL AS VARCHAR)
                AS PAIS,

            0 AS ES_ENTIDAD,
            1 AS ES_PROVEEDOR,

            ANIO_ARCHIVO,
            MES_ARCHIVO,

            1 AS PRIORIDAD_FUENTE,

            (
                CASE WHEN RUC_PROVEEDOR IS NOT NULL
                    THEN 1 ELSE 0 END
                +
                CASE WHEN NOMBRE_PROVEEDOR_ESTANDAR IS NOT NULL
                    THEN 1 ELSE 0 END
            ) AS COMPLETITUD

        FROM stg_proveedores

        WHERE ID_PROVEEDOR IS NOT NULL;
        """
    )


# ==========================================================
# DIMENSIÓN ORGANIZACIÓN
# ==========================================================

def crear_dim_organizacion(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_organizacion AS

        WITH candidatos AS (
            SELECT
                *,

                (
                    ANIO_ARCHIVO * 1000000
                    +
                    MES_ARCHIVO * 10000
                    +
                    PRIORIDAD_FUENTE * 100
                    +
                    COMPLETITUD
                ) AS ORDEN_VERSION

            FROM organizacion_candidato

            WHERE CLAVE_ORGANIZACION IS NOT NULL
        ),

        agregados AS (
            SELECT
                CLAVE_ORGANIZACION,

                ARG_MAX(
                    ID_ORGANIZACION_ORIGINAL,
                    ORDEN_VERSION
                ) AS ID_ORGANIZACION_ORIGINAL,

                ARG_MAX(
                    RUC,
                    ORDEN_VERSION
                ) AS RUC,

                ARG_MAX(
                    RUC_VALIDO,
                    ORDEN_VERSION
                ) AS RUC_VALIDO,

                ARG_MAX(
                    NOMBRE_ORIGINAL,
                    ORDEN_VERSION
                ) AS NOMBRE_ORIGINAL,

                ARG_MAX(
                    NOMBRE_ESTANDAR,
                    ORDEN_VERSION
                ) AS NOMBRE_ESTANDAR,

                ARG_MAX(
                    DIRECCION,
                    ORDEN_VERSION
                ) AS DIRECCION,

                ARG_MAX(
                    LOCALIDAD,
                    ORDEN_VERSION
                ) AS LOCALIDAD,

                ARG_MAX(
                    REGION,
                    ORDEN_VERSION
                ) AS REGION,

                ARG_MAX(
                    DEPARTAMENTO,
                    ORDEN_VERSION
                ) AS DEPARTAMENTO,

                ARG_MAX(
                    PAIS,
                    ORDEN_VERSION
                ) AS PAIS,

                MAX(ES_ENTIDAD)
                    AS ES_ENTIDAD,

                MAX(ES_PROVEEDOR)
                    AS ES_PROVEEDOR,

                MIN(
                    ANIO_ARCHIVO * 100
                    + MES_ARCHIVO
                ) AS PERIODO_PRIMERA_APARICION,

                MAX(
                    ANIO_ARCHIVO * 100
                    + MES_ARCHIVO
                ) AS PERIODO_ULTIMA_APARICION

            FROM candidatos

            GROUP BY CLAVE_ORGANIZACION
        )

        SELECT
            MD5(CLAVE_ORGANIZACION)
                AS ORGANIZACION_HASH,

            CLAVE_ORGANIZACION,

            ID_ORGANIZACION_ORIGINAL,
            RUC,
            RUC_VALIDO,
            NOMBRE_ORIGINAL,
            NOMBRE_ESTANDAR,

            CASE
                WHEN ES_ENTIDAD = 1
                     AND ES_PROVEEDOR = 1
                    THEN 'ENTIDAD_Y_PROVEEDOR'

                WHEN ES_ENTIDAD = 1
                    THEN 'ENTIDAD'

                WHEN ES_PROVEEDOR = 1
                    THEN 'PROVEEDOR'

                ELSE 'OTRO'
            END AS TIPO_ORGANIZACION,

            DIRECCION,
            LOCALIDAD,
            REGION,
            DEPARTAMENTO,
            PAIS,

            CASE
                WHEN
                    PAIS IS NOT NULL
                    OR DEPARTAMENTO IS NOT NULL
                    OR REGION IS NOT NULL
                    OR LOCALIDAD IS NOT NULL

                THEN MD5(
                    CONCAT_WS(
                        '|',
                        COALESCE(PAIS, ''),
                        COALESCE(DEPARTAMENTO, ''),
                        COALESCE(REGION, ''),
                        COALESCE(LOCALIDAD, '')
                    )
                )

                ELSE NULL
            END AS CLAVE_UBICACION,

            ES_ENTIDAD,
            ES_PROVEEDOR,

            PERIODO_PRIMERA_APARICION,
            PERIODO_ULTIMA_APARICION

        FROM agregados;
        """
    )


# ==========================================================
# DIMENSIÓN UBICACIÓN
# ==========================================================

def crear_dim_ubicacion(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_ubicacion AS

        SELECT DISTINCT
            CLAVE_UBICACION,
            PAIS,
            DEPARTAMENTO,
            REGION,
            LOCALIDAD

        FROM dim_organizacion

        WHERE CLAVE_UBICACION IS NOT NULL;
        """
    )


# ==========================================================
# TABLA DE HECHOS
# ==========================================================

def crear_fact_contrato(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE fact_contrato_actual AS

        WITH enriquecido AS (
            SELECT
                c.*,

                CASE
                    WHEN re.OCID IS NOT NULL
                        THEN 'OCID_Y_ENTREGA'

                    WHEN ro.OCID IS NOT NULL
                        THEN 'RESPALDO_OCID'

                    ELSE 'SIN_REGISTRO'
                END AS TIPO_RELACION_REGISTRO,

                COALESCE(
                    re.RUC_ENTIDAD_CONTRATANTE,
                    re.RUC_COMPRADOR,
                    ro.RUC_ENTIDAD_CONTRATANTE,
                    ro.RUC_COMPRADOR
                ) AS RUC_ENTIDAD,

                COALESCE(
                    re.ID_ENTIDAD_CONTRATANTE,
                    re.ID_COMPRADOR,
                    ro.ID_ENTIDAD_CONTRATANTE,
                    ro.ID_COMPRADOR
                ) AS ID_ENTIDAD,

                COALESCE(
                    re.NOMBRE_ENTIDAD_CONTRATANTE_ORIGINAL,
                    re.NOMBRE_COMPRADOR_ORIGINAL,
                    ro.NOMBRE_ENTIDAD_CONTRATANTE_ORIGINAL,
                    ro.NOMBRE_COMPRADOR_ORIGINAL
                ) AS NOMBRE_ENTIDAD_ORIGINAL,

                COALESCE(
                    re.NOMBRE_ENTIDAD_ESTANDAR,
                    re.NOMBRE_COMPRADOR_ESTANDAR,
                    ro.NOMBRE_ENTIDAD_ESTANDAR,
                    ro.NOMBRE_COMPRADOR_ESTANDAR
                ) AS NOMBRE_ENTIDAD_ESTANDAR,

                COALESCE(
                    re.ID_LICITACION,
                    ro.ID_LICITACION
                ) AS ID_LICITACION,

                COALESCE(
                    re.TITULO_LICITACION,
                    ro.TITULO_LICITACION
                ) AS TITULO_LICITACION,

                COALESCE(
                    re.DESCRIPCION_LICITACION,
                    ro.DESCRIPCION_LICITACION
                ) AS DESCRIPCION_LICITACION,

                COALESCE(
                    re.METODO_CONTRATACION,
                    ro.METODO_CONTRATACION
                ) AS METODO_CONTRATACION,

                COALESCE(
                    re.DETALLE_METODO_CONTRATACION,
                    ro.DETALLE_METODO_CONTRATACION
                ) AS DETALLE_METODO_CONTRATACION,

                COALESCE(
                    re.CATEGORIA_PRINCIPAL,
                    ro.CATEGORIA_PRINCIPAL
                ) AS CATEGORIA_PRINCIPAL,

                COALESCE(
                    re.MONTO_LICITACION,
                    ro.MONTO_LICITACION
                ) AS MONTO_LICITACION,

                COALESCE(
                    re.MONEDA_LICITACION,
                    ro.MONEDA_LICITACION
                ) AS MONEDA_LICITACION,

                COALESCE(
                    re.NOMBRE_MONEDA_LICITACION,
                    ro.NOMBRE_MONEDA_LICITACION
                ) AS NOMBRE_MONEDA_LICITACION

            FROM contrato_actual AS c

            LEFT JOIN registro_actual_clave AS re
                ON c.OCID = re.OCID
                AND c.ID_ENTREGA = re.ID_ENTREGA

            LEFT JOIN registro_actual_ocid AS ro
                ON c.OCID = ro.OCID
        ),

        llaves AS (
            SELECT
                *,

                CASE
                    WHEN RUC_ENTIDAD IS NOT NULL
                        THEN 'RUC|' || RUC_ENTIDAD

                    WHEN ID_ENTIDAD IS NOT NULL
                        THEN 'ID|' || ID_ENTIDAD

                    ELSE NULL
                END AS CLAVE_ENTIDAD,

                CASE
                    WHEN COALESCE(
                        MONEDA_CONTRATO,
                        NOMBRE_MONEDA_CONTRATO
                    ) IS NOT NULL

                    THEN
                        'MONEDA|'
                        ||
                        UPPER(
                            TRIM(
                                COALESCE(
                                    MONEDA_CONTRATO,
                                    NOMBRE_MONEDA_CONTRATO
                                )
                            )
                        )

                    ELSE 'MONEDA|SIN_MONEDA'
                END AS CLAVE_MONEDA,

                CASE
                    WHEN CATEGORIA_PRINCIPAL IS NOT NULL
                    THEN
                        MD5(
                            UPPER(
                                TRIM(CATEGORIA_PRINCIPAL)
                            )
                        )

                    ELSE NULL
                END AS CLAVE_CATEGORIA,

                CASE
                    WHEN
                        METODO_CONTRATACION IS NOT NULL
                        OR DETALLE_METODO_CONTRATACION IS NOT NULL

                    THEN
                        MD5(
                            CONCAT_WS(
                                '|',
                                UPPER(
                                    TRIM(
                                        COALESCE(
                                            METODO_CONTRATACION,
                                            ''
                                        )
                                    )
                                ),
                                UPPER(
                                    TRIM(
                                        COALESCE(
                                            DETALLE_METODO_CONTRATACION,
                                            ''
                                        )
                                    )
                                )
                            )
                        )

                    ELSE NULL
                END AS CLAVE_METODO

            FROM enriquecido
        )

        SELECT
            MD5(CLAVE_CONTRATO)
                AS CONTRATO_HASH,

            CLAVE_CONTRATO,

            OCID,
            ID_ENTREGA,
            ID_CONTRATO,
            ID_ADJUDICACION,

            CLAVE_ENTIDAD,

            organizacion.CLAVE_UBICACION
                AS CLAVE_UBICACION_ENTIDAD,

            CLAVE_MONEDA,
            CLAVE_CATEGORIA,
            CLAVE_METODO,

            CASE
                WHEN FECHA_FIRMA IS NOT NULL
                THEN CAST(
                    STRFTIME(
                        FECHA_FIRMA,
                        '%Y%m%d'
                    ) AS INTEGER
                )
                ELSE NULL
            END AS FECHA_FIRMA_KEY,

            CASE
                WHEN FECHA_INICIO IS NOT NULL
                THEN CAST(
                    STRFTIME(
                        FECHA_INICIO,
                        '%Y%m%d'
                    ) AS INTEGER
                )
                ELSE NULL
            END AS FECHA_INICIO_KEY,

            CASE
                WHEN FECHA_FIN IS NOT NULL
                THEN CAST(
                    STRFTIME(
                        FECHA_FIN,
                        '%Y%m%d'
                    ) AS INTEGER
                )
                ELSE NULL
            END AS FECHA_FIN_KEY,

            TITULO_CONTRATO,
            DESCRIPCION_CONTRATO,

            FECHA_FIRMA,
            FECHA_INICIO,
            FECHA_FIN,
            FECHA_FIN_IMPLEMENTACION,

            DURACION_DIAS,
            MONTO_CONTRATO,
            MONEDA_CONTRATO,
            NOMBRE_MONEDA_CONTRATO,

            MONTO_FINAL,
            MONEDA_FINAL,
            NOMBRE_MONEDA_FINAL,

            RUC_ENTIDAD,
            ID_ENTIDAD,
            NOMBRE_ENTIDAD_ORIGINAL,
            NOMBRE_ENTIDAD_ESTANDAR,

            ID_LICITACION,
            TITULO_LICITACION,
            DESCRIPCION_LICITACION,

            METODO_CONTRATACION,
            DETALLE_METODO_CONTRATACION,
            CATEGORIA_PRINCIPAL,

            MONTO_LICITACION,
            MONEDA_LICITACION,
            NOMBRE_MONEDA_LICITACION,

            TIPO_RELACION_REGISTRO,

            ANIO_ARCHIVO,
            MES_ARCHIVO,
            ARCHIVO_ORIGEN,

            MD5(
                CONCAT_WS(
                    '|',
                    COALESCE(CLAVE_CONTRATO, ''),
                    COALESCE(CLAVE_ENTIDAD, ''),
                    COALESCE(CAST(MONTO_CONTRATO AS VARCHAR), ''),
                    COALESCE(MONEDA_CONTRATO, ''),
                    COALESCE(CAST(FECHA_FIRMA AS VARCHAR), ''),
                    COALESCE(CAST(FECHA_INICIO AS VARCHAR), ''),
                    COALESCE(CAST(FECHA_FIN AS VARCHAR), ''),
                    COALESCE(DESCRIPCION_CONTRATO, '')
                )
            ) AS HASH_FILA

        FROM llaves

        LEFT JOIN dim_organizacion AS organizacion
            ON llaves.CLAVE_ENTIDAD
            = organizacion.CLAVE_ORGANIZACION;
        """
    )


# ==========================================================
# PUENTE CONTRATO-PROVEEDOR
# ==========================================================

def crear_bridge(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE bridge_contrato_proveedor AS

        SELECT DISTINCT
            c.CLAVE_CONTRATO,
            p.CLAVE_PROVEEDOR,

            p.ID_PROVEEDOR,
            p.RUC_PROVEEDOR,

            p.NOMBRE_PROVEEDOR_ORIGINAL,
            p.NOMBRE_PROVEEDOR_ESTANDAR

        FROM contrato_actual AS c

        INNER JOIN proveedor_actual AS p
            ON c.OCID = p.OCID
            AND c.ID_ADJUDICACION
                = p.ID_ADJUDICACION

        WHERE p.CLAVE_PROVEEDOR IS NOT NULL;
        """
    )


# ==========================================================
# DIMENSIONES POR ROL
# ==========================================================

def crear_dimensiones_rol(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_entidad AS

        SELECT DISTINCT
            organizacion.*

        FROM dim_organizacion AS organizacion

        INNER JOIN (
            SELECT DISTINCT
                CLAVE_ENTIDAD

            FROM fact_contrato_actual

            WHERE CLAVE_ENTIDAD IS NOT NULL
        ) AS entidades_usadas

            ON organizacion.CLAVE_ORGANIZACION
            = entidades_usadas.CLAVE_ENTIDAD;
        """
    )

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_proveedor AS

        SELECT DISTINCT
            organizacion.*

        FROM dim_organizacion AS organizacion

        INNER JOIN (
            SELECT DISTINCT
                CLAVE_PROVEEDOR

            FROM bridge_contrato_proveedor

            WHERE CLAVE_PROVEEDOR IS NOT NULL
        ) AS proveedores_usados

            ON organizacion.CLAVE_ORGANIZACION
            = proveedores_usados.CLAVE_PROVEEDOR;
        """
    )


# ==========================================================
# DIMENSIONES COMPLEMENTARIAS
# ==========================================================

def crear_dimensiones_complementarias(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_moneda AS

        SELECT DISTINCT
            CLAVE_MONEDA,

            MONEDA_CONTRATO
                AS CODIGO_MONEDA,

            NOMBRE_MONEDA_CONTRATO
                AS NOMBRE_MONEDA

        FROM fact_contrato_actual;
        """
    )

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_categoria AS

        SELECT DISTINCT
            CLAVE_CATEGORIA,
            CATEGORIA_PRINCIPAL

        FROM fact_contrato_actual

        WHERE CLAVE_CATEGORIA IS NOT NULL;
        """
    )

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_metodo AS

        SELECT DISTINCT
            CLAVE_METODO,
            METODO_CONTRATACION,
            DETALLE_METODO_CONTRATACION

        FROM fact_contrato_actual

        WHERE CLAVE_METODO IS NOT NULL;
        """
    )


# ==========================================================
# DIMENSIÓN FECHA
# ==========================================================

def crear_dim_fecha(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    resultado = conexion.execute(
        """
        SELECT
            MIN(FECHA) AS FECHA_MINIMA,
            MAX(FECHA) AS FECHA_MAXIMA

        FROM (
            SELECT FECHA_FIRMA AS FECHA
            FROM fact_contrato_actual

            UNION ALL

            SELECT FECHA_INICIO AS FECHA
            FROM fact_contrato_actual

            UNION ALL

            SELECT FECHA_FIN AS FECHA
            FROM fact_contrato_actual
        )

        WHERE
            FECHA BETWEEN
                DATE '2000-01-01'
                AND DATE '2100-12-31';
        """
    ).fetchone()

    fecha_minima = resultado[0]
    fecha_maxima = resultado[1]

    if fecha_minima is None:
        fecha_minima = pd.Timestamp(
            "2022-01-01"
        )

    if fecha_maxima is None:
        fecha_maxima = pd.Timestamp(
            "2030-12-31"
        )

    fecha_minima = pd.Timestamp(
        fecha_minima
    )

    fecha_maxima = pd.Timestamp(
        fecha_maxima
    )

    fechas = pd.date_range(
        fecha_minima,
        fecha_maxima,
        freq="D"
    )

    meses = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre",
    }

    dias = {
        0: "Lunes",
        1: "Martes",
        2: "Miércoles",
        3: "Jueves",
        4: "Viernes",
        5: "Sábado",
        6: "Domingo",
    }

    df_fecha = pd.DataFrame({
        "FECHA": fechas
    })

    df_fecha["FECHA_KEY"] = (
        df_fecha["FECHA"]
        .dt.strftime("%Y%m%d")
        .astype(int)
    )

    df_fecha["ANIO"] = (
        df_fecha["FECHA"].dt.year
    )

    df_fecha["SEMESTRE"] = (
        (df_fecha["FECHA"].dt.month - 1)
        // 6
        + 1
    )

    df_fecha["TRIMESTRE"] = (
        df_fecha["FECHA"].dt.quarter
    )

    df_fecha["MES_NUMERO"] = (
        df_fecha["FECHA"].dt.month
    )

    df_fecha["MES_NOMBRE"] = (
        df_fecha["MES_NUMERO"]
        .map(meses)
    )

    df_fecha["ANIO_MES"] = (
        df_fecha["FECHA"]
        .dt.strftime("%Y-%m")
    )

    df_fecha["DIA"] = (
        df_fecha["FECHA"].dt.day
    )

    df_fecha["DIA_SEMANA_NUMERO"] = (
        df_fecha["FECHA"].dt.dayofweek + 1
    )

    df_fecha["DIA_SEMANA_NOMBRE"] = (
        df_fecha["FECHA"]
        .dt.dayofweek
        .map(dias)
    )

    df_fecha["ES_FIN_SEMANA"] = (
        df_fecha["FECHA"]
        .dt.dayofweek
        .isin([5, 6])
    )

    conexion.register(
        "df_fecha_temporal",
        df_fecha
    )

    conexion.execute(
        """
        CREATE OR REPLACE TABLE dim_fecha AS
        SELECT *
        FROM df_fecha_temporal;
        """
    )

    conexion.unregister(
        "df_fecha_temporal"
    )


# ==========================================================
# EXPORTACIÓN
# ==========================================================

def exportar_tabla(
    conexion: duckdb.DuckDBPyConnection,
    tabla: str,
    nombre: str,
    generar_csv: bool
) -> None:

    salida_parquet = (
        CARPETA_CURATED
        / f"{nombre}.parquet"
    )

    if salida_parquet.exists():
        salida_parquet.unlink()

    conexion.execute(
        f"""
        COPY {tabla}
        TO '{ruta_sql(salida_parquet)}'
        (
            FORMAT PARQUET,
            COMPRESSION ZSTD
        );
        """
    )

    if generar_csv:
        salida_csv = (
            CARPETA_CURATED
            / f"{nombre}.csv"
        )

        if salida_csv.exists():
            salida_csv.unlink()

        conexion.execute(
            f"""
            COPY {tabla}
            TO '{ruta_sql(salida_csv)}'
            (
                FORMAT CSV,
                HEADER TRUE,
                DELIMITER '|',
                QUOTE '"',
                ESCAPE '"'
            );
            """
        )


def exportar_modelo(
    conexion: duckdb.DuckDBPyConnection,
    generar_csv: bool
) -> None:

    print("\nExportando tablas curadas...")

    for tabla, nombre in TABLAS_EXPORTACION.items():
        exportar_tabla(
            conexion,
            tabla,
            nombre,
            generar_csv
        )

        cantidad = conexion.execute(
            f"""
            SELECT COUNT(*)
            FROM {tabla};
            """
        ).fetchone()[0]

        print(
            f"{nombre}: {cantidad:,} filas"
        )


# ==========================================================
# CONTROL DE CALIDAD
# ==========================================================

def crear_control_calidad(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    indicadores = []

    consultas = {
        "Filas staging contratos":
            "SELECT COUNT(*) FROM stg_contratos",

        "Contratos actuales":
            "SELECT COUNT(*) FROM fact_contrato_actual",

        "Claves de contrato únicas":
            """
            SELECT COUNT(DISTINCT CLAVE_CONTRATO)
            FROM fact_contrato_actual
            """,

        "Contratos sin entidad":
            """
            SELECT COUNT(*)
            FROM fact_contrato_actual
            WHERE CLAVE_ENTIDAD IS NULL
            """,

        "Contratos con respaldo por OCID":
            """
            SELECT COUNT(*)
            FROM fact_contrato_actual
            WHERE TIPO_RELACION_REGISTRO = 'RESPALDO_OCID'
            """,

        "Contratos sin registro relacionado":
            """
            SELECT COUNT(*)
            FROM fact_contrato_actual
            WHERE TIPO_RELACION_REGISTRO = 'SIN_REGISTRO'
            """,

        "Relaciones contrato-proveedor":
            """
            SELECT COUNT(*)
            FROM bridge_contrato_proveedor
            """,

        "Contratos con proveedor":
            """
            SELECT COUNT(
                DISTINCT CLAVE_CONTRATO
            )
            FROM bridge_contrato_proveedor
            """,

        "Organizaciones únicas":
            """
            SELECT COUNT(*)
            FROM dim_organizacion
            """,

        "Entidades únicas":
            """
            SELECT COUNT(*)
            FROM dim_entidad
            """,

        "Proveedores únicos":
            """
            SELECT COUNT(*)
            FROM dim_proveedor
            """,

        "Organizaciones con ubicación":
            """
            SELECT COUNT(*)
            FROM dim_organizacion
            WHERE CLAVE_UBICACION IS NOT NULL
            """,

        "Entidades con ubicación":
            """
            SELECT COUNT(*)
            FROM dim_entidad
            WHERE CLAVE_UBICACION IS NOT NULL
            """,

        "Proveedores con ubicación":
            """
            SELECT COUNT(*)
            FROM dim_proveedor
            WHERE CLAVE_UBICACION IS NOT NULL
            """,

        "Ubicaciones únicas":
            """
            SELECT COUNT(*)
            FROM dim_ubicacion
            """,

        "Monedas":
            """
            SELECT COUNT(*)
            FROM dim_moneda
            """,

        "Categorías":
            """
            SELECT COUNT(*)
            FROM dim_categoria
            """,

        "Métodos de contratación":
            """
            SELECT COUNT(*)
            FROM dim_metodo
            """,
    }

    for indicador, consulta in consultas.items():
        valor = conexion.execute(
            consulta
        ).fetchone()[0]

        indicadores.append({
            "INDICADOR": indicador,
            "VALOR": valor
        })

    df_indicadores = pd.DataFrame(
        indicadores
    )

    df_sin_entidad = conexion.execute(
        """
        SELECT
            CLAVE_CONTRATO,
            OCID,
            ID_CONTRATO,
            TITULO_CONTRATO,
            MONTO_CONTRATO,
            ANIO_ARCHIVO,
            MES_ARCHIVO,
            TIPO_RELACION_REGISTRO

        FROM fact_contrato_actual

        WHERE CLAVE_ENTIDAD IS NULL

        LIMIT 1000;
        """
    ).df()

    df_sin_proveedor = conexion.execute(
        """
        SELECT
            f.CLAVE_CONTRATO,
            f.OCID,
            f.ID_CONTRATO,
            f.ID_ADJUDICACION,
            f.TITULO_CONTRATO,
            f.MONTO_CONTRATO,
            f.NOMBRE_ENTIDAD_ESTANDAR

        FROM fact_contrato_actual AS f

        LEFT JOIN bridge_contrato_proveedor AS b
            ON f.CLAVE_CONTRATO
            = b.CLAVE_CONTRATO

        WHERE b.CLAVE_CONTRATO IS NULL

        LIMIT 1000;
        """
    ).df()

    df_geo_entidades = conexion.execute(
        """
        SELECT
            DEPARTAMENTO,
            REGION,
            LOCALIDAD,
            COUNT(*) AS CANTIDAD_ENTIDADES

        FROM dim_entidad

        WHERE CLAVE_UBICACION IS NOT NULL

        GROUP BY
            DEPARTAMENTO,
            REGION,
            LOCALIDAD

        ORDER BY
            CANTIDAD_ENTIDADES DESC;
        """
    ).df()

    df_geo_proveedores = conexion.execute(
        """
        SELECT
            DEPARTAMENTO,
            REGION,
            LOCALIDAD,
            COUNT(*) AS CANTIDAD_PROVEEDORES

        FROM dim_proveedor

        WHERE CLAVE_UBICACION IS NOT NULL

        GROUP BY
            DEPARTAMENTO,
            REGION,
            LOCALIDAD

        ORDER BY
            CANTIDAD_PROVEEDORES DESC;
        """
    ).df()

    df_resumen_anual = conexion.execute(
        """
        SELECT
            COALESCE(
                EXTRACT(
                    YEAR FROM FECHA_FIRMA
                ),
                ANIO_ARCHIVO
            ) AS ANIO,

            COUNT(*) AS CONTRATOS,

            SUM(
                COALESCE(
                    MONTO_CONTRATO,
                    0
                )
            ) AS MONTO_TOTAL

        FROM fact_contrato_actual

        GROUP BY ANIO

        ORDER BY ANIO;
        """
    ).df()

    with pd.ExcelWriter(
        ARCHIVO_CONTROL,
        engine="openpyxl"
    ) as writer:

        df_indicadores.to_excel(
            writer,
            sheet_name="INDICADORES",
            index=False
        )

        df_resumen_anual.to_excel(
            writer,
            sheet_name="RESUMEN_ANUAL",
            index=False
        )

        df_sin_entidad.to_excel(
            writer,
            sheet_name="MUESTRA_SIN_ENTIDAD",
            index=False
        )

        df_sin_proveedor.to_excel(
            writer,
            sheet_name="MUESTRA_SIN_PROVEEDOR",
            index=False
        )

        df_geo_entidades.to_excel(
            writer,
            sheet_name="GEO_ENTIDADES",
            index=False
        )

        df_geo_proveedores.to_excel(
            writer,
            sheet_name="GEO_PROVEEDORES",
            index=False
        )


# ==========================================================
# DICCIONARIO DE DATOS
# ==========================================================

def crear_diccionario(
    conexion: duckdb.DuckDBPyConnection
) -> None:

    relaciones = pd.DataFrame([
        {
            "TABLA_ORIGEN": "FACT_CONTRATO_ACTUAL",
            "COLUMNA_ORIGEN": "CLAVE_ENTIDAD",
            "TABLA_DESTINO": "DIM_ENTIDAD",
            "COLUMNA_DESTINO": "CLAVE_ORGANIZACION",
            "CARDINALIDAD": "Muchos a uno",
        },
        {
            "TABLA_ORIGEN": "FACT_CONTRATO_ACTUAL",
            "COLUMNA_ORIGEN": "CLAVE_MONEDA",
            "TABLA_DESTINO": "DIM_MONEDA",
            "COLUMNA_DESTINO": "CLAVE_MONEDA",
            "CARDINALIDAD": "Muchos a uno",
        },
        {
            "TABLA_ORIGEN": "FACT_CONTRATO_ACTUAL",
            "COLUMNA_ORIGEN": "CLAVE_CATEGORIA",
            "TABLA_DESTINO": "DIM_CATEGORIA",
            "COLUMNA_DESTINO": "CLAVE_CATEGORIA",
            "CARDINALIDAD": "Muchos a uno",
        },
        {
            "TABLA_ORIGEN": "FACT_CONTRATO_ACTUAL",
            "COLUMNA_ORIGEN": "CLAVE_METODO",
            "TABLA_DESTINO": "DIM_METODO_CONTRATACION",
            "COLUMNA_DESTINO": "CLAVE_METODO",
            "CARDINALIDAD": "Muchos a uno",
        },
        {
            "TABLA_ORIGEN": "FACT_CONTRATO_ACTUAL",
            "COLUMNA_ORIGEN": "FECHA_FIRMA_KEY",
            "TABLA_DESTINO": "DIM_FECHA",
            "COLUMNA_DESTINO": "FECHA_KEY",
            "CARDINALIDAD": "Muchos a uno",
        },
        {
            "TABLA_ORIGEN": "FACT_CONTRATO_ACTUAL",
            "COLUMNA_ORIGEN": "CLAVE_CONTRATO",
            "TABLA_DESTINO": "BRIDGE_CONTRATO_PROVEEDOR",
            "COLUMNA_DESTINO": "CLAVE_CONTRATO",
            "CARDINALIDAD": "Uno a muchos",
        },
        {
            "TABLA_ORIGEN": "BRIDGE_CONTRATO_PROVEEDOR",
            "COLUMNA_ORIGEN": "CLAVE_PROVEEDOR",
            "TABLA_DESTINO": "DIM_PROVEEDOR",
            "COLUMNA_DESTINO": "CLAVE_ORGANIZACION",
            "CARDINALIDAD": "Muchos a uno",
        },
        {
            "TABLA_ORIGEN": "DIM_ENTIDAD",
            "COLUMNA_ORIGEN": "CLAVE_UBICACION",
            "TABLA_DESTINO": "DIM_UBICACION",
            "COLUMNA_DESTINO": "CLAVE_UBICACION",
            "CARDINALIDAD": "Muchos a uno",
        },
        {
            "TABLA_ORIGEN": "DIM_PROVEEDOR",
            "COLUMNA_ORIGEN": "CLAVE_UBICACION",
            "TABLA_DESTINO": "DIM_UBICACION",
            "COLUMNA_DESTINO": "CLAVE_UBICACION",
            "CARDINALIDAD": "Muchos a uno",
        },
    ])

    with pd.ExcelWriter(
        ARCHIVO_DICCIONARIO,
        engine="openpyxl"
    ) as writer:

        relaciones.to_excel(
            writer,
            sheet_name="RELACIONES",
            index=False
        )

        for tabla, nombre in TABLAS_EXPORTACION.items():
            descripcion = conexion.execute(
                f"""
                DESCRIBE {tabla};
                """
            ).df()

            nombre_hoja = nombre[:31]

            descripcion.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False
            )


# ==========================================================
# DOCUMENTO DE ENTREGA PARA BASE DE DATOS
# ==========================================================

def crear_handoff() -> None:

    contenido = """
ENTREGA DE MODELO CURADO OECE
===============================================================================

OBJETIVO
Los archivos de esta carpeta corresponden a la capa curada y deduplicada.
Son los archivos que deben cargarse al modelo analítico en la base de datos.

NO CARGAR DIRECTAMENTE
- Los Excel mensuales originales.
- Los 4 millones de filas de partes involucradas.
- Los archivos staging como tablas finales del modelo estrella.

FORMATO PRINCIPAL
Parquet con compresión ZSTD.

FORMATO ALTERNATIVO
CSV separado por barra vertical | cuando se ejecuta el programa con --csv.

ORDEN SUGERIDO DE CARGA
1. DIM_FECHA
2. DIM_UBICACION
3. DIM_ENTIDAD
4. DIM_PROVEEDOR
5. DIM_MONEDA
6. DIM_CATEGORIA
7. DIM_METODO_CONTRATACION
8. FACT_CONTRATO_ACTUAL
9. BRIDGE_CONTRATO_PROVEEDOR

LLAVES DE NEGOCIO
DIM_ENTIDAD:
    CLAVE_ORGANIZACION

DIM_PROVEEDOR:
    CLAVE_ORGANIZACION

DIM_UBICACION:
    CLAVE_UBICACION

FACT_CONTRATO_ACTUAL:
    CLAVE_CONTRATO

BRIDGE_CONTRATO_PROVEEDOR:
    CLAVE_CONTRATO + CLAVE_PROVEEDOR

IMPORTANTE
La base de datos puede generar llaves sustitutas numéricas mediante IDENTITY,
SEQUENCE o el mecanismo equivalente del motor seleccionado.

La carga debe buscar primero por la llave de negocio:
- Si no existe: INSERT.
- Si existe y HASH_FILA cambió: UPDATE.
- Si existe y HASH_FILA no cambió: no realizar cambios.

ACTUALIZACIÓN INCREMENTAL
Cuando un archivo mensual cambie:
1. Se actualiza su partición staging.
2. Se reconstruye o actualiza la capa curada afectada.
3. Se comparan CLAVE_CONTRATO y HASH_FILA.
4. Se insertan contratos nuevos.
5. Se actualizan contratos modificados.
6. Se actualiza el puente contrato-proveedor.
7. Se registra la ejecución en una tabla de control.

POWER BI
Power BI debe conectarse a las tablas del modelo estrella almacenadas
en la base de datos, no a los Excel mensuales.
""".strip()

    with open(
        ARCHIVO_HANDOFF,
        "w",
        encoding="utf-8"
    ) as archivo:

        archivo.write(contenido)


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def construir_modelo(
    generar_csv: bool
) -> None:

    inicio = datetime.now()

    print("========================================")
    print("CONSTRUCCIÓN DEL MODELO CURADO OECE")
    print("========================================")

    conexion = duckdb.connect(
        str(BASE_DUCKDB)
    )

    conexion.execute(
        "PRAGMA threads = 4;"
    )

    try:
        print("\n1. Creando vistas staging...")
        crear_vistas_staging(
            conexion
        )

        print("\n2. Seleccionando contratos actuales...")
        crear_contratos_actuales(
            conexion
        )

        print("\n3. Seleccionando registros actuales...")
        crear_registros_actuales(
            conexion
        )

        print("\n4. Seleccionando proveedores actuales...")
        crear_proveedores_actuales(
            conexion
        )

        print("\n5. Construyendo organizaciones maestras...")
        crear_candidatos_organizacion(
            conexion
        )

        crear_dim_organizacion(
            conexion
        )

        crear_dim_ubicacion(
            conexion
        )

        print("\n6. Construyendo tabla de hechos...")
        crear_fact_contrato(
            conexion
        )

        print("\n7. Construyendo puente de proveedores...")
        crear_bridge(
            conexion
        )

        print("\n8. Construyendo dimensiones...")
        crear_dimensiones_rol(
            conexion
        )

        crear_dimensiones_complementarias(
            conexion
        )

        crear_dim_fecha(
            conexion
        )

        print("\n9. Exportando modelo...")
        exportar_modelo(
            conexion,
            generar_csv
        )

        print("\n10. Generando controles...")
        crear_control_calidad(
            conexion
        )

        crear_diccionario(
            conexion
        )

        crear_handoff()

    finally:
        conexion.close()

    fin = datetime.now()

    print("\n========================================")
    print("MODELO CURADO TERMINADO")
    print("========================================")
    print(
        f"Duración: {fin - inicio}"
    )
    print(
        f"Carpeta curada: {CARPETA_CURATED}"
    )
    print(
        f"Control de calidad: {ARCHIVO_CONTROL}"
    )
    print(
        f"Diccionario: {ARCHIVO_DICCIONARIO}"
    )
    print(
        f"Entrega para BD: {ARCHIVO_HANDOFF}"
    )


def main() -> None:

    parser = argparse.ArgumentParser(
        description=(
            "Construye el modelo curado y "
            "deduplicado de contratos OECE."
        )
    )

    parser.add_argument(
        "--csv",
        action="store_true",
        help=(
            "Genera también archivos CSV "
            "separados por |."
        )
    )

    argumentos = parser.parse_args()

    construir_modelo(
        generar_csv=argumentos.csv
    )


if __name__ == "__main__":
    main()
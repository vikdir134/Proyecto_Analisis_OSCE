from __future__ import annotations

import argparse
import hashlib
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


# ==========================================================
# RUTAS
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

CARPETA_RAW = (
    BASE_DIR
    / "data"
    / "raw"
    / "oece"
)

CARPETA_STAGING = (
    BASE_DIR
    / "data"
    / "staging"
    / "oece"
)

CARPETA_PARTES = (
    CARPETA_STAGING
    / "partes"
)

CARPETA_CALIDAD = (
    BASE_DIR
    / "data"
    / "quality"
    / "oece"
)

CARPETA_PARTES.mkdir(
    parents=True,
    exist_ok=True
)

CARPETA_CALIDAD.mkdir(
    parents=True,
    exist_ok=True
)

ARCHIVO_MANIFEST = (
    CARPETA_STAGING
    / "manifest_partes.json"
)

ARCHIVO_CONTROL = (
    CARPETA_CALIDAD
    / "CONTROL_STAGING_PARTES_OECE.xlsx"
)

ARCHIVO_CATALOGO_GEO = (
    CARPETA_STAGING
    / "catalogo_geografico_oece.parquet"
)


# ==========================================================
# MAPEO DE COLUMNAS
# ==========================================================

MAPA_PARTES = {
    "Open Contracting ID":
        "OCID",

    "Entrega compilada:ID de Entrega":
        "ID_ENTREGA",

    "Entrega compilada:Partes involucradas:ID de Entidad":
        "ID_ORGANIZACION",

    "Entrega compilada:Partes involucradas:Nombre común":
        "NOMBRE_COMUN_ORIGINAL",

    "Entrega compilada:Partes involucradas:Identificador principal:ID":
        "IDENTIFICADOR_PRINCIPAL",

    "Entrega compilada:Partes involucradas:Identificador principal:Esquema":
        "ESQUEMA_IDENTIFICADOR",

    "Entrega compilada:Partes involucradas:Identificador principal:Nombre Legal":
        "NOMBRE_LEGAL_ORIGINAL",

    "Entrega compilada:Partes involucradas:Dirección:Dirección":
        "DIRECCION_ORIGINAL",

    "Entrega compilada:Partes involucradas:Dirección:Localidad":
        "LOCALIDAD_ORIGINAL",

    "Entrega compilada:Partes involucradas:Dirección:Región":
        "REGION_ORIGINAL",

    "Entrega compilada:Partes involucradas:Dirección:Departamento":
        "DEPARTAMENTO_ORIGINAL",

    "Entrega compilada:Partes involucradas:Dirección:País":
        "PAIS_ORIGINAL",

    "Entrega compilada:Partes involucradas:Punto de contacto:Teléfono":
        "TELEFONO",

    "Entrega compilada:Partes involucradas:Roles de las partes":
        "ROLES_ORIGINAL",

    "Entrega compilada:Partes involucradas:Punto de contacto:Correo electrónico":
        "CORREO",

    "Entrega compilada:Partes involucradas:Punto de contacto:URL":
        "URL_CONTACTO",
}


COLUMNAS_PARTES = [
    "OCID",
    "ID_ENTREGA",
    "ID_ORGANIZACION",
    "NOMBRE_COMUN_ORIGINAL",
    "IDENTIFICADOR_PRINCIPAL",
    "ESQUEMA_IDENTIFICADOR",
    "NOMBRE_LEGAL_ORIGINAL",
    "DIRECCION_ORIGINAL",
    "LOCALIDAD_ORIGINAL",
    "REGION_ORIGINAL",
    "DEPARTAMENTO_ORIGINAL",
    "PAIS_ORIGINAL",
    "TELEFONO",
    "ROLES_ORIGINAL",
    "CORREO",
    "URL_CONTACTO",
]


# ==========================================================
# FUNCIONES DE LIMPIEZA
# ==========================================================

def normalizar_encabezado(valor: Any) -> str:
    if valor is None:
        return ""

    texto = str(valor)

    texto = texto.replace("\n", " ")
    texto = texto.replace("\r", " ")
    texto = texto.replace("\t", " ")

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


def limpiar_texto(
    serie: pd.Series
) -> pd.Series:

    serie = serie.astype("string")

    serie = (
        serie
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.replace("\t", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    return serie.replace({
        "": pd.NA,
        "nan": pd.NA,
        "None": pd.NA,
        "<NA>": pd.NA,
    })


def normalizar_id(
    serie: pd.Series
) -> pd.Series:

    serie = limpiar_texto(serie)

    return serie.str.replace(
        r"^(\d+)\.0$",
        r"\1",
        regex=True
    )


def normalizar_catalogo(
    valor: Any
) -> Any:

    if pd.isna(valor):
        return pd.NA

    texto = str(valor).strip().upper()

    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    texto = re.sub(
        r"[^A-Z0-9Ñ&/ -]+",
        " ",
        texto
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip() or pd.NA


def extraer_ruc(
    valor: Any
) -> Any:

    if pd.isna(valor):
        return pd.NA

    texto = str(valor).strip()

    coincidencia = re.search(
        r"(?<!\d)(\d{11})(?!\d)",
        texto
    )

    if coincidencia:
        return coincidencia.group(1)

    solo_digitos = re.sub(
        r"\D",
        "",
        texto
    )

    if len(solo_digitos) == 11:
        return solo_digitos

    return pd.NA


def validar_ruc_peru(
    ruc: Any
) -> Any:

    if pd.isna(ruc):
        return pd.NA

    ruc = str(ruc)

    if not re.fullmatch(r"\d{11}", ruc):
        return False

    factores = [
        5, 4, 3, 2, 7,
        6, 5, 4, 3, 2
    ]

    suma = sum(
        int(ruc[i]) * factores[i]
        for i in range(10)
    )

    digito = 11 - (suma % 11)

    if digito == 10:
        digito = 0

    elif digito == 11:
        digito = 1

    return digito == int(ruc[10])


def agregar_columnas_faltantes(
    df: pd.DataFrame
) -> pd.DataFrame:

    for columna in COLUMNAS_PARTES:
        if columna not in df.columns:
            df[columna] = pd.NA

    return df[COLUMNAS_PARTES].copy()


# ==========================================================
# CONTROL INCREMENTAL
# ==========================================================

def calcular_sha256(
    archivo: Path
) -> str:

    resultado = hashlib.sha256()

    with open(archivo, "rb") as entrada:
        while True:
            bloque = entrada.read(
                1024 * 1024
            )

            if not bloque:
                break

            resultado.update(bloque)

    return resultado.hexdigest()


def cargar_manifest() -> dict:
    if not ARCHIVO_MANIFEST.exists():
        return {}

    try:
        with open(
            ARCHIVO_MANIFEST,
            "r",
            encoding="utf-8"
        ) as archivo:

            return json.load(archivo)

    except (json.JSONDecodeError, OSError):
        return {}


def guardar_manifest(
    manifest: dict
) -> None:

    with open(
        ARCHIVO_MANIFEST,
        "w",
        encoding="utf-8"
    ) as archivo:

        json.dump(
            manifest,
            archivo,
            ensure_ascii=False,
            indent=2
        )


# ==========================================================
# RUTAS PARTICIONADAS
# ==========================================================

def obtener_anio_mes(
    archivo: Path
) -> tuple[str, str]:

    relativa = archivo.relative_to(
        CARPETA_RAW
    )

    partes = relativa.parts

    if len(partes) < 4:
        raise RuntimeError(
            f"Ruta inesperada: {relativa}"
        )

    anio = partes[0]
    mes = partes[1].zfill(2)

    return anio, mes


def obtener_ruta_salida(
    anio: str,
    mes: str
) -> Path:

    carpeta = (
        CARPETA_PARTES
        / f"anio={anio}"
        / f"mes={mes}"
    )

    carpeta.mkdir(
        parents=True,
        exist_ok=True
    )

    return carpeta / "partes.parquet"


def guardar_parquet_seguro(
    df: pd.DataFrame,
    ruta: Path
) -> None:

    temporal = ruta.with_suffix(
        ".tmp.parquet"
    )

    df.to_parquet(
        temporal,
        index=False,
        engine="pyarrow",
        compression="snappy"
    )

    if ruta.exists():
        ruta.unlink()

    temporal.replace(ruta)


# ==========================================================
# TRANSFORMACIÓN
# ==========================================================

def transformar_partes(
    df: pd.DataFrame,
    anio: str,
    mes: str,
    archivo: Path
) -> tuple[pd.DataFrame, dict]:

    filas_origen = len(df)

    df.columns = [
        normalizar_encabezado(columna)
        for columna in df.columns
    ]

    df = df.rename(
        columns=MAPA_PARTES
    )

    df = agregar_columnas_faltantes(df)

    columnas_id = [
        "OCID",
        "ID_ENTREGA",
        "ID_ORGANIZACION",
        "IDENTIFICADOR_PRINCIPAL",
    ]

    for columna in columnas_id:
        df[columna] = normalizar_id(
            df[columna]
        )

    columnas_texto = [
        columna
        for columna in COLUMNAS_PARTES
        if columna not in columnas_id
    ]

    for columna in columnas_texto:
        df[columna] = limpiar_texto(
            df[columna]
        )

    # Nombre principal de la organización
    df["NOMBRE_ORGANIZACION_ORIGINAL"] = (
        df["NOMBRE_LEGAL_ORIGINAL"]
        .fillna(df["NOMBRE_COMUN_ORIGINAL"])
    )

    df["NOMBRE_ORGANIZACION_ESTANDAR"] = (
        df["NOMBRE_ORGANIZACION_ORIGINAL"]
        .map(normalizar_catalogo)
    )

    # RUC
    df["RUC_ORGANIZACION"] = (
        df["IDENTIFICADOR_PRINCIPAL"]
        .map(extraer_ruc)
    )

    mascara_sin_ruc = (
        df["RUC_ORGANIZACION"].isna()
    )

    df.loc[
        mascara_sin_ruc,
        "RUC_ORGANIZACION"
    ] = (
        df.loc[
            mascara_sin_ruc,
            "ID_ORGANIZACION"
        ]
        .map(extraer_ruc)
    )

    df["RUC_VALIDO"] = (
        df["RUC_ORGANIZACION"]
        .map(validar_ruc_peru)
        .astype("boolean")
    )

    # Geografía
    df["DIRECCION_ESTANDAR"] = (
        df["DIRECCION_ORIGINAL"]
        .map(normalizar_catalogo)
    )

    df["LOCALIDAD_ESTANDAR"] = (
        df["LOCALIDAD_ORIGINAL"]
        .map(normalizar_catalogo)
    )

    df["REGION_ESTANDAR"] = (
        df["REGION_ORIGINAL"]
        .map(normalizar_catalogo)
    )

    df["DEPARTAMENTO_ESTANDAR"] = (
        df["DEPARTAMENTO_ORIGINAL"]
        .map(normalizar_catalogo)
    )

    df["PAIS_ESTANDAR"] = (
        df["PAIS_ORIGINAL"]
        .map(normalizar_catalogo)
    )

    # Roles
    df["ROLES_ESTANDAR"] = (
        df["ROLES_ORIGINAL"]
        .map(normalizar_catalogo)
    )

    df["ES_COMPRADOR_O_ENTIDAD"] = (
        df["ROLES_ESTANDAR"]
        .str.contains(
            (
                r"BUYER|PROCURINGENTITY|"
                r"COMPRADOR|ENTIDAD CONTRATANTE"
            ),
            regex=True,
            na=False
        )
    )

    df["ES_PROVEEDOR"] = (
        df["ROLES_ESTANDAR"]
        .str.contains(
            r"SUPPLIER|PROVEEDOR",
            regex=True,
            na=False
        )
    )

    # Exigir identificación mínima del registro
    df = df[
        df["OCID"].notna()
        & df["ID_ENTREGA"].notna()
        & (
            df["ID_ORGANIZACION"].notna()
            | df["IDENTIFICADOR_PRINCIPAL"].notna()
            | df["NOMBRE_ORGANIZACION_ORIGINAL"].notna()
        )
    ].copy()

    # Llave maestra de organización
    df["CLAVE_ORGANIZACION"] = pd.NA

    mascara_ruc = (
        df["RUC_ORGANIZACION"].notna()
    )

    df.loc[
        mascara_ruc,
        "CLAVE_ORGANIZACION"
    ] = (
        "RUC|"
        + df.loc[
            mascara_ruc,
            "RUC_ORGANIZACION"
        ].astype("string")
    )

    mascara_id = (
        df["CLAVE_ORGANIZACION"].isna()
        & df["ID_ORGANIZACION"].notna()
    )

    df.loc[
        mascara_id,
        "CLAVE_ORGANIZACION"
    ] = (
        "ID|"
        + df.loc[
            mascara_id,
            "ID_ORGANIZACION"
        ].astype("string")
    )

    # Llave dentro del proceso OCDS
    identificador_parte = (
        df["ID_ORGANIZACION"]
        .fillna(df["IDENTIFICADOR_PRINCIPAL"])
        .fillna(df["NOMBRE_ORGANIZACION_ESTANDAR"])
    )

    df["CLAVE_PARTE"] = (
        df["OCID"].astype("string")
        + "|"
        + df["ID_ENTREGA"].astype("string")
        + "|"
        + identificador_parte.astype("string")
    )

    df["ES_DUPLICADO_CLAVE"] = (
        df.duplicated(
            subset=["CLAVE_PARTE"],
            keep=False
        )
    )

    df["ANIO_ARCHIVO"] = int(anio)
    df["MES_ARCHIVO"] = int(mes)
    df["ARCHIVO_ORIGEN"] = archivo.name

    control = {
        "ANIO": anio,
        "MES": mes,
        "ARCHIVO": archivo.name,
        "FILAS_ORIGEN": filas_origen,
        "FILAS_SALIDA": len(df),
        "FILAS_DESCARTADAS": (
            filas_origen - len(df)
        ),
        "RUC_PRESENTES": int(
            df["RUC_ORGANIZACION"]
            .notna()
            .sum()
        ),
        "RUC_VALIDOS": int(
            df["RUC_VALIDO"]
            .fillna(False)
            .sum()
        ),
        "CON_DEPARTAMENTO": int(
            df["DEPARTAMENTO_ESTANDAR"]
            .notna()
            .sum()
        ),
        "CON_REGION": int(
            df["REGION_ESTANDAR"]
            .notna()
            .sum()
        ),
        "CON_LOCALIDAD": int(
            df["LOCALIDAD_ESTANDAR"]
            .notna()
            .sum()
        ),
        "CON_PAIS": int(
            df["PAIS_ESTANDAR"]
            .notna()
            .sum()
        ),
        "ROL_ENTIDAD": int(
            df["ES_COMPRADOR_O_ENTIDAD"]
            .sum()
        ),
        "ROL_PROVEEDOR": int(
            df["ES_PROVEEDOR"]
            .sum()
        ),
    }

    return df, control


# ==========================================================
# PERFIL GLOBAL
# ==========================================================

def generar_perfil_global() -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame
]:

    controles = []
    catalogos_geo = []
    roles_acumulados = []

    archivos = sorted(
        CARPETA_PARTES.rglob(
            "partes.parquet"
        )
    )

    for archivo in archivos:
        df = pd.read_parquet(
            archivo,
            columns=[
                "ANIO_ARCHIVO",
                "MES_ARCHIVO",
                "RUC_ORGANIZACION",
                "RUC_VALIDO",
                "DEPARTAMENTO_ESTANDAR",
                "REGION_ESTANDAR",
                "LOCALIDAD_ESTANDAR",
                "PAIS_ESTANDAR",
                "ROLES_ESTANDAR",
                "ES_COMPRADOR_O_ENTIDAD",
                "ES_PROVEEDOR",
            ]
        )

        anio = int(
            df["ANIO_ARCHIVO"].iloc[0]
        ) if not df.empty else None

        mes = int(
            df["MES_ARCHIVO"].iloc[0]
        ) if not df.empty else None

        controles.append({
            "ANIO": anio,
            "MES": mes,
            "FILAS": len(df),
            "RUC_PRESENTES": int(
                df["RUC_ORGANIZACION"]
                .notna()
                .sum()
            ),
            "RUC_VALIDOS": int(
                df["RUC_VALIDO"]
                .fillna(False)
                .sum()
            ),
            "CON_DEPARTAMENTO": int(
                df["DEPARTAMENTO_ESTANDAR"]
                .notna()
                .sum()
            ),
            "CON_REGION": int(
                df["REGION_ESTANDAR"]
                .notna()
                .sum()
            ),
            "CON_LOCALIDAD": int(
                df["LOCALIDAD_ESTANDAR"]
                .notna()
                .sum()
            ),
            "CON_PAIS": int(
                df["PAIS_ESTANDAR"]
                .notna()
                .sum()
            ),
            "ROL_ENTIDAD": int(
                df["ES_COMPRADOR_O_ENTIDAD"]
                .sum()
            ),
            "ROL_PROVEEDOR": int(
                df["ES_PROVEEDOR"]
                .sum()
            ),
        })

        catalogo_mes = (
            df[[
                "PAIS_ESTANDAR",
                "DEPARTAMENTO_ESTANDAR",
                "REGION_ESTANDAR",
                "LOCALIDAD_ESTANDAR",
            ]]
            .dropna(how="all")
            .drop_duplicates()
        )

        catalogos_geo.append(
            catalogo_mes
        )

        roles_mes = (
            df["ROLES_ESTANDAR"]
            .dropna()
            .value_counts()
            .rename_axis("ROL")
            .reset_index(name="CANTIDAD")
        )

        roles_acumulados.append(
            roles_mes
        )

    df_control = pd.DataFrame(
        controles
    )

    if catalogos_geo:
        df_geo = (
            pd.concat(
                catalogos_geo,
                ignore_index=True
            )
            .drop_duplicates()
            .sort_values(
                by=[
                    "PAIS_ESTANDAR",
                    "DEPARTAMENTO_ESTANDAR",
                    "REGION_ESTANDAR",
                    "LOCALIDAD_ESTANDAR",
                ],
                na_position="last"
            )
            .reset_index(drop=True)
        )
    else:
        df_geo = pd.DataFrame()

    if roles_acumulados:
        df_roles = (
            pd.concat(
                roles_acumulados,
                ignore_index=True
            )
            .groupby(
                "ROL",
                as_index=False
            )["CANTIDAD"]
            .sum()
            .sort_values(
                "CANTIDAD",
                ascending=False
            )
        )
    else:
        df_roles = pd.DataFrame()

    return (
        df_control,
        df_geo,
        df_roles
    )


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def construir_staging_partes(
    forzar: bool
) -> None:

    archivos = sorted(
        archivo
        for archivo in CARPETA_RAW.glob(
            "*/*/extraido/*.xlsx"
        )
        if archivo.is_file()
        and not archivo.name.startswith("~$")
    )

    if not archivos:
        print(
            "No se encontraron archivos mensuales."
        )
        return

    manifest = cargar_manifest()

    controles_ejecucion = []
    errores = []

    procesados = 0
    omitidos = 0

    print("========================================")
    print("STAGING DE PARTES Y GEOGRAFÍA")
    print("========================================")
    print(f"Archivos encontrados: {len(archivos)}")

    for archivo in archivos:
        relativa = str(
            archivo.relative_to(
                CARPETA_RAW
            )
        ).replace("\\", "/")

        try:
            anio, mes = obtener_anio_mes(
                archivo
            )

            salida = obtener_ruta_salida(
                anio,
                mes
            )

            firma = calcular_sha256(
                archivo
            )

            anterior = manifest.get(
                relativa,
                {}
            )

            if (
                not forzar
                and anterior.get("sha256") == firma
                and salida.exists()
            ):
                print(
                    f"OMITIDO SIN CAMBIOS: "
                    f"{anio}-{mes}"
                )

                omitidos += 1
                continue

            print(
                f"Procesando {anio}-{mes}: "
                f"{archivo.name}"
            )

            df_origen = pd.read_excel(
                archivo,
                sheet_name="Ent_PartesInvolucradas",
                dtype=str,
                engine="openpyxl"
            )

            df_salida, control = transformar_partes(
                df_origen,
                anio,
                mes,
                archivo
            )

            guardar_parquet_seguro(
                df_salida,
                salida
            )

            controles_ejecucion.append(
                control
            )

            manifest[relativa] = {
                "sha256": firma,
                "procesado_en": (
                    datetime.now()
                    .astimezone()
                    .isoformat()
                ),
                "salida": str(
                    salida.relative_to(
                        CARPETA_STAGING
                    )
                ).replace("\\", "/"),
                "filas": len(df_salida),
            }

            guardar_manifest(
                manifest
            )

            procesados += 1

            print(
                f"PARTES: {len(df_salida)} filas"
            )

        except Exception as error:
            print(
                f"ERROR en {relativa}: {error}"
            )

            errores.append({
                "ARCHIVO": relativa,
                "ERROR": str(error),
            })

    print(
        "\nGenerando perfil geográfico global..."
    )

    (
        df_cobertura,
        df_geografia,
        df_roles
    ) = generar_perfil_global()

    if not df_geografia.empty:
        df_geografia.insert(
            0,
            "UBICACION_KEY",
            range(
                1,
                len(df_geografia) + 1
            )
        )

        df_geografia.to_parquet(
            ARCHIVO_CATALOGO_GEO,
            index=False,
            engine="pyarrow",
            compression="snappy"
        )

    total_filas = int(
        df_cobertura["FILAS"].sum()
    ) if not df_cobertura.empty else 0

    total_departamento = int(
        df_cobertura[
            "CON_DEPARTAMENTO"
        ].sum()
    ) if not df_cobertura.empty else 0

    total_region = int(
        df_cobertura[
            "CON_REGION"
        ].sum()
    ) if not df_cobertura.empty else 0

    total_localidad = int(
        df_cobertura[
            "CON_LOCALIDAD"
        ].sum()
    ) if not df_cobertura.empty else 0

    resumen = pd.DataFrame([
        {
            "INDICADOR": "Archivos mensuales encontrados",
            "VALOR": len(archivos)
        },
        {
            "INDICADOR": "Archivos procesados",
            "VALOR": procesados
        },
        {
            "INDICADOR": "Archivos omitidos sin cambios",
            "VALOR": omitidos
        },
        {
            "INDICADOR": "Archivos con error",
            "VALOR": len(errores)
        },
        {
            "INDICADOR": "Total de partes",
            "VALOR": total_filas
        },
        {
            "INDICADOR": "Registros con departamento",
            "VALOR": total_departamento
        },
        {
            "INDICADOR": "Registros con región",
            "VALOR": total_region
        },
        {
            "INDICADOR": "Registros con localidad",
            "VALOR": total_localidad
        },
        {
            "INDICADOR": "Ubicaciones distintas",
            "VALOR": len(df_geografia)
        },
    ])

    df_errores = pd.DataFrame(
        errores,
        columns=[
            "ARCHIVO",
            "ERROR"
        ]
    )

    df_control_ejecucion = pd.DataFrame(
        controles_ejecucion
    )

    with pd.ExcelWriter(
        ARCHIVO_CONTROL,
        engine="openpyxl"
    ) as writer:

        resumen.to_excel(
            writer,
            sheet_name="RESUMEN",
            index=False
        )

        df_cobertura.to_excel(
            writer,
            sheet_name="COBERTURA_MENSUAL",
            index=False
        )

        df_geografia.to_excel(
            writer,
            sheet_name="CATALOGO_GEOGRAFICO",
            index=False
        )

        df_roles.to_excel(
            writer,
            sheet_name="ROLES_ORGANIZACION",
            index=False
        )

        df_control_ejecucion.to_excel(
            writer,
            sheet_name="PROCESADOS_EJECUCION",
            index=False
        )

        df_errores.to_excel(
            writer,
            sheet_name="ERRORES",
            index=False
        )

    print("\n========================================")
    print("STAGING DE PARTES TERMINADO")
    print("========================================")
    print(f"Procesados: {procesados}")
    print(f"Sin cambios: {omitidos}")
    print(f"Errores: {len(errores)}")
    print(f"Total de partes: {total_filas}")
    print(
        f"Con departamento: "
        f"{total_departamento}"
    )
    print(f"Con región: {total_region}")
    print(
        f"Con localidad: "
        f"{total_localidad}"
    )
    print(
        f"Ubicaciones distintas: "
        f"{len(df_geografia)}"
    )
    print(f"\nControl: {ARCHIVO_CONTROL}")
    print(
        f"Catálogo geográfico: "
        f"{ARCHIVO_CATALOGO_GEO}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Construye el staging de partes, "
            "organizaciones y geografía OECE."
        )
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help=(
            "Procesa nuevamente todos los meses."
        )
    )

    argumentos = parser.parse_args()

    construir_staging_partes(
        forzar=argumentos.force
    )


if __name__ == "__main__":
    main()
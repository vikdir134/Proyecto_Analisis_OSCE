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
import pyarrow.parquet as pq


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

CARPETA_CALIDAD = (
    BASE_DIR
    / "data"
    / "quality"
    / "oece"
)

CARPETA_STAGING.mkdir(parents=True, exist_ok=True)
CARPETA_CALIDAD.mkdir(parents=True, exist_ok=True)

ARCHIVO_MANIFEST = (
    CARPETA_STAGING
    / "manifest_staging.json"
)

ARCHIVO_CONTROL = (
    CARPETA_CALIDAD
    / "CONTROL_STAGING_OECE.xlsx"
)

ARCHIVO_HISTORIAL = (
    CARPETA_CALIDAD
    / "HISTORIAL_EJECUCIONES_STAGING.csv"
)

ANIOS_PERMITIDOS = {
    "2022",
    "2023",
    "2024",
    "2025",
    "2026",
}


# ==========================================================
# MAPEO: REGISTROS
# ==========================================================

MAPA_REGISTROS = {
    "Open Contracting ID":
        "OCID",

    "Entrega compilada:ID de Entrega":
        "ID_ENTREGA",

    "Entrega compilada:Etiqueta de entrega":
        "ETIQUETA_ENTREGA",

    "Entrega compilada:Fecha de entrega":
        "FECHA_ENTREGA",

    "compiledrelease/publisheddate":
        "FECHA_PUBLICACION",

    "Entrega compilada:Comprador:ID de Organización":
        "ID_COMPRADOR_ORIGINAL",

    "Entrega compilada:Comprador:Nombre de la Organización":
        "NOMBRE_COMPRADOR_ORIGINAL",

    "Entrega compilada:Licitación:ID de licitación":
        "ID_LICITACION",

    "Entrega compilada:Licitación:Título de la licitación":
        "TITULO_LICITACION",

    "Entrega compilada:Licitación:Descripción de la licitación":
        "DESCRIPCION_LICITACION",

    "Entrega compilada:Licitación:Entidad contratante:ID de Organización":
        "ID_ENTIDAD_CONTRATANTE_ORIGINAL",

    "Entrega compilada:Licitación:Entidad contratante:Nombre de la Organización":
        "NOMBRE_ENTIDAD_CONTRATANTE_ORIGINAL",

    "compiledrelease/tender/datepublished":
        "FECHA_PUBLICACION_LICITACION",

    "Entrega compilada:Licitación:Método de contratación":
        "METODO_CONTRATACION",

    "Entrega compilada:Licitación:Detalles del método de contratación":
        "DETALLE_METODO_CONTRATACION",

    "Entrega compilada:Licitación:Categoría principal de contratación":
        "CATEGORIA_PRINCIPAL",

    "Entrega compilada:Licitación:Categorías adicionales de contratación":
        "CATEGORIAS_ADICIONALES",

    "Entrega compilada:Licitación:Valor:Monto":
        "MONTO_LICITACION",

    "Entrega compilada:Licitación:Valor:Moneda":
        "MONEDA_LICITACION",

    "Entrega compilada:Licitación:Valor:Nombre de Moneda":
        "NOMBRE_MONEDA_LICITACION",

    "compiledrelease/tender/value/amount_pen":
        "MONTO_LICITACION_PEN",
}


COLUMNAS_REGISTROS = [
    "OCID",
    "ID_ENTREGA",
    "ETIQUETA_ENTREGA",
    "FECHA_ENTREGA",
    "FECHA_PUBLICACION",
    "ID_COMPRADOR_ORIGINAL",
    "NOMBRE_COMPRADOR_ORIGINAL",
    "ID_LICITACION",
    "TITULO_LICITACION",
    "DESCRIPCION_LICITACION",
    "ID_ENTIDAD_CONTRATANTE_ORIGINAL",
    "NOMBRE_ENTIDAD_CONTRATANTE_ORIGINAL",
    "FECHA_PUBLICACION_LICITACION",
    "METODO_CONTRATACION",
    "DETALLE_METODO_CONTRATACION",
    "CATEGORIA_PRINCIPAL",
    "CATEGORIAS_ADICIONALES",
    "MONTO_LICITACION",
    "MONEDA_LICITACION",
    "NOMBRE_MONEDA_LICITACION",
    "MONTO_LICITACION_PEN",
]


# ==========================================================
# MAPEO: CONTRATOS
# ==========================================================

MAPA_CONTRATOS = {
    "Open Contracting ID":
        "OCID",

    "Entrega compilada:ID de Entrega":
        "ID_ENTREGA",

    "Entrega compilada:Contratos:ID del Contrato":
        "ID_CONTRATO",

    "Entrega compilada:Contratos:ID de Adjudicación":
        "ID_ADJUDICACION",

    "Entrega compilada:Contratos:Título del contrato":
        "TITULO_CONTRATO",

    "Entrega compilada:Contratos:Descripción del contrato":
        "DESCRIPCION_CONTRATO",

    "Entrega compilada:Contratos:Periodo:Fecha de inicio":
        "FECHA_INICIO",

    "Entrega compilada:Contratos:Periodo:Fecha de fin":
        "FECHA_FIN",

    "Entrega compilada:Contratos:Periodo:Duración (días)":
        "DURACION_DIAS",

    "Entrega compilada:Contratos:Valor:Monto":
        "MONTO_CONTRATO",

    "Entrega compilada:Contratos:Valor:Moneda":
        "MONEDA_CONTRATO",

    "Entrega compilada:Contratos:Valor:Nombre de Moneda":
        "NOMBRE_MONEDA_CONTRATO",

    "Entrega compilada:Contratos:Fecha de firma":
        "FECHA_FIRMA",

    "Entrega compilada:Contratos:Implementación:Valor final:Monto":
        "MONTO_FINAL",

    "Entrega compilada:Contratos:Implementación:Valor final:Moneda":
        "MONEDA_FINAL",

    "Entrega compilada:Contratos:Implementación:Valor final:Nombre de Moneda":
        "NOMBRE_MONEDA_FINAL",

    "Entrega compilada:Contratos:Implementación:Fecha de fin":
        "FECHA_FIN_IMPLEMENTACION",
}


COLUMNAS_CONTRATOS = [
    "OCID",
    "ID_ENTREGA",
    "ID_CONTRATO",
    "ID_ADJUDICACION",
    "TITULO_CONTRATO",
    "DESCRIPCION_CONTRATO",
    "FECHA_INICIO",
    "FECHA_FIN",
    "DURACION_DIAS",
    "MONTO_CONTRATO",
    "MONEDA_CONTRATO",
    "NOMBRE_MONEDA_CONTRATO",
    "FECHA_FIRMA",
    "MONTO_FINAL",
    "MONEDA_FINAL",
    "NOMBRE_MONEDA_FINAL",
    "FECHA_FIN_IMPLEMENTACION",
]


# ==========================================================
# MAPEO: PROVEEDORES
# ==========================================================

MAPA_PROVEEDORES = {
    "Open Contracting ID":
        "OCID",

    "Entrega compilada:ID de Entrega":
        "ID_ENTREGA",

    "Entrega compilada:Adjudicaciones:ID de Adjudicación":
        "ID_ADJUDICACION",

    "Entrega compilada:Adjudicaciones:Proveedores:ID de Organización":
        "ID_PROVEEDOR_ORIGINAL",

    "Entrega compilada:Adjudicaciones:Proveedores:Nombre de la Organización":
        "NOMBRE_PROVEEDOR_ORIGINAL",
}


COLUMNAS_PROVEEDORES = [
    "OCID",
    "ID_ENTREGA",
    "ID_ADJUDICACION",
    "ID_PROVEEDOR_ORIGINAL",
    "NOMBRE_PROVEEDOR_ORIGINAL",
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
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def limpiar_texto(serie: pd.Series) -> pd.Series:
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


def normalizar_id(serie: pd.Series) -> pd.Series:
    serie = limpiar_texto(serie)

    # Corrige identificadores convertidos por Excel:
    # 20549203931.0 -> 20549203931
    serie = serie.str.replace(
        r"^(\d+)\.0$",
        r"\1",
        regex=True,
    )

    return serie


def normalizar_nombre(valor: Any) -> Any:
    if pd.isna(valor):
        return pd.NA

    texto = str(valor).strip().upper()

    texto = unicodedata.normalize(
        "NFKD",
        texto,
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    texto = re.sub(
        r"[^A-Z0-9Ñ& ]+",
        " ",
        texto,
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto,
    )

    return texto.strip() or pd.NA


def extraer_ruc(valor: Any) -> Any:
    """
    Busca una secuencia de exactamente 11 dígitos.

    También funciona con identificadores como:
    PE-RUC-20549203931
    """

    if pd.isna(valor):
        return pd.NA

    texto = str(valor).strip()

    coincidencia = re.search(
        r"(?<!\d)(\d{11})(?!\d)",
        texto,
    )

    if coincidencia:
        return coincidencia.group(1)

    solo_digitos = re.sub(
        r"\D",
        "",
        texto,
    )

    if len(solo_digitos) == 11:
        return solo_digitos

    return pd.NA


def validar_ruc_peru(ruc: Any) -> Any:
    """
    Valida el dígito verificador de un RUC peruano.
    """

    if pd.isna(ruc):
        return pd.NA

    ruc = str(ruc)

    if not re.fullmatch(r"\d{11}", ruc):
        return False

    factores = [
        5, 4, 3, 2, 7,
        6, 5, 4, 3, 2,
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


def convertir_numero(valor: Any) -> float | None:
    if pd.isna(valor):
        return None

    texto = str(valor).strip()

    if not texto:
        return None

    texto = texto.replace("\u00a0", "")
    texto = texto.replace(" ", "")
    texto = texto.replace("S/", "")
    texto = texto.replace("$", "")

    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "")
            texto = texto.replace(",", ".")
        else:
            texto = texto.replace(",", "")

    elif "," in texto:
        texto = texto.replace(",", ".")

    texto = re.sub(
        r"[^0-9.\-]",
        "",
        texto,
    )

    try:
        return float(texto)
    except ValueError:
        return None


def convertir_fecha(serie: pd.Series) -> pd.Series:
    return pd.to_datetime(
        serie,
        errors="coerce",
    )


def agregar_columnas_faltantes(
    df: pd.DataFrame,
    columnas: list[str],
) -> pd.DataFrame:

    for columna in columnas:
        if columna not in df.columns:
            df[columna] = pd.NA

    return df[columnas].copy()


# ==========================================================
# FIRMA DEL ARCHIVO
# ==========================================================

def calcular_sha256(archivo: Path) -> str:
    hash_archivo = hashlib.sha256()

    with open(archivo, "rb") as entrada:
        while True:
            bloque = entrada.read(
                1024 * 1024
            )

            if not bloque:
                break

            hash_archivo.update(bloque)

    return hash_archivo.hexdigest()


def cargar_manifest() -> dict[str, Any]:
    if not ARCHIVO_MANIFEST.exists():
        return {}

    try:
        with open(
            ARCHIVO_MANIFEST,
            "r",
            encoding="utf-8",
        ) as archivo:

            return json.load(archivo)

    except (json.JSONDecodeError, OSError):
        return {}


def guardar_manifest(
    manifest: dict[str, Any],
) -> None:

    with open(
        ARCHIVO_MANIFEST,
        "w",
        encoding="utf-8",
    ) as archivo:

        json.dump(
            manifest,
            archivo,
            ensure_ascii=False,
            indent=2,
        )


# ==========================================================
# IDENTIFICACIÓN DEL ARCHIVO
# ==========================================================

def obtener_anio_mes(
    archivo: Path,
) -> tuple[str, str]:

    ruta_relativa = archivo.relative_to(
        CARPETA_RAW
    )

    partes = ruta_relativa.parts

    if len(partes) < 4:
        raise RuntimeError(
            f"Ruta inesperada: {ruta_relativa}"
        )

    anio = partes[0]
    mes = partes[1].zfill(2)

    if anio not in ANIOS_PERMITIDOS:
        raise RuntimeError(
            f"Año no permitido: {anio}"
        )

    if not mes.isdigit():
        raise RuntimeError(
            f"Mes inválido: {mes}"
        )

    return anio, mes


# ==========================================================
# TRANSFORMACIÓN: REGISTROS
# ==========================================================

def transformar_registros(
    df: pd.DataFrame,
    anio: str,
    mes: str,
    archivo: Path,
) -> tuple[pd.DataFrame, dict[str, Any]]:

    filas_origen = len(df)

    df.columns = [
        normalizar_encabezado(columna)
        for columna in df.columns
    ]

    df = df.rename(
        columns=MAPA_REGISTROS
    )

    df = agregar_columnas_faltantes(
        df,
        COLUMNAS_REGISTROS,
    )

    columnas_id = [
        "OCID",
        "ID_ENTREGA",
        "ID_COMPRADOR_ORIGINAL",
        "ID_LICITACION",
        "ID_ENTIDAD_CONTRATANTE_ORIGINAL",
    ]

    for columna in columnas_id:
        df[columna] = normalizar_id(
            df[columna]
        )

    columnas_texto = [
        columna
        for columna in COLUMNAS_REGISTROS
        if columna not in columnas_id
    ]

    for columna in columnas_texto:
        df[columna] = limpiar_texto(
            df[columna]
        )

    df["FECHA_ENTREGA"] = convertir_fecha(
        df["FECHA_ENTREGA"]
    )

    df["FECHA_PUBLICACION"] = convertir_fecha(
        df["FECHA_PUBLICACION"]
    )

    df["FECHA_PUBLICACION_LICITACION"] = convertir_fecha(
        df["FECHA_PUBLICACION_LICITACION"]
    )

    df["MONTO_LICITACION"] = (
        df["MONTO_LICITACION"]
        .map(convertir_numero)
    )

    df["MONTO_LICITACION_PEN"] = (
        df["MONTO_LICITACION_PEN"]
        .map(convertir_numero)
    )

    df["ID_COMPRADOR"] = (
        df["ID_COMPRADOR_ORIGINAL"]
    )

    df["ID_ENTIDAD_CONTRATANTE"] = (
        df["ID_ENTIDAD_CONTRATANTE_ORIGINAL"]
    )

    df["RUC_COMPRADOR"] = (
        df["ID_COMPRADOR_ORIGINAL"]
        .map(extraer_ruc)
    )

    df["RUC_ENTIDAD_CONTRATANTE"] = (
        df["ID_ENTIDAD_CONTRATANTE_ORIGINAL"]
        .map(extraer_ruc)
    )

    df["RUC_COMPRADOR_VALIDO"] = (
        df["RUC_COMPRADOR"]
        .map(validar_ruc_peru)
        .astype("boolean")
    )

    df["RUC_ENTIDAD_VALIDO"] = (
        df["RUC_ENTIDAD_CONTRATANTE"]
        .map(validar_ruc_peru)
        .astype("boolean")
    )

    df["NOMBRE_COMPRADOR_ESTANDAR"] = (
        df["NOMBRE_COMPRADOR_ORIGINAL"]
        .map(normalizar_nombre)
    )

    df["NOMBRE_ENTIDAD_ESTANDAR"] = (
        df["NOMBRE_ENTIDAD_CONTRATANTE_ORIGINAL"]
        .map(normalizar_nombre)
    )

    df = df[
        df["OCID"].notna()
        & df["ID_ENTREGA"].notna()
    ].copy()

    df["CLAVE_REGISTRO"] = (
        df["OCID"].astype("string")
        + "|"
        + df["ID_ENTREGA"].astype("string")
    )

    df["ES_DUPLICADO_CLAVE"] = (
        df.duplicated(
            subset=["CLAVE_REGISTRO"],
            keep=False,
        )
    )

    df["ANIO_ARCHIVO"] = int(anio)
    df["MES_ARCHIVO"] = int(mes)
    df["ARCHIVO_ORIGEN"] = archivo.name

    control = {
        "TABLA": "REGISTROS",
        "ANIO": anio,
        "MES": mes,
        "FILAS_ORIGEN": filas_origen,
        "FILAS_SALIDA": len(df),
        "FILAS_DESCARTADAS": filas_origen - len(df),
        "CLAVES_DUPLICADAS": int(
            df["ES_DUPLICADO_CLAVE"].sum()
        ),
        "RUC_PRESENTES": int(
            df["RUC_ENTIDAD_CONTRATANTE"]
            .notna()
            .sum()
        ),
        "RUC_VALIDOS": int(
            df["RUC_ENTIDAD_VALIDO"]
            .fillna(False)
            .sum()
        ),
    }

    return df, control


# ==========================================================
# TRANSFORMACIÓN: CONTRATOS
# ==========================================================

def transformar_contratos(
    df: pd.DataFrame,
    anio: str,
    mes: str,
    archivo: Path,
) -> tuple[pd.DataFrame, dict[str, Any]]:

    filas_origen = len(df)

    df.columns = [
        normalizar_encabezado(columna)
        for columna in df.columns
    ]

    df = df.rename(
        columns=MAPA_CONTRATOS
    )

    df = agregar_columnas_faltantes(
        df,
        COLUMNAS_CONTRATOS,
    )

    columnas_id = [
        "OCID",
        "ID_ENTREGA",
        "ID_CONTRATO",
        "ID_ADJUDICACION",
    ]

    for columna in columnas_id:
        df[columna] = normalizar_id(
            df[columna]
        )

    columnas_texto = [
        "TITULO_CONTRATO",
        "DESCRIPCION_CONTRATO",
        "MONEDA_CONTRATO",
        "NOMBRE_MONEDA_CONTRATO",
        "MONEDA_FINAL",
        "NOMBRE_MONEDA_FINAL",
    ]

    for columna in columnas_texto:
        df[columna] = limpiar_texto(
            df[columna]
        )

    columnas_fecha = [
        "FECHA_INICIO",
        "FECHA_FIN",
        "FECHA_FIRMA",
        "FECHA_FIN_IMPLEMENTACION",
    ]

    for columna in columnas_fecha:
        df[columna] = convertir_fecha(
            df[columna]
        )

    columnas_numericas = [
        "DURACION_DIAS",
        "MONTO_CONTRATO",
        "MONTO_FINAL",
    ]

    for columna in columnas_numericas:
        df[columna] = (
            df[columna]
            .map(convertir_numero)
        )

    df = df[
        df["OCID"].notna()
        & df["ID_CONTRATO"].notna()
    ].copy()

    df["CLAVE_CONTRATO"] = (
        df["OCID"].astype("string")
        + "|"
        + df["ID_CONTRATO"].astype("string")
    )

    df["ES_DUPLICADO_CLAVE"] = (
        df.duplicated(
            subset=["CLAVE_CONTRATO"],
            keep=False,
        )
    )

    df["ANIO_ARCHIVO"] = int(anio)
    df["MES_ARCHIVO"] = int(mes)
    df["ARCHIVO_ORIGEN"] = archivo.name

    control = {
        "TABLA": "CONTRATOS",
        "ANIO": anio,
        "MES": mes,
        "FILAS_ORIGEN": filas_origen,
        "FILAS_SALIDA": len(df),
        "FILAS_DESCARTADAS": filas_origen - len(df),
        "CLAVES_DUPLICADAS": int(
            df["ES_DUPLICADO_CLAVE"].sum()
        ),
        "RUC_PRESENTES": 0,
        "RUC_VALIDOS": 0,
    }

    return df, control


# ==========================================================
# TRANSFORMACIÓN: PROVEEDORES
# ==========================================================

def transformar_proveedores(
    df: pd.DataFrame,
    anio: str,
    mes: str,
    archivo: Path,
) -> tuple[pd.DataFrame, dict[str, Any]]:

    filas_origen = len(df)

    df.columns = [
        normalizar_encabezado(columna)
        for columna in df.columns
    ]

    df = df.rename(
        columns=MAPA_PROVEEDORES
    )

    df = agregar_columnas_faltantes(
        df,
        COLUMNAS_PROVEEDORES,
    )

    columnas_id = [
        "OCID",
        "ID_ENTREGA",
        "ID_ADJUDICACION",
        "ID_PROVEEDOR_ORIGINAL",
    ]

    for columna in columnas_id:
        df[columna] = normalizar_id(
            df[columna]
        )

    df["NOMBRE_PROVEEDOR_ORIGINAL"] = limpiar_texto(
        df["NOMBRE_PROVEEDOR_ORIGINAL"]
    )

    df["ID_PROVEEDOR"] = (
        df["ID_PROVEEDOR_ORIGINAL"]
    )

    df["RUC_PROVEEDOR"] = (
        df["ID_PROVEEDOR_ORIGINAL"]
        .map(extraer_ruc)
    )

    df["RUC_PROVEEDOR_VALIDO"] = (
        df["RUC_PROVEEDOR"]
        .map(validar_ruc_peru)
        .astype("boolean")
    )

    df["NOMBRE_PROVEEDOR_ESTANDAR"] = (
        df["NOMBRE_PROVEEDOR_ORIGINAL"]
        .map(normalizar_nombre)
    )

    df = df[
        df["OCID"].notna()
        & df["ID_ENTREGA"].notna()
        & df["ID_ADJUDICACION"].notna()
        & df["ID_PROVEEDOR"].notna()
    ].copy()

    df["CLAVE_PROVEEDOR_ADJUDICACION"] = (
        df["OCID"].astype("string")
        + "|"
        + df["ID_ADJUDICACION"].astype("string")
        + "|"
        + df["ID_PROVEEDOR"].astype("string")
    )

    df["ES_DUPLICADO_CLAVE"] = (
        df.duplicated(
            subset=["CLAVE_PROVEEDOR_ADJUDICACION"],
            keep=False,
        )
    )

    df["ANIO_ARCHIVO"] = int(anio)
    df["MES_ARCHIVO"] = int(mes)
    df["ARCHIVO_ORIGEN"] = archivo.name

    control = {
        "TABLA": "PROVEEDORES",
        "ANIO": anio,
        "MES": mes,
        "FILAS_ORIGEN": filas_origen,
        "FILAS_SALIDA": len(df),
        "FILAS_DESCARTADAS": filas_origen - len(df),
        "CLAVES_DUPLICADAS": int(
            df["ES_DUPLICADO_CLAVE"].sum()
        ),
        "RUC_PRESENTES": int(
            df["RUC_PROVEEDOR"]
            .notna()
            .sum()
        ),
        "RUC_VALIDOS": int(
            df["RUC_PROVEEDOR_VALIDO"]
            .fillna(False)
            .sum()
        ),
    }

    return df, control


# ==========================================================
# GUARDAR PARQUET PARTICIONADO
# ==========================================================

def ruta_parquet(
    tabla: str,
    anio: str,
    mes: str,
) -> Path:

    carpeta = (
        CARPETA_STAGING
        / tabla.lower()
        / f"anio={anio}"
        / f"mes={mes}"
    )

    carpeta.mkdir(
        parents=True,
        exist_ok=True,
    )

    return carpeta / f"{tabla.lower()}.parquet"


def guardar_parquet_seguro(
    df: pd.DataFrame,
    ruta_final: Path,
) -> None:

    ruta_temporal = ruta_final.with_suffix(
        ".tmp.parquet"
    )

    df.to_parquet(
        ruta_temporal,
        index=False,
        engine="pyarrow",
        compression="snappy",
    )

    if ruta_final.exists():
        ruta_final.unlink()

    ruta_temporal.replace(
        ruta_final
    )


# ==========================================================
# PROCESAR UN ARCHIVO MENSUAL
# ==========================================================

def procesar_archivo(
    archivo: Path,
) -> tuple[
    list[dict[str, Any]],
    list[str],
]:

    anio, mes = obtener_anio_mes(
        archivo
    )

    print("\n----------------------------------------")
    print(f"Procesando: {anio}-{mes}")
    print(f"Archivo: {archivo.name}")

    registros_origen = pd.read_excel(
        archivo,
        sheet_name="Registros",
        dtype=str,
        engine="openpyxl",
    )

    contratos_origen = pd.read_excel(
        archivo,
        sheet_name="Ent_Contratos",
        dtype=str,
        engine="openpyxl",
    )

    proveedores_origen = pd.read_excel(
        archivo,
        sheet_name="Ent_Adj_Proveedores",
        dtype=str,
        engine="openpyxl",
    )

    registros, control_registros = transformar_registros(
        registros_origen,
        anio,
        mes,
        archivo,
    )

    contratos, control_contratos = transformar_contratos(
        contratos_origen,
        anio,
        mes,
        archivo,
    )

    proveedores, control_proveedores = transformar_proveedores(
        proveedores_origen,
        anio,
        mes,
        archivo,
    )

    salidas = {
        "REGISTROS": (
            registros,
            ruta_parquet(
                "REGISTROS",
                anio,
                mes,
            ),
        ),

        "CONTRATOS": (
            contratos,
            ruta_parquet(
                "CONTRATOS",
                anio,
                mes,
            ),
        ),

        "PROVEEDORES": (
            proveedores,
            ruta_parquet(
                "PROVEEDORES",
                anio,
                mes,
            ),
        ),
    }

    rutas_generadas = []

    for tabla, (df, ruta) in salidas.items():
        guardar_parquet_seguro(
            df,
            ruta,
        )

        rutas_generadas.append(
            str(
                ruta.relative_to(
                    CARPETA_STAGING
                )
            ).replace("\\", "/")
        )

        print(
            f"{tabla}: {len(df)} filas"
        )

    return (
        [
            control_registros,
            control_contratos,
            control_proveedores,
        ],
        rutas_generadas,
    )


# ==========================================================
# RESUMEN GLOBAL DE PARQUET
# ==========================================================

def generar_resumen_global() -> pd.DataFrame:
    filas = []

    for tabla in [
        "registros",
        "contratos",
        "proveedores",
    ]:
        archivos = sorted(
            (
                CARPETA_STAGING
                / tabla
            ).rglob("*.parquet")
        )

        for archivo in archivos:
            partes = archivo.parts

            anio = next(
                (
                    parte.replace("anio=", "")
                    for parte in partes
                    if parte.startswith("anio=")
                ),
                "",
            )

            mes = next(
                (
                    parte.replace("mes=", "")
                    for parte in partes
                    if parte.startswith("mes=")
                ),
                "",
            )

            try:
                cantidad = (
                    pq.ParquetFile(
                        archivo
                    )
                    .metadata
                    .num_rows
                )

                estado = "OK"

            except Exception as error:
                cantidad = 0
                estado = f"ERROR: {error}"

            filas.append({
                "TABLA": tabla.upper(),
                "ANIO": anio,
                "MES": mes,
                "ARCHIVO": str(
                    archivo.relative_to(
                        CARPETA_STAGING
                    )
                ),
                "FILAS": cantidad,
                "ESTADO": estado,
            })

    return pd.DataFrame(filas)


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def construir_staging(
    forzar: bool,
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
            "No se encontraron archivos XLSX "
            "en la Landing Zone."
        )
        return

    print("========================================")
    print("CONSTRUCCIÓN DE STAGING OECE")
    print("========================================")
    print(f"Archivos encontrados: {len(archivos)}")

    manifest = cargar_manifest()

    controles = []
    errores = []

    procesados = 0
    omitidos = 0

    inicio = datetime.now().astimezone()

    for archivo in archivos:
        ruta_relativa = str(
            archivo.relative_to(
                CARPETA_RAW
            )
        ).replace("\\", "/")

        try:
            firma = calcular_sha256(
                archivo
            )

            estado_anterior = manifest.get(
                ruta_relativa,
                {},
            )

            if (
                not forzar
                and estado_anterior.get("sha256") == firma
            ):
                print(
                    f"OMITIDO SIN CAMBIOS: "
                    f"{ruta_relativa}"
                )

                omitidos += 1
                continue

            control_archivo, salidas = procesar_archivo(
                archivo
            )

            controles.extend(
                control_archivo
            )

            manifest[ruta_relativa] = {
                "sha256": firma,
                "procesado_en": (
                    datetime.now()
                    .astimezone()
                    .isoformat()
                ),
                "salidas": salidas,
            }

            guardar_manifest(
                manifest
            )

            procesados += 1

        except Exception as error:
            print(
                f"ERROR en {ruta_relativa}: "
                f"{error}"
            )

            errores.append({
                "ARCHIVO": ruta_relativa,
                "ERROR": str(error),
            })

    fin = datetime.now().astimezone()

    df_control = pd.DataFrame(
        controles
    )

    df_errores = pd.DataFrame(
        errores,
        columns=[
            "ARCHIVO",
            "ERROR",
        ],
    )

    df_resumen_global = generar_resumen_global()

    if not df_resumen_global.empty:
        resumen_tablas = (
            df_resumen_global
            .groupby(
                "TABLA",
                as_index=False,
            )
            .agg(
                ARCHIVOS_PARQUET=(
                    "ARCHIVO",
                    "count",
                ),
                TOTAL_FILAS=(
                    "FILAS",
                    "sum",
                ),
            )
        )
    else:
        resumen_tablas = pd.DataFrame(
            columns=[
                "TABLA",
                "ARCHIVOS_PARQUET",
                "TOTAL_FILAS",
            ]
        )

    resumen_ejecucion = pd.DataFrame([
        {
            "FECHA_INICIO": inicio.isoformat(),
            "FECHA_FIN": fin.isoformat(),
            "ARCHIVOS_ENCONTRADOS": len(archivos),
            "ARCHIVOS_PROCESADOS": procesados,
            "ARCHIVOS_OMITIDOS": omitidos,
            "ARCHIVOS_CON_ERROR": len(errores),
            "FORZADO": "SI" if forzar else "NO",
        }
    ])

    with pd.ExcelWriter(
        ARCHIVO_CONTROL,
        engine="openpyxl",
    ) as writer:

        resumen_ejecucion.to_excel(
            writer,
            sheet_name="RESUMEN_EJECUCION",
            index=False,
        )

        resumen_tablas.to_excel(
            writer,
            sheet_name="RESUMEN_TABLAS",
            index=False,
        )

        df_resumen_global.to_excel(
            writer,
            sheet_name="DETALLE_PARTICIONES",
            index=False,
        )

        df_control.to_excel(
            writer,
            sheet_name="CONTROL_PROCESAMIENTO",
            index=False,
        )

        df_errores.to_excel(
            writer,
            sheet_name="ERRORES",
            index=False,
        )

    # Guardar historial acumulativo
    escribir_encabezado = (
        not ARCHIVO_HISTORIAL.exists()
    )

    resumen_ejecucion.to_csv(
        ARCHIVO_HISTORIAL,
        mode="a",
        index=False,
        header=escribir_encabezado,
        encoding="utf-8-sig",
    )

    print("\n========================================")
    print("STAGING TERMINADO")
    print("========================================")
    print(f"Archivos encontrados: {len(archivos)}")
    print(f"Procesados: {procesados}")
    print(f"Sin cambios: {omitidos}")
    print(f"Errores: {len(errores)}")

    print("\nResumen de tablas:")

    if resumen_tablas.empty:
        print("No se generaron tablas.")
    else:
        print(
            resumen_tablas.to_string(
                index=False
            )
        )

    print("\nCarpeta staging:")
    print(CARPETA_STAGING)

    print("\nReporte de control:")
    print(ARCHIVO_CONTROL)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Construye las tablas staging "
            "históricas OCDS/OECE."
        )
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help=(
            "Procesa nuevamente todos los meses, "
            "aunque el archivo no haya cambiado."
        ),
    )

    argumentos = parser.parse_args()

    construir_staging(
        forzar=argumentos.force
    )


if __name__ == "__main__":
    main()
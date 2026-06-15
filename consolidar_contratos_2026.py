from pathlib import Path
import re

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
    / "2026"
)

CARPETA_PROCESSED = (
    BASE_DIR
    / "data"
    / "processed"
    / "oece"
    / "2026"
)

CARPETA_PROCESSED.mkdir(parents=True, exist_ok=True)

SALIDA_XLSX = (
    CARPETA_PROCESSED
    / "CONTRATOS_2026_CONSOLIDADO.xlsx"
)

SALIDA_REPORTE = (
    CARPETA_PROCESSED
    / "REPORTE_CONSOLIDACION_CONTRATOS_2026.txt"
)

HOJA_CONTRATOS = "Ent_Contratos"


# ==========================================================
# EQUIVALENCIA DE COLUMNAS
# ==========================================================

MAPA_COLUMNAS = {
    "Open Contracting ID": "OCID",

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
        "MONEDA",

    "Entrega compilada:Contratos:Valor:Nombre de Moneda":
        "NOMBRE_MONEDA",

    "Entrega compilada:Contratos:Fecha de firma":
        "FECHA_FIRMA",

    "Entrega compilada:Contratos:Implementación:Valor final:Monto":
        "MONTO_FINAL",

    "Entrega compilada:Contratos:Implementación:Valor final:Moneda":
        "MONEDA_FINAL",

    "Entrega compilada:Contratos:Implementación:Valor final:Nombre de Moneda":
        "NOMBRE_MONEDA_FINAL",

    "Entrega compilada:Contratos:Implementación:Fecha de fin":
        "FECHA_FIN_IMPLEMENTACION"
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
    "MONEDA",
    "NOMBRE_MONEDA",
    "FECHA_FIRMA",
    "MONTO_FINAL",
    "MONEDA_FINAL",
    "NOMBRE_MONEDA_FINAL",
    "FECHA_FIN_IMPLEMENTACION"
]


# ==========================================================
# FUNCIONES DE LIMPIEZA
# ==========================================================

def normalizar_encabezado(encabezado):
    """
    Limpia espacios, saltos de línea y tabulaciones del encabezado.

    No cambia el contenido de los datos.
    """

    if encabezado is None:
        return ""

    encabezado = str(encabezado)
    encabezado = encabezado.replace("\n", " ")
    encabezado = encabezado.replace("\r", " ")
    encabezado = encabezado.replace("\t", " ")
    encabezado = re.sub(r"\s+", " ", encabezado)

    return encabezado.strip()


def limpiar_celdas_texto(df):
    """
    Limpia espacios externos y convierte textos vacíos en valores nulos.

    No elimina comas.
    No reemplaza punto y coma.
    No modifica descripciones.
    """

    for columna in df.columns:
        df[columna] = df[columna].astype("string").str.strip()

        df[columna] = df[columna].replace(
            {
                "": pd.NA,
                "nan": pd.NA,
                "None": pd.NA,
                "<NA>": pd.NA
            }
        )

    return df


def obtener_mes_desde_ruta(archivo):
    """
    Ruta esperada:
    2026/01/extraido/archivo.xlsx
    """

    return archivo.parent.parent.name


# ==========================================================
# LEER UN MES
# ==========================================================

def leer_contratos_mes(archivo):
    mes = obtener_mes_desde_ruta(archivo)

    print("\n----------------------------------------")
    print(f"Archivo: {archivo.name}")
    print(f"Mes: {mes}")

    try:
        df = pd.read_excel(
            archivo,
            sheet_name=HOJA_CONTRATOS,
            dtype=str,
            engine="openpyxl"
        )

    except ValueError as error:
        raise RuntimeError(
            f"El archivo {archivo.name} no tiene "
            f"la hoja {HOJA_CONTRATOS}."
        ) from error

    filas_originales = len(df)

    # Limpiar encabezados
    df.columns = [
        normalizar_encabezado(columna)
        for columna in df.columns
    ]

    # Renombrar por el nombre del encabezado, no por posición
    df = df.rename(columns=MAPA_COLUMNAS)

    # Validar columnas principales
    columnas_criticas = [
        "OCID",
        "ID_CONTRATO"
    ]

    faltantes_criticas = [
        columna
        for columna in columnas_criticas
        if columna not in df.columns
    ]

    if faltantes_criticas:
        raise RuntimeError(
            f"Faltan columnas críticas en {archivo.name}: "
            f"{faltantes_criticas}"
        )

    # Crear como vacías las columnas opcionales que no existan
    for columna in COLUMNAS_CONTRATOS:
        if columna not in df.columns:
            df[columna] = pd.NA

    # Mantener únicamente la estructura estandarizada
    df = df[COLUMNAS_CONTRATOS].copy()

    df = limpiar_celdas_texto(df)

    # Eliminar filas completamente vacías
    df = df.dropna(how="all")

    # Un contrato válido debe tener OCID e ID del contrato
    filas_sin_ocid = int(df["OCID"].isna().sum())
    filas_sin_id_contrato = int(df["ID_CONTRATO"].isna().sum())

    df = df[
        df["OCID"].notna()
        & df["ID_CONTRATO"].notna()
    ].copy()

    # Trazabilidad
    df["ANIO_ARCHIVO"] = "2026"
    df["MES_ARCHIVO"] = mes
    df["ARCHIVO_ORIGEN"] = archivo.name

    # Llave recomendada
    df["CLAVE_CONTRATO"] = (
        df["OCID"].astype("string")
        + "|"
        + df["ID_CONTRATO"].astype("string")
    )

    resultado = {
        "archivo": archivo.name,
        "mes": mes,
        "filas_originales": filas_originales,
        "filas_validas": len(df),
        "filas_sin_ocid": filas_sin_ocid,
        "filas_sin_id_contrato": filas_sin_id_contrato,
        "columnas_originales": len(MAPA_COLUMNAS)
    }

    print(f"Filas originales: {filas_originales}")
    print(f"Contratos válidos: {len(df)}")
    print(f"Filas sin OCID: {filas_sin_ocid}")
    print(
        f"Filas sin ID de contrato: "
        f"{filas_sin_id_contrato}"
    )

    return df, resultado


# ==========================================================
# CONSOLIDACIÓN
# ==========================================================

def consolidar_contratos():
    archivos = sorted(
        CARPETA_RAW.glob(
            "*/extraido/*_seace_v3_es.xlsx"
        )
    )

    if not archivos:
        print(
            "No se encontraron archivos mensuales en:"
        )
        print(CARPETA_RAW)
        return

    print("Archivos encontrados:")

    for archivo in archivos:
        print(f"- {archivo}")

    dataframes = []
    resumen_archivos = []
    errores = []

    for archivo in archivos:
        try:
            df_mes, resumen_mes = leer_contratos_mes(
                archivo
            )

            dataframes.append(df_mes)
            resumen_archivos.append(resumen_mes)

        except Exception as error:
            print(
                f"ERROR procesando {archivo.name}: "
                f"{error}"
            )

            errores.append({
                "archivo": archivo.name,
                "error": str(error)
            })

    if not dataframes:
        print(
            "No se pudo procesar ningún archivo."
        )
        return

    # Todas las apariciones mensuales
    historico = pd.concat(
        dataframes,
        ignore_index=True,
        sort=False
    )

    historico["MES_NUMERO"] = pd.to_numeric(
        historico["MES_ARCHIVO"],
        errors="coerce"
    )

    # Ordenar para que la aparición más reciente quede al final
    historico = historico.sort_values(
        by=[
            "MES_NUMERO",
            "ARCHIVO_ORIGEN"
        ],
        ascending=True
    ).reset_index(drop=True)

    duplicados_clave = int(
        historico.duplicated(
            subset=["CLAVE_CONTRATO"],
            keep=False
        ).sum()
    )

    # Vista actual: una fila por contrato
    contratos_actuales = (
        historico
        .drop_duplicates(
            subset=["CLAVE_CONTRATO"],
            keep="last"
        )
        .copy()
    )

    contratos_actuales = contratos_actuales.sort_values(
        by=[
            "MES_NUMERO",
            "OCID",
            "ID_CONTRATO"
        ]
    ).reset_index(drop=True)

    # Resumen por mes
    resumen_mensual = (
        historico
        .groupby(
            ["ANIO_ARCHIVO", "MES_ARCHIVO"],
            dropna=False
        )
        .agg(
            REGISTROS_MENSUALES=(
                "CLAVE_CONTRATO",
                "size"
            ),
            CONTRATOS_UNICOS=(
                "CLAVE_CONTRATO",
                "nunique"
            ),
            MONTO_INFORMADO=(
                "MONTO_CONTRATO",
                lambda serie: int(
                    serie.notna().sum()
                )
            ),
            FECHA_FIRMA_INFORMADA=(
                "FECHA_FIRMA",
                lambda serie: int(
                    serie.notna().sum()
                )
            )
        )
        .reset_index()
        .sort_values("MES_ARCHIVO")
    )

    resumen_general = pd.DataFrame([
        {
            "ARCHIVOS_PROCESADOS": len(dataframes),
            "ARCHIVOS_CON_ERROR": len(errores),
            "FILAS_HISTORICAS": len(historico),
            "CONTRATOS_ACTUALES_UNICOS": len(
                contratos_actuales
            ),
            "FILAS_EN_CLAVES_REPETIDAS": duplicados_clave
        }
    ])

    resumen_archivos_df = pd.DataFrame(
        resumen_archivos
    )

    errores_df = pd.DataFrame(
        errores,
        columns=["archivo", "error"]
    )

    # Retirar columna auxiliar
    historico = historico.drop(
        columns=["MES_NUMERO"]
    )

    contratos_actuales = contratos_actuales.drop(
        columns=["MES_NUMERO"]
    )

    print("\nGuardando Excel consolidado...")

    with pd.ExcelWriter(
        SALIDA_XLSX,
        engine="openpyxl"
    ) as writer:

        contratos_actuales.to_excel(
            writer,
            sheet_name="CONTRATOS_ACTUALES",
            index=False
        )

        historico.to_excel(
            writer,
            sheet_name="HISTORICO_MENSUAL",
            index=False
        )

        resumen_mensual.to_excel(
            writer,
            sheet_name="RESUMEN_MENSUAL",
            index=False
        )

        resumen_archivos_df.to_excel(
            writer,
            sheet_name="ARCHIVOS_PROCESADOS",
            index=False
        )

        resumen_general.to_excel(
            writer,
            sheet_name="RESUMEN_GENERAL",
            index=False
        )

        if not errores_df.empty:
            errores_df.to_excel(
                writer,
                sheet_name="ERRORES",
                index=False
            )

    # Reporte de texto
    with open(
        SALIDA_REPORTE,
        "w",
        encoding="utf-8"
    ) as archivo:

        archivo.write(
            "CONSOLIDACIÓN DE CONTRATOS OECE 2026\n"
        )

        archivo.write("=" * 80)
        archivo.write("\n")

        archivo.write(
            f"Archivos encontrados: {len(archivos)}\n"
        )

        archivo.write(
            f"Archivos procesados: {len(dataframes)}\n"
        )

        archivo.write(
            f"Archivos con error: {len(errores)}\n"
        )

        archivo.write(
            f"Filas históricas: {len(historico)}\n"
        )

        archivo.write(
            "Contratos actuales únicos: "
            f"{len(contratos_actuales)}\n"
        )

        archivo.write(
            "Filas pertenecientes a claves repetidas: "
            f"{duplicados_clave}\n"
        )

        if errores:
            archivo.write("\nERRORES\n")

            for error in errores:
                archivo.write(
                    f"- {error['archivo']}: "
                    f"{error['error']}\n"
                )

    print("\n========================================")
    print("CONSOLIDACIÓN TERMINADA")
    print("========================================")
    print(f"Archivos procesados: {len(dataframes)}")
    print(f"Filas históricas: {len(historico)}")
    print(
        f"Contratos únicos actuales: "
        f"{len(contratos_actuales)}"
    )
    print(
        f"Filas en claves repetidas: "
        f"{duplicados_clave}"
    )
    print(f"Excel generado: {SALIDA_XLSX}")
    print(f"Reporte generado: {SALIDA_REPORTE}")


if __name__ == "__main__":
    consolidar_contratos()
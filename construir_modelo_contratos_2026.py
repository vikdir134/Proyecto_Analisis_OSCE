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

ARCHIVO_CONTRATOS = (
    CARPETA_PROCESSED
    / "CONTRATOS_2026_CONSOLIDADO.xlsx"
)

SALIDA_MODELO = (
    CARPETA_PROCESSED
    / "MODELO_CONTRATOS_OECE_2026.xlsx"
)

SALIDA_REPORTE = (
    CARPETA_PROCESSED
    / "REPORTE_MODELO_CONTRATOS_2026.txt"
)


# ==========================================================
# COLUMNAS DE REGISTROS
# ==========================================================

MAPA_REGISTROS = {
    "Open Contracting ID":
        "OCID",

    "Entrega compilada:ID de Entrega":
        "ID_ENTREGA",

    "Entrega compilada:Fecha de entrega":
        "FECHA_ENTREGA",

    "compiledrelease/publisheddate":
        "FECHA_PUBLICACION",

    "Entrega compilada:Comprador:ID de Organización":
        "ID_COMPRADOR",

    "Entrega compilada:Comprador:Nombre de la Organización":
        "NOMBRE_COMPRADOR",

    "Entrega compilada:Licitación:ID de licitación":
        "ID_LICITACION",

    "Entrega compilada:Licitación:Título de la licitación":
        "TITULO_LICITACION",

    "Entrega compilada:Licitación:Descripción de la licitación":
        "DESCRIPCION_LICITACION",

    "Entrega compilada:Licitación:Entidad contratante:ID de Organización":
        "ID_ENTIDAD_CONTRATANTE",

    "Entrega compilada:Licitación:Entidad contratante:Nombre de la Organización":
        "NOMBRE_ENTIDAD_CONTRATANTE",

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
        "MONTO_LICITACION_PEN"
}


COLUMNAS_REGISTROS = [
    "OCID",
    "ID_ENTREGA",
    "FECHA_ENTREGA",
    "FECHA_PUBLICACION",
    "ID_COMPRADOR",
    "NOMBRE_COMPRADOR",
    "ID_LICITACION",
    "TITULO_LICITACION",
    "DESCRIPCION_LICITACION",
    "ID_ENTIDAD_CONTRATANTE",
    "NOMBRE_ENTIDAD_CONTRATANTE",
    "FECHA_PUBLICACION_LICITACION",
    "METODO_CONTRATACION",
    "DETALLE_METODO_CONTRATACION",
    "CATEGORIA_PRINCIPAL",
    "CATEGORIAS_ADICIONALES",
    "MONTO_LICITACION",
    "MONEDA_LICITACION",
    "NOMBRE_MONEDA_LICITACION",
    "MONTO_LICITACION_PEN"
]


# ==========================================================
# COLUMNAS DE PROVEEDORES
# ==========================================================

MAPA_PROVEEDORES = {
    "Open Contracting ID":
        "OCID",

    "Entrega compilada:ID de Entrega":
        "ID_ENTREGA",

    "Entrega compilada:Adjudicaciones:ID de Adjudicación":
        "ID_ADJUDICACION",

    "Entrega compilada:Adjudicaciones:Proveedores:ID de Organización":
        "ID_PROVEEDOR",

    "Entrega compilada:Adjudicaciones:Proveedores:Nombre de la Organización":
        "NOMBRE_PROVEEDOR"
}


COLUMNAS_PROVEEDORES = [
    "OCID",
    "ID_ENTREGA",
    "ID_ADJUDICACION",
    "ID_PROVEEDOR",
    "NOMBRE_PROVEEDOR"
]


# ==========================================================
# LIMPIEZA
# ==========================================================

def normalizar_encabezado(valor):
    if valor is None:
        return ""

    valor = str(valor)
    valor = valor.replace("\n", " ")
    valor = valor.replace("\r", " ")
    valor = valor.replace("\t", " ")
    valor = re.sub(r"\s+", " ", valor)

    return valor.strip()


def normalizar_id(serie):
    """
    Normaliza identificadores sin eliminar ceros iniciales.

    Solo elimina espacios y terminaciones .0 generadas por Excel.
    """

    serie = serie.astype("string").str.strip()

    serie = serie.replace({
        "": pd.NA,
        "nan": pd.NA,
        "None": pd.NA,
        "<NA>": pd.NA
    })

    serie = serie.str.replace(
        r"^(\d+)\.0$",
        r"\1",
        regex=True
    )

    return serie


def limpiar_texto(serie):
    serie = serie.astype("string").str.strip()

    serie = serie.replace({
        "": pd.NA,
        "nan": pd.NA,
        "None": pd.NA,
        "<NA>": pd.NA
    })

    return serie


def agregar_columnas_faltantes(df, columnas):
    for columna in columnas:
        if columna not in df.columns:
            df[columna] = pd.NA

    return df[columnas].copy()


def obtener_mes(archivo):
    """
    Ruta esperada:
    data/raw/oece/2026/01/extraido/archivo.xlsx
    """

    return archivo.parent.parent.name


# ==========================================================
# LEER HOJA REGISTROS
# ==========================================================

def leer_registros(archivo):
    mes = obtener_mes(archivo)

    df = pd.read_excel(
        archivo,
        sheet_name="Registros",
        dtype=str,
        engine="openpyxl"
    )

    df.columns = [
        normalizar_encabezado(columna)
        for columna in df.columns
    ]

    df = df.rename(columns=MAPA_REGISTROS)
    df = agregar_columnas_faltantes(
        df,
        COLUMNAS_REGISTROS
    )

    for columna in [
        "OCID",
        "ID_ENTREGA",
        "ID_COMPRADOR",
        "ID_LICITACION",
        "ID_ENTIDAD_CONTRATANTE"
    ]:
        df[columna] = normalizar_id(df[columna])

    for columna in df.columns:
        if columna not in [
            "OCID",
            "ID_ENTREGA",
            "ID_COMPRADOR",
            "ID_LICITACION",
            "ID_ENTIDAD_CONTRATANTE"
        ]:
            df[columna] = limpiar_texto(df[columna])

    df = df[
        df["OCID"].notna()
        & df["ID_ENTREGA"].notna()
    ].copy()

    df["MES_ARCHIVO"] = mes
    df["MES_NUMERO"] = pd.to_numeric(
        mes,
        errors="coerce"
    )
    df["ARCHIVO_ORIGEN_REGISTRO"] = archivo.name

    return df


# ==========================================================
# LEER PROVEEDORES
# ==========================================================

def leer_proveedores(archivo):
    mes = obtener_mes(archivo)

    df = pd.read_excel(
        archivo,
        sheet_name="Ent_Adj_Proveedores",
        dtype=str,
        engine="openpyxl"
    )

    df.columns = [
        normalizar_encabezado(columna)
        for columna in df.columns
    ]

    df = df.rename(columns=MAPA_PROVEEDORES)
    df = agregar_columnas_faltantes(
        df,
        COLUMNAS_PROVEEDORES
    )

    for columna in [
        "OCID",
        "ID_ENTREGA",
        "ID_ADJUDICACION",
        "ID_PROVEEDOR"
    ]:
        df[columna] = normalizar_id(df[columna])

    df["NOMBRE_PROVEEDOR"] = limpiar_texto(
        df["NOMBRE_PROVEEDOR"]
    )

    df = df[
        df["OCID"].notna()
        & df["ID_ENTREGA"].notna()
        & df["ID_ADJUDICACION"].notna()
        & df["ID_PROVEEDOR"].notna()
    ].copy()

    df["MES_ARCHIVO"] = mes
    df["MES_NUMERO"] = pd.to_numeric(
        mes,
        errors="coerce"
    )
    df["ARCHIVO_ORIGEN_PROVEEDOR"] = archivo.name

    return df


# ==========================================================
# CARGAR CONTRATOS CONSOLIDADOS
# ==========================================================

def cargar_contratos():
    if not ARCHIVO_CONTRATOS.exists():
        raise FileNotFoundError(
            f"No existe el archivo:\n{ARCHIVO_CONTRATOS}"
        )

    contratos = pd.read_excel(
        ARCHIVO_CONTRATOS,
        sheet_name="CONTRATOS_ACTUALES",
        dtype=str,
        engine="openpyxl"
    )

    for columna in [
        "OCID",
        "ID_ENTREGA",
        "ID_CONTRATO",
        "ID_ADJUDICACION",
        "CLAVE_CONTRATO"
    ]:
        if columna not in contratos.columns:
            raise RuntimeError(
                f"Falta la columna {columna} "
                "en CONTRATOS_ACTUALES."
            )

        contratos[columna] = normalizar_id(
            contratos[columna]
        )

    return contratos


# ==========================================================
# FUNCIONES DE AGREGACIÓN
# ==========================================================

def concatenar_unicos(serie):
    valores = (
        serie
        .dropna()
        .astype(str)
        .str.strip()
    )

    valores = [
        valor
        for valor in valores.unique()
        if valor
    ]

    return " | ".join(sorted(valores))


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def construir_modelo():
    archivos = sorted(
        CARPETA_RAW.glob(
            "*/extraido/*_seace_v3_es.xlsx"
        )
    )

    if not archivos:
        print(
            "No se encontraron archivos mensuales."
        )
        return

    print("Archivos encontrados:")

    for archivo in archivos:
        print(f"- {archivo}")

    contratos = cargar_contratos()

    registros_lista = []
    proveedores_lista = []
    control = []
    errores = []

    for archivo in archivos:
        mes = obtener_mes(archivo)

        print("\n----------------------------------------")
        print(f"Procesando mes {mes}: {archivo.name}")

        try:
            registros_mes = leer_registros(archivo)
            proveedores_mes = leer_proveedores(archivo)

            registros_lista.append(registros_mes)
            proveedores_lista.append(proveedores_mes)

            control.append({
                "MES": mes,
                "ARCHIVO": archivo.name,
                "REGISTROS": len(registros_mes),
                "PROVEEDORES": len(proveedores_mes),
                "ESTADO": "OK"
            })

            print(
                f"Registros válidos: "
                f"{len(registros_mes)}"
            )

            print(
                f"Proveedores válidos: "
                f"{len(proveedores_mes)}"
            )

        except Exception as error:
            print(f"ERROR: {error}")

            errores.append({
                "MES": mes,
                "ARCHIVO": archivo.name,
                "ERROR": str(error)
            })

            control.append({
                "MES": mes,
                "ARCHIVO": archivo.name,
                "REGISTROS": 0,
                "PROVEEDORES": 0,
                "ESTADO": f"ERROR: {error}"
            })

    if not registros_lista:
        print(
            "No se pudo leer ninguna hoja Registros."
        )
        return

    registros_historicos = pd.concat(
        registros_lista,
        ignore_index=True,
        sort=False
    )

    proveedores_historicos = pd.concat(
        proveedores_lista,
        ignore_index=True,
        sort=False
    )

    # ======================================================
    # ÚLTIMO REGISTRO DISPONIBLE
    # ======================================================

    registros_actuales = (
        registros_historicos
        .sort_values(
            by=[
                "MES_NUMERO",
                "ARCHIVO_ORIGEN_REGISTRO"
            ]
        )
        .drop_duplicates(
            subset=[
                "OCID",
                "ID_ENTREGA"
            ],
            keep="last"
        )
        .copy()
    )

    # ======================================================
    # ÚLTIMA VERSIÓN DE PROVEEDORES
    # ======================================================

    proveedores_actuales = (
        proveedores_historicos
        .sort_values(
            by=[
                "MES_NUMERO",
                "ARCHIVO_ORIGEN_PROVEEDOR"
            ]
        )
        .drop_duplicates(
            subset=[
                "OCID",
                "ID_ENTREGA",
                "ID_ADJUDICACION",
                "ID_PROVEEDOR"
            ],
            keep="last"
        )
        .copy()
    )

    # ======================================================
    # FACT CONTRATOS + REGISTROS
    # ======================================================

    columnas_registro_fact = [
        "OCID",
        "ID_ENTREGA",
        "FECHA_ENTREGA",
        "FECHA_PUBLICACION",
        "ID_COMPRADOR",
        "NOMBRE_COMPRADOR",
        "ID_LICITACION",
        "TITULO_LICITACION",
        "DESCRIPCION_LICITACION",
        "ID_ENTIDAD_CONTRATANTE",
        "NOMBRE_ENTIDAD_CONTRATANTE",
        "FECHA_PUBLICACION_LICITACION",
        "METODO_CONTRATACION",
        "DETALLE_METODO_CONTRATACION",
        "CATEGORIA_PRINCIPAL",
        "CATEGORIAS_ADICIONALES",
        "MONTO_LICITACION",
        "MONEDA_LICITACION",
        "NOMBRE_MONEDA_LICITACION",
        "MONTO_LICITACION_PEN"
    ]

    fact_contratos = contratos.merge(
        registros_actuales[columnas_registro_fact],
        how="left",
        on=[
            "OCID",
            "ID_ENTREGA"
        ],
        validate="many_to_one"
    )

    fact_contratos["REGISTRO_RELACIONADO"] = (
        fact_contratos["ID_LICITACION"]
        .notna()
        .map({
            True: "SI",
            False: "NO"
        })
    )

    # Usar entidad contratante.
    # Si está vacía, usar comprador como respaldo.
    fact_contratos["ID_ENTIDAD_MODELO"] = (
        fact_contratos["ID_ENTIDAD_CONTRATANTE"]
        .fillna(fact_contratos["ID_COMPRADOR"])
    )

    fact_contratos["NOMBRE_ENTIDAD_MODELO"] = (
        fact_contratos["NOMBRE_ENTIDAD_CONTRATANTE"]
        .fillna(fact_contratos["NOMBRE_COMPRADOR"])
    )

    fact_contratos["ORIGEN_ENTIDAD_MODELO"] = "SIN_ENTIDAD"

    mascara_entidad = (
        fact_contratos["ID_ENTIDAD_CONTRATANTE"]
        .notna()
    )

    mascara_comprador = (
        fact_contratos["ID_ENTIDAD_CONTRATANTE"]
        .isna()
        & fact_contratos["ID_COMPRADOR"].notna()
    )

    fact_contratos.loc[
        mascara_entidad,
        "ORIGEN_ENTIDAD_MODELO"
    ] = "ENTIDAD_CONTRATANTE"

    fact_contratos.loc[
        mascara_comprador,
        "ORIGEN_ENTIDAD_MODELO"
    ] = "COMPRADOR_RESPALDO"

    # ======================================================
    # TABLA PUENTE CONTRATO-PROVEEDOR
    # ======================================================

    claves_contrato = fact_contratos[[
        "CLAVE_CONTRATO",
        "OCID",
        "ID_ENTREGA",
        "ID_ADJUDICACION"
    ]].drop_duplicates()

    bridge_contrato_proveedor = claves_contrato.merge(
        proveedores_actuales[[
            "OCID",
            "ID_ENTREGA",
            "ID_ADJUDICACION",
            "ID_PROVEEDOR",
            "NOMBRE_PROVEEDOR",
            "MES_ARCHIVO",
            "ARCHIVO_ORIGEN_PROVEEDOR"
        ]],
        how="inner",
        on=[
            "OCID",
            "ID_ENTREGA",
            "ID_ADJUDICACION"
        ],
        validate="many_to_many"
    )

    bridge_contrato_proveedor = (
        bridge_contrato_proveedor
        .drop_duplicates(
            subset=[
                "CLAVE_CONTRATO",
                "ID_PROVEEDOR"
            ]
        )
        .reset_index(drop=True)
    )

    # ======================================================
    # RESUMEN DE PROVEEDORES POR CONTRATO
    # ======================================================

    if not bridge_contrato_proveedor.empty:
        resumen_proveedores = (
            bridge_contrato_proveedor
            .groupby(
                "CLAVE_CONTRATO",
                dropna=False
            )
            .agg(
                CANTIDAD_PROVEEDORES=(
                    "ID_PROVEEDOR",
                    "nunique"
                ),
                IDS_PROVEEDORES=(
                    "ID_PROVEEDOR",
                    concatenar_unicos
                ),
                NOMBRES_PROVEEDORES=(
                    "NOMBRE_PROVEEDOR",
                    concatenar_unicos
                )
            )
            .reset_index()
        )

        fact_contratos = fact_contratos.merge(
            resumen_proveedores,
            how="left",
            on="CLAVE_CONTRATO",
            validate="one_to_one"
        )

    else:
        fact_contratos["CANTIDAD_PROVEEDORES"] = 0
        fact_contratos["IDS_PROVEEDORES"] = pd.NA
        fact_contratos["NOMBRES_PROVEEDORES"] = pd.NA

    fact_contratos["CANTIDAD_PROVEEDORES"] = (
        pd.to_numeric(
            fact_contratos["CANTIDAD_PROVEEDORES"],
            errors="coerce"
        )
        .fillna(0)
        .astype(int)
    )

    fact_contratos["TIENE_PROVEEDOR"] = (
        fact_contratos["CANTIDAD_PROVEEDORES"]
        .gt(0)
        .map({
            True: "SI",
            False: "NO"
        })
    )

    # ======================================================
    # DIMENSIÓN PROVEEDORES
    # ======================================================

    dim_proveedores = (
        proveedores_actuales
        .sort_values(
            by=[
                "MES_NUMERO",
                "ARCHIVO_ORIGEN_PROVEEDOR"
            ]
        )
        .dropna(
            subset=["ID_PROVEEDOR"]
        )
        .drop_duplicates(
            subset=["ID_PROVEEDOR"],
            keep="last"
        )[[
            "ID_PROVEEDOR",
            "NOMBRE_PROVEEDOR"
        ]]
        .sort_values("ID_PROVEEDOR")
        .reset_index(drop=True)
    )

    # ======================================================
    # DIMENSIÓN ENTIDADES
    # ======================================================

    entidades = registros_actuales.copy()

    entidades["ID_ENTIDAD_MODELO"] = (
        entidades["ID_ENTIDAD_CONTRATANTE"]
        .fillna(entidades["ID_COMPRADOR"])
    )

    entidades["NOMBRE_ENTIDAD_MODELO"] = (
        entidades["NOMBRE_ENTIDAD_CONTRATANTE"]
        .fillna(entidades["NOMBRE_COMPRADOR"])
    )

    entidades["ORIGEN_ENTIDAD_MODELO"] = (
        entidades["ID_ENTIDAD_CONTRATANTE"]
        .notna()
        .map({
            True: "ENTIDAD_CONTRATANTE",
            False: "COMPRADOR_RESPALDO"
        })
    )

    dim_entidades = (
        entidades
        .sort_values(
            by=[
                "MES_NUMERO",
                "ARCHIVO_ORIGEN_REGISTRO"
            ]
        )
        .dropna(
            subset=["ID_ENTIDAD_MODELO"]
        )
        .drop_duplicates(
            subset=["ID_ENTIDAD_MODELO"],
            keep="last"
        )[[
            "ID_ENTIDAD_MODELO",
            "NOMBRE_ENTIDAD_MODELO",
            "ORIGEN_ENTIDAD_MODELO"
        ]]
        .sort_values("ID_ENTIDAD_MODELO")
        .reset_index(drop=True)
    )

    # ======================================================
    # CALIDAD
    # ======================================================

    total_contratos = len(fact_contratos)

    calidad = pd.DataFrame([
        {
            "INDICADOR": "Total de contratos",
            "CANTIDAD": total_contratos
        },
        {
            "INDICADOR": "Claves de contrato únicas",
            "CANTIDAD": fact_contratos[
                "CLAVE_CONTRATO"
            ].nunique()
        },
        {
            "INDICADOR": "Filas duplicadas por clave",
            "CANTIDAD": int(
                fact_contratos.duplicated(
                    subset=["CLAVE_CONTRATO"]
                ).sum()
            )
        },
        {
            "INDICADOR": "Contratos sin registro relacionado",
            "CANTIDAD": int(
                fact_contratos[
                    "ID_LICITACION"
                ].isna().sum()
            )
        },
        {
            "INDICADOR": "Contratos sin entidad",
            "CANTIDAD": int(
                fact_contratos[
                    "ID_ENTIDAD_MODELO"
                ].isna().sum()
            )
        },
        {
            "INDICADOR": "Contratos sin proveedor",
            "CANTIDAD": int(
                fact_contratos[
                    "CANTIDAD_PROVEEDORES"
                ].eq(0).sum()
            )
        },
        {
            "INDICADOR": "Contratos con varios proveedores",
            "CANTIDAD": int(
                fact_contratos[
                    "CANTIDAD_PROVEEDORES"
                ].gt(1).sum()
            )
        },
        {
            "INDICADOR": "Relaciones contrato-proveedor",
            "CANTIDAD": len(
                bridge_contrato_proveedor
            )
        },
        {
            "INDICADOR": "Proveedores únicos",
            "CANTIDAD": len(dim_proveedores)
        },
        {
            "INDICADOR": "Entidades únicas",
            "CANTIDAD": len(dim_entidades)
        }
    ])

    control_df = pd.DataFrame(control)

    errores_df = pd.DataFrame(
        errores,
        columns=[
            "MES",
            "ARCHIVO",
            "ERROR"
        ]
    )

    # Retirar columnas auxiliares
    registros_actuales_exportar = (
        registros_actuales
        .drop(
            columns=["MES_NUMERO"],
            errors="ignore"
        )
    )

    proveedores_actuales_exportar = (
        proveedores_actuales
        .drop(
            columns=["MES_NUMERO"],
            errors="ignore"
        )
    )

    # ======================================================
    # GUARDAR EXCEL
    # ======================================================

    print("\nGuardando modelo dimensional...")

    with pd.ExcelWriter(
        SALIDA_MODELO,
        engine="openpyxl"
    ) as writer:

        fact_contratos.to_excel(
            writer,
            sheet_name="FACT_CONTRATOS",
            index=False
        )

        bridge_contrato_proveedor.to_excel(
            writer,
            sheet_name="BRIDGE_CONTRATO_PROV",
            index=False
        )

        dim_proveedores.to_excel(
            writer,
            sheet_name="DIM_PROVEEDORES",
            index=False
        )

        dim_entidades.to_excel(
            writer,
            sheet_name="DIM_ENTIDADES",
            index=False
        )

        registros_actuales_exportar.to_excel(
            writer,
            sheet_name="REGISTROS_ACTUALES",
            index=False
        )

        proveedores_actuales_exportar.to_excel(
            writer,
            sheet_name="PROVEEDORES_ACTUALES",
            index=False
        )

        calidad.to_excel(
            writer,
            sheet_name="CALIDAD",
            index=False
        )

        control_df.to_excel(
            writer,
            sheet_name="CONTROL_ARCHIVOS",
            index=False
        )

        if not errores_df.empty:
            errores_df.to_excel(
                writer,
                sheet_name="ERRORES",
                index=False
            )

    # ======================================================
    # REPORTE
    # ======================================================

    with open(
        SALIDA_REPORTE,
        "w",
        encoding="utf-8"
    ) as reporte:

        reporte.write(
            "MODELO DE CONTRATOS OECE 2026\n"
        )
        reporte.write("=" * 80)
        reporte.write("\n")

        for _, fila in calidad.iterrows():
            reporte.write(
                f"{fila['INDICADOR']}: "
                f"{fila['CANTIDAD']}\n"
            )

        if errores:
            reporte.write("\nERRORES\n")

            for error in errores:
                reporte.write(
                    f"- Mes {error['MES']} | "
                    f"{error['ARCHIVO']} | "
                    f"{error['ERROR']}\n"
                )

    print("\n========================================")
    print("MODELO GENERADO")
    print("========================================")
    print(f"Contratos: {len(fact_contratos)}")
    print(
        "Relaciones contrato-proveedor: "
        f"{len(bridge_contrato_proveedor)}"
    )
    print(
        f"Proveedores: {len(dim_proveedores)}"
    )
    print(
        f"Entidades: {len(dim_entidades)}"
    )
    print(f"Archivo: {SALIDA_MODELO}")
    print(f"Reporte: {SALIDA_REPORTE}")


if __name__ == "__main__":
    construir_modelo()
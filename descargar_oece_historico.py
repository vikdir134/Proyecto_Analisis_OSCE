import argparse
import hashlib
import json
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote

import requests


# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

CARPETA_OECE = (
    BASE_DIR
    / "data"
    / "raw"
    / "oece"
)

CARPETA_OECE.mkdir(parents=True, exist_ok=True)

ARCHIVO_ESTADO_GLOBAL = (
    CARPETA_OECE
    / "estado_descargas_2022_2026.json"
)

URL_CATALOGO = (
    "https://contratacionesabiertas.oece.gob.pe/"
    "api/v1/files"
)

ANIOS_PREDETERMINADOS = [
    2022,
    2023,
    2024,
    2025,
    2026
]

FUENTE = "seace_v3"

FORMATOS_PREFERIDOS = [
    "xlsx_es",
    "xlsx"
]

REGISTROS_POR_PAGINA = 100

HEADERS_GENERALES = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/149 Safari/537.36"
    ),
    "Accept-Language": "es-PE,es;q=0.9"
}


# ==========================================================
# ESTADO DE DESCARGAS
# ==========================================================

def leer_json(ruta):
    if not ruta.exists():
        return {}

    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            return json.load(archivo)

    except (json.JSONDecodeError, OSError):
        return {}


def normalizar_ruta_estado(ruta, anio):
    """
    Convierte rutas antiguas relativas al año en rutas relativas
    a la carpeta general data/raw/oece.
    """

    if not ruta:
        return ruta

    ruta = str(ruta).replace("\\", "/")

    if ruta.startswith(f"{anio}/"):
        return ruta

    return f"{anio}/{ruta}"


def migrar_estado_anterior(estado, anios):
    """
    Reutiliza archivos descargados por el programa anterior,
    por ejemplo:

    data/raw/oece/2026/estado_descargas_2026.json
    """

    for anio in anios:
        archivo_anterior = (
            CARPETA_OECE
            / str(anio)
            / f"estado_descargas_{anio}.json"
        )

        datos_anteriores = leer_json(archivo_anterior)

        if not isinstance(datos_anteriores, dict):
            continue

        for identificador, registro in datos_anteriores.items():
            if not isinstance(registro, dict):
                continue

            if identificador in estado:
                continue

            registro_migrado = dict(registro)

            registro_migrado["file"] = normalizar_ruta_estado(
                registro_migrado.get("file"),
                anio
            )

            registro_migrado["extracted_folder"] = normalizar_ruta_estado(
                registro_migrado.get("extracted_folder"),
                anio
            )

            carpeta_extraida = registro_migrado.get(
                "extracted_folder"
            )

            archivos_extraidos = registro_migrado.get(
                "extracted_files",
                []
            )

            rutas_completas = []

            if carpeta_extraida:
                for archivo in archivos_extraidos:
                    ruta = (
                        Path(carpeta_extraida)
                        / archivo
                    )

                    rutas_completas.append(
                        str(ruta).replace("\\", "/")
                    )

            registro_migrado["extracted_files"] = rutas_completas

            estado[identificador] = registro_migrado

    return estado


def cargar_estado(anios):
    contenido = leer_json(ARCHIVO_ESTADO_GLOBAL)

    if isinstance(contenido, dict) and "paquetes" in contenido:
        estado = contenido.get("paquetes", {})
    elif isinstance(contenido, dict):
        estado = contenido
    else:
        estado = {}

    estado = migrar_estado_anterior(
        estado,
        anios
    )

    return estado


def guardar_estado(estado):
    contenido = {
        "ultima_ejecucion": datetime.now()
        .astimezone()
        .isoformat(),

        "fuente": FUENTE,

        "paquetes": estado
    }

    with open(
        ARCHIVO_ESTADO_GLOBAL,
        "w",
        encoding="utf-8"
    ) as archivo:

        json.dump(
            contenido,
            archivo,
            ensure_ascii=False,
            indent=2
        )


# ==========================================================
# CONSULTA DEL CATÁLOGO
# ==========================================================

def encontrar_paquetes(objeto):
    """
    Busca recursivamente paquetes con estructura:

    {
        "id": "seace_v3-2026-06",
        "files": {...},
        "year": "2026",
        "month": "06",
        "source": "seace_v3"
    }
    """

    encontrados = []

    if isinstance(objeto, dict):
        es_paquete = (
            isinstance(objeto.get("files"), dict)
            and "year" in objeto
            and "month" in objeto
            and "source" in objeto
        )

        if es_paquete:
            encontrados.append(objeto)

        for valor in objeto.values():
            encontrados.extend(
                encontrar_paquetes(valor)
            )

    elif isinstance(objeto, list):
        for elemento in objeto:
            encontrados.extend(
                encontrar_paquetes(elemento)
            )

    return encontrados


def elegir_formato(paquete):
    archivos = paquete.get("files", {})

    for formato in FORMATOS_PREFERIDOS:
        url = archivos.get(formato)

        if url:
            return formato, url

    return None, None


def consultar_paquetes_anio(session, anio):
    """
    Consulta todos los paquetes de un año.
    Incluye paginación y protección contra páginas repetidas.
    """

    paquetes_encontrados = {}
    pagina = 1

    print("\n========================================")
    print(f"CONSULTANDO AÑO {anio}")
    print("========================================")

    while True:
        parametros = {
            "page": pagina,
            "paginateBy": REGISTROS_POR_PAGINA,
            "year": anio,
            "format": "json"
        }

        respuesta = session.get(
            URL_CATALOGO,
            params=parametros,
            headers={
                "Accept": "application/json"
            },
            timeout=90
        )

        print(
            f"Página {pagina} | "
            f"HTTP {respuesta.status_code}"
        )

        respuesta.raise_for_status()

        try:
            contenido = respuesta.json()

        except ValueError as error:
            raise RuntimeError(
                f"El catálogo del año {anio} "
                "no devolvió JSON."
            ) from error

        paquetes_pagina = encontrar_paquetes(
            contenido
        )

        if not paquetes_pagina:
            break

        nuevos_en_pagina = 0

        for paquete in paquetes_pagina:
            identificador = paquete.get("id")

            if not identificador:
                identificador = (
                    f"{paquete.get('source')}-"
                    f"{paquete.get('year')}-"
                    f"{paquete.get('month')}"
                )

            if identificador not in paquetes_encontrados:
                nuevos_en_pagina += 1

            paquetes_encontrados[
                identificador
            ] = paquete

        # Evita un ciclo infinito si la API repite la página.
        if nuevos_en_pagina == 0:
            break

        if len(paquetes_pagina) < REGISTROS_POR_PAGINA:
            break

        pagina += 1

    resultado = []

    for paquete in paquetes_encontrados.values():
        if str(paquete.get("year")) != str(anio):
            continue

        if paquete.get("source") != FUENTE:
            continue

        formato, url = elegir_formato(paquete)

        if not formato or not url:
            print(
                f"Sin XLSX disponible: "
                f"{paquete.get('id')}"
            )
            continue

        paquete_preparado = dict(paquete)
        paquete_preparado["_formato_elegido"] = formato
        paquete_preparado["_url_descarga"] = url

        resultado.append(paquete_preparado)

    resultado.sort(
        key=lambda paquete: int(
            paquete.get("month", 0)
        )
    )

    print(
        f"Paquetes válidos encontrados para {anio}: "
        f"{len(resultado)}"
    )

    for paquete in resultado:
        print(
            f"- {paquete.get('id')} | "
            f"mes {paquete.get('month')} | "
            f"{paquete.get('_formato_elegido')}"
        )

    return resultado


# ==========================================================
# NOMBRES Y ARCHIVOS
# ==========================================================

def limpiar_nombre_archivo(nombre):
    nombre = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        nombre
    )

    return nombre.strip().strip(".")


def obtener_nombre_servidor(respuesta):
    content_disposition = respuesta.headers.get(
        "Content-Disposition",
        ""
    )

    coincidencia_utf8 = re.search(
        r"filename\*=UTF-8''([^;]+)",
        content_disposition,
        flags=re.IGNORECASE
    )

    if coincidencia_utf8:
        return limpiar_nombre_archivo(
            unquote(
                coincidencia_utf8.group(1)
            )
        )

    coincidencia_normal = re.search(
        r'filename="?([^";]+)"?',
        content_disposition,
        flags=re.IGNORECASE
    )

    if coincidencia_normal:
        return limpiar_nombre_archivo(
            coincidencia_normal.group(1)
        )

    return None


def determinar_extension(
    respuesta,
    nombre_servidor
):
    if nombre_servidor:
        extension = Path(
            nombre_servidor
        ).suffix.lower()

        if extension:
            return extension

    content_type = respuesta.headers.get(
        "Content-Type",
        ""
    ).lower()

    if "zip" in content_type:
        return ".zip"

    if "spreadsheetml" in content_type:
        return ".xlsx"

    if "excel" in content_type:
        return ".xlsx"

    return ".zip"


def es_archivo_xlsx_interno(ruta):
    """
    Tanto un ZIP como un XLSX usan estructura ZIP.

    Esta función identifica si el archivo ZIP realmente
    es un XLSX revisando su estructura interna.
    """

    if not zipfile.is_zipfile(ruta):
        return False

    with zipfile.ZipFile(ruta, "r") as archivo_zip:
        nombres = set(
            archivo_zip.namelist()
        )

    tiene_content_types = (
        "[Content_Types].xml" in nombres
    )

    tiene_carpeta_excel = any(
        nombre.startswith("xl/")
        for nombre in nombres
    )

    return (
        tiene_content_types
        and tiene_carpeta_excel
    )


def extraer_zip_seguro(ruta_zip, destino):
    """
    Extrae el ZIP evitando rutas fuera de la carpeta destino.
    """

    destino = destino.resolve()

    with zipfile.ZipFile(
        ruta_zip,
        "r"
    ) as archivo_zip:

        for miembro in archivo_zip.infolist():
            ruta_destino = (
                destino
                / miembro.filename
            ).resolve()

            if (
                ruta_destino != destino
                and destino not in ruta_destino.parents
            ):
                raise RuntimeError(
                    "El ZIP contiene una ruta insegura: "
                    f"{miembro.filename}"
                )

        archivo_zip.extractall(destino)


# ==========================================================
# VALIDACIÓN DEL ESTADO
# ==========================================================

def descarga_anterior_valida(
    registro,
    timestamp_api
):
    if not registro:
        return False

    mismo_timestamp = (
        registro.get("timestamp")
        == timestamp_api
    )

    if not mismo_timestamp:
        return False

    archivo_relativo = registro.get("file")

    if not archivo_relativo:
        return False

    ruta_archivo = (
        CARPETA_OECE
        / archivo_relativo
    )

    if not ruta_archivo.exists():
        return False

    archivos_extraidos = registro.get(
        "extracted_files",
        []
    )

    if archivos_extraidos:
        for archivo in archivos_extraidos:
            ruta = (
                CARPETA_OECE
                / archivo
            )

            if not ruta.exists():
                return False

        return True

    carpeta_extraida = registro.get(
        "extracted_folder"
    )

    if carpeta_extraida:
        ruta_extraida = (
            CARPETA_OECE
            / carpeta_extraida
        )

        if ruta_extraida.exists():
            excels = list(
                ruta_extraida.rglob("*.xlsx")
            )

            if excels:
                return True

    return False


# ==========================================================
# DESCARGA
# ==========================================================

def descargar_paquete(
    session,
    paquete
):
    identificador = paquete.get("id")
    anio = str(paquete.get("year"))
    mes = str(
        paquete.get("month")
    ).zfill(2)

    timestamp = paquete.get(
        "timestamp",
        ""
    )

    formato = paquete.get(
        "_formato_elegido"
    )

    url = paquete.get(
        "_url_descarga"
    )

    carpeta_mes = (
        CARPETA_OECE
        / anio
        / mes
    )

    carpeta_descarga = (
        carpeta_mes
        / "descarga"
    )

    carpeta_extraida = (
        carpeta_mes
        / "extraido"
    )

    carpeta_descarga.mkdir(
        parents=True,
        exist_ok=True
    )

    print("\n----------------------------------------")
    print(f"Paquete: {identificador}")
    print(f"Año: {anio}")
    print(f"Mes: {mes}")
    print(f"Formato: {formato}")
    print(f"Timestamp: {timestamp}")
    print(f"URL: {url}")

    respuesta = session.get(
        url,
        stream=True,
        timeout=(30, 1200),
        allow_redirects=True,
        headers={
            # El servidor devuelve un ZIP aunque pidamos XLSX.
            "Accept": "*/*",

            "Referer": (
                "https://contratacionesabiertas."
                "oece.gob.pe/descargas"
            )
        }
    )

    print(
        f"Estado HTTP: "
        f"{respuesta.status_code}"
    )

    print(
        "Content-Type:",
        respuesta.headers.get(
            "Content-Type",
            ""
        )
    )

    print(
        "Content-Disposition:",
        respuesta.headers.get(
            "Content-Disposition",
            ""
        )
    )

    if respuesta.status_code != 200:
        try:
            detalle = respuesta.text[:500]
            print(
                f"Respuesta del servidor: "
                f"{detalle}"
            )
        except Exception:
            pass

        respuesta.raise_for_status()

    content_type = respuesta.headers.get(
        "Content-Type",
        ""
    ).lower()

    if "text/html" in content_type:
        raise RuntimeError(
            f"{identificador} devolvió HTML."
        )

    if "application/json" in content_type:
        raise RuntimeError(
            f"{identificador} devolvió JSON "
            "en lugar del archivo."
        )

    nombre_servidor = obtener_nombre_servidor(
        respuesta
    )

    extension = determinar_extension(
        respuesta,
        nombre_servidor
    )

    if nombre_servidor:
        nombre_archivo = nombre_servidor
    else:
        nombre_archivo = (
            f"{anio}-{mes}_{FUENTE}_es"
            f"{extension}"
        )

    ruta_temporal = (
        carpeta_descarga
        / f"{nombre_archivo}.part"
    )

    ruta_inicial = (
        carpeta_descarga
        / nombre_archivo
    )

    sha256 = hashlib.sha256()
    bytes_descargados = 0

    with open(
        ruta_temporal,
        "wb"
    ) as archivo:

        for bloque in respuesta.iter_content(
            chunk_size=1024 * 1024
        ):
            if not bloque:
                continue

            archivo.write(bloque)
            sha256.update(bloque)

            bytes_descargados += len(bloque)

    if bytes_descargados == 0:
        ruta_temporal.unlink(
            missing_ok=True
        )

        raise RuntimeError(
            f"{identificador} llegó vacío."
        )

    ruta_temporal.replace(
        ruta_inicial
    )

    # Detectar el tipo real después de descargar.
    if (
        ruta_inicial.suffix.lower() not in {
            ".zip",
            ".xlsx"
        }
        and zipfile.is_zipfile(ruta_inicial)
    ):
        if es_archivo_xlsx_interno(
            ruta_inicial
        ):
            ruta_correcta = (
                ruta_inicial
                .with_suffix(".xlsx")
            )
        else:
            ruta_correcta = (
                ruta_inicial
                .with_suffix(".zip")
            )

        ruta_inicial.replace(
            ruta_correcta
        )

        ruta_final = ruta_correcta

    else:
        ruta_final = ruta_inicial

    # Si el servidor dijo ZIP, verificar si en realidad
    # se trata de un XLSX directo.
    if (
        ruta_final.suffix.lower() == ".zip"
        and es_archivo_xlsx_interno(
            ruta_final
        )
    ):
        ruta_xlsx = (
            ruta_final
            .with_suffix(".xlsx")
        )

        ruta_final.replace(
            ruta_xlsx
        )

        ruta_final = ruta_xlsx

    archivos_extraidos = []

    if ruta_final.suffix.lower() == ".zip":
        print(
            "Archivo ZIP detectado. "
            "Descomprimiendo..."
        )

        if not zipfile.is_zipfile(
            ruta_final
        ):
            raise RuntimeError(
                f"{ruta_final.name} no es "
                "un ZIP válido."
            )

        if carpeta_extraida.exists():
            shutil.rmtree(
                carpeta_extraida
            )

        carpeta_extraida.mkdir(
            parents=True,
            exist_ok=True
        )

        extraer_zip_seguro(
            ruta_final,
            carpeta_extraida
        )

        archivos_extraidos = [
            archivo
            for archivo in carpeta_extraida.rglob("*")
            if archivo.is_file()
        ]

    elif ruta_final.suffix.lower() == ".xlsx":
        print(
            "XLSX directo detectado."
        )

        if carpeta_extraida.exists():
            shutil.rmtree(
                carpeta_extraida
            )

        carpeta_extraida.mkdir(
            parents=True,
            exist_ok=True
        )

        ruta_copia = (
            carpeta_extraida
            / ruta_final.name
        )

        shutil.copy2(
            ruta_final,
            ruta_copia
        )

        archivos_extraidos = [
            ruta_copia
        ]

    else:
        raise RuntimeError(
            "Formato descargado no reconocido: "
            f"{ruta_final.suffix}"
        )

    excels_extraidos = [
        archivo
        for archivo in archivos_extraidos
        if archivo.suffix.lower() == ".xlsx"
    ]

    if not excels_extraidos:
        raise RuntimeError(
            f"{identificador} no produjo "
            "ningún archivo XLSX."
        )

    print(
        f"Archivos extraídos: "
        f"{len(archivos_extraidos)}"
    )

    for archivo in archivos_extraidos:
        print(
            "-",
            archivo.relative_to(
                CARPETA_OECE
            )
        )

    return {
        "id": identificador,
        "source": FUENTE,
        "year": anio,
        "month": mes,
        "timestamp": timestamp,
        "format": formato,
        "url": url,

        "file": str(
            ruta_final.relative_to(
                CARPETA_OECE
            )
        ).replace("\\", "/"),

        "size_bytes": bytes_descargados,

        "sha256": sha256.hexdigest(),

        "extracted_folder": str(
            carpeta_extraida.relative_to(
                CARPETA_OECE
            )
        ).replace("\\", "/"),

        "extracted_files": [
            str(
                archivo.relative_to(
                    CARPETA_OECE
                )
            ).replace("\\", "/")
            for archivo in archivos_extraidos
        ],

        "xlsx_files": [
            str(
                archivo.relative_to(
                    CARPETA_OECE
                )
            ).replace("\\", "/")
            for archivo in excels_extraidos
        ],

        "downloaded_at": datetime.now()
        .astimezone()
        .isoformat()
    }


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Descarga los paquetes mensuales "
            "OCDS/OECE."
        )
    )

    parser.add_argument(
        "--year",
        action="append",
        type=int,
        help=(
            "Año que se desea procesar. "
            "Puede repetirse."
        )
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help=(
            "Descarga nuevamente todos los paquetes, "
            "aunque no hayan cambiado."
        )
    )

    argumentos = parser.parse_args()

    if argumentos.year:
        anios = sorted(
            set(argumentos.year)
        )
    else:
        anios = ANIOS_PREDETERMINADOS

    print("Años que se procesarán:")
    print(anios)

    estado = cargar_estado(
        anios
    )

    guardar_estado(
        estado
    )

    session = requests.Session()

    session.headers.update(
        HEADERS_GENERALES
    )

    todos_los_paquetes = []
    errores_catalogo = []

    for anio in anios:
        try:
            paquetes = consultar_paquetes_anio(
                session,
                anio
            )

            todos_los_paquetes.extend(
                paquetes
            )

        except (
            requests.RequestException,
            RuntimeError
        ) as error:

            print(
                f"ERROR consultando {anio}: "
                f"{error}"
            )

            errores_catalogo.append({
                "year": anio,
                "error": str(error)
            })

    print("\n========================================")
    print("TOTAL DE PAQUETES ENCONTRADOS")
    print("========================================")
    print(len(todos_los_paquetes))

    descargados = 0
    omitidos = 0
    errores_descarga = 0

    resumen_anual = {}

    for paquete in todos_los_paquetes:
        identificador = paquete.get("id")
        anio = str(paquete.get("year"))

        if anio not in resumen_anual:
            resumen_anual[anio] = {
                "encontrados": 0,
                "descargados": 0,
                "omitidos": 0,
                "errores": 0
            }

        resumen_anual[anio][
            "encontrados"
        ] += 1

        timestamp_api = paquete.get(
            "timestamp",
            ""
        )

        estado_anterior = estado.get(
            identificador,
            {}
        )

        if (
            not argumentos.force
            and descarga_anterior_valida(
                estado_anterior,
                timestamp_api
            )
        ):
            print(
                f"\nOMITIDO: {identificador}. "
                "No presenta cambios."
            )

            omitidos += 1

            resumen_anual[anio][
                "omitidos"
            ] += 1

            continue

        try:
            resultado = descargar_paquete(
                session,
                paquete
            )

            estado[
                identificador
            ] = resultado

            guardar_estado(
                estado
            )

            tamaño_mb = (
                resultado["size_bytes"]
                / 1024
                / 1024
            )

            print(
                f"Descargado: "
                f"{resultado['file']}"
            )

            print(
                f"Tamaño: "
                f"{tamaño_mb:.2f} MB"
            )

            print(
                f"SHA-256: "
                f"{resultado['sha256']}"
            )

            descargados += 1

            resumen_anual[anio][
                "descargados"
            ] += 1

        except (
            requests.RequestException,
            RuntimeError,
            OSError,
            zipfile.BadZipFile
        ) as error:

            print(
                f"ERROR descargando "
                f"{identificador}: {error}"
            )

            errores_descarga += 1

            resumen_anual[anio][
                "errores"
            ] += 1

    guardar_estado(
        estado
    )

    print("\n========================================")
    print("RESUMEN GENERAL")
    print("========================================")

    print(
        f"Paquetes encontrados: "
        f"{len(todos_los_paquetes)}"
    )

    print(
        f"Descargados o actualizados: "
        f"{descargados}"
    )

    print(
        f"Sin cambios: "
        f"{omitidos}"
    )

    print(
        f"Errores de descarga: "
        f"{errores_descarga}"
    )

    print(
        f"Errores de catálogo: "
        f"{len(errores_catalogo)}"
    )

    print("\nRESUMEN POR AÑO")

    for anio, valores in sorted(
        resumen_anual.items()
    ):
        print(
            f"{anio} | "
            f"Encontrados: {valores['encontrados']} | "
            f"Descargados: {valores['descargados']} | "
            f"Sin cambios: {valores['omitidos']} | "
            f"Errores: {valores['errores']}"
        )

    print("\nCarpeta principal:")
    print(CARPETA_OECE)

    print("\nArchivo de control:")
    print(ARCHIVO_ESTADO_GLOBAL)


if __name__ == "__main__":
    main()
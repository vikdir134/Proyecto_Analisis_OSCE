import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path
from urllib.parse import unquote

import requests


# ==========================================================
# CONFIGURACIÓN
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

CARPETA_ANIO = (
    BASE_DIR
    / "data"
    / "raw"
    / "oece"
    / "2026"
)

CARPETA_ANIO.mkdir(parents=True, exist_ok=True)

ARCHIVO_ESTADO = CARPETA_ANIO / "estado_descargas_2026.json"

URL_CATALOGO = (
    "https://contratacionesabiertas.oece.gob.pe/api/v1/files"
)

ANIO = "2026"
FUENTE = "seace_v3"
FORMATO_DESCARGA = "xlsx_es"

HEADERS_GENERALES = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/149 Safari/537.36"
    ),
    "Accept-Language": "es-PE,es;q=0.9"
}


# ==========================================================
# ESTADO
# ==========================================================

def cargar_estado():
    if not ARCHIVO_ESTADO.exists():
        return {}

    try:
        with open(ARCHIVO_ESTADO, "r", encoding="utf-8") as archivo:
            return json.load(archivo)

    except (json.JSONDecodeError, OSError):
        return {}


def guardar_estado(estado):
    with open(ARCHIVO_ESTADO, "w", encoding="utf-8") as archivo:
        json.dump(
            estado,
            archivo,
            ensure_ascii=False,
            indent=2
        )


# ==========================================================
# CATÁLOGO
# ==========================================================

def encontrar_paquetes(objeto):
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
            encontrados.extend(encontrar_paquetes(valor))

    elif isinstance(objeto, list):
        for elemento in objeto:
            encontrados.extend(encontrar_paquetes(elemento))

    return encontrados


def consultar_paquetes_2026(session):
    parametros = {
        "page": 1,
        "paginateBy": 100,
        "year": ANIO,
        "format": "json"
    }

    print("Consultando catálogo OECE 2026...")

    respuesta = session.get(
        URL_CATALOGO,
        params=parametros,
        headers={
            "Accept": "application/json"
        },
        timeout=60
    )

    print(f"Estado HTTP del catálogo: {respuesta.status_code}")

    respuesta.raise_for_status()

    contenido = respuesta.json()
    paquetes = encontrar_paquetes(contenido)

    paquetes_unicos = {}

    for paquete in paquetes:
        identificador = paquete.get("id")

        if not identificador:
            identificador = (
                f"{paquete.get('source')}-"
                f"{paquete.get('year')}-"
                f"{paquete.get('month')}"
            )

        paquetes_unicos[identificador] = paquete

    resultado = []

    for paquete in paquetes_unicos.values():
        if str(paquete.get("year")) != ANIO:
            continue

        if paquete.get("source") != FUENTE:
            continue

        archivos = paquete.get("files", {})

        if FORMATO_DESCARGA not in archivos:
            continue

        resultado.append(paquete)

    resultado.sort(
        key=lambda paquete: int(paquete.get("month", 0))
    )

    return resultado


# ==========================================================
# NOMBRES DE ARCHIVO
# ==========================================================

def limpiar_nombre_archivo(nombre):
    nombre = re.sub(r'[<>:"/\\|?*]', "_", nombre)
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
            unquote(coincidencia_utf8.group(1))
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


def determinar_extension(respuesta, nombre_servidor):
    if nombre_servidor:
        extension = Path(nombre_servidor).suffix.lower()

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


# ==========================================================
# DESCARGA Y EXTRACCIÓN
# ==========================================================

def descargar_paquete(session, paquete):
    identificador = paquete["id"]
    mes = str(paquete["month"]).zfill(2)
    timestamp = paquete.get("timestamp", "")
    url = paquete["files"][FORMATO_DESCARGA]

    carpeta_mes = CARPETA_ANIO / mes
    carpeta_descarga = carpeta_mes / "descarga"
    carpeta_extraida = carpeta_mes / "extraido"

    carpeta_descarga.mkdir(parents=True, exist_ok=True)

    print("\n----------------------------------------")
    print(f"Paquete: {identificador}")
    print(f"Mes: {mes}")
    print(f"Timestamp: {timestamp}")
    print(f"URL: {url}")

    # Accept */* permite que el servidor responda con ZIP
    respuesta = session.get(
        url,
        stream=True,
        timeout=(30, 900),
        allow_redirects=True,
        headers={
            "Accept": "*/*",
            "Referer": (
                "https://contratacionesabiertas."
                "oece.gob.pe/descargas"
            )
        }
    )

    print(f"Estado HTTP: {respuesta.status_code}")
    print(
        "Content-Type:",
        respuesta.headers.get("Content-Type", "")
    )
    print(
        "Content-Disposition:",
        respuesta.headers.get("Content-Disposition", "")
    )

    if respuesta.status_code != 200:
        try:
            detalle = respuesta.text[:500]
            print(f"Respuesta del servidor: {detalle}")
        except Exception:
            pass

        respuesta.raise_for_status()

    content_type = respuesta.headers.get(
        "Content-Type",
        ""
    ).lower()

    if "text/html" in content_type:
        raise RuntimeError(
            f"{identificador} devolvió HTML en vez de un archivo."
        )

    if "application/json" in content_type:
        raise RuntimeError(
            f"{identificador} devolvió JSON en vez de un archivo: "
            f"{respuesta.text[:300]}"
        )

    nombre_servidor = obtener_nombre_servidor(respuesta)
    extension = determinar_extension(
        respuesta,
        nombre_servidor
    )

    if nombre_servidor:
        nombre_archivo = nombre_servidor
    else:
        nombre_archivo = (
            f"oece_{FUENTE}_{ANIO}_{mes}{extension}"
        )

    ruta_final = carpeta_descarga / nombre_archivo
    ruta_temporal = carpeta_descarga / (
        nombre_archivo + ".part"
    )

    sha256 = hashlib.sha256()
    bytes_descargados = 0

    with open(ruta_temporal, "wb") as archivo:
        for bloque in respuesta.iter_content(
            chunk_size=1024 * 1024
        ):
            if not bloque:
                continue

            archivo.write(bloque)
            sha256.update(bloque)
            bytes_descargados += len(bloque)

    if bytes_descargados == 0:
        ruta_temporal.unlink(missing_ok=True)

        raise RuntimeError(
            f"El archivo de {identificador} llegó vacío."
        )

    ruta_temporal.replace(ruta_final)

    archivos_extraidos = []

    # Descomprimir únicamente cuando realmente sea un ZIP
    if ruta_final.suffix.lower() == ".zip":
        print("El archivo es ZIP. Descomprimiendo...")

        if not zipfile.is_zipfile(ruta_final):
            raise RuntimeError(
                f"El archivo {ruta_final.name} tiene extensión ZIP, "
                "pero su contenido no es un ZIP válido."
            )

        if carpeta_extraida.exists():
            shutil.rmtree(carpeta_extraida)

        carpeta_extraida.mkdir(
            parents=True,
            exist_ok=True
        )

        with zipfile.ZipFile(ruta_final, "r") as zip_archivo:
            zip_archivo.extractall(carpeta_extraida)

            archivos_extraidos = [
                elemento
                for elemento in zip_archivo.namelist()
                if not elemento.endswith("/")
            ]

        print(
            f"Archivos extraídos: {len(archivos_extraidos)}"
        )

        for archivo in archivos_extraidos:
            print(f"- {archivo}")

    elif ruta_final.suffix.lower() == ".xlsx":
        print("El servidor entregó un XLSX directamente.")

    else:
        print(
            f"Archivo guardado con extensión "
            f"{ruta_final.suffix}"
        )

    return {
        "id": identificador,
        "year": ANIO,
        "month": mes,
        "timestamp": timestamp,
        "url": url,
        "file": str(
            ruta_final.relative_to(CARPETA_ANIO)
        ),
        "size_bytes": bytes_descargados,
        "sha256": sha256.hexdigest(),
        "extracted_folder": (
            str(carpeta_extraida.relative_to(CARPETA_ANIO))
            if archivos_extraidos
            else None
        ),
        "extracted_files": archivos_extraidos
    }


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def main():
    estado = cargar_estado()

    session = requests.Session()
    session.headers.update(HEADERS_GENERALES)

    try:
        paquetes = consultar_paquetes_2026(session)

    except requests.RequestException as error:
        print(f"Error consultando catálogo: {error}")
        return

    except ValueError:
        print("El catálogo no devolvió un JSON válido.")
        return

    print("\n========================================")
    print("PAQUETES DE 2026 ENCONTRADOS")
    print("========================================")
    print(f"Cantidad: {len(paquetes)}")

    for paquete in paquetes:
        print(
            f"- {paquete.get('id')} | "
            f"mes {paquete.get('month')} | "
            f"{paquete.get('timestamp')}"
        )

    descargados = 0
    omitidos = 0
    errores = 0

    for paquete in paquetes:
        identificador = paquete["id"]
        timestamp_api = paquete.get("timestamp", "")

        estado_anterior = estado.get(
            identificador,
            {}
        )

        archivo_relativo = estado_anterior.get("file")

        archivo_existe = False

        if archivo_relativo:
            archivo_existe = (
                CARPETA_ANIO / archivo_relativo
            ).exists()

        mismo_timestamp = (
            estado_anterior.get("timestamp")
            == timestamp_api
        )

        if archivo_existe and mismo_timestamp:
            print(
                f"\nOMITIDO: {identificador}. "
                "El timestamp no cambió."
            )

            omitidos += 1
            continue

        try:
            resultado = descargar_paquete(
                session,
                paquete
            )

            estado[identificador] = resultado
            guardar_estado(estado)

            tamaño_mb = (
                resultado["size_bytes"]
                / 1024
                / 1024
            )

            print(
                f"Descargado: {resultado['file']}"
            )
            print(f"Tamaño: {tamaño_mb:.2f} MB")
            print(f"SHA-256: {resultado['sha256']}")

            descargados += 1

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

            errores += 1

    print("\n========================================")
    print("RESUMEN")
    print("========================================")
    print(f"Paquetes encontrados: {len(paquetes)}")
    print(f"Descargados: {descargados}")
    print(f"Sin cambios: {omitidos}")
    print(f"Errores: {errores}")
    print(f"Carpeta: {CARPETA_ANIO}")
    print(f"Estado: {ARCHIVO_ESTADO}")


if __name__ == "__main__":
    main()
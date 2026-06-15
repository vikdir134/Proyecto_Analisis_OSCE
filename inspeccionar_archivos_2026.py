import json
from pathlib import Path
from urllib.parse import urlparse

import requests


# ==========================================
# CONFIGURACIÓN
# ==========================================

BASE_DIR = Path(__file__).resolve().parent
CARPETA_SALIDA = BASE_DIR / "API_OECE_2026"

CARPETA_SALIDA.mkdir(exist_ok=True)

URL_API = "https://contratacionesabiertas.oece.gob.pe/api/v1/files"

PARAMETROS = {
    "page": 1,
    "paginateBy": 100,
    "year": 2026,
    "format": "json"
}

HEADERS = {
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 OECE-ETL-Prueba"
}


# ==========================================
# FUNCIONES
# ==========================================

def guardar_json(nombre_archivo, contenido):
    ruta = CARPETA_SALIDA / nombre_archivo

    with open(ruta, "w", encoding="utf-8") as archivo:
        json.dump(
            contenido,
            archivo,
            ensure_ascii=False,
            indent=2
        )

    return ruta


def encontrar_lista_registros(objeto):
    """
    Busca una lista de registros dentro de respuestas como:

    {
        "data": [...]
    }

    {
        "results": [...]
    }

    {
        "data": {
            "items": [...]
        }
    }
    """

    if isinstance(objeto, list):
        return objeto

    if not isinstance(objeto, dict):
        return []

    nombres_probables = [
        "data",
        "files",
        "items",
        "results",
        "records",
        "rows",
        "content"
    ]

    for nombre in nombres_probables:
        valor = objeto.get(nombre)

        if isinstance(valor, list):
            return valor

        if isinstance(valor, dict):
            resultado = encontrar_lista_registros(valor)

            if resultado:
                return resultado

    # Buscar en cualquier otra propiedad
    for valor in objeto.values():
        if isinstance(valor, dict):
            resultado = encontrar_lista_registros(valor)

            if resultado:
                return resultado

    return []


def buscar_urls(objeto, ruta=""):
    """
    Busca todas las URL encontradas dentro del JSON.
    Devuelve también el nombre del campo donde apareció.
    """

    resultados = []

    if isinstance(objeto, dict):
        for clave, valor in objeto.items():
            nueva_ruta = f"{ruta}.{clave}" if ruta else clave
            resultados.extend(buscar_urls(valor, nueva_ruta))

    elif isinstance(objeto, list):
        for indice, valor in enumerate(objeto):
            nueva_ruta = f"{ruta}[{indice}]"
            resultados.extend(buscar_urls(valor, nueva_ruta))

    elif isinstance(objeto, str):
        valor = objeto.strip()

        if valor.startswith("http://") or valor.startswith("https://"):
            resultados.append({
                "campo": ruta,
                "url": valor
            })

    return resultados


def parece_archivo_descargable(url):
    """
    Detecta URL que probablemente corresponda a un archivo.
    """

    ruta = urlparse(url).path.lower()

    extensiones = (
        ".json",
        ".csv",
        ".xlsx",
        ".xls",
        ".zip",
        ".gz",
        ".sha",
        ".txt"
    )

    return ruta.endswith(extensiones)


def main():
    print("Consultando catálogo de archivos de 2026...")
    print(f"Página: {PARAMETROS['page']}")
    print(f"Registros solicitados: {PARAMETROS['paginateBy']}")

    try:
        respuesta = requests.get(
            URL_API,
            params=PARAMETROS,
            headers=HEADERS,
            timeout=60
        )

        print(f"\nEstado HTTP: {respuesta.status_code}")
        print(
            "Tipo de contenido:",
            respuesta.headers.get("Content-Type", "No informado")
        )

        respuesta.raise_for_status()

    except requests.RequestException as error:
        print(f"\nError consultando la API: {error}")
        return

    try:
        contenido = respuesta.json()

    except ValueError:
        ruta_html = CARPETA_SALIDA / "respuesta_no_json.html"

        with open(ruta_html, "w", encoding="utf-8") as archivo:
            archivo.write(respuesta.text)

        print("\nLa respuesta no fue JSON.")
        print(f"Se guardó para revisión en: {ruta_html}")
        return

    ruta_respuesta = guardar_json(
        "respuesta_files_2026_pagina_1.json",
        contenido
    )

    print(f"\nRespuesta completa guardada en:")
    print(ruta_respuesta)

    print("\nTipo de estructura principal:")
    print(type(contenido).__name__)

    if isinstance(contenido, dict):
        print("\nCampos principales encontrados:")

        for clave in contenido.keys():
            print(f"- {clave}")

    registros = encontrar_lista_registros(contenido)

    print(f"\nRegistros encontrados en la respuesta: {len(registros)}")

    if registros:
        primer_registro = registros[0]

        ruta_primer_registro = guardar_json(
            "primer_registro_2026.json",
            primer_registro
        )

        print("\nPrimer registro guardado en:")
        print(ruta_primer_registro)

        if isinstance(primer_registro, dict):
            print("\nCampos del primer registro:")

            for campo, valor in primer_registro.items():
                texto = str(valor)

                if len(texto) > 150:
                    texto = texto[:150] + "..."

                print(f"- {campo}: {texto}")

    else:
        print("\nNo se encontró automáticamente la lista de archivos.")
        print("Revisa respuesta_files_2026_pagina_1.json.")

    urls = buscar_urls(contenido)

    ruta_urls = CARPETA_SALIDA / "urls_encontradas.txt"

    with open(ruta_urls, "w", encoding="utf-8") as archivo:
        for elemento in urls:
            archivo.write(
                f"{elemento['campo']} | {elemento['url']}\n"
            )

    print(f"\nURL encontradas: {len(urls)}")
    print(f"Listado guardado en: {ruta_urls}")

    urls_descarga = [
        elemento
        for elemento in urls
        if parece_archivo_descargable(elemento["url"])
    ]

    ruta_descargas = CARPETA_SALIDA / "posibles_descargas.txt"

    with open(ruta_descargas, "w", encoding="utf-8") as archivo:
        for elemento in urls_descarga:
            archivo.write(
                f"{elemento['campo']} | {elemento['url']}\n"
            )

    print(f"\nPosibles archivos descargables: {len(urls_descarga)}")
    print(f"Listado guardado en: {ruta_descargas}")

    print("\nProceso terminado correctamente.")


if __name__ == "__main__":
    main()
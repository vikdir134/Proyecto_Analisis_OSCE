import requests
import json
from pathlib import Path
from urllib.parse import urljoin

# Carpeta de salida
BASE_DIR = Path(__file__).resolve().parent
SALIDA = BASE_DIR / "api_pruebas"
SALIDA.mkdir(exist_ok=True)

# Dominios posibles del portal OCDS/OECE
BASES = [
    "https://contratacionesabiertas.oece.gob.pe/",
    "https://contratacionesabiertas.osce.gob.pe/",
]

# Rutas comunes para descubrir documentación/API
RUTAS_PRUEBA = [
    "",
    "api",
    "api/",
    "api/docs",
    "api/swagger",
    "api/swagger.json",
    "api/openapi.json",
    "openapi.json",
    "swagger.json",
    "descargas",
    "downloads",
    "files",
    "api/files",
    "api/releases",
    "api/records",
    "api/release",
    "api/record",
]


def pedir_url(url):
    """
    Hace una petición GET simple.
    """
    try:
        response = requests.get(
            url,
            timeout=30,
            headers={
                "User-Agent": "Mozilla/5.0 prueba-api-osce"
            }
        )

        content_type = response.headers.get("content-type", "")

        return {
            "url": url,
            "status_code": response.status_code,
            "content_type": content_type,
            "texto": response.text[:1000],
            "json": intentar_json(response),
        }

    except Exception as e:
        return {
            "url": url,
            "status_code": "ERROR",
            "content_type": "",
            "texto": str(e),
            "json": None,
        }


def intentar_json(response):
    """
    Intenta convertir la respuesta a JSON.
    Si no se puede, devuelve None.
    """
    try:
        return response.json()
    except Exception:
        return None


def guardar_resultado(nombre, data):
    archivo = SALIDA / nombre
    with open(archivo, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def main():
    resultados = []

    print("Probando posibles endpoints de la API OCDS/OECE...\n")

    for base in BASES:
        print(f"BASE: {base}")

        for ruta in RUTAS_PRUEBA:
            url = urljoin(base, ruta)
            resultado = pedir_url(url)

            resultados.append({
                "url": resultado["url"],
                "status_code": resultado["status_code"],
                "content_type": resultado["content_type"],
            })

            print(
                f"{resultado['status_code']} | "
                f"{resultado['content_type'][:40]:40} | "
                f"{url}"
            )

            # Guardar solo respuestas útiles
            if resultado["status_code"] == 200:
                nombre_archivo = (
                    url.replace("https://", "")
                       .replace("http://", "")
                       .replace("/", "_")
                       .replace("?", "_")
                       .replace("&", "_")
                       .replace("=", "_")
                )

                if resultado["json"] is not None:
                    guardar_resultado(f"{nombre_archivo}.json", resultado["json"])
                else:
                    guardar_resultado(f"{nombre_archivo}.txt.json", {
                        "url": url,
                        "content_type": resultado["content_type"],
                        "preview": resultado["texto"]
                    })

        print()

    guardar_resultado("resumen_endpoints_probados.json", resultados)

    print("\nListo.")
    print(f"Revisa la carpeta: {SALIDA}")
    print("Archivo principal: resumen_endpoints_probados.json")


if __name__ == "__main__":
    main()
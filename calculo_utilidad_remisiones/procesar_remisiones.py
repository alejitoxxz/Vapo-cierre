"""Script para calcular utilidad de remisiones a partir de PDFs.

Este módulo lee todos los archivos PDF en la carpeta ``data/`` y extrae
las líneas que representan ítems de venta usando expresiones regulares.
Con base en una tabla fija de costos, calcula las ganancias por producto,
por ítem y por remisión. Finalmente, genera un archivo Excel en la
carpeta ``output/`` con las hojas ``Resumen``, ``Detalle_Items`` y
``Pendientes``.
"""

import os
import re
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import pdfplumber

# Tabla fija de costos (precios de entrada)
COSTOS: Dict[str, int] = {
    "PRIV BAR": 13500,
    "SPACEMAN": 10500,
    "VELOCITY": 11000,
    "CHRIS BROWN 15000": 11000,
    "SOLARIS": 14500,
    "DEATH ROW": 5000,
    "DEATH ROW 5K": 4800,
    "SNOOPYSMOKE": 9000,
    "BUGATTI": 5000,
    "MTRX25K": 11600,
    "MTRX12K": 9300,
    "MOVEMENT": 8300,
    "LOST MARY": 5000,
    "ORION BAR": 8300,
    "IJOY": 14000,
    "AIRFUZE": 13700,
    "JUDO": 13500,
    "VERA": 10300,
    "CZAR": 8500,
    "MINTOPIA": 8500,
    "SNOOPY30K": 15000,
    "SPACEMAN 20K": 12500,
    "CONNECT": 16000,
    "HELLO SYNIX": 18100,
    "CAPSULA THC": 8800,
    "EQUATOR": 15000,
    "URANUS": 10500,
    "BATERIA THC": 7500,
    "LENTES": 7500,
    "ELF THC": 24000,
    "MOTI": 8500,
    "HUKMANIA": 8000,
    "YOVO": 8000,
    "PULSE": 45000,
    "AIRPODS": 28000,
    "CRAZYACE": 0,
    "ELFBARTE": 5500,
    "AIRBAR": 7500,
    "EASE": 5000,
    "LIGHT RISE": 0,
    "SMARTH TC": 13000,
    "NOS KYLINBAR": 12500,
    "NICKYJAM": 9000,
    "FUMEDESTILADO": 21500,
}

# Palabras típicas de encabezados o pies que se deben ignorar al procesar líneas.
LINEA_NO_ITEM_PREFIXES: Tuple[str, ...] = (
    "SEÑOR",
    "SENOR",
    "DIRECCIÓN",
    "DIRECCION",
    "CIUDAD",
    "TELÉFONO",
    "TELEFONO",
    "FECHA",
    "REMISIÓN",
    "REMISION",
    "ÍTEM",
    "ITEM",
    "ELABORADO",
    "SUBTOTAL",
    "TOTAL",
    "NIT",
    "NO.",
)

# Expresión regular para detectar las líneas de ítems.
LINEA_ITEM_REGEX = re.compile(
    r"^\s*"
    r"(?P<prod>.+?)"  # nombre de producto (relajado, hasta el primer precio)
    r"\s+\$?(?P<p_unit>[\d\.,]+)"  # precio unitario con separadores
    r"\s+(?P<cant>\d+)"  # cantidad entera
    r"\s+\S+"  # descuento (cualquier token)
    r"\s+\$?(?P<total>[\d\.,]+)"  # total de la línea
    r"\s*$",
    flags=re.IGNORECASE,
)


def parse_entero(valor_texto: str) -> Optional[int]:
    """Convierte un string numérico con separadores a entero.

    Se eliminan símbolos como ``$``, ``.`` y ``,`` antes de convertir. Si la
    conversión falla, se devuelve ``None`` para permitir un manejo seguro.
    """

    texto_limpio = re.sub(r"[^\d]", "", valor_texto)
    if not texto_limpio:
        return None
    try:
        return int(texto_limpio)
    except ValueError:
        return None


def normalizar_texto_producto(nombre: str) -> str:
    """Normaliza el nombre del producto para buscarlo en ``COSTOS``.

    - Convierte a mayúsculas.
    - Elimina espacios al inicio y final.
    - Normaliza múltiples espacios a uno solo.
    """

    mayusculas = nombre.upper().strip()
    # Reemplaza secuencias de espacios por un solo espacio.
    return re.sub(r"\s+", " ", mayusculas)


def es_linea_item(linea: str) -> bool:
    """Determina si una línea puede representar un ítem de venta.

    Se descartan las líneas que no contienen el símbolo ``$`` o que empiezan
    con prefijos típicos de encabezados o pies. Si pasa los filtros, se
    evalúa contra la expresión regular de ítems.
    """

    if "$" not in linea:
        return False

    contenido = linea.strip().upper()
    if any(contenido.startswith(prefijo) for prefijo in LINEA_NO_ITEM_PREFIXES):
        return False

    return bool(LINEA_ITEM_REGEX.match(linea))


def extraer_items_desde_pdf(ruta_pdf: str) -> List[Dict[str, object]]:
    """Extrae los ítems de un archivo PDF.

    Recorre cada página, divide el texto en líneas y aplica la expresión
    regular para capturar nombre de producto, precio unitario, cantidad y
    total de venta. Devuelve una lista de diccionarios con los datos
    numéricos ya convertidos a enteros.
    """

    items: List[Dict[str, object]] = []

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                for linea in texto.split("\n"):
                    if not es_linea_item(linea):
                        continue

                    coincidencia = LINEA_ITEM_REGEX.match(linea)
                    if not coincidencia:
                        continue

                    producto = coincidencia.group("prod").strip()
                    precio_unitario = parse_entero(coincidencia.group("p_unit"))
                    cantidad = parse_entero(coincidencia.group("cant"))
                    total_venta = parse_entero(coincidencia.group("total"))

                    # Solo agregar si los valores numéricos son válidos.
                    if None in (precio_unitario, cantidad, total_venta):
                        continue

                    items.append(
                        {
                            "Producto": producto,
                            "Precio_Venta_Unitario": precio_unitario,
                            "Cantidad": cantidad,
                            "Total_Venta_Item": total_venta,
                        }
                    )
    except Exception as exc:  # pragma: no cover - protección mínima
        print(f"No se pudo procesar el PDF '{ruta_pdf}': {exc}")

    return items


def procesar_remisiones(
    carpeta_data: str = "data",
    carpeta_output: str = "output",
    nombre_archivo_excel: str = "resumen_remisiones.xlsx",
) -> None:
    """Orquesta la lectura de PDFs, cálculo de utilidades y escritura del Excel."""

    # Asegurar que la carpeta de salida exista.
    os.makedirs(carpeta_output, exist_ok=True)

    resumen_remisiones: List[Dict[str, object]] = []
    detalle_items: List[Dict[str, object]] = []
    productos_desconocidos: Set[str] = set()

    # Recorre todos los PDFs en la carpeta de datos.
    for nombre_archivo in sorted(os.listdir(carpeta_data)):
        if not nombre_archivo.lower().endswith(".pdf"):
            continue

        ruta_pdf = os.path.join(carpeta_data, nombre_archivo)
        nombre_remision = os.path.splitext(nombre_archivo)[0]
        print(f"Procesando remisión: {nombre_remision}")

        items_pdf = extraer_items_desde_pdf(ruta_pdf)
        if not items_pdf:
            print(f"  - No se encontraron ítems en {nombre_archivo}. Se omite del resumen.")
            continue

        total_ventas_remision = 0
        total_ganancias_remision = 0
        productos_en_remision: Set[str] = set()

        for item in items_pdf:
            producto_original = item["Producto"]
            producto_normalizado = normalizar_texto_producto(producto_original)
            costo_unitario = COSTOS.get(producto_normalizado)
            precio_venta_unitario = int(item["Precio_Venta_Unitario"])
            cantidad = int(item["Cantidad"])
            total_venta_item = int(item["Total_Venta_Item"])

            ganancia_unidad: Optional[int] = None
            ganancia_item: Optional[int] = None

            if costo_unitario is not None:
                ganancia_unidad = precio_venta_unitario - costo_unitario
                ganancia_item = ganancia_unidad * cantidad
                total_ganancias_remision += ganancia_item
            else:
                productos_desconocidos.add(producto_normalizado)

            total_ventas_remision += total_venta_item
            productos_en_remision.add(producto_normalizado)

            detalle_items.append(
                {
                    "Remision": nombre_remision,
                    "Producto": producto_original,
                    "Cantidad": cantidad,
                    "Precio_Venta_Unitario": precio_venta_unitario,
                    "Total_Venta_Item": total_venta_item,
                    "Costo_Unitario": costo_unitario,
                    "Ganancia_Unidad": ganancia_unidad,
                    "Ganancia_Item": ganancia_item,
                }
            )

        # Construye el detalle de productos únicos por remisión.
        detalle_productos = " + ".join(sorted(productos_en_remision))
        resumen_remisiones.append(
            {
                "Remision": nombre_remision,
                "Total_Ventas_COP": total_ventas_remision,
                "Total_Ganancias_COP": total_ganancias_remision,
                "Detalle_de_Items": detalle_productos,
            }
        )

        print(
            f"  - Total ventas: {total_ventas_remision:,} COP | "
            f"Ganancias: {total_ganancias_remision:,} COP"
        )

    if not resumen_remisiones:
        print("No se procesaron remisiones. Verifique que existan PDFs en la carpeta de datos.")
        return

    # Crear DataFrame de resumen y agregar fila TOTAL al final.
    df_resumen = pd.DataFrame(resumen_remisiones)
    total_ventas_global = int(df_resumen["Total_Ventas_COP"].sum())
    total_ganancias_global = int(df_resumen["Total_Ganancias_COP"].sum())
    fila_total = {
        "Remision": f"TOTAL ({len(df_resumen)} remisiones)",
        "Total_Ventas_COP": total_ventas_global,
        "Total_Ganancias_COP": total_ganancias_global,
        "Detalle_de_Items": f"{len(df_resumen)} remisiones",
    }
    df_resumen = pd.concat([df_resumen, pd.DataFrame([fila_total])], ignore_index=True)

    # DataFrame de detalle de ítems.
    df_detalle = pd.DataFrame(detalle_items)

    # Preparar DataFrame de productos pendientes si aplica.
    hojas_excel = {"Resumen": df_resumen, "Detalle_Items": df_detalle}
    if productos_desconocidos:
        df_pendientes = pd.DataFrame(
            sorted(productos_desconocidos), columns=["Producto_no_en_tabla_costos"]
        )
        hojas_excel["Pendientes"] = df_pendientes

    # Escribir el archivo Excel con openpyxl como motor.
    ruta_excel = os.path.join(carpeta_output, nombre_archivo_excel)
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        for nombre_hoja, df in hojas_excel.items():
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)

    print(f"Archivo Excel generado en: {ruta_excel}")
    if productos_desconocidos:
        print(
            "Productos sin costo definido (ver hoja 'Pendientes'): "
            + ", ".join(sorted(productos_desconocidos))
        )


if __name__ == "__main__":
    procesar_remisiones()

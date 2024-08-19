import os
import mimetypes
from datetime import datetime
from pathlib import Path
from PIL import Image, ExifTags
import fitz  # PyMuPDF para manejar archivos PDF
import docx
import openpyxl

# Colores para el texto
verde_limon = "\033[92m"
purpura = "\033[95m"
azul = "\033[94m"
amarillo = "\033[93m"
reset_color = "\033[0m"

def mostrar_texto_intro():
    texto_intro = f"""
{verde_limon}
                     __             .___          __           
______    ____      |__|  ____    __| _/_____   _/  |_   ______
\____ \ _/ __ \     |  |_/ __ \  / __ | \__  \  \   __\ /  ___/
|  |_> >\  ___/     |  |\  ___/ / /_/ |  / __ \_ |  |   \___ \ 
|   __/  \___  >/\__|  | \___  >\____ | (____  / |__|  /____  >
|__|         \/ \______|     \/      \/      \/             \/ 
{reset_color}
{purpura}Esta herramienta fue creada con fines educativos, cualquier mal uso de la misma es bajo tu responsabilidad - 9israel6{reset_color}
{azul}https://www.facebook.com/9isra6{reset_color}
{azul}https://www.instagram.com/9israel6{reset_color}
    """
    print(texto_intro)

def eliminar_metadatos_imagen(foto):
    try:
        with Image.open(foto) as img:
            img_no_exif = Image.new(img.mode, img.size)
            img_no_exif.putdata(list(img.getdata()))
            output_filename = f"sin_metadatos_{Path(foto).name}"
            img_no_exif.save(output_filename)
            print(f"Metadatos eliminados. Imagen guardada como {output_filename}")
    except Exception as e:
        print(f"Error al procesar la imagen: {e}")

def eliminar_metadatos_pdf(pdf):
    try:
        doc = fitz.open(pdf)
        output_filename = f"sin_metadatos_{Path(pdf).name}"
        doc.save(output_filename, deflate=True)
        doc.close()
        print(f"Metadatos eliminados. PDF guardado como {output_filename}")
    except Exception as e:
        print(f"Error al procesar el PDF: {e}")

def eliminar_metadatos_documento(documento):
    try:
        extension = Path(documento).suffix.lower()
        output_filename = f"sin_metadatos_{Path(documento).name}"
        
        if extension == '.docx':
            doc = docx.Document(documento)
            doc.save(output_filename)
        elif extension == '.xlsx':
            wb = openpyxl.load_workbook(documento)
            wb.save(output_filename)
        else:
            print("Formato de documento no soportado para eliminación de metadatos.")
            return
        
        print(f"Metadatos eliminados del documento. Guardado como {output_filename}")
    except Exception as e:
        print(f"Error al procesar el documento: {e}")

def eliminar_metadatos():
    while True:
        print(f"\n{amarillo}Submenú: Elimine sus metadatos aquí{reset_color}")
        print(f"{amarillo}1. Eliminar metadatos de imagen{reset_color}")
        print(f"{amarillo}2. Eliminar metadatos de PDF{reset_color}")
        print(f"{amarillo}3. Eliminar metadatos de documentos (Word, Excel){reset_color}")
        print(f"{amarillo}4. Volver al menú principal{reset_color}")
        opcion = input("Seleccione una opción (1-4): ")
        if opcion == '1':
            imagen = input("Ingrese el nombre del archivo de imagen (BMP, GIF, JPG, TIF, PNG) o 'volver' para regresar al menú principal: ")
            if imagen.lower() == 'volver':
                break
            if os.path.isfile(imagen):
                eliminar_metadatos_imagen(imagen)
            else:
                print(f"Error: El archivo {imagen} no existe.")
        elif opcion == '2':
            pdf = input("Ingrese el nombre del archivo PDF o 'volver' para regresar al menú principal: ")
            if pdf.lower() == 'volver':
                break
            if os.path.isfile(pdf):
                eliminar_metadatos_pdf(pdf)
            else:
                print(f"Error: El archivo {pdf} no existe.")
        elif opcion == '3':
            documento = input("Ingrese el nombre del archivo de documento (Word, Excel) o 'volver' para regresar al menú principal: ")
            if documento.lower() == 'volver':
                break
            if os.path.isfile(documento):
                eliminar_metadatos_documento(documento)
            else:
                print(f"Error: El archivo {documento} no existe.")
        elif opcion == '4':
            break
        else:
            print("Opción no válida, intente de nuevo.")

def ingresar_archivos():
    while True:
        print(f"\n{amarillo}Submenú: Ingresar archivos{reset_color}")
        archivo = input("Ingrese el nombre del archivo para analizar o 'volver' para regresar al menú principal: ")
        if archivo.lower() == 'volver':
            break
        if os.path.isfile(archivo):
            try:
                ruta = Path(archivo).resolve()
                stat = ruta.stat()
                tipo_mime, _ = mimetypes.guess_type(archivo)
                
                print(f"\nMetadatos del archivo: {archivo}")
                print(f"Ubicación: {ruta}")
                print(f"Fecha de creación: {datetime.fromtimestamp(stat.st_ctime)}")
                print(f"Fecha de modificación: {datetime.fromtimestamp(stat.st_mtime)}")
                print(f"Tamaño del archivo: {stat.st_size} bytes")
                print(f"Tipo de archivo: {tipo_mime}")
                
            except Exception as e:
                print(f"Error al obtener metadatos: {e}")
        else:
            print(f"Error: El archivo {archivo} no existe.")

def ingresar_foto():
    while True:
        print(f"\n{amarillo}Submenú: Ingresar foto{reset_color}")
        foto = input("Ingrese el nombre del archivo de imagen (BMP, GIF, JPG, TIF, PNG) o 'volver' para regresar al menú principal: ")
        if foto.lower() == 'volver':
            break
        if os.path.isfile(foto):
            try:
                with Image.open(foto) as img:
                    exif_data = img._getexif() or {}
                    ancho, alto = img.size
                    megapixeles = (ancho * alto) / 1_000_000
                    tamano = os.path.getsize(foto)  # Peso del archivo en bytes

                    ubicacion = "No disponible"
                    hora = exif_data.get(36867, "No disponible")
                    dispositivo = exif_data.get(271, "No disponible")

                    # Mapeo de etiquetas GPS a valores humanos
                    gps_info = {}
                    for tag, value in exif_data.items():
                        tag_name = ExifTags.TAGS.get(tag, tag)
                        if tag_name == "GPSInfo":
                            for key in value.keys():
                                gps_info[ExifTags.GPSTAGS.get(key, key)] = value[key]

                    if gps_info:
                        latitud = gps_info.get("GPSLatitude", "No disponible")
                        longitud = gps_info.get("GPSLongitude", "No disponible")
                        ubicacion = f"Lat: {latitud}, Lon: {longitud}"

                    print(f"\nMetadatos de la imagen: {foto}")
                    print(f"Ubicación: {ubicacion}")
                    print(f"Hora de la foto: {hora}")
                    print(f"Fecha de la foto: {hora[:10] if hora != 'No disponible' else 'No disponible'}")
                    print(f"Resolución: {ancho}x{alto} píxeles")
                    print(f"Megapíxeles: {megapixeles:.2f}")
                    print(f"Dispositivo: {dispositivo}")
                    print(f"Tamaño de la imagen: {ancho}x{alto}")
                    print(f"Peso de la imagen: {tamano} bytes")

            except Exception as e:
                print(f"Error al abrir la imagen: {e}")
        else:
            print(f"Error: El archivo {foto} no existe.")

def mostrar_menu():
    print(f"\n{amarillo}Menú Pejedats{reset_color}")
    print(f"{amarillo}1. Archivos{reset_color}")
    print(f"{amarillo}2. Fotos{reset_color}")
    print(f"{amarillo}3. Eliminar metadatos{reset_color}")
    print(f"{amarillo}4. Salir{reset_color}")
    return input("Seleccione una opción (1-4): ")

def main():
    mostrar_texto_intro()
    while True:
        opcion = mostrar_menu()
        if opcion == '1':
            ingresar_archivos()
        elif opcion == '2':
            ingresar_foto()
        elif opcion == '3':
            eliminar_metadatos()
        elif opcion == '4':
            print("Saliendo del programa...")
            break
        else:
            print("Opción no válida, intente de nuevo.")

if __

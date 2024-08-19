import os
import mimetypes
from datetime import datetime
from pathlib import Path
from PIL import Image, ExifTags
import fitz  # PyMuPDF para manejar archivos PDF
import docx
import openpyxl

def mostrar_texto_intro():
    texto_intro = """
                     __             .___          __           
______    ____      |__|  ____    __| _/_____   _/  |_   ______
\____ \ _/ __ \     |  |_/ __ \  / __ | \__  \  \   __\ /  ___/
|  |_> >\  ___/     |  |\  ___/ / /_/ |  / __ \_ |  |   \___ \ 
|   __/  \___  >/\__|  | \___  >\____ | (____  / |__|  /____  >
|__|         \/ \______|     \/      \/      \/             \/ 
    """
    print(texto_intro)

def eliminar_metadatos_imagen(foto):
    try:
        with Image.open(foto) as img:
            # Crear una nueva imagen sin metadatos
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
    print("Submenú: Elimine sus metadatos aquí")
    while True:
        print("1. Eliminar metadatos de imagen")
        print("2. Eliminar metadatos de PDF")
        print("3. Eliminar metadatos de documentos (Word, Excel)")
        print("4. Volver al menú principal")
        opcion = input("Seleccione una opción (1-4): ")
        if opcion == '1':
            imagen = input("Ingrese el nombre del archivo de imagen (BMP, GIF, JPG, TIF, PNG): ")
            if os.path.isfile(imagen):
                eliminar_metadatos_imagen(imagen)
            else:
                print(f"Error: El archivo {imagen} no existe.")
        elif opcion == '2':
            pdf = input("Ingrese el nombre del archivo PDF: ")
            if os.path.isfile(pdf):
                eliminar_metadatos_pdf(pdf)
            else:
                print(f"Error: El archivo {pdf} no existe.")
        elif opcion == '3':
            documento = input("Ingrese el nombre del archivo de documento (Word, Excel): ")
            if os.path.isfile(documento):
                eliminar_metadatos_documento(documento)
            else:
                print(f"Error: El archivo {documento} no existe.")
        elif opcion == '4':
            break
        else:
            print("Opción no válida, intente de nuevo.")

def ingresar_archivos():
    print("Submenú: Ingresar archivos")
    archivo = input("Ingrese el nombre del archivo para analizar: ")
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
    print("Submenú: Ingresar foto")
    foto = input("Ingrese el nombre del archivo de imagen (BMP, GIF, JPG, TIF, PNG): ")
    if os.path.isfile(foto):
        try:
            with Image.open(foto) as img:
                exif_data = img._getexif() or {}
                ancho, alto = img.size
                megapixeles = (ancho * alto) / 1_000_000
                tamano = os.path.getsize(foto)  # Peso del archivo en bytes

                ubicacion = "No disponible"
                hora = exif_data.get(ExifTags.TAGS.get('DateTime', 36867), "No disponible")
                dispositivo = exif_data.get(ExifTags.TAGS.get('Make', 271), "No disponible")

                if 'GPSInfo' in exif_data:
                    gps_info = exif_data['GPSInfo']
                    latitude = gps_info.get(2, "No disponible")
                    longitude = gps_info.get(4, "No disponible")
                    ubicacion = f"Lat: {latitude}, Lon: {longitude}"

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
    print("\nMenú Pejedats")
    print("1. Archivos")
    print("2. Fotos")
    print("3. Eliminar metadatos")
    print("4. Salir")
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

if __name__ == "__main__":
    main()

import fitz  # Importar la parte de PyMuPDF para manejar PDFs
import os

""" Ejecución 3 """


def graphTransform():
    # Ruta donde se encuentran los archivos PDF
    pdf_directory = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeGeninformes\Data\DataGeneticaGenReport\Origen\GraphReport"

    # Ruta donde se guardarán las imágenes PNG
    output_directory = os.path.join(pdf_directory, "PNG_Images")

    # Crear el directorio para las imágenes PNG si no existe
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Obtener la lista de archivos en el directorio de PDFs
    pdf_files = [f for f in os.listdir(pdf_directory) if f.lower().endswith(".pdf")]

    # Convertir los archivos PDF en imágenes PNG
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_directory, pdf_file)

        # Crear un objeto PyMuPDF para el archivo PDF
        pdf_document = fitz.open(pdf_path)

        # Extraer la página única del PDF
        page = pdf_document.load_page(0)  # Página 0 (la única página)

        # Renderizar la página como imagen
        pixmap = page.get_pixmap()

        # Limpiar el nombre del archivo y extraer la parte después del último guión bajo ('_')
        clean_name = pdf_file.rsplit("_", 1)[
            -1
        ]  # Tomar la parte después del último guión bajo ('_')
        clean_name = os.path.splitext(clean_name)[0]  # Eliminar la extensión PDF

        # Guardar la imagen como archivo PNG
        image_filename = f"{clean_name}.png"
        image_path = os.path.join(output_directory, image_filename)
        pixmap.save(image_path, "PNG")

        # Cerrar el objeto PyMuPDF
        pdf_document.close()

    print("\n" + "=" * 100)
    print("Conversión de las graficas completa.")
    print("=" * 100 + "\n")


graphTransform()

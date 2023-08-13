from win32com import client
import os


def convertXlsxToPdf():
    # Ruta de la carpeta con los archivos .xlsx
    folder_path = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeInfGenMolecular\Data\Salida"

    # Crea una instancia del objeto Excel.Application
    xlApp = client.Dispatch("Excel.Application")

    # Itera sobre los archivos en la carpeta
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            # Ruta completa al archivo Excel
            excel_file = os.path.join(folder_path, filename)

            # Carga el archivo Excel
            books = xlApp.Workbooks.Open(excel_file)

            # Itera sobre las hojas en el archivo Excel
            for ws in books.Worksheets:
                ws.Visible = 1
                pdf_file = os.path.join(folder_path, filename.replace(".xlsx", ".pdf"))
                ws.ExportAsFixedFormat(0, pdf_file)

            # Cierra el archivo Excel
            books.Close()

    # Cierra la instancia de Excel.Application
    xlApp.Quit()

    print("Conversión de XLSX a PDF finalizada.")

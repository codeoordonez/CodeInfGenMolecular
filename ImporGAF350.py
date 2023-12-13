import os
import pandas as pd
import shutil
from datetime import datetime

""" Ejecución 2 Betacaseina """


def importFinalReportDF():
    # Ruta del archivo Excel que copiaste
    ruta_archivo_excel = r"\\comospfs01\15.Molecular\Reporte resultados\2023 Formato reporte de resultados Genetica Animal Betacaseina.xlsx"

    # Ruta de la carpeta salida
    carpeta2 = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeGeninformes\Data\DataGeneticaGenReport"

    # Nombre de la hoja que contiene la tabla
    nombre_hoja = "GA-F-##"  # Reemplaza 'BD Molecular Animal' con el nombre de la hoja que contiene la tabla

    # Rango de celdas donde se encuentra la tabla "MolecularAnimal2023"
    rango_tabla = "A9:K20008"  # Reemplaza con el rango de celdas que contiene la tabla

    # Leer solo la tabla "MolecularAnimal2023" dentro de la hoja
    consolidado = pd.read_excel(
        ruta_archivo_excel, sheet_name=nombre_hoja, skiprows=6, header=0, usecols="A:BI"
    )

    # Filtrar las filas con valores NaN en la columna 'Variante B-caseina determinada por Sanger'
    consolidado_sin_nan = consolidado.dropna(
        subset=["Variante B-caseina determinada por Sanger"]
    )

    print(consolidado_sin_nan)
    print(consolidado_sin_nan.dtypes)

    # Obtener la fecha y hora actual
    fecha_hora_actual = datetime.now().strftime("%Y-%m-%d Hour %H-%M-%S")

    # Ruta donde quieres guardar el archivo CSV exportado
    ruta_archivo_csv_normal = os.path.join(carpeta2, "ExportBetacaseina.csv")

    # Crear el nombre del archivo con fecha y hora
    nombre_archivo_fecha_hora = f"ExportFinalReportBetacaseina_{fecha_hora_actual}.csv"
    ruta_archivo_csv_fecha_hora = os.path.join(carpeta2, nombre_archivo_fecha_hora)

    # Exportar el DataFrame filtrado a un archivo CSV
    consolidado_sin_nan.to_csv(ruta_archivo_csv_normal, index=False)
    consolidado_sin_nan.to_csv(
        ruta_archivo_csv_fecha_hora, index=False, encoding="utf-8-sig"
    )  # Archivo con fecha y hora


importFinalReportDF()

import pandas as pd
import os
from datetime import datetime


def mergeDfReceptionFinalReport():
    # Definir las rutas de los archivos CSV
    ruta_archivo1 = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeInfGenMolecular\Data\DataGeneticaGenReport\ExportFinalReport.csv"
    ruta_archivo2 = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeInfGenMolecular\Data\GenImportRecepCSV\ExportInfoSample.csv"

    # Cargar los archivos CSV en DataFrames
    df1 = pd.read_csv(ruta_archivo1, encoding="utf-8-sig")
    df2 = pd.read_csv(ruta_archivo2, encoding="utf-8-sig")

    # Unir los DataFrames usando la columna "idan" y "Identificación de la muestra" como llaves
    df_merged = pd.merge(
        df1, df2, left_on="Código de muestra", right_on="Código de muestra", how="left"
    )

    # Obtener la fecha actual en formato YYYY-MM-DD
    now = datetime.now().strftime("%Y-%m-%d")

    # Agregar una columna llamada "Now" con la fecha actual al formato YYYY-MM-DD al DataFrame fusionado
    df_merged["Now"] = now

    # Convertir la columna "No de solicitud" a tipo de dato entero o dejarla vacía si no es un número
    df_merged["No de solicitud"] = pd.to_numeric(
        df_merged["No de solicitud"], errors="coerce"
    )
    df_merged["No de solicitud"] = df_merged["No de solicitud"].astype("Int64")

    # Obtener la ruta donde se guardarán los archivos
    ruta_salida = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeInfGenMolecular\Data\Salida\DfFinal"

    # Exportar el DataFrame final a un archivo CSV sin la fecha y hora
    ruta_salida_normal = os.path.join(ruta_salida, "DataFrameFinal.csv")
    df_merged.to_csv(ruta_salida_normal, index=False, encoding="utf-8-sig")

    # Exportar el DataFrame final a un archivo CSV con la fecha y hora
    fecha_hora_actual = datetime.now().strftime(" %Y-%m-%d Hour %H-%M-%S")
    nombre_archivo_fecha_hora = f"DataFrameFinal{fecha_hora_actual}.csv"
    ruta_salida_fecha_hora = os.path.join(ruta_salida, nombre_archivo_fecha_hora)
    df_merged.to_csv(ruta_salida_fecha_hora, index=False, encoding="utf-8-sig")

    # Imprimir el mensaje de éxito
    print("\n" + "=" * 100)
    print("DataFrames fusionados exportados exitosamente!")
    print(f"Archivo sin fecha y hora guardado en: '{ruta_salida_normal}'")
    print(f"Archivo con fecha y hora guardado en: '{ruta_salida_fecha_hora}'")
    print("=" * 100 + "\n")


# mergeDfReceptionFinalReport()

import os
import pandas as pd
import shutil
from datetime import datetime


def importFinalReportDF():
    # Ruta de la carpeta original
    carpeta = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeInfGenMolecular\Data\DataGeneticaGenReport\Origen\FinalReportCSV"

    # Ruta de la carpeta salida
    carpeta2 = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeInfGenMolecular\Data\DataGeneticaGenReport"

    nombres_archivos = os.listdir(carpeta)

    if len(nombres_archivos) == 1:
        nombre_archivo = nombres_archivos[0]
        ruta_archivo_origen = os.path.join(carpeta, nombre_archivo)

        # Verificar si el archivo existe antes de moverlo
        if os.path.exists(ruta_archivo_origen):
            # Mover el archivo a la carpeta de destino
            carpeta_destino = os.path.join(carpeta2, "Origen\ProcessFile")
            ruta_archivo_destino = os.path.join(carpeta_destino, nombre_archivo)
            shutil.move(ruta_archivo_origen, ruta_archivo_destino)

            print("\n" + "=" * 100)
            print("Archivo movido exitosamente a la carpeta de destino. ProcessFile")
            print("=" * 100 + "\n")

            # Leer el archivo movido y crear el DataFrame
            ruta_archivo_nuevo = os.path.join(carpeta_destino, nombre_archivo)
            for codificacion in ["utf-8", "latin-1", "ISO-8859-1"]:
                try:
                    df = pd.read_csv(
                        ruta_archivo_nuevo, sep="\t", encoding=codificacion
                    )
                    print(df.head())
                    break  # Detener el bucle si la lectura tiene éxito
                except UnicodeDecodeError:
                    pass  # Inténtalo con la siguiente codificación si hay un error
            else:
                print("No se pudo leer el archivo con ninguna codificación.")

            # Mapeo de valores
            valor_mapping = {0: "--", 1: "*-", 2: "**", "0": "--", "1": "*-", "2": "**"}

            # Aplicar el mapeo en las columnas específicas
            columnas_a_mapear = [
                "Calpain_316_C",
                "Calpain_530_G",
                "EXON2FB_A",
                "Leptin_A1457G_G",
                "Leptin_C963T_A",
                "Leptin_T945M_G",
                "LeptinA59V_G",
                "WSUCAST_A",
                "UMD3.39263696_1_G",
                "UMD3.39339348_1_A",
            ]

            for columna in columnas_a_mapear:
                df[columna] = df[columna].replace(valor_mapping)

            # Convertir la columna 'Fecha' a tipo datetime
            df["Fecha Fin"] = pd.to_datetime(df["Fecha Fin"], errors="coerce")

            # Cambiar el formato de la columna 'Fecha' a "AAAA-MM-DD"
            df["Fecha Fin"] = df["Fecha Fin"].dt.strftime("%Y-%m-%d")

            # Ruta donde quieres guardar los archivos CSV exportados en la carpeta de origen
            ruta_archivo_csv_normal = os.path.join(carpeta2, "ExportFinalReport.csv")

            # Obtener la fecha y hora actual
            fecha_hora_actual = datetime.now().strftime("%Y-%m-%d Hour %H-%M-%S")

            print(df)

            # Crear el nombre del archivo con fecha y hora
            nombre_archivo_fecha_hora = f"ExportFinalReport_{fecha_hora_actual}.csv"
            ruta_archivo_csv_fecha_hora = os.path.join(
                carpeta2, nombre_archivo_fecha_hora
            )

            # Exportar el DataFrame a los archivos CSV en la carpeta de origen
            df.to_csv(
                ruta_archivo_csv_normal, index=False, encoding="utf-8- sig"
            )  # Archivo normal
            df.to_csv(
                ruta_archivo_csv_fecha_hora, index=False, encoding="utf-8-sig"
            )  # Archivo con fecha y hora
            print("\n" + "=" * 100)
            print(
                "DataFrames exportados exitosamente a archivos CSV en la carpeta DataGeneticaGenReport."
            )
            print("=" * 100 + "\n")

        else:
            print("\n" + "=" * 100)
            print("El archivo no existe en la carpeta original.")
            print("=" * 100 + "\n")
    else:
        print("\n" + "=" * 100)
        print("La Carpeta puede estar vacia o contener mas de un archivo.")
        print("=" * 100 + "\n")


# importFinalReportDF()

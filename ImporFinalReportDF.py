import os
import pandas as pd
import shutil
from datetime import datetime

""" Ejecución 2 """


def importFinalReportDF():
    # Ruta de la carpeta original
    carpeta = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeGeninformes\Data\DataGeneticaGenReport\Origen\FinalReportCSV"

    # Ruta de la carpeta salida
    carpeta2 = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeGeninformes\Data\DataGeneticaGenReport"

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

            # Convertir la columna 'FecAnalisis' a tipo datetime y cambiar el formato
            df["FecAnalisis"] = pd.to_datetime(
                df["FecAnalisis"], format="%d/%m/%Y", errors="coerce"
            ).dt.strftime("%Y-%m-%d")

            if pd.api.types.is_datetime64_ns_dtype(df["FecAnalisis"]):
                # Cambiar el formato de la columna 'FecAnalisis' a "AAAA-MM-DD"
                df["FecAnalisis"] = df["FecAnalisis"].dt.strftime("%Y-%m-%d")
            else:
                print("Error: La columna 'FecAnalisis' no es de tipo datetime.")

            print(df["FecAnalisis"])

            """ # Convertir la columna 'Fecha' a tipo datetime
            df["FecAnalisis"] = pd.to_datetime(df["FecAnalisis"], errors="coerce")

            # Cambiar el formato de la columna 'Fecha' a "AAAA-MM-DD"
            df["FecAnalisis"] = df["FecAnalisis"].dt.strftime("%Y-%m-%d") """

            # Ruta donde quieres guardar los archivos CSV exportados en la carpeta de origen
            ruta_archivo_csv_normal = os.path.join(carpeta2, "ExportFinalReport.csv")

            # Obtener la fecha y hora actual
            fecha_hora_actual = datetime.now().strftime("%Y-%m-%d Hour %H-%M-%S")

            # Agrega la columna 'FlagRaza' en el DataFrame
            df["FlagRaza"] = df["Raza"].apply(
                lambda x: "UNK" if x == "UNK" else ("CRUCE" if len(x) == 6 else "PURA")
            )

            # Divide la columna 'Raza' en 'Raza1' y 'Raza2 solo cuando tenga 6 caracteres
            df["Raza1"] = df["Raza"].str[:3][df["FlagRaza"] == "CRUCE"]
            df["Raza2"] = df["Raza"].str[3:][df["FlagRaza"] == "CRUCE"]

            # Itera a través de las filas del DataFrame
            for index, row in df.iterrows():
                grup_gen = row["GrupGen"]
                raza = row["FlagRaza"]

                # Verifica las condiciones para asignar las cadenas de texto a las columnas correspondientes
                if grup_gen == "MILK":
                    if raza == "PURA":
                        df.at[index, "Enc1"] = (
                            "Comparación con individuos de la población "
                            + df.at[index, "Raza"]
                        )
                        df.at[index, "Enc2"] = "Comparación con individuos cebuinos"
                        df.at[index, "DatTBL1"] = (
                            "Similitud con la población " + df.at[index, "Raza"]
                        )
                        df.at[index, "DatTBL2"] = (
                            "Similitud con la población " + df.at[index, "Raza"]
                        )
                        df.at[index, "DatTBL3"] = "Componente genético adicional"
                        df.at[index, "DatTBL4"] = "Componente genético cebuíno"
                    elif raza == "CRUCE":
                        df.at[index, "Enc1"] = (
                            "Comparación con individuos de la población "
                            + df.at[index, "Raza"]
                        )
                        df.at[index, "Enc2"] = "Comparación con individuos cebuinos"
                        df.at[index, "DatTBL1"] = (
                            "Similitud con la población " + df.at[index, "Raza1"]
                        )
                        df.at[index, "DatTBL2"] = (
                            "Similitud con la población " + df.at[index, "Raza"]
                        )
                        df.at[index, "DatTBL3"] = (
                            "Similitud con la población " + df.at[index, "Raza2"]
                        )
                        df.at[index, "DatTBL4"] = "Componente genético cebuíno"
                    elif raza == "UNK":  # no utilizado
                        df.at[index, "Enc1"] = "Texto para MILK UNK en Enc1"
                        df.at[index, "Enc2"] = "Texto para MILK UNK en Enc2"
                        df.at[index, "DatTBL1"] = "Texto para MILK UNK en DatTBL1"
                        df.at[index, "DatTBL2"] = "Texto para MILK UNK en DatTBL2"
                        df.at[index, "DatTBL3"] = "Texto para MILK UNK en DatTBL3"
                        df.at[index, "DatTBL4"] = "Texto para MILK UNK en DatTBL4"
                elif grup_gen == "MEAT":
                    if raza == "PURA":
                        df.at[index, "Enc1"] = "Comparación con individuos taurinos"
                        df.at[index, "Enc2"] = "Comparación con individuos cebuinos"
                        df.at[
                            index, "DatTBL1"
                        ] = "Similitud con alelos presentes en poblaciones taurinas"
                        df.at[
                            index, "DatTBL2"
                        ] = "Similitud de alelos presentes en poblaciones cebuinas"
                        df.at[index, "DatTBL3"] = "Componente genético adicional"
                        df.at[
                            index, "DatTBL4"
                        ] = "Similitud de alelos presentes en un grupo taurino"
                    elif raza == "CRUCE":
                        df.at[index, "Enc1"] = (
                            "Comparación con individuos de la población "
                            + df.at[index, "Raza"]
                        )
                        df.at[index, "Enc2"] = "Texto para MEAT CRUCE en Enc2"
                        df.at[index, "DatTBL1"] = "Similitud con la población <<Raza>>"
                        df.at[
                            index, "DatTBL2"
                        ] = "Similitud de alelos presentes en poblaciones cebuinas"
                        df.at[index, "DatTBL3"] = "Componente genético adiciona"
                        df.at[index, "DatTBL4"] = "Similitud con la población <<Raza>>"
                    elif raza == "UNK":  # no utilizado
                        df.at[index, "Enc1"] = "Texto para MEAT UNK en Enc1"
                        df.at[index, "Enc2"] = "Texto para MEAT UNK en Enc2"
                        df.at[index, "DatTBL1"] = "Texto para MEAT UNK en DatTBL1"
                        df.at[index, "DatTBL2"] = "Texto para MEAT UNK en DatTBL2"
                        df.at[index, "DatTBL3"] = "Texto para MEAT UNK en DatTBL3"
                        df.at[index, "DatTBL4"] = "Texto para MEAT UNK en DatTBL4"
                elif grup_gen == "ZEB":
                    if raza == "PURA":
                        df.at[index, "Enc1"] = "Comparación con individuos taurinos"
                        df.at[index, "Enc2"] = "Comparación con individuos cebuinos"
                        df.at[index, "DatTBL1"] = (
                            "Similitud con la población " + df.at[index, "Raza"]
                        )
                        df.at[
                            index, "DatTBL2"
                        ] = "Similitud de alelos presentes en poblaciones cebuinas"
                        df.at[index, "DatTBL3"] = "Componente genético adicional"
                        df.at[index, "DatTBL4"] = (
                            "Similitud con la población " + df.at[index, "Raza"]
                        )
                    elif raza == "CRUCE":
                        df.at[index, "Enc1"] = (
                            "Comparación con individuos de la población "
                            + df.at[index, "Raza"]
                        )
                        df.at[index, "Enc2"] = "Comparación con individuos cebuinos"
                        df.at[index, "DatTBL1"] = "Texto para ZEB CRUCE en DatTBL1"
                        df.at[index, "DatTBL2"] = "Texto para ZEB CRUCE en DatTBL2"
                        df.at[index, "DatTBL3"] = "Texto para ZEB CRUCE en DatTBL3"
                        df.at[index, "DatTBL4"] = "Texto para ZEB CRUCE en DatTBL4"
                    elif raza == "UNK":
                        df.at[index, "Enc1"] = "Texto para ZEB UNK en Enc1"
                        df.at[index, "Enc2"] = "Texto para ZEB UNK en Enc2"
                        df.at[index, "DatTBL1"] = "Texto para ZEB UNK en DatTBL1"
                        df.at[index, "DatTBL2"] = "Texto para ZEB UNK en DatTBL2"
                        df.at[index, "DatTBL3"] = "Texto para ZEB UNK en DatTBL3"
                        df.at[index, "DatTBL4"] = "Texto para ZEB UNK en DatTBL4"
                elif grup_gen == "CRE":
                    if raza == "PURA":
                        df.at[index, "Enc1"] = (
                            "Comparación con individuos de la población "
                            + df.at[index, "Raza"]
                        )
                        df.at[index, "Enc2"] = "Comparación con individuos cebuinos"
                        df.at[index, "DatTBL1"] = (
                            "Similitud con la población " + df.at[index, "Raza"]
                        )
                        df.at[
                            index, "DatTBL2"
                        ] = "Similitud de alelos presentes en poblaciones cebuinas"
                        df.at[index, "DatTBL3"] = "Componente genético adicional"
                        df.at[index, "DatTBL4"] = (
                            "Similitud con la población " + df.at[index, "Raza"]
                        )
                    elif raza == "CRUCE":  # no utilizado
                        df.at[index, "Enc1"] = "Texto para CRE CRUCE en Enc1"
                        df.at[index, "Enc2"] = "Texto para CRE CRUCE en Enc2"
                        df.at[index, "DatTBL1"] = "Texto para CRE CRUCE en DatTBL1"
                        df.at[index, "DatTBL2"] = "Texto para CRE CRUCE en DatTBL2"
                        df.at[index, "DatTBL3"] = "Texto para CRE CRUCE en DatTBL3"
                        df.at[index, "DatTBL4"] = "Texto para CRE CRUCE en DatTBL4"
                    elif raza == "UNK":  # no utilizado
                        df.at[index, "Enc1"] = "Texto para CRE UNK en Enc1"
                        df.at[index, "Enc2"] = "Texto para CRE UNK en Enc2"
                        df.at[index, "DatTBL1"] = "Texto para CRE UNK en DatTBL1"
                        df.at[index, "DatTBL2"] = "Texto para CRE UNK en DatTBL2"
                        df.at[index, "DatTBL3"] = "Texto para CRE UNK en DatTBL3"
                        df.at[index, "DatTBL4"] = "Texto para CRE UNK en DatTBL4"
                else:
                    print(
                        f"Error: Combinación no válida de GrupGen ({grup_gen}) y Raza ({raza}) en la fila {index}"
                    )

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


importFinalReportDF()

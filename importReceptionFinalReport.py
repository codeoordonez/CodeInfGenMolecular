import pandas as pd
import os
import datetime
import shutil

""" Ejecución 4 """


def importReceptionFinalReport():
    # Ruta del archivo Excel que copiaste
    ruta_archivo_excel = "D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeGeninformes\Data\GA-F -71 2023 BD Unidad de Laboratorio de Servicios Molecular 2023-01-02.xlsx"

    # Nombre de la hoja que contiene la tabla
    nombre_hoja = "BD Molecular Animal"  # Reemplaza 'BD Molecular Animal' con el nombre de la hoja que contiene la tabla

    # Rango de celdas donde se encuentra la tabla "MolecularAnimal2023"
    rango_tabla = "A9:BI1920"  # Reemplaza con el rango de celdas que contiene la tabla

    # Leer solo la tabla "MolecularAnimal2023" dentro de la hoja
    tabla_molecular_animal_2023 = pd.read_excel(
        ruta_archivo_excel, sheet_name=nombre_hoja, skiprows=8, usecols="A:BI"
    )

    numero_filas = tabla_molecular_animal_2023.shape[0]
    print(f"El DataFrame tiene antes {numero_filas} filas.")

    # Eliminar todas las filas que contienen el valor "N/D" en la columna "No de solicitud"
    tabla_molecular_animal_2023 = tabla_molecular_animal_2023.dropna(
        subset=["No de solicitud"]
    )

    numero_filas = tabla_molecular_animal_2023.shape[0]
    print(f"El DataFrame tiene despues {numero_filas} filas.")

    # Limpiar los espacios en los nombres de las columnas
    tabla_molecular_animal_2023.columns = (
        tabla_molecular_animal_2023.columns.str.strip()
    )

    # Obtener los nombres de las columnas en una lista
    nombres_columnas = tabla_molecular_animal_2023.columns.tolist()

    # Lista de columnas que deseas mantener en el DataFrame
    columnas_deseadas = [
        "Código de muestra",
        "Fecha de llegada al laboratorio. Año- Mes-Día",
        "Nombre de Usuario",
        "Cédula",
        "Dirección",
        "Departamento",
        "Municipio",
        "Teléfono",
        "No de solicitud",
        "Finca",
        "Vereda",
        "Identificación de la muestra",
        "Matriz",
        "Fecha de toma de muestra  Año- Mes- Día",
        "Observación",
        "Tipo de análisis Compilado",
    ]

    # Función para aplicar el formato de capitalización a las columnas seleccionadas
    def aplicar_formato_capitalizacion(valor, columna):
        if isinstance(valor, str) and columna not in [
            "Código de muestra",
            "Tipo de análisis Compilado",
        ]:
            return valor.title()
        else:
            return valor

    # Aplicar el formato de capitalización a las columnas seleccionadas del DataFrame
    for columna in columnas_deseadas:
        tabla_molecular_animal_2023[columna] = tabla_molecular_animal_2023[
            columna
        ].apply(lambda x: aplicar_formato_capitalizacion(x, columna))

    # Convertir la columna "Fecha de toma de muestra Año- Mes- Día" a tipo datetime
    tabla_molecular_animal_2023[
        "Fecha de toma de muestra  Año- Mes- Día"
    ] = pd.to_datetime(
        tabla_molecular_animal_2023["Fecha de toma de muestra  Año- Mes- Día"],
        errors="coerce",
    )

    """oordonez, esta funcion podria ser mas utili si  # Convertir la columna "Fecha de toma de muestra Año- Mes- Día" a tipo datetime no funciona correctamente
    
        def custom_datetime_conversion(value):
        
        try:
            return pd.to_datetime(value)
        except (ValueError, TypeError):
            return pd.NaT

        tabla_molecular_animal_2023["Fecha de toma de muestra  Año- Mes- Día"] = tabla_molecular_animal_2023["Fecha de toma de muestra  Año- Mes- Día"].apply(custom_datetime_conversion) """

    # Aplicar una función para extraer solo la fecha (día-mes-año) sin la hora o retornar "No Indica" si el valor es NaT
    def extraer_fecha_sin_hora(fecha):
        if not pd.isnull(fecha):
            return fecha.date()
        else:
            return "No Indica"

    # Aplicar la función a la columna para extraer la fecha sin la hora o "No Indica"
    tabla_molecular_animal_2023[
        "Fecha de toma de muestra  Año- Mes- Día"
    ] = tabla_molecular_animal_2023["Fecha de toma de muestra  Año- Mes- Día"].apply(
        extraer_fecha_sin_hora
    )
    # Convertir la columna "No de solicitud" a valores enteros
    tabla_molecular_animal_2023["No de solicitud"] = tabla_molecular_animal_2023[
        "No de solicitud"
    ].astype(int)

    # Exportar el DataFr

    # Seleccionar solo las columnas deseadas en el DataFrame
    tabla_molecular_animal_2023 = tabla_molecular_animal_2023[columnas_deseadas]

    # Obtener la fecha de creación del archivo en formato YYYY-MM-DD
    fecha_creacion = datetime.datetime.now().strftime("%Y-%m-%d Hour %H-%M-%S")

    # Directorio donde se guardará el archivo CSV
    directorio_destino = r"D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeGeninformes\Data\GenImportRecepCSV"

    # Nombre del archivo CSV que incluye la fecha de creación
    nombre_archivo_csv = f"ExportInfoSample.csv"

    # Nombre del archivo CSV que incluye la fecha de creación
    nombre_archivo_csv1 = f"ExportInfoSample_{fecha_creacion}.csv"

    # Ruta completa del archivo CSV a exportar
    ruta_archivo_csv = os.path.join(directorio_destino, nombre_archivo_csv)

    # Ruta completa del archivo CSV a exportar
    ruta_archivo_csv1 = os.path.join(directorio_destino, nombre_archivo_csv1)

    # Exportar el DataFrame a un archivo CSV en la ruta especificada
    tabla_molecular_animal_2023.to_csv(
        ruta_archivo_csv, index=False, encoding="utf-8-sig"
    )

    # Exportar el DataFrame a un archivo CSV en la ruta especificada
    tabla_molecular_animal_2023.to_csv(
        ruta_archivo_csv1, index=False, encoding="utf-8-sig"
    )
    # Confirmar que el archivo CSV ha sido creado
    print("\n" + "=" * 100)
    print(
        f"El archivo CSV {nombre_archivo_csv} ha sido creado exitosamente en la ruta: {ruta_archivo_csv}."
    )
    print("=" * 100 + "\n")
    print("\n" + "=" * 100)
    print(
        f"El archivo CSV {nombre_archivo_csv1} ha sido creado exitosamente en la ruta: {ruta_archivo_csv1}."
    )
    print("=" * 100 + "\n")


importReceptionFinalReport()

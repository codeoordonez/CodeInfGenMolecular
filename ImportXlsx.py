import shutil


def importXlsx():
    # Ruta de origen del archivo Excel1
    ruta_origen = r"\\comospfs01\4.Recepcion\Base de datos\GA-F -71 2023 BD Unidad de Laboratorio de Servicios Molecular 2023-01-02.xlsx"

    # Ruta de destino donde se copiará el archivo
    ruta_destino = "D:\OneDrive - AGROSAVIA - CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA\SampleManager\Desarrollos\InfGenMolecular\CodeInfGenMolecular\Data\GA-F -71 2023 BD Unidad de Laboratorio de Servicios Molecular 2023-01-02.xlsx"

    # Copiar el archivo desde la ruta de origen a la ruta de destino
    shutil.copy(ruta_origen, ruta_destino)

    print("\n" + "=" * 100)
    print(f"El archivo Excel1 se ha copiado exitosamente a la carpeta de destino.")
    print("=" * 100 + "\n")


# importXlsx()

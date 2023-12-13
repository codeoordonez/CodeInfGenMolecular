""" oordonez """


from ImportXlsx import importXlsx
from ImporFinalReportDF import importFinalReportDF
from GraphTransform import graphTransform
from importReceptionFinalReport import importReceptionFinalReport
from MergeDFReceptionFinalReport import mergeDfReceptionFinalReport
from mergeFinal import mergeFinal
from convertXlsxToPdf import convertXlsxToPdf

importXlsx()


def preguntar_continuar(tarea):
    while True:
        respuesta = input(f"¿Deseas {tarea}? (S/N): ").strip().upper()
        if respuesta == "S":
            return True
        elif respuesta == "N":
            return False
        else:
            print("Por favor, ingresa 'S' para sí o 'N' para no.")


if preguntar_continuar("importar el reporte final"):
    importFinalReportDF()

if preguntar_continuar("realizar la transformación de gráficos"):
    graphTransform()

if preguntar_continuar("importar el informe final de recepción"):
    importReceptionFinalReport()

if preguntar_continuar("fusionar el informe de recepción final"):
    mergeDfReceptionFinalReport()

if preguntar_continuar("realizar la mezcla final"):
    mergeFinal()

if preguntar_continuar("convertir archivos Excel a PDF"):
    convertXlsxToPdf()

print("Proceso finalizado.")

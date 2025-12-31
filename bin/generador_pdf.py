import os
import io
import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

current_folder = os.path.dirname(__file__)
parent_folder = os.path.dirname(current_folder)
files_folder = os.path.join(parent_folder, "files")
data = os.path.join(files_folder, "Data.xlsx")
original_pdf = os.path.join(current_folder, "juicio.pdf")


def generatePDF(
    no_expediente,
    no_intento,
    fecha_convocatoria,
    fecha_emision,
    nombre,
    dni,
    basico_teorico,
    basico_teorico_practico,
    basico_practico,
    automatizacion_teorico,
    automatizacion_teorico_practico,
    automatizacion_practico,
    redes_teorico,
    redes_teorico_practico,
    redes_practico,
    riesgo_teorico,
    riesgo_teorico_practico,
    riesgo_practico,
    quirofano_teorico,
    quirofano_teorico_practico,
    quirofano_practico,
    lampara_teorico,
    lampara_teorico_practico,
    lampara_practico,
    generadora_teorico,
    generadora_teorico_practico,
    generadora_practico,
    rite_teorico,
    rite_teorico_practico,
    rite_practico_clima,
    rite_practico_calefaccion,
    examinador,
    comentario1,
    comentario2,
    comentario3,
):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont("times", "times.ttf"))
    pdfmetrics.registerFont(TTFont("timesbd", "timesbd.ttf"))
    pdfmetrics.registerFont(TTFont("arial", "arial.ttf"))

    c = canvas.Canvas(packet, letter)

    # Página 1

    c.setFont("arial", 10)
    c.drawString(183, 737, fecha_convocatoria)
    c.drawString(352, 737, str(int(no_intento)))
    c.drawString(107, 715, nombre)
    c.drawString(352, 715, dni)
    c.drawString(120, 286, examinador)
    c.drawString(20, 209, comentario1)
    c.drawString(20, 186, comentario2)
    c.drawString(20, 162, comentario3)
    c.drawString(270, 110, examinador)
    c.drawString(290, 83, fecha_emision)

    c.setFont("timesbd", 9)
    # Basico
    if basico_teorico == "Apto":
        c.drawString(242, 592, "X")
    elif basico_teorico == "No Apto":
        c.drawString(282, 592, "X")
    else:
        c.drawString(323, 592, "X")

    if basico_teorico_practico == "Apto":
        c.drawString(362, 592, "X")
    elif basico_teorico_practico == "No Apto":
        c.drawString(401, 592, "X")
    else:
        c.drawString(443, 592, "X")

    if basico_practico == "Apto":
        c.drawString(482, 592, "X")
    elif basico_practico == "No Apto":
        c.drawString(520, 592, "X")
    else:
        c.drawString(559, 592, "X")

    # Automatización
    if automatizacion_teorico == "Apto":
        c.drawString(242, 574, "X")
    elif automatizacion_teorico == "No Apto":
        c.drawString(282, 574, "X")
    else:
        c.drawString(323, 574, "X")

    if automatizacion_teorico_practico == "Apto":
        c.drawString(362, 574, "X")
    elif automatizacion_teorico_practico == "No Apto":
        c.drawString(401, 574, "X")
    else:
        c.drawString(443, 574, "X")

    if automatizacion_practico == "Apto":
        c.drawString(482, 574, "X")
    elif automatizacion_practico == "No Apto":
        c.drawString(520, 574, "X")
    else:
        c.drawString(559, 574, "X")

    # Redes
    if redes_teorico == "Apto":
        c.drawString(242, 556, "X")
    elif redes_teorico == "No Apto":
        c.drawString(282, 556, "X")
    else:
        c.drawString(323, 556, "X")

    if redes_teorico_practico == "Apto":
        c.drawString(362, 556, "X")
    elif redes_teorico_practico == "No Apto":
        c.drawString(401, 556, "X")
    else:
        c.drawString(443, 556, "X")

    if redes_practico == "Apto":
        c.drawString(482, 556, "X")
    elif redes_practico == "No Apto":
        c.drawString(520, 556, "X")
    else:
        c.drawString(559, 556, "X")

    # Riesgo
    if riesgo_teorico == "Apto":
        c.drawString(242, 531.5, "X")
    elif riesgo_teorico == "No Apto":
        c.drawString(282, 531.5, "X")
    else:
        c.drawString(323, 531.5, "X")

    if riesgo_teorico_practico == "Apto":
        c.drawString(362, 531.5, "X")
    elif riesgo_teorico_practico == "No Apto":
        c.drawString(401, 531.5, "X")
    else:
        c.drawString(443, 531.5, "X")

    if riesgo_practico == "Apto":
        c.drawString(482, 531.5, "X")
    elif riesgo_practico == "No Apto":
        c.drawString(520, 531.5, "X")
    else:
        c.drawString(559, 531.5, "X")

    # Quirofano
    if quirofano_teorico == "Apto":
        c.drawString(242, 500, "X")
    elif quirofano_teorico == "No Apto":
        c.drawString(282, 500, "X")
    else:
        c.drawString(323, 500, "X")

    if quirofano_teorico_practico == "Apto":
        c.drawString(362, 500, "X")
    elif quirofano_teorico_practico == "No Apto":
        c.drawString(401, 500, "X")
    else:
        c.drawString(443, 500, "X")

    if quirofano_practico == "Apto":
        c.drawString(482, 500, "X")
    elif quirofano_practico == "No Apto":
        c.drawString(520, 500, "X")
    else:
        c.drawString(559, 500, "X")

    # Generadora
    if generadora_teorico == "Apto":
        c.drawString(242, 468.5, "X")
    elif generadora_teorico == "No Apto":
        c.drawString(282, 468.5, "X")
    else:
        c.drawString(323, 468.5, "X")

    if generadora_teorico_practico == "Apto":
        c.drawString(362, 468.5, "X")
    elif generadora_teorico_practico == "No Apto":
        c.drawString(401, 468.5, "X")
    else:
        c.drawString(443, 468.5, "X")

    if generadora_practico == "Apto":
        c.drawString(482, 468.5, "X")
    elif generadora_practico == "No Apto":
        c.drawString(520, 468.5, "X")
    else:
        c.drawString(559, 468.5, "X")

    # Lampara
    if lampara_teorico == "Apto":
        c.drawString(242, 437, "X")
    elif lampara_teorico == "No Apto":
        c.drawString(282, 437, "X")
    else:
        c.drawString(323, 437, "X")

    if lampara_teorico_practico == "Apto":
        c.drawString(362, 437, "X")
    elif lampara_teorico_practico == "No Apto":
        c.drawString(401, 437, "X")
    else:
        c.drawString(443, 437, "X")

    if lampara_practico == "Apto":
        c.drawString(482, 437, "X")
    elif lampara_practico == "No Apto":
        c.drawString(520, 437, "X")
    else:
        c.drawString(559, 437, "X")

    # RITE
    if rite_teorico == "Apto":
        c.drawString(136, 331, "X")
    elif rite_teorico == "No Apto":
        c.drawString(175, 331, "X")
    else:
        c.drawString(213, 331, "X")

    if rite_teorico_practico == "Apto":
        c.drawString(251, 331, "X")
    elif rite_teorico_practico == "No Apto":
        c.drawString(290, 331, "X")
    else:
        c.drawString(328, 331, "X")

    if rite_practico_clima == "Apto":
        c.drawString(366, 331, "X")
    elif rite_practico_clima == "No Apto":
        c.drawString(405, 331, "X")
    else:
        c.drawString(443, 331, "X")

    if rite_practico_calefaccion == "Apto":
        c.drawString(481, 331, "X")
    elif rite_practico_calefaccion == "No Apto":
        c.drawString(520, 331, "X")
    else:
        c.drawString(558, 331, "X")

    c.showPage()
    c.save()

    packet.seek(0)

    new_pdf = PdfFileReader(packet)

    existing_pdf = PdfFileReader(open(original_pdf, "rb"))
    output = PdfFileWriter()

    # Creación página
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    new_pdf = os.path.join(
        files_folder, f"Juicio de Competencia_{int(no_expediente)}.pdf"
    )
    output_stream = open(new_pdf, "wb")
    output.write(output_stream)
    output_stream.close()


wb = xlrd.open_workbook(data)

hoja = wb.sheet_by_index(0)
for i in range(4, hoja.nrows):

    no_expediente = hoja.cell_value(i, 0)

    no_intento = hoja.cell_value(i, 1)

    fecha_segmentada_1 = hoja.cell_value(i, 2).split(" del ")
    fecha_segmentada_2 = hoja.cell_value(i, 3).split(" del ")

    fecha_convocatoria = (
        fecha_segmentada_1[0]
        + "/"
        + fecha_segmentada_1[1]
        + "/"
        + fecha_segmentada_1[2]
    )
    fecha_emision = (
        fecha_segmentada_2[0]
        + "/"
        + fecha_segmentada_2[1]
        + "/"
        + fecha_segmentada_2[2]
    )

    nombre = hoja.cell_value(i, 4)

    dni = hoja.cell_value(i, 5)

    # Basico
    if hoja.cell_value(i, 6) == "X":
        basico_teorico = "Apto"
    elif hoja.cell_value(i, 7) == "X":
        basico_teorico = "No Apto"
    else:
        basico_teorico = "NA"

    if hoja.cell_value(i, 9) == "X":
        basico_teorico_practico = "Apto"
    elif hoja.cell_value(i, 10) == "X":
        basico_teorico_practico = "No Apto"
    else:
        basico_teorico_practico = "NA"

    if hoja.cell_value(i, 12) == "X":
        basico_practico = "Apto"
    elif hoja.cell_value(i, 13) == "X":
        basico_practico = "No Apto"
    else:
        basico_practico = "NA"

    # Automatizacion
    if hoja.cell_value(i, 15) == "X":
        automatizacion_teorico = "Apto"
    elif hoja.cell_value(i, 16) == "X":
        automatizacion_teorico = "No Apto"
    else:
        automatizacion_teorico = "NA"

    if hoja.cell_value(i, 18) == "X":
        automatizacion_teorico_practico = "Apto"
    elif hoja.cell_value(i, 19) == "X":
        automatizacion_teorico_practico = "No Apto"
    else:
        automatizacion_teorico_practico = "NA"

    if hoja.cell_value(i, 21) == "X":
        automatizacion_practico = "Apto"
    elif hoja.cell_value(i, 22) == "X":
        automatizacion_practico = "No Apto"
    else:
        automatizacion_practico = "NA"

    # Redes
    if hoja.cell_value(i, 24) == "X":
        redes_teorico = "Apto"
    elif hoja.cell_value(i, 25) == "X":
        redes_teorico = "No Apto"
    else:
        redes_teorico = "NA"

    if hoja.cell_value(i, 27) == "X":
        redes_teorico_practico = "Apto"
    elif hoja.cell_value(i, 28) == "X":
        redes_teorico_practico = "No Apto"
    else:
        redes_teorico_practico = "NA"

    if hoja.cell_value(i, 30) == "X":
        redes_practico = "Apto"
    elif hoja.cell_value(i, 31) == "X":
        redes_practico = "No Apto"
    else:
        redes_practico = "NA"

    # Riesgo
    if hoja.cell_value(i, 33) == "X":
        riesgo_teorico = "Apto"
    elif hoja.cell_value(i, 34) == "X":
        riesgo_teorico = "No Apto"
    else:
        riesgo_teorico = "NA"

    if hoja.cell_value(i, 36) == "X":
        riesgo_teorico_practico = "Apto"
    elif hoja.cell_value(i, 37) == "X":
        riesgo_teorico_practico = "No Apto"
    else:
        riesgo_teorico_practico = "NA"

    if hoja.cell_value(i, 39) == "X":
        riesgo_practico = "Apto"
    elif hoja.cell_value(i, 40) == "X":
        riesgo_practico = "No Apto"
    else:
        riesgo_practico = "NA"

    # Quirofano
    if hoja.cell_value(i, 42) == "X":
        quirofano_teorico = "Apto"
    elif hoja.cell_value(i, 43) == "X":
        quirofano_teorico = "No Apto"
    else:
        quirofano_teorico = "NA"

    if hoja.cell_value(i, 45) == "X":
        quirofano_teorico_practico = "Apto"
    elif hoja.cell_value(i, 46) == "X":
        quirofano_teorico_practico = "No Apto"
    else:
        quirofano_teorico_practico = "NA"

    if hoja.cell_value(i, 48) == "X":
        quirofano_practico = "Apto"
    elif hoja.cell_value(i, 49) == "X":
        quirofano_practico = "No Apto"
    else:
        quirofano_practico = "NA"

    # Lampara
    if hoja.cell_value(i, 51) == "X":
        print("Apto Teórico")
        lampara_teorico = "Apto"
    elif hoja.cell_value(i, 52) == "X":
        print("No Apto Teórico")
        lampara_teorico = "No Apto"
    else:
        print("NA Apto Teórico")
        lampara_teorico = "NA"

    if hoja.cell_value(i, 54) == "X":
        print("Apto Práctico")
        lampara_teorico_practico = "Apto"
    elif hoja.cell_value(i, 55) == "X":
        print("No Apto Práctico")
        lampara_teorico_practico = "No Apto"
    else:
        print("NA Apto Práctico")
        lampara_teorico_practico = "NA"

    if hoja.cell_value(i, 57) == "X":
        print("Apto Práctico")
        lampara_practico = "Apto"
    elif hoja.cell_value(i, 58) == "X":
        print("No Apto Práctico")
        lampara_practico = "No Apto"
    else:
        print("NA Apto Práctico")
        lampara_practico = "NA"

    # Generadora
    if hoja.cell_value(i, 60) == "X":
        generadora_teorico = "Apto"
    elif hoja.cell_value(i, 61) == "X":
        generadora_teorico = "No Apto"
    else:
        generadora_teorico = "NA"

    if hoja.cell_value(i, 63) == "X":
        generadora_teorico_practico = "Apto"
    elif hoja.cell_value(i, 64) == "X":
        generadora_teorico_practico = "No Apto"
    else:
        generadora_teorico_practico = "NA"

    if hoja.cell_value(i, 66) == "X":
        generadora_practico = "Apto"
    elif hoja.cell_value(i, 67) == "X":
        generadora_practico = "No Apto"
    else:
        generadora_practico = "NA"

    # RITE
    if hoja.cell_value(i, 69) == "X":
        rite_teorico = "Apto"
    elif hoja.cell_value(i, 70) == "X":
        rite_teorico = "No Apto"
    else:
        rite_teorico = "NA"

    if hoja.cell_value(i, 72) == "X":
        rite_teorico_practico = "Apto"
    elif hoja.cell_value(i, 73) == "X":
        rite_teorico_practico = "No Apto"
    else:
        rite_teorico_practico = "NA"

    if hoja.cell_value(i, 75) == "X":
        rite_practico_clima = "Apto"
    elif hoja.cell_value(i, 76) == "X":
        rite_practico_clima = "No Apto"
    else:
        rite_practico_clima = "NA"

    if hoja.cell_value(i, 78) == "X":
        rite_practico_calefaccion = "Apto"
    elif hoja.cell_value(i, 79) == "X":
        rite_practico_calefaccion = "No Apto"
    else:
        rite_practico_calefaccion = "NA"

    examinador = hoja.cell_value(i, 81)

    comentario1 = hoja.cell_value(i, 82)

    comentario2 = hoja.cell_value(i, 83)

    comentario3 = hoja.cell_value(i, 84)

    print(examinador)

    generatePDF(
        no_expediente,
        no_intento,
        fecha_convocatoria,
        fecha_emision,
        nombre,
        dni,
        basico_teorico,
        basico_teorico_practico,
        basico_practico,
        automatizacion_teorico,
        automatizacion_teorico_practico,
        automatizacion_practico,
        redes_teorico,
        redes_teorico_practico,
        redes_practico,
        riesgo_teorico,
        riesgo_teorico_practico,
        riesgo_practico,
        quirofano_teorico,
        quirofano_teorico_practico,
        quirofano_practico,
        lampara_teorico,
        lampara_teorico_practico,
        lampara_practico,
        generadora_teorico,
        generadora_teorico_practico,
        generadora_practico,
        rite_teorico,
        rite_teorico_practico,
        rite_practico_clima,
        rite_practico_calefaccion,
        examinador,
        comentario1,
        comentario2,
        comentario3,
    )
print("Documentos generados correctamente")
input()

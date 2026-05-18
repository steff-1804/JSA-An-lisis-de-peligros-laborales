from flask import Flask, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
from io import BytesIO
import base64

app = Flask(__name__)

# Aumentar límite de subida
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/exportar", methods=["POST"])
def exportar():

    datos = request.json

    wb = Workbook()
    ws = wb.active
    ws.title = "Matriz JSA"

    # =========================
    # ESTILOS
    # =========================

    azul = PatternFill(start_color="0B1F3A", end_color="0B1F3A", fill_type="solid")
    verde = PatternFill(start_color="D9F99D", end_color="D9F99D", fill_type="solid")
    amarillo = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
    rojo = PatternFill(start_color="EF9A9A", end_color="EF9A9A", fill_type="solid")

    borde = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    titulo_font = Font(bold=True, size=14, color="FFFFFF")
    header_font = Font(bold=True)

    # =========================
    # ENCABEZADO
    # =========================

    ws.merge_cells("A1:C3")
    ws["A1"] = "TASK DESCRIPTION / DESCRIPCIÓN DEL TRABAJO"
    ws["A1"].font = header_font

    ws.merge_cells("D1:J1")
    ws["D1"] = "Risk Assessment / Job Hazard Analysis"
    ws["D1"].font = Font(bold=True, size=16)

    ws["K1"] = "Version"
    ws["L1"] = "Mar 2026"

    ws["D2"] = "TITLE:"
    ws["E2"] = datos["titulo"]

    ws["D3"] = "DOCUMENT REFERENCE:"
    ws["E3"] = datos["documento"]

    ws["D4"] = "TEAM MEMBERS:"
    ws["E4"] = datos["miembros"]

    ws["D5"] = "DATE:"
    ws["E5"] = datos["fecha"]

    # =========================
    # TABLA
    # =========================

    headers = [
        "Foto",
        "Actividad / Paso de trabajo",
        "Área",
        "Peligro y consecuencia potencial",
        "Tipo riesgo",
        "SEV",
        "LIKL",
        "CONT",
        "RPN",
        "Tipo control",
        "Existing Controls",
        "Recommended Controls"
    ]

    fila_inicio = 7

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=fila_inicio, column=col)
        cell.value = header
        cell.fill = azul
        cell.font = titulo_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borde

    # =========================
    # DATOS
    # =========================

    fila = fila_inicio + 1

    for item in datos["matriz"]:

        rpn = int(item["sev"]) * int(item["likl"]) * int(item["cont"])

        ws.row_dimensions[fila].height = 120

        valores = [
            "",
            item["actividad"],
            item["area"],
            item["peligro"],
            item["tipo_riesgo"],
            item["sev"],
            item["likl"],
            item["cont"],
            rpn,
            item["tipo_control"],
            item["existing_controls"],
            item["recommended_controls"]
        ]

        for col, valor in enumerate(valores, 1):

            cell = ws.cell(row=fila, column=col)
            cell.value = valor
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center",
                wrap_text=True
            )
            cell.border = borde

        # COLOR RPN
        rpn_cell = ws.cell(row=fila, column=9)

        if rpn <= 10:
            rpn_cell.fill = verde
        elif rpn <= 30:
            rpn_cell.fill = amarillo
        else:
            rpn_cell.fill = rojo

        # =========================
        # INSERTAR IMAGEN
        # =========================

        if item["foto"]:

            try:

                img_data = item["foto"].split(",")[1]
                image_bytes = base64.b64decode(img_data)

                image = Image.open(BytesIO(image_bytes))

                image.thumbnail((120, 120))

                temp = BytesIO()
                image.save(temp, format="PNG")
                temp.seek(0)

                excel_img = ExcelImage(temp)
                excel_img.width = 100
                excel_img.height = 100

                ws.add_image(excel_img, f"A{fila}")

            except:
                pass

        fila += 1

    # =========================
    # ANCHOS
    # =========================

    widths = {
        "A": 18,
        "B": 35,
        "C": 20,
        "D": 40,
        "E": 15,
        "F": 10,
        "G": 10,
        "H": 10,
        "I": 12,
        "J": 18,
        "K": 40,
        "L": 40
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # =========================
    # EXPORTAR
    # =========================

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Matriz_JSA.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True)

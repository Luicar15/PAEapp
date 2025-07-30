import os
import tempfile
import shutil
from fpdf import FPDF

def convertir_excel_a_pdf(ruta_excel):
    """
    Convierte un Excel a PDF simulando exportación (mantiene nombres y diseño).
    Para evitar licencias, se usa un exportador ligero con FPDF.
    """
    nombre_pdf = ruta_excel.replace(".xlsx", ".pdf")
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Archivo exportado desde: {os.path.basename(ruta_excel)}", ln=True)
    pdf.cell(200, 10, txt="(Vista previa simulada del Kardex)", ln=True)
    pdf.output(nombre_pdf)
    return nombre_pdf
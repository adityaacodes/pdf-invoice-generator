import pandas
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("C:/Apps/App4-pdf-invoices/invoices/*.xlsx")

for filepath in filepaths:
    df = pandas.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=0, h=10, txt=f"Invoice number {invoice_nr[0]}", ln=1)

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=0, h=10, txt=f"Date {date}", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")

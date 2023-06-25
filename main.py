import pandas
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("C:/Apps/App4-pdf-invoices/invoices/*.xlsx")

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=0, h=10, txt=f"Invoice number {invoice_nr}", ln=1)

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=0, h=10, txt=f"Date {date}", ln=2)

    df = pandas.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]

    # Add a header
    pdf.cell(w=30, h=8, txt=columns[0], align="C", border=1)
    pdf.cell(w=70, h=8, txt=columns[1], align="C", border=1)
    pdf.cell(w=31, h=8, txt=columns[2], align="C", border=1)
    pdf.cell(w=30, h=8, txt=columns[3], align="C", border=1)
    pdf.cell(w=30, h=8, txt=columns[4], align="C", border=1, ln=1)

    # Adding rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=11)
        pdf.set_text_color(30, 30, 30)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), align="C", border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), align="C", border=1)
        pdf.cell(w=31, h=8, txt=str(row['amount_purchased']), align="C", border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), align="C", border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), align="C", ln=1, border=1)

    # Add total price cell
    total_price = df['total_price'].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", align="C", border=1)
    pdf.cell(w=70, h=8, txt="", align="C", border=1)
    pdf.cell(w=31, h=8, txt="", align="C", border=1)
    pdf.cell(w=30, h=8, txt="", align="C", border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), align="C", ln=1, border=1)

    # Add total price sentence
    pdf.set_font(family="Times", style="B", size=15)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=20, txt=f"The total price is {total_price} euros.", align="L")

    # To add company name and logo

    # pdf.set_font(family="Times", style="B", size=15)
    # pdf.set_text_color(0, 0, 0)
    # pdf.cell(w=0, h=20, txt=f"The total price is {total_price} euros.", align="L")
    # pdf.image("{image_url}")

    pdf.output(f"PDFs/{filename}.pdf")

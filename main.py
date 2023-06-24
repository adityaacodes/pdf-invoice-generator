import pandas
import glob
from fpdf import FPDF

pdf = FPDF(orientation="P", unit="mm", format="A4")

filepaths = glob.glob("C:/Apps/App4-pdf-invoices-generator/invoices/*.xlsx")
for file in filepaths:
    df = pandas.read_excel(file, sheet_name="Sheet 1")
    print(df)

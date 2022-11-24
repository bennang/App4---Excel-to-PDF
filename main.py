import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    filename = Path(filepath).stem
    invoice_no, invoice_date = filename.split("-")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_no}", ln=1)

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    headers = list(df.columns)
    headers = [item.replace('_', ' ').title() for item in headers]
    pdf.set_font(family='Times', size=10, style="B")
    pdf.cell(w=30, h=8, txt=headers[0], border=1)
    pdf.cell(w=70, h=8, txt=headers[1], border=1)
    pdf.cell(w=30, h=8, txt=headers[2], border=1)
    pdf.cell(w=30, h=8, txt=headers[3], border=1)
    pdf.cell(w=30, h=8, txt=headers[4], border=1, ln=1)

    for index, rows in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.cell(w=30, h=8, txt=str(rows['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(rows['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(rows['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(rows['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(rows['total_price']), border=1, ln=1)

    total_sum = df['total_price'].sum()
    pdf.set_font(family='Times', size=10)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=70, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    pdf.ln(10)

    pdf.set_font(family='Times', size=12, style="B")
    pdf.cell(w=30, h=8, txt=f"The total due amount is {total_sum} euros", ln=1)
    pdf.cell(w=25, h=8, txt="PythonHow")
    pdf.image('pythonhow.png', w=10)
    pdf.output(f"PDFs/{filename}.pdf")


import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")


for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    pdf.add_page()
    filename = Path(filepath).stem
    invno = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Invoice no. {invno}", ln=1)        # Add invoice no
    pdf.cell(w=0, h=8, txt=f"Date: {date}", ln=1)                     # Add date

    # Add header to the table
    column = list(df.columns)
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, border=1, txt=column[0])
    pdf.cell(w=60, h=8, border=1, txt=column[1])
    pdf.cell(w=40, h=8, border=1, txt=column[2])
    pdf.cell(w=40, h=8, border=1, txt=column[3])
    pdf.cell(w=40, h=8, border=1, txt=column[4], ln=1)

    total_sum = df['total_price'].sum()

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1, align='l')
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=1, align='l')
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1, align='l')
        pdf.cell(w=40, h=8, txt=str(row['price_per_unit']), border=1, align='l')
        pdf.cell(w=40, h=8, txt=str(row['total_price']), border=1, align='l', ln=1)

    # Add total price column
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt=str(total_sum), ln=1, border=1)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=0, h=10, txt=f"The total price is {total_sum} rupees.")

    pdf.output(f"PDFs/{filename}.pdf")

from pathlib import Path
import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="p", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invorice = filename.split("-")[0]
    pdf.set_font(family="Times", size=16,style="B")
    pdf.cell(w=50, h=8 , txt=f"invoice nr.{invorice}",ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date:",ln=1)



    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #add a header
    columns = df.columns
    columns =  [item.replace("_"," ").title() for item in columns]
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt= columns[0], border=1)
    pdf.cell(w=50, h=8, txt= columns[1], border=1)
    pdf.cell(w=30, h=8, txt= columns[2], border=1)
    pdf.cell(w=30, h=8, txt= columns[3], border=1)
    pdf.cell(w=30, h=8, txt= columns[4], border=1, ln=1)
# add rows
    for index, row in df.iterrows():
          pdf.set_font(family="Times", size=10)
          pdf.set_text_color(80, 80, 80)
          pdf.cell(w=30, h=8,txt=str(row["product_id"]), border=1)
          pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
          pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
          pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
          pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    totalsum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(totalsum), border=1, ln=1)
    pdf.cell(w=30, h=8, txt="TOTAL SUM:")
    pdf.cell(w=30, h=8, txt=str(totalsum), ln=1)
    pdf.cell(w=20,h=8, txt=f"pythonHOW")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFS/{filename}.pdf")


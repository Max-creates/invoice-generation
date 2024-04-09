import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    filename = Path(filepath).stem
    invoice_number, date = filename.split("-")
    
    pdf.set_font("Times", "B", 16)
    pdf.cell(50, 8, f"Invoice nr.{invoice_number}", ln=1)
    
    pdf.set_font("Times", "B", 16)
    pdf.cell(50, 8, f"Date: {date}", ln=2)
    
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # Add header
    columns = df.columns
    columns = [i.replace("_", " ").title() for i in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)
    
    # Add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=f"{row['product_id']}", border=1)
        pdf.cell(w=70, h=8, txt=f"{row['product_name']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['total_price']}", border=1, ln=1)
    
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=f"{total_sum}", border=1, ln=1)
    
    # Add total sum sentence
    
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)
    
    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"CompanyName")
    pdf.image("pythonhow.png", w=10, x=43)
    
    
    pdf.output(f"PDFs/{filename}.pdf")
    
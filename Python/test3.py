from fpdf import FPDF, XPos, YPos

pdf = FPDF()
pdf.add_page()
pdf.set_font("Helvetica", size=14)
pdf.cell(200, 10, txt="Hello from fpdf2!", ln=True)
pdf.output("hello.pdf")

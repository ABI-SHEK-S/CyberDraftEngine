from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def generate_letter(case, officer, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    c.setFont("Helvetica", 12)
    y = 750
    c.drawString(50, y, f"To,")
    y -= 20
    c.drawString(50, y, case['RecipientName'])
    y -= 20
    c.drawString(50, y, f"Subject: Request for Information Regarding Case No. {case['CaseNumber']}")
    y -= 40
    c.drawString(50, y, "Dear Sir/Madam,")
    y -= 20
  
    c.drawString(50, y, f"Case No. {case['CaseNumber']}, reported on {case['RequestDate']}")
    y -= 20
    c.drawString(50, y, "Please provide the relevant details at the earliest convenience.")
    y -= 40
    c.drawString(50, y, f"Officer: {officer['OfficerName']}")
    y -= 20
    c.drawString(50, y, f"From: {officer['Address']}")
    y -= 20
    c.drawString(50, y, "Regards,")
    y -= 20
    c.drawString(50, y, "Cyber Crime Department")
    c.save()
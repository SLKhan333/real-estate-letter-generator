import os
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
import streamlit as st

# Set up style
styles = getSampleStyleSheet()
body_style = ParagraphStyle(
    'Body',
    parent=styles['Normal'],
    fontName='Times-Roman',
    fontSize=12,
    leading=16,
    alignment=TA_LEFT
)

def generate_letter_pdf(owner_name, full_address, logo_bytes, signature_bytes):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=LETTER,
                            rightMargin=inch, leftMargin=inch,
                            topMargin=inch, bottomMargin=inch)

    letter_lines = f"""
    Dear {owner_name},<br/><br/>
    My name is Scott Leung, and I’m a real estate agent actively working with a qualified buyer looking for a home in your neighborhood.<br/><br/>
    After reviewing recent sales and available properties, your home at <strong>{full_address}</strong> stood out as a potential perfect fit.<br/><br/>
    Would you consider a confidential conversation about selling your home — either now or in the near future? My client is flexible on timelines and is looking for the right property, not just any listing.<br/><br/>
    If you’ve ever thought about selling, or would simply like to explore your options, please don’t hesitate to reach out.<br/><br/>
    You can call or text me directly at <strong>(415) 994-5946</strong> or email me at <strong>scott@theownteam.com</strong>.<br/><br/>
    Looking forward to connecting,<br/><br/>
    Sincerely,<br/><br/>
    """

    elements = []
    elements.append(Image(logo_bytes, width=1.5 * inch, height=1.5 * inch))
    elements.append(Spacer(1, 0.2 * inch))
    elements.append(Paragraph(letter_lines, body_style))
    elements.append(Spacer(1, 0.1 * inch))
    elements.append(Image(signature_bytes, width=2.2 * inch, height=0.9 * inch, hAlign='LEFT'))
    elements.append(Spacer(1, 0.1 * inch))
    elements.append(Paragraph("Scott Leung<br/>Realtor® - DRE #02218857<br/>theOWNteam.com<br/>(415) 994-5946<br/>scott@theownteam.com", body_style))

    doc.build(elements)
    buffer.seek(0)
    return buffer

def main():
    st.title("🏡 Real Estate Letter Generator")
    st.write("Upload your CSV and branding files to generate personalized seller letters.")

    logo_file = st.file_uploader("Upload Your Logo (JPG/PNG)", type=['jpg', 'jpeg', 'png'])
    signature_file = st.file_uploader("Upload Your Signature Image (PNG)", type=['png'])
    csv_file = st.file_uploader("Upload Owner CSV File", type=['csv'])

    if logo_file and signature_file and csv_file:
        df = pd.read_csv(csv_file)
        zip_buffer = BytesIO()

        with ZipFile(zip_buffer, 'w') as zipf:
            for idx, row in df.iterrows():
                try:
                    owner_name = str(row[7]).strip()
                    street = str(row[16]).strip()
                    city = str(row[17]).strip()
                    state = str(row[18]).strip()
                    zip_code = str(row[19]).strip()
                    full_address = f"{street}, {city}, {state} {zip_code}"

                    pdf_buffer = generate_letter_pdf(owner_name, full_address, logo_file, signature_file)
                    safe_name = owner_name.replace(" ", "_").replace("/", "_")
                    file_name = f"{safe_name}_{zip_code}.pdf"
                    zipf.writestr(file_name, pdf_buffer.read())

                except Exception as e:
                    st.warning(f"Error processing row {idx}: {e}")

        zip_buffer.seek(0)
        st.download_button("📥 Download ZIP of Letters", zip_buffer, "personalized_letters.zip")

if __name__ == "__main__":
    main()

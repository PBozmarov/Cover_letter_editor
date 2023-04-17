import os
from docx import Document
from docx.shared import Pt
from docx2pdf import convert
from datetime import datetime
import streamlit as st

def replace_text(document, old_text, new_text, bold=False):
    for paragraph in document.paragraphs:
        if old_text in paragraph.text:
            paragraph.text = paragraph.text.replace(old_text, new_text)
            run = paragraph.runs[0]
            run.text = run.text.replace(old_text, new_text)
            run.font.size = Pt(12)
            if bold == True:
                run.bold = True

def generate_cover_letter(position, company, city, country, sector):
    
    if sector == 'AI':
        template='/Users/pavelbozmarov/Desktop/Desktop/Documents/Personal/cv/Latest_Cover_letters/Cover_Letter_ML.docx'
    elif sector =='Quant':
        template='/Users/pavelbozmarov/Desktop/Desktop/Documents/Personal/cv/Latest_Cover_letters/Cover_Letter_Algo.docx'

    today_date = datetime.today()
    today_date = today_date.strftime('%d/%m/%y')
    document = Document(template)

    replace_text(document, 'POSITION', position)
    replace_text(document, 'DATE', today_date, bold=True)
    replace_text(document, 'COMPANY', company)
    replace_text(document, 'CITY', city)
    replace_text(document, 'COUNTRY', country)

    output_filename = f'/Users/pavelbozmarov/Desktop/Desktop/Documents/Personal/cv/cover_letter_{position}_{company}.docx'
    document.save(output_filename)
    convert(output_filename)
    os.remove(output_filename)
    return f'Cover letter generated as cover_letter_{position}_{company}.pdf'

st.title("Cover Letter Generator")

position = st.text_input("Enter the position:")
company = st.text_input("Enter the company name:")
city = st.text_input("Enter the city:")
country = st.text_input("Enter the country:")
sector = st.selectbox("Select sector:", ["AI", "Quant"])

if st.button("Generate Cover Letter"):
    result = generate_cover_letter(position, company, city, country, sector)
    st.success(result)
    st.balloons()

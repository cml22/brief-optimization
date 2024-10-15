import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
from docx.oxml import parse_xml

def create_word_file(filename, content):
    document = Document()

    # Titre du document
    document.add_heading('Contenu extrait', level=1)

    # Ajout du contenu
    for part in content:
        paragraph = document.add_paragraph()
        add_hyperlink(paragraph, part['url'], part['text'])

    # Enregistrement du document
    document.save(filename)

def add_hyperlink(paragraph, url, text):
    # Add the hyperlink run
    run = paragraph.add_run(text)
    run.font.color.rgb = RGBColor(0, 0, 255)  # Set hyperlink color to blue
    run.font.underline = True  # Underline for hyperlink

    # Add the hyperlink relationship
    r_id = paragraph.part.rels.add_relationship(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        url,
        'hyperlink',
        target_mode='External'
    )

    # Set hyperlink properties in the XML
    hyperlink = parse_xml(r'<w:hyperlink r:id="{}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'.format(r_id))
    paragraph._element.append(hyperlink)
    run._element.get_or_add_rPr().append(parse_xml(r'<w:rStyle w:val="Hyperlink"/>'))

def extract_content(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Extraire le contenu, modifie cela selon tes besoins
    content = []
    for link in soup.find_all('a'):
        content.append({
            'text': link.get_text(),
            'url': link.get('href')
        })
    
    return content

# Streamlit application
st.title('HTML Content Extractor to Word')
url_input = st.text_input('Enter the page URL')
jira_input = st.text_input('Add the JIRA link (TT - Traffic Team)')
if st.button('Create Word File'):
    content = extract_content(url_input)
    filename = 'extracted_content.docx'
    create_word_file(filename, content)
    st.success(f'Word file created: {filename}')

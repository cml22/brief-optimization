import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
from docx.oxml import OxmlElement

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

    # Create the hyperlink relationship
    r_id = paragraph.part.rels.add_relationship(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        url,
        'hyperlink',
        target_mode='External'
    )

    # Create the hyperlink XML element
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', r_id)
    hyperlink.append(run._element)  # Add the run to the hyperlink

    # Append hyperlink to the paragraph
    paragraph._element.append(hyperlink)

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

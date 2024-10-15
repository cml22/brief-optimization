import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
from docx.oxml import OxmlElement  # Import direct d'OxmlElement
from docx.oxml.ns import qn  # Pour les namespaces

def create_word_file(filename, content):
    document = Document()
    document.add_heading('Contenu extrait', level=1)

    # Ajout du contenu
    for part in content:
        paragraph = document.add_paragraph()
        add_hyperlink(document, paragraph, part['url'], part['text'])

    # Enregistrement du document
    document.save(filename)

def add_hyperlink(doc, paragraph, url, text):
    """
    Ajoute un hyperlien à un paragraphe dans un document Word.
    """
    # Crée un run pour l'hyperlien
    run = paragraph.add_run(text)
    run.font.color.rgb = RGBColor(0, 0, 255)  # Couleur bleue pour le lien
    run.font.underline = True  # Souligner le lien

    # Ajouter la relation du lien hypertexte
    part = paragraph.part
    r_id = part.rels.add_relationship(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        url,
        is_external=True
    )

    # Crée l'élément XML pour l'hyperlien
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    new_run.append(rPr)

    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._element.append(hyperlink)

def extract_content(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Extraire le contenu
    content = []
    for link in soup.find_all('a'):
        content.append({
            'text': link.get_text() or "Lien sans texte",  # Gérer les liens sans texte
            'url': link.get('href')
        })
    
    return content

# Application Streamlit
st.title('HTML Content Extractor to Word')
url_input = st.text_input('Enter the page URL')
jira_input = st.text_input('Add the JIRA link (TT - Traffic Team)')
if st.button('Create Word File'):
    content = extract_content(url_input)
    if content:  # Vérifie que du contenu a été extrait
        filename = 'extracted_content.docx'
        create_word_file(filename, content)
        st.success(f'Word file created: {filename}')
    else:
        st.error('Aucun lien trouvé sur la page.')

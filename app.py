import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
from docx.oxml import OxmlElement

def create_word_file(filename, content):
    document = Document()
    document.add_heading('Contenu extrait', level=1)

    # Ajout du contenu
    for part in content:
        paragraph = document.add_paragraph()
        add_hyperlink(paragraph, part['url'], part['text'])

    # Enregistrement du document
    document.save(filename)

def add_hyperlink(paragraph, url, text):
    # Créer un nouveau run pour le texte du lien
    run = paragraph.add_run(text)
    run.font.color.rgb = RGBColor(0, 0, 255)  # Couleur bleue pour le lien
    run.font.underline = True  # Souligner le lien

    # Créer la relation d'hyperlien
    r_id = paragraph.part.rels.add_relationship(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        url,
        'hyperlink',
        target_mode='External'
    )

    # Créer l'élément hyperlink et l'ajouter au paragraphe
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', r_id)
    
    # Ajouter le texte du lien à l'élément hyperlink
    hyperlink.append(run._element)
    
    # Ajouter l'élément hyperlink au paragraphe
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

import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor

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
    """
    Ajoute un lien hypertexte à un paragraphe dans un document Word.
    """
    # Ajouter le texte avec formatage pour simuler un lien hypertexte
    run = paragraph.add_run(text)
    run.font.color.rgb = RGBColor(0, 0, 255)  # Couleur bleue pour les liens
    run.font.underline = True  # Souligner le texte du lien

    # Ajoute l'URL entre parenthèses à côté du texte
    paragraph.add_run(f" ({url})")

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

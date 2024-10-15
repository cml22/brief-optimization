import streamlit as st
from bs4 import BeautifulSoup
import requests
from docx import Document
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

def add_hyperlink(paragraph, url, text):
    """
    Ajoute un hyperlien à un paragraphe dans un document Word.
    """
    r_id = paragraph.part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    run_properties = OxmlElement('w:rPr')

    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    run_properties.append(underline)

    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')  # Blue color for links
    run_properties.append(color)

    new_run.append(run_properties)
    text_element = OxmlElement('w:t')
    text_element.text = text
    new_run.append(text_element)

    hyperlink.append(new_run)
    paragraph._element.append(hyperlink)

def create_word_file_from_url(filename, url):
    """
    Récupère le contenu d'une URL, extrait les titres et les liens, puis génère un fichier Word formaté.
    """
    # Envoyer une requête HTTP pour récupérer la page
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    document = Document()
    
    # Récupérer le titre principal (H1) et ajouter dans le document
    h1 = soup.find('h1')
    if h1:
        document.add_heading(f"TT={h1.get_text().strip()}", level=1)
    
    # Récupérer les autres titres (H2, H3, ...) et les paragraphes
    for tag in soup.find_all(['h2', 'h3', 'h4', 'p', 'a']):
        if tag.name == 'h2':
            document.add_heading(f"TT={tag.get_text().strip()}", level=2)
        elif tag.name == 'h3':
            document.add_heading(f"TT={tag.get_text().strip()}", level=3)
        elif tag.name == 'h4':
            document.add_heading(f"TT={tag.get_text().strip()}", level=4)
        elif tag.name == 'p':
            paragraph = document.add_paragraph(tag.get_text().strip())
            # Rechercher les liens dans le paragraphe
            for link in tag.find_all('a'):
                add_hyperlink(paragraph, link['href'], link.get_text())
    
    # Sauvegarder le fichier Word
    document.save(filename)

# Interface utilisateur avec Streamlit
st.title("Générateur de document Word à partir d'une URL")

# Entrée pour l'URL
url = st.text_input("Entrez l'URL de la page web :")

# Entrée pour le nom du fichier
filename = st.text_input("Nom du fichier Word (sans extension) :")
filename = filename + ".docx" if filename else ""

# Bouton pour créer le document
if st.button("Créer le document"):
    if url and filename:
        create_word_file_from_url(filename, url)
        st.success(f"Document '{filename}' créé avec succès à partir de l'URL : {url}")
    else:
        st.error("Veuillez fournir une URL et un nom de fichier valide.")

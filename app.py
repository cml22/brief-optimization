import requests
from bs4 import BeautifulSoup
from docx import Document
import streamlit as st
import html

# Fonction pour extraire le contenu à partir de <h1> jusqu'aux balises <h6>, <p> et les liens <a>
def extract_content_from_url(url):
    response = requests.get(url)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')

    content = []
    start = False
    for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p']):
        if element.name == 'h1':
            start = True
        if start:
            text = element.get_text().strip()
            text = html.unescape(text)  # Convertit les entités HTML en caractères normaux
            if text:
                if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    content.append({'type': 'heading', 'level': element.name, 'text': text})
                elif element.name == 'p':
                    paragraph = ""
                    for sub_element in element:
                        if sub_element.name == 'a' and sub_element.get('href'):
                            paragraph += f'{sub_element.get_text()} ({sub_element.get("href")}) '
                        else:
                            paragraph += sub_element.string if sub_element.string else ''
                    content.append({'type': 'paragraph', 'text': paragraph.strip()})

    return content

# Fonction pour créer un document Word à partir du contenu extrait
def create_word_file(jira_link, content):
    ticket_number = jira_link[-4:]
    filename = f"Brief SEO Optimization - TT-{ticket_number}.docx"

    doc = Document()
    doc.add_heading(f'Brief SEO Optimization - TT-{ticket_number}', 0)

    # Ajouter le contenu extrait dans le fichier Word avec formatage
    for element in content:
        if element['type'] == 'heading':
            level = int(element['level'][1])  # Niveau de heading (h1 = 1, h2 = 2, etc.)
            doc.add_heading(element['text'], level=level)
        elif element['type'] == 'paragraph':
            doc.add_paragraph(element['text'])

    # Sauvegarder le fichier Word
    doc.save(filename)
    return filename

# Interface Streamlit
st.title("Extracteur de contenu HTML vers Word")
url = st.text_input("Insérer l'URL de la page")
jira_link = st.text_input("Ajouter le lien JIRA")

if st.button("Générer le fichier Word"):
    if url and jira_link:
        # Extraire le contenu à partir de l'URL
        content = extract_content_from_url(url)
        if content:
            # Créer le fichier Word
            filename = create_word_file(jira_link, content)
            with open(filename, "rb") as file:
                btn = st.download_button(
                    label="Télécharger le fichier",
                    data=file,
                    file_name=filename
                )
        else:
            st.error("Impossible d'extraire le contenu de cette URL.")
    else:
        st.error("Veuillez remplir tous les champs.")

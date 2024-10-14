import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
import re
import io

def extract_content(url):
    """Extrait le contenu HTML d'une page web et les titres."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        content = []
        
        # Extraire les titres et le contenu
        for element in soup.body.find_all(['h1', 'h2', 'h3', 'p']):
            if element.name == 'h1':
                content.append((element.get_text(), 1))  # Titre de niveau 1
            elif element.name == 'h2':
                content.append((element.get_text(), 2))  # Titre de niveau 2
            elif element.name == 'h3':
                content.append((element.get_text(), 3))  # Titre de niveau 3
            elif element.name == 'p':
                content.append((element.get_text(), 4))  # Paragraphe

        return content
    except requests.RequestException as e:
        st.error(f"Erreur lors de la récupération de l'URL: {e}")
        return None

def create_word_file(content, jira_link):
    """Crée un fichier Word avec le contenu extrait et le lien JIRA."""
    jira_number = re.search(r'TT-(\d{4})', jira_link)
    if jira_number:
        title = f"Brief SEO Optimization - TT-{jira_number.group(1)}"
    else:
        st.error("Le lien JIRA doit contenir le format 'TT-XXXX'.")
        return None, None

    doc = Document()
    doc.add_heading(title, level=1)

    # Ajouter le contenu au document Word
    for text, level in content:
        if level == 1:
            doc.add_heading(text, level=1)
        elif level == 2:
            doc.add_heading(text, level=2)
        elif level == 3:
            doc.add_heading(text, level=3)
        elif level == 4:
            doc.add_paragraph(text)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output, title

# Configuration de l'application Streamlit
st.title("Extracteur de contenu HTML pour WordPress")

url = st.text_input("Insérer l'URL de la page:")
jira_link = st.text_input("Ajouter le lien JIRA:")

if st.button("Extraire et créer le fichier Word"):
    if url and jira_link:
        content = extract_content(url)
        if content:
            word_file, title = create_word_file(content, jira_link)
            if word_file:
                st.success(f"Fichier Word créé: {title}.docx")
                st.download_button(
                    label="Télécharger le fichier",
                    data=word_file,
                    file_name=f"{title}.docx",  # Correction ici : 'file_name'
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.warning("Veuillez remplir tous les champs.")

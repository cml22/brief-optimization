import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
import re
import io

def extract_content(url):
    """Extrait le contenu HTML d'une page web."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extraire le texte du body
        content = soup.body.get_text(separator='\n')
        return content
    except requests.RequestException as e:
        st.error(f"Erreur lors de la récupération de l'URL: {e}")
        return None

def create_word_file(content, jira_link):
    """Crée un fichier Word avec le contenu extrait et le lien JIRA."""
    # Extraire le numéro JIRA
    jira_number = re.search(r'TT-(\d{4})', jira_link)
    if jira_number:
        title = f"Brief SEO Optimization - TT-{jira_number.group(1)}"
    else:
        st.error("Le lien JIRA doit contenir le format 'TT-XXXX'.")
        return None, None

    # Créer un document Word dans un objet BytesIO
    doc = Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(content)

    # Enregistrer le document dans un objet BytesIO
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)  # Remettre le curseur au début de l'objet BytesIO
    return output, title

# Configuration de l'application Streamlit
st.title("Extracteur de contenu HTML 

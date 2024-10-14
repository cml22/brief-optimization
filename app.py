import requests
from bs4 import BeautifulSoup
from docx import Document
import streamlit as st

# Fonction pour extraire le texte à partir de la balise <h1> et au-delà
def extract_content_from_url(url):
    response = requests.get(url)
    # Spécifier l'encodage correct
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')

    # Trouver le contenu à partir du premier <h1>
    content = ""
    start = False
    for element in soup.find_all(['h1', 'h2', 'h3', 'p']):
        if element.name == 'h1':
            start = True
        if start:
            if element.name == 'h1':
                content += f'\n# {element.get_text()}\n'
            elif element.name == 'h2':
                content += f'\n## {element.get_text()}\n'
            elif element.name == 'h3':
                content += f'\n### {element.get_text()}\n'
            elif element.name == 'p':
                content += f'{element.get_text()}\n'
    return content

# Fonction pour créer un document Word
def create_word_file(jira_link, content):
    # Extraire les 4 derniers chiffres du lien JIRA
    ticket_number = jira_link[-4:]
    filename = f"Brief SEO Optimization - TT-{ticket_number}.docx"

    doc = Document()
    doc.add_heading(f'Brief SEO Optimization - TT-{ticket_number}', 0)

    # Ajouter le contenu extrait au fichier Word
    for line in content.split("\n"):
        if line.startswith("# "):
            doc.add_heading(line[2:], level=1)
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)
        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        else:
            doc.add_paragraph(line)
    
    # Sauvegarder le document
    doc.save(filename)
    return filename

# Interface Streamlit
st.title("Extracteur de contenu HTML vers Word")
url = st.text_input("Insérer l'URL de la page")
jira_link = st.text_input("Ajouter le lien JIRA")

if st.button("Générer le fichier Word"):
    if url and jira_link:
        # Extraire le contenu de l'URL
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

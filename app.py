import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
import os

def create_word_file(filename, content, jira_link):
    document = Document()
    document.add_heading('Contenu extrait', level=1)

    # Ajout du contenu avec les liens simulés
    for part in content:
        paragraph = document.add_paragraph()
        add_hyperlink(paragraph, part['url'], part['text'])

    # Ajouter le lien JIRA à la fin du document
    if jira_link:
        document.add_paragraph('Lien JIRA :')
        paragraph = document.add_paragraph()
        add_hyperlink(paragraph, jira_link, jira_link)

    # Enregistrement du document
    document.save(filename)

def add_hyperlink(paragraph, url, text):
    """
    Simule un lien hypertexte dans un document Word en ajoutant du texte bleu et souligné.
    """
    run = paragraph.add_run(text)
    run.font.color.rgb = RGBColor(0, 0, 255)  # Couleur bleue pour le texte du lien
    run.font.underline = True  # Texte souligné pour simuler un lien hypertexte

    # Ajoute l'URL entre parenthèses après le texte du lien
    paragraph.add_run(f" ({url})")

def extract_content(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Extraire tous les liens
    content = []
    for link in soup.find_all('a'):
        content.append({
            'text': link.get_text() or "Lien sans texte",  # Gérer les liens sans texte
            'url': link.get('href') or "#"  # Gérer les liens sans URL
        })
    
    return content

# Application Streamlit
st.title('Extraction de contenu HTML vers Word')
url_input = st.text_input('Entrez l\'URL de la page')
jira_input = st.text_input('Ajouter le lien JIRA (TT - Traffic Team)')
if st.button('Créer le fichier Word'):
    if url_input:
        content = extract_content(url_input)
        if content:  # Vérifie que du contenu a été extrait
            filename = 'extracted_content.docx'
            create_word_file(filename, content, jira_input)
            st.success(f'Fichier Word créé : {filename}')
            
            # Offrir le fichier à télécharger
            with open(filename, 'rb') as f:
                st.download_button('Télécharger le fichier Word', f, file_name=filename)

            # Supprimer le fichier local après l'avoir téléchargé
            os.remove(filename)
        else:
            st.error('Aucun lien trouvé sur la page.')
    else:
        st.error('Veuillez entrer une URL.')

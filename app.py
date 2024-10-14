import requests
from bs4 import BeautifulSoup
from docx import Document
import streamlit as st
import html

# Fonction pour extraire le contenu à partir des balises <h1> jusqu'à <h6> et les paragraphes <p>, tout en conservant les liens <a>
def extract_content_from_url(url):
    response = requests.get(url)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')

    content = []
    start = False
    # Ignorer la div spécifique contenant le texte à retirer
    for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'div']):
        if element.name == 'h1':
            start = True
        if start:
            # Vérifie si l'élément est la div à ignorer
            if element.name == 'div' and 'container text-center text-level--sm pb-4' in element['class']:
                continue
            
            text = element.get_text().strip()
            text = html.unescape(text)  # Convertit les entités HTML en caractères normaux
            if text:
                if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    content.append({'type': 'heading', 'level': element.name, 'text': text, 'element': element})
                elif element.name == 'p':
                    paragraph = ""
                    for sub_element in element:
                        if sub_element.name == 'a' and sub_element.get('href'):
                            paragraph += f'{sub_element.get_text()} ({sub_element.get("href")}) '
                        else:
                            paragraph += sub_element.string if sub_element.string else ''
                    content.append({'type': 'paragraph', 'text': paragraph.strip(), 'element': element})

    return content

# Fonction pour créer un document Word à partir du contenu extrait
def create_word_file(jira_link, content):
    ticket_number = jira_link[-4:]
    filename = f"Brief SEO Optimization - TT-{ticket_number}.docx"

    doc = Document()

    # Commencer directement avec le contenu de H1 comme titre principal
    for element in content:
        if element['type'] == 'heading' and element['level'] == 'h1':
            doc.add_heading(element['text'], level=0)  # Titre principal H1
            break  # Ne pas ajouter d'autre titre principal

    # Ajouter les autres headings et paragraphes dans le fichier Word avec formatage
    for element in content:
        if element['type'] == 'heading' and element['level'] != 'h1':
            level = int(element['level'][1])  # Niveau de heading (h2 = 2, h3 = 3, etc.)
            doc.add_heading(element['text'], level=level)
        elif element['type'] == 'paragraph':
            p = doc.add_paragraph()
            for sub_element in element['element']:
                if sub_element.name == 'a' and sub_element.get('href'):
                    # Ajout des liens cliquables sans l'URL affichée
                    add_hyperlink(p, sub_element.get('href'), sub_element.get_text())
                else:
                    p.add_run(sub_element.string if sub_element.string else '')

    # Sauvegarder le fichier Word
    doc.save(filename)
    return filename

# Fonction pour ajouter un lien hypertexte dans un paragraphe Word
def add_hyperlink(paragraph, url, text):
    part = paragraph.add_run(text)
    part.font.color.theme_color = 10  # Couleur du lien
    part.font.underline = True
    paragraph.add_run(" ")

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

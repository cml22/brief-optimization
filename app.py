import requests
from bs4 import BeautifulSoup
from docx import Document
import streamlit as st

# Fonction pour extraire le contenu à partir de <h1> jusqu'aux balises <h6>, <p> et les liens <a>
def extract_content_from_url(url):
    response = requests.get(url)
    # Utilisation de l'encodage UTF-8 pour éviter les caractères mal encodés
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')

    # Extraire le contenu du body et le nettoyer
    content = []
    start = False
    for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p']):
        if element.name == 'h1':
            start = True
        if start:
            text = element.get_text().strip()  # Retirer les espaces et sauts de ligne inutiles
            if text:  # Si le texte n'est pas vide
                if element.name == 'h1':
                    content.append(f'# {text}')
                elif element.name == 'h2':
                    content.append(f'## {text}')
                elif element.name == 'h3':
                    content.append(f'### {text}')
                elif element.name == 'h4':
                    content.append(f'#### {text}')
                elif element.name == 'h5':
                    content.append(f'##### {text}')
                elif element.name == 'h6':
                    content.append(f'###### {text}')
                elif element.name == 'p':
                    # Ajouter du texte avec les liens intégrés
                    paragraph = ""
                    for sub_element in element:
                        if sub_element.name == 'a' and sub_element.get('href'):
                            # Ajouter le lien sous la forme [texte](URL)
                            paragraph += f'{sub_element.get_text()} ({sub_element.get("href")}) '
                        else:
                            paragraph += sub_element.string if sub_element.string else ''
                    content.append(paragraph.strip())
    
    return "\n".join(content)

# Fonction pour créer un document Word à partir du contenu extrait
def create_word_file(jira_link, content):
    # Extraire les 4 derniers chiffres du lien JIRA
    ticket_number = jira_link[-4:]
    filename = f"Brief SEO Optimization - TT-{ticket_number}.docx"

    doc = Document()
    doc.add_heading(f'Brief SEO Optimization - TT-{ticket_number}', 0)

    # Ajouter le contenu extrait dans le fichier Word avec formatage
    for line in content.split("\n"):
        line = line.strip()  # Supprimer les espaces inutiles
        if line.startswith("# "):
            doc.add_heading(line[2:], level=1)
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)
        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        elif line.startswith("#### "):
            doc.add_heading(line[5:], level=4)
        elif line.startswith("##### "):
            doc.add_heading(line[6:], level=5)
        elif line.startswith("###### "):
            doc.add_heading(line[7:], level=6)
        else:
            # Ajouter un paragraphe pour le texte normal sans saut de ligne inutile
            if line:
                doc.add_paragraph(line)

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

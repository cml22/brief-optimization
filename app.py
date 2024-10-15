import streamlit as st
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def add_hyperlink(paragraph, url, text):
    """Ajoute un lien hypertexte à un paragraphe."""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Créer un élément XML pour le lien hypertexte
    hyperlink = parse_xml(r'<w:hyperlink {}>'.format(nsdecls('w')))
    hyperlink.set('r:id', r_id)

    # Créer un run pour le texte du lien
    run = paragraph.add_run(text)
    run.font.color.rgb = (0, 0, 255)  # Couleur bleue
    run.font.underline = True  # Souligner pour indiquer un lien

    # Ajouter le run à l'élément hyperlink
    paragraph._element.append(hyperlink)

def create_word_file(filename, content):
    """Crée un fichier Word avec le contenu spécifié."""
    document = Document()

    for part in content:
        paragraph = document.add_paragraph()
        text = part[0]
        url = part[1]

        # Gestion des liens hypertexte uniquement s'il y a un texte et une URL
        if url and text:
            add_hyperlink(paragraph, url, text)
        else:
            paragraph.add_run(text)  # Ajouter le texte sans lien si pas d'URL

    document.save(filename)

# Interface Streamlit
st.title("Générateur de document Word avec liens hypertextes")

# Entrée pour le nom du fichier
filename = st.text_input("Nom du fichier Word (sans extension) :")
filename = filename + ".docx" if filename else ""

# Entrée pour le contenu
content = []
num_entries = st.number_input("Nombre d'entrées de texte :", min_value=1, value=1)

for i in range(num_entries):
    text = st.text_input(f"Texte {i + 1} :")
    url = st.text_input(f"URL {i + 1} :")
    content.append((text, url))

# Bouton pour générer le document
if st.button("Créer le document"):
    if filename and content:
        create_word_file(filename, content)
        st.success(f"Document '{filename}' créé avec succès !")
    else:
        st.error("Veuillez fournir un nom de fichier et au moins une entrée de texte.")

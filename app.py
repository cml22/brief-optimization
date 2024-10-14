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
            p = doc.add_paragraph()
            paragraph = element['text']
            for sub_element in element.get('element', []):
                if sub_element.name == 'a' and sub_element.get('href'):
                    add_hyperlink(p, sub_element.get('href'), sub_element.get_text())
                else:
                    p.add_run(sub_element.string if sub_element.string else '')

    # Sauvegarder le fichier Word
    doc.save(filename)
    return filename

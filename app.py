from docx import Document
from docx.shared import Inches

# Function to create a Word document from the extracted content
def create_word_file(file_name, content, url):
    doc = Document()
    
    # Add the URL as the first paragraph
    doc.add_paragraph(f"URL: {url}")

    # Add extracted content to the Word file with formatting
    for element in content:
        if element['type'] == 'heading':
            level = int(element['level'][1])  # Heading level (h1 = 1, h2 = 2, etc.)
            doc.add_heading(element['text'], level=level)
        elif element['type'] == 'paragraph':
            # Add paragraph text, replacing the hyperlink as anchor text
            paragraph = doc.add_paragraph()
            for sub_element in element['text'].split(' '):
                # Check if the part is a hyperlink (contains " (")
                if '(' in sub_element and sub_element.endswith(')'):
                    # Extract the text and URL
                    anchor_text = sub_element[:sub_element.index(' (')]
                    url = sub_element[sub_element.index('(') + 1:-1]  # Extract URL
                    run = paragraph.add_run(anchor_text)
                    run.bold = True  # Optionally, make the link bold
                    # Create hyperlink
                    hyperlink = doc.add_paragraph().add_run(anchor_text)
                    hyperlink.font.color.rgb = (0, 0, 255)  # Blue color for hyperlinks
                    hyperlink.font.underline = True  # Underline for hyperlinks
                    # This method does not create a functional hyperlink; it's just for display
                    # You may need to use python-docx with an additional library or workaround to create actual hyperlinks.
                else:
                    paragraph.add_run(sub_element + ' ')

    # Save the Word file
    doc.save(file_name)
    return file_name

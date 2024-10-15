import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import streamlit as st
import html

# Function to extract content from <h1> to <h6>, <p> tags and <a> links
def extract_content_from_url(url):
    response = requests.get(url)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')

    content = []
    h1_text = ""
    start = False
    for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p']):
        if element.name == 'h1':
            h1_text = element.get_text().strip()
            start = True
        if start:
            text = element.get_text().strip()
            text = html.unescape(text)  # Converts HTML entities into normal characters
            if text:
                if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    content.append({'type': 'heading', 'level': element.name, 'text': text})
                elif element.name == 'p':
                    paragraph = []
                    for sub_element in element:
                        if sub_element.name == 'a' and sub_element.get('href'):
                            anchor_text = sub_element.get_text()
                            link_url = sub_element.get("href")
                            paragraph.append(('link', anchor_text, link_url))
                        else:
                            if sub_element.string:
                                paragraph.append(('text', sub_element.string))
                    content.append({'type': 'paragraph', 'content': paragraph})

    return content, h1_text

# Function to add a hyperlink to a Word document
def add_hyperlink(paragraph, url, text):
    # This will create a hyperlink
    part = paragraph.add_run(text)
    r_id = paragraph.part.relate_to(url, "hyperlink", is_external=True)
    part._element.get_or_add_rPr().append(parse_xml(r'<a:blip r:embed="{}"/>'.format(r_id)))
    part.font.color.rgb = RGBColor(0, 0, 255)  # Set the hyperlink color to blue
    part.font.underline = True  # Underline the hyperlink

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
            paragraph = doc.add_paragraph()
            for part in element['content']:
                if part[0] == 'text':
                    paragraph.add_run(part[1])
                elif part[0] == 'link':
                    # Add hyperlink
                    add_hyperlink(paragraph, part[2], part[1])  # Add hyperlink with URL and text

    # Save the Word file
    doc.save(file_name)
    return file_name

# Streamlit interface
st.title("HTML Content Extractor to Word")
url = st.text_input("Enter the page URL")
jira_link = st.text_input("Add the JIRA link (TT - Traffic Team)")

if st.button("Generate Word File"):
    if url:
        content, h1_text = extract_content_from_url(url)
        if content:
            if jira_link:
                ticket_number = jira_link[-4:]
                filename = f"Brief SEO Optimization - TT-{ticket_number}.docx"
            else:
                filename = f"{h1_text}.docx"

            create_word_file(filename, content, url)

            with open(filename, "rb") as file:
                st.download_button(
                    label="Your file is ready! Click to download",
                    data=file,
                    file_name=filename
                )
        else:
            st.error("Unable to extract content from this URL.")
    else:
        st.error("Please fill out the URL field.")

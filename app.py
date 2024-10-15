import requests
from bs4 import BeautifulSoup
from docx import Document
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
            h1_text = element.get_text().strip()  # Capture the H1 text
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
                            # Append link text and URL as a tuple
                            anchor_text = sub_element.get_text()
                            url = sub_element.get("href")
                            paragraph.append({'text': anchor_text, 'url': url})
                        elif sub_element.string:
                            paragraph.append({'text': sub_element.string})
                    content.append({'type': 'paragraph', 'text': paragraph})

    return content, h1_text  # Return the H1 text as well

# Function to add a hyperlink to a Word document
def add_hyperlink(paragraph, url, text):
    # Add the hyperlink as text with an underlying link
    run = paragraph.add_run(text)
    run.bold = True  # Optionally make the hyperlink bold
    # Create the hyperlink
    r_id = paragraph.part.relate_to(url, "hyperlink", is_external=True)
    run._element.get_or_add_rPr().append(parse_xml(r'<a:blip r:link="{}"/>'.format(r_id)))

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
            # Loop through the paragraph content to add text and hyperlinks
            for part in element['text']:
                if 'url' in part:
                    # If there's a URL, add it as a hyperlink
                    add_hyperlink(paragraph, part['url'], part['text'])
                else:
                    # Otherwise, just add the text
                    paragraph.add_run(part['text'])
    
    # Save the Word file
    doc.save(file_name)
    return file_name

# Streamlit interface
st.title("HTML Content Extractor to Word")
url = st.text_input("Enter the page URL")
jira_link = st.text_input("Add the JIRA link (TT - Traffic Team)")

if st.button("Generate Word File"):
    if url:
        # Extract content from the URL
        content, h1_text = extract_content_from_url(url)
        if content:
            # Determine the file name based on the JIRA link
            if jira_link:
                ticket_number = jira_link[-4:]
                filename = f"Brief SEO Optimization - TT-{ticket_number}.docx"
            else:
                filename = f"{h1_text}.docx"  # Use H1 text if JIRA link is empty

            # Create the Word file
            create_word_file(filename, content, url)  # Pass URL to the function

            with open(filename, "rb") as file:
                # Automatically trigger the download
                st.download_button(
                    label="Your file is ready! Click to download",
                    data=file,
                    file_name=filename
                )
        else:
            st.error("Unable to extract content from this URL.")
    else:
        st.error("Please fill out the URL field.")

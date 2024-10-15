import requests
from bs4 import BeautifulSoup
from htmldocx import HtmlToDocx
import streamlit as st

# Function to extract HTML content from the URL
def extract_html_from_url(url):
    response = requests.get(url)
    response.encoding = 'utf-8'
    return response.text

# Function to create a Word document from the HTML content
def create_word_from_html(file_name, html_content):
    new_parser = HtmlToDocx()
    docx_content = new_parser.parse_html_string(html_content)
    docx_content.save(file_name)
    return file_name

# Streamlit interface
st.title("HTML Content Extractor to Word")
url = st.text_input("Enter the page URL")
jira_link = st.text_input("Add the JIRA link (TT - Traffic Team)")

if st.button("Generate Word File"):
    if url:
        # Extract HTML content from the URL
        html_content = extract_html_from_url(url)
        if html_content:
            # Determine the file name based on the JIRA link
            if jira_link:
                ticket_number = jira_link[-4:]
                filename = f"Brief SEO Optimization - TT-{ticket_number}.docx"
            else:
                filename = f"content.docx"

            # Create the Word file
            create_word_from_html(filename, html_content)

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

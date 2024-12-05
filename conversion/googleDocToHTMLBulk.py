import os
import pypandoc
import bs4
import re
import base64
from docx import Document
from io import BytesIO
from PIL import Image

def convert_image_to_base64(image_bytes):
    """Convert image bytes to a base64 string."""
    image_stream = BytesIO(image_bytes)
    img = Image.open(image_stream)

    # Convert image to base64
    buffer = BytesIO()
    img.save(buffer, format=img.format)
    encoded_image = base64.b64encode(buffer.getvalue()).decode('utf-8')
    
    return f"data:image/{img.format.lower()};base64,{encoded_image}"

def extract_images_from_docx(docx_file, soup):
    """Extract images from .docx and embed them into the HTML as base64."""
    doc = Document(docx_file)

    # Get all the image relationships from the document
    image_relations = {}
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            # Read the image binary data
            image_data = rel.target_part.blob
            base64_image = convert_image_to_base64(image_data)
            image_relations[rel.target_ref] = base64_image

    # Replace <img> tags in HTML with base64-encoded images
    for img_tag in soup.find_all('img'):
        if img_tag.get('src') in image_relations:
            img_tag['src'] = image_relations[img_tag['src']]

def convert_docx_to_html(docx_file, output_folder, disruptive_string=None):
    # Convert .docx to HTML using pypandoc
    output = pypandoc.convert_file(docx_file, 'html')

    # Parse the HTML output using BeautifulSoup
    soup = bs4.BeautifulSoup(output, 'html.parser')

    # Check if there is already an <h1> tag
    h1_exists = soup.find('h1') is not None

    # Only convert the first <p> tag to <h1> if there's no existing <h1>
    if not h1_exists:
        first_p = soup.find('p')
        if first_p:
            first_p.name = 'h1'

    # Remove titles matching the pattern 'Title 1', 'Title 2', etc.
    title_pattern = re.compile(r'Title\s*\d+', re.IGNORECASE)
    for h1 in soup.find_all('h1'):
        if title_pattern.match(h1.get_text(strip=True)):
            h1.decompose()  # Remove the <h1> tag containing the title

    # Embed images as base64 in HTML
    extract_images_from_docx(docx_file, soup)

    # Clean up unwanted characters (e.g., \r or \n) and disruptive strings
    cleaned_output = str(soup).replace('\r', ' ').replace('\n', '').replace('<br>', '<br>')

    # If there's a specific disruptive string, remove it
    if disruptive_string:
        cleaned_output = cleaned_output.replace(disruptive_string, '')

    # Define the output file path
    html_file = os.path.join(output_folder, os.path.splitext(os.path.basename(docx_file))[0] + '.html')

    # Write the cleaned HTML to a file
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(cleaned_output)

    print(f"Converted {docx_file} to {html_file}")

def bulk_convert_docx_to_html(docx_folder, output_folder, disruptive_string=None):
    for docx_file in os.listdir(docx_folder):
        if docx_file.endswith(".docx"):
            convert_docx_to_html(os.path.join(docx_folder, docx_file), output_folder, disruptive_string)
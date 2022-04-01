from PIL import Image
from pdf2image import convert_from_path
import pytesseract as tess
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docxcompose.composer import Composer
import os

#pytesseract path
tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

file_list = []
image_names = []

# converts pdf to image 
def pdf_to_image(pdfs, page_numbers):
    global string_name
    pages = convert_from_path(pdfs, 350)
    name = pdfs.split('\\')
    nameWithext = name[-1].split('.')
    string_name = nameWithext[0]

    i = 1

    for page in pages:
        print('------------------------------------------')
        print('Creating jpg file...')
        image_name = string_name + " Page_" + str(i) + ".jpg"  
        page.save(image_name, "JPEG")
        i = i+1  
        image_names.append(image_name)
        image_to_text(image_name, page_numbers)

# converts image produced by previous function to readable text through pytesseract
def image_to_text(imageName, page_numbers):
    print('------------------------------------------')
    print('Converting image to text...')
    img = Image.open(imageName)
    text = tess.image_to_string(img)
    text_to_document(imageName, text, page_numbers)

# creates a new document from the text produced through python-docx module
def text_to_document(imageName, text, page_numbers):
    print('------------------------------------------')
    print('Creating document from text...')
    document = Document()
    document_name = imageName +'.docx'
    file_list.append(document_name)
    paragraph = document.add_paragraph(f"""{text}""")
    paragraph.style = document.styles.add_style('Style Name', WD_STYLE_TYPE.PARAGRAPH)
    font = paragraph.style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    document.save(document_name)
    
    # if the length of the list of file_names is not equal to the page numbers, 
    # do not proceed to next function
    if len(file_list) == int(page_numbers):
        combine_all_docx(file_list)

# merges all documents into one file through docxcompose module
def combine_all_docx(files_list):
    print('------------------------------------------')
    print('Combining produced documents...')
    number_of_sections=len(files_list)
    master = Document(f'{files_list[0]}')
    composer = Composer(master)
    for i in range(1, number_of_sections):
        doc_temp = Document(files_list[i])
        composer.append(doc_temp)
    composer.save(f"{string_name}.docx")
    delete_jpg_and_pdf(files_list, number_of_sections)

# deletes jpgs and pdfs created earlier
def delete_jpg_and_pdf(files_list, number_of_sections):
    print('------------------------------------------')
    print('Deleting duplicates...')
    for i in range(0, number_of_sections):
        os.remove(files_list[i])
        os.remove(image_names[i])
    print('------------------------------------------')
    print('Task completed')

if __name__ == '__main__':
    pdf_to_image(input("   Image PDF to DOCX  \n-----------------------\nEnter Directory of File:  \n"), int(input('How many pages does it have?: ')))
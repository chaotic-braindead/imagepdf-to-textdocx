from PIL import Image as Img
from pdf2image import convert_from_path
import pytesseract as tess
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
import os
from tkinter import *
from tkinter import filedialog

#pytesseract path
tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

file_list = []
image_names = []

# converts pdf to image 
def pdf_to_image(pdfs):
    global name
    global pages
    pages = convert_from_path(pdfs, 350)
    name = pdfs.split('\\')[-1].split('.')[0]
    i = 1

    for page in pages:
        print('------------------------------------------')
        print('Creating jpg file...')
        image_name = name + " Page_" + str(i) + ".jpg"  
        page.save(image_name, "JPEG")
        i = i+1  
        image_names.append(image_name)
        image_to_text(image_name)

# converts image produced by previous function to readable text through pytesseract
def image_to_text(imageName):
    print('------------------------------------------')
    print('Converting image to text...')
    img = Img.open(imageName)
    text = tess.image_to_string(img)
    text_to_document(imageName, text)

# creates a new document from the text produced through python-docx module
def text_to_document(imageName, text):
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
    if len(file_list) == len(pages):
        combine_word_document(file_list)

def combine_word_document(files_list):
    merged_document = Document()
    for index, file in enumerate(files_list):
        sub_doc = Document(file)
        if index < len(files_list)-1:
            sub_doc.add_page_break()

        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    merged_document.save(f'{name}.docx')
    delete_jpg_and_pdf(files_list)


# deletes jpgs and pdfs created earlier
def delete_jpg_and_pdf(files_list):
    print('------------------------------------------')
    print('Deleting duplicates...')
    for i in range(len(files_list)):
        os.remove(files_list[i])
        os.remove(image_names[i])
    print('------------------------------------------')
    print('Task completed')


if __name__ == '__main__':
    pdf_to_image(input("   Image PDF to DOCX  \n-----------------------\nEnter Directory of File:  \n"))

from PIL import Image as Img
from pdf2image import convert_from_path
import pdf2image
import pytesseract as tess
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
import os
from tkinter import *
from tkinter import filedialog
import time
import threading

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

    files_list.clear()
    image_names.clear()
    print('------------------------------------------')
    print('Task completed')

if __name__ == '__main__':
    root = Tk()
    root.title('PDF to Selectable Text DOCX Converter')
    root.iconbitmap(r'C:\Users\raf\Downloads\folder.ico')
    root.geometry('500x200')
    root.configure(bg='#4f4d4d')
    my_text = Text(root, width=50, height=1, font=('Segoe', 10))
    my_text.grid(row=6, column=8, sticky="nsew", padx=2, pady=2)
    my_text.place(x=25, y=25)
    # my_text.pack(pady=25)
    my_text.configure(state='disabled',)

    def getlabel(status):
        if status == 'Converting...':
            my_label.place(x=205, y=75)
            my_label.config(text=status, bg='#ffffff')
        elif status == 'Select a file to convert':
            my_label.place(x=193, y=75)
            my_label.config(text=status, bg='#ffffff')      
        else:
            my_label.place(x=157, y=75)
            my_label.config(text=status, bg='#53fc8e')
            

    def bttn(width, height, x, y, text, bcolor, fcolor, cmd):

        button = Button(root, width=width, height=height, text=text,
        fg=bcolor, bg=fcolor, border=0, activeforeground=fcolor,
        command=cmd, font=('Helvetica', 10, 'bold'))

        button.place(x=x, y=y)

    def getpdf():
        my_text.configure(state='normal')
        path = filedialog.askopenfilename(initialdir=r'C:\Users\raf\Downloads', title='Select a PDF file to convert', filetypes=(('pdf files','*.pdf'), ('All files', '*.*')))
        my_text.delete(1.0, END)
        my_text.insert('1.0', path)
        my_text.configure(state='disabled')
        getlabel('Select a file to convert')

    def convert():
        path = my_text.get('1.0', END).rstrip()
        getlabel('Converting...')
        pdf_to_image(path)
        getlabel('Successfully converted .pdf to .docx!')
    

    bttn(10, 1, 395, 23, 'Browse', '#ffffff', '#050404', getpdf)
    bttn(20, 2, 166, 120, 'Convert', '#ffffff', '#050404', convert)
    
    my_label = Label(root, font=('Segoe', 8), text='Select a file to convert')
    my_label.place(x=193, y=75)

    
    root.mainloop()
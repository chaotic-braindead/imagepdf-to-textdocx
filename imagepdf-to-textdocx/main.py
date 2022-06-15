from PIL import Image as Img
from pdf2image import convert_from_path
import pytesseract as tess
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
import os
from tkinter import *
from tkinter import filedialog

# pytesseract path
tess.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

file_list = []

# converts pdf to image
class PDFtoDocx:
    def __init__(self, dir: str) -> None:
        self.__dir = dir

    @property
    def dir(self) -> str:
        return self.__dir

    @property
    def name(self) -> str:
        return self.dir.split("\\")[-1].split(".")[0]

    def pdf_to_image(self, export_as_docx=True) -> None:
        pages = convert_from_path(self.dir, 350)
        i = 1
        image_names = []
        name = self.name

        print("------------------------------------------")
        print("Creating jpg file...")

        if len(pages) > 1:
            for page in pages:
                image_name = name + " Page_" + str(i) + ".jpg"
                page.save(image_name, "JPEG")
                i = i + 1
                image_names.append(image_name)
        else:
            page.save(image_name, "JPEG")
            image_names.append(image_name)

        if export_as_docx is True:
            self.image_to_docx(image_names)

            print("------------------------------------------")
            print("Deleting duplicates...")
            for i in range(len(image_names)):
                os.remove(image_names[i])
            print("------------------------------------------")
            print("Task completed")

    def image_to_docx(self, image_names):
        print("------------------------------------------")
        print("Converting image to text...")

        text = ""
        print(image_names)
        for image in image_names:
            img = Img.open(image)
            text += f"{tess.image_to_string(img)}\n"

        print("------------------------------------------")
        print("Creating document from text...")

        document = Document()
        document_name = self.name + ".docx"

        paragraph = document.add_paragraph(f"""{text}""")
        paragraph.style = document.styles.add_style(
            "Style Name", WD_STYLE_TYPE.PARAGRAPH
        )
        font = paragraph.style.font
        font.name = "Times New Roman"
        font.size = Pt(12)
        document.save(document_name)


if __name__ == "__main__":

    root = Tk()
    root.title("PDF to Selectable Text DOCX Converter")
    root.iconbitmap(r"C:\Users\raf\Downloads\folder.ico")
    root.geometry("500x200")
    root.configure(bg="#4f4d4d")
    my_text = Text(root, width=50, height=1, font=("Segoe", 10))
    my_text.grid(row=6, column=8, sticky="nsew", padx=2, pady=2)
    my_text.place(x=25, y=25)
    # my_text.pack(pady=25)
    my_text.configure(
        state="disabled",
    )

    def getlabel(status):
        if status == "Converting...":
            my_label.place(x=205, y=75)
            my_label.config(text=status, bg="#ffffff")
        elif status == "Select a file to convert":
            my_label.place(x=193, y=75)
            my_label.config(text=status, bg="#ffffff")
        else:
            my_label.place(x=157, y=75)
            my_label.config(text=status, bg="#53fc8e")

    def bttn(width, height, x, y, text, bcolor, fcolor, cmd):

        button = Button(
            root,
            width=width,
            height=height,
            text=text,
            fg=bcolor,
            bg=fcolor,
            border=0,
            activeforeground=fcolor,
            command=cmd,
            font=("Helvetica", 10, "bold"),
        )

        button.place(x=x, y=y)

    def getpdf():
        my_text.configure(state="normal")
        path = filedialog.askopenfilename(
            initialdir=r"C:\Users\raf\Downloads",
            title="Select a PDF file to convert",
            filetypes=(("pdf files", "*.pdf"), ("All files", "*.*")),
        )
        my_text.delete(1.0, END)
        my_text.insert("1.0", path)
        my_text.configure(state="disabled")
        getlabel("Select a file to convert")

    def convert():
        path = my_text.get("1.0", END).rstrip()
        getlabel("Converting...")
        conv = PDFtoDocx(path)
        conv.pdf_to_image()
        getlabel("Successfully converted .pdf to .docx!")

    bttn(10, 1, 395, 23, "Browse", "#ffffff", "#050404", getpdf)
    bttn(20, 2, 166, 120, "Convert", "#ffffff", "#050404", convert)

    my_label = Label(root, font=("Segoe", 8), text="Select a file to convert")
    my_label.place(x=193, y=75)

    root.mainloop()

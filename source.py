from tkinter import *
from tkinter import filedialog, messagebox  # file processing lib
from PIL import Image  # image processing lib
from docx.shared import Inches  # image in doc to pdf process
from reportlab.pdfgen import canvas  # more GUI widgets
from reportlab.lib.pagesizes import letter  # more GUI widgets
from docx2pdf import convert  # image in word convertion with LibreOffice


def convert_images_to_pdf():
    images = select_images()
    if images:
        pdf_jpg = select_save_pdf()
        if pdf_jpg:
            try:
                # Create a new PDF file
                pdf = Image.open(images[0])
                pdf.save(pdf_jpg, "PDF", resolution=100.0,
                         save_all=True, append_images=images[1:])
                messagebox.showinfo(
                    "Success", "Images have been successfully converted to PDF.")
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to convert to PDF.\nError: {str(e)}")


def convert_doc_to_pdf():
    word_document = select_doc()
    if word_document:
        pdf_jpg = select_save_pdf()
        if pdf_jpg:
            try:
                convert(word_document)
                messagebox.showinfo(
                    "Success", "Word Documents have been successfully converted to PDF.")
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to convert to PDF.\nError: {str(e)}")


def select_doc():
    doc = filedialog.askopenfilename(
        title="Select Doc",
        defaultextension=".docx",
        initialdir="C:/",
        filetypes=(("Doc files", "*.docx"), ("All files", "*.*"))
    )
    return doc


def select_images():
    images = filedialog.askopenfilenames(
        title="Select Images",
        filetypes=(("Image files", "*.jpg;*.jpeg;*.png"),
                   ("All files", "*.*")),
        initialdir="C:/")
    return images


def select_save_pdf():
    save_pdf = filedialog.asksaveasfilename(title="Save PDF as", defaultextension=".pdf",
                                            initialdir="C:/", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
    return save_pdf


# Frontend
background = 'Yellow'
button_font = ("Times New Roman", 13)
button_background = 'Turquoise'


root = Tk()
root.geometry('250x400')
root.resizable(0, 0)
root.config(bg=background)

root.title("PDF converter")
root.iconphoto(False, PhotoImage(file='icon.png'))

Button(root, text='DOC to PDF', width=20, font=button_font,
       bg=button_background, command=convert_doc_to_pdf).place(x=30, y=50)
Button(root, text='JPG to PDF', width=20, font=button_font,
       bg=button_background, command=convert_images_to_pdf).place(x=30, y=90)


# Finalize window
root.update()
root.mainloop()

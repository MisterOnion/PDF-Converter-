import tkinter as tk
from tkinter import filedialog, messagebox  # file processing lib
from PIL import Image  # image processing lib
from docx import Document  # doc to pdf process
from docx.shared import Inches  # image in doc to pdf process
from io import BytesIO  # binary conversion
from reportlab.pdfgen import canvas  # more GUI widgets
from reportlab.lib.pagesizes import letter  # more GUI widgets
from docx2pdf import convert # image in word convertion with LibreOffice


class PDFConverter:
    def __init__(self, master):
        self.master = master
        master.title("JPG to PDF converter")
        master.iconphoto(False, tk.PhotoImage(file='icon.png'))

        # Menu bar
        # menubar = tk.Menu(master)
        # master.config(menu=menubar)

        # Create Menu Widgets
        # file_menu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label="File", menu=file_menu)

        # Canvas
        self.canvas = tk.Canvas(master, width=300, height=200)
        self.convert_JPG_btn = tk.Button(
            master, text="Convert JPG", command=self.convert_images_to_pdf)

        self.canvas = tk.Canvas(master, width=300, height=200)
        self.convert_DOC_btn = tk.Button(
            master, text="Convert DOC", command=self.convert_doc_to_pdf)

    def convert_images_to_pdf(self):
        images = self.select_images()
        if images:
            pdf_name = self.select_image_pdf()
            if pdf_name:
                try:
                    # Create a new PDF file
                    pdf = Image.open(images[0])
                    pdf.save(pdf_name, "PDF", resolution=100.0,
                             save_all=True, append_images=images[1:])
                    messagebox.showinfo(
                        "Success", "Images have been successfully converted to PDF.")
                except Exception as e:
                    messagebox.showerror(
                        "Error", f"Failed to convert to PDF.\nError: {str(e)}")

    def convert_doc_to_pdf(self):
        word_document = self.select_doc()

        try:
            convert(word_document)
            word_document = self.select_doc_pdf()
            messagebox.showinfo("Success", "Word Documents have been successfully converted to PDF.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert to PDF.\nError: {str(e)}")



     

    def select_doc(self):
        doc_pdf = filedialog.asksaveasfilename(
            title="Save PDF as",
            defaultextension=".pdf",
            initialdir="C:/",
            filetypes=(("Doc files", "*.docx"), ("All files", "*.*"))
        )
        return doc_pdf
    

    def select_images(self):
        images = filedialog.askopenfilenames(title="Select Images", filetypes=(
            ("Image files", "*.jpg;*.jpeg;*.png"), ("All files", "*.*")), initialdir="C:/")
        return images

    def select_image_pdf(self):
        image_pdf = filedialog.asksaveasfilename(title="Save PDF as", defaultextension=".pdf",
                                           initialdir="C:/", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
        return image_pdf
    
    def select_doc_pdf(self):
        doc_pdf = filedialog.asksaveasfilename(title="Save PDF as", defaultextension=".pdf",
            initialdir="C:/", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
        return doc_pdf

    def run(self):
        self.canvas.pack()
        self.convert_DOC_btn.pack()
        self.convert_JPG_btn.pack()
        self.master.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverter(root)
    app.run()

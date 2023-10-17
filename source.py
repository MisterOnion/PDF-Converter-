import tkinter as tk
from tkinter import filedialog, messagebox  # file processing lib
from PIL import Image  # image processing lib
from docx import Document  # doc to pdf process
from docx.shared import Inches  # image in doc to pdf process
import io  # manages input/ouput low level binary data
from io import BytesIO  # binary conversion
from reportlab.pdfgen import canvas  # more GUI widgets
from reportlab.lib.pagesizes import letter  # more GUI widgets



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
            pdf_name = self.select_pdf()
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

        if word_document:
            try:
                doc = Document(word_document.name)
                doc_name = self.select_pdf()

                if doc_name:
                    c = canvas.Canvas(doc_name, pagesize=letter)

                    for paragraph in doc.paragraphs:
                        for run in paragraph.runs:
                            for inline_shape in run._r.inline_shapes:
                                image_stream = BytesIO(inline_shape._inline._blip)
                                image = Image.open(image_stream)

                                # Generate a unique image name using the element ID
                                image_path = f"image_{inline_shape._inline.attrib['id']}.jpg"
                                image.save(image_path)

                                # Insert the image into the PDF
                                c.drawInlineImage(image_path, 0, 0)

                    c.save()
                    messagebox.showinfo("Success", "Word Documents have been successfully converted to PDF.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to convert to PDF.\nError: {str(e)}")

    def select_doc(self):
        doc = filedialog.askopenfile(title="Doc File", filetypes=(
            ("Word Documents", "*.docx"), ("All files", "*.*")), initialdir="C:/")
        return doc

    def select_images(self):
        images = filedialog.askopenfilenames(title="Select Images", filetypes=(
            ("Image files", "*.jpg;*.jpeg;*.png"), ("All files", "*.*")), initialdir="C:/")
        return images

    def select_pdf(self):
        pdf = filedialog.asksaveasfilename(title="Save PDF as", defaultextension=".pdf",
                                           initialdir="C:/", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
        return pdf

    def run(self):
        self.canvas.pack()
        self.convert_DOC_btn.pack()
        self.convert_JPG_btn.pack()
        self.master.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverter(root)
    app.run()

import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image  # image processing lib


class ImageToPDFConverter:
    def __init__(self, master):
        self.master = master
        master.title("JPG to PDF converter")
        master.iconphoto(False, tk.PhotoImage(file='icon.png'))

        self.canvas = tk.Canvas(master, width=300, height=300)
        self.convert_btn = tk.Button(
            master, text="Convert", command=self.convert_images_to_pdf)

    def convert_images_to_pdf(self):
        images = self.select_images()
        if images:
            pdf_name = self.select_pdf()
            if pdf_name:
                try:
                    # Createa new PDF file
                    pdf = Image.open(images[0])
                    pdf.save(pdf_name, "PDF", resolution=100.0,
                             save_all=True, append_images=images[1:])
                    messagebox.showinfo(
                        "Success", "Images have been successfully converted to PDF.")
                except Exception as e:
                    messagebox.showerror(
                        "Error", f"Failed to convert to PDF.\nError: {str(e)}")

    def select_images(self):
        images = filedialog.askopenfilenames(title="Select Images", filetypes=(
            ("Image files", "*.jpg;*.jpeg;*.png"), ("All files", "*.*")), initialdir="C:/")
        return images

    def select_pdf(self):
        pdf = filedialog.asksaveasfilename(title="Save PDF as", defaultextension=".pdf",
                                           initialr="C:/", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
        return pdf

    def run(self):
        self.canvas.pack()
        self.convert_btn.pack()
        self.master.mainloop()



if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToPDFConverter(root)
    app.run()

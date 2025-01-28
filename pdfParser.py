import tkinter as tk
from tkinter import filedialog, messagebox, Text
from spire.pdf.common import *
from spire.pdf import *
from pypdf import PdfReader

def process_pdf_extraction():
    try:
        pdf_file = pdf_file_path.get()
        if not pdf_file:
            messagebox.showerror("Error", "No PDF file selected. Please upload a file.")
            return

        pdf_document = PdfDocument()
        pdf_document.LoadFromFile(pdf_file)
        pdf_reader = PdfReader(pdf_file)

        selected_page = int(page_input.get()) - 1  # Adjust to zero-indexed

        extraction_type = extraction_choice.get()

        if extraction_type == "Image":

            page = pdf_document.Pages.get_Item(selected_page)

            image_helper = PdfImageHelper()
            images_info = image_helper.GetImagesInfo(page)

            if not images_info:
                messagebox.showinfo("Info", f"No images found on page {selected_page + 1}.")
            else:
                for idx, image_info in enumerate(images_info):
                    image = image_info.Image
                    output_image = f"Extracted_images/Page-{selected_page + 1}-Image-{idx + 1}.png"
                    image.Save(output_image)

                messagebox.showinfo("Success", f"Extracted {len(images_info)} images from page {selected_page + 1}.")

        elif extraction_type == "Text":
            page = pdf_reader.pages[selected_page]
            extracted_text = page.extract_text()

            if extracted_text:

                text_window = tk.Toplevel(main_window)
                text_window.title("Extracted Text")
                text_display = Text(text_window, wrap=tk.WORD, width=80, height=20)
                text_display.insert(tk.END, extracted_text)
                text_display.pack(fill=tk.BOTH, expand=True)
            else:
                messagebox.showinfo("Info", f"No text found on page {selected_page + 1}.")

        else:
            messagebox.showerror("Error", "Invalid selection. Choose either 'Text' or 'Image'.")


        pdf_document.Dispose()

    except Exception as error:
        messagebox.showerror("Error", f"An error occurred: {error}")

def select_pdf_file():
    selected_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if selected_file:
        pdf_file_path.set(selected_file)

main_window = tk.Tk()
main_window.title("PDF Content Extractor")

tk.Label(main_window, text="Select a PDF file:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
pdf_file_path = tk.StringVar()
file_display = tk.Entry(main_window, textvariable=pdf_file_path, width=50, state="readonly")
file_display.grid(row=0, column=1, padx=5, pady=5)
tk.Button(main_window, text="Browse", command=select_pdf_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(main_window, text="Page number:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
page_input = tk.StringVar()
tk.Entry(main_window, textvariable=page_input, width=10).grid(row=1, column=1, padx=5, pady=5, sticky="w")


tk.Label(main_window, text="Extraction type:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
extraction_choice = tk.StringVar(value="Text")
tk.Radiobutton(main_window, text="Text", variable=extraction_choice, value="Text").grid(row=2, column=1, padx=5, pady=5, sticky="w")
tk.Radiobutton(main_window, text="Image", variable=extraction_choice, value="Image").grid(row=2, column=1, padx=100, pady=5, sticky="w")

tk.Button(main_window, text="Extract", command=process_pdf_extraction).grid(row=3, column=0, columnspan=3, pady=10)

main_window.mainloop()
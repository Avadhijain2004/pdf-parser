import os
from spire.pdf.common import *
from spire.pdf import *
from pypdf import PdfReader
import pdfplumber
from tabulate import tabulate

def extractImages(pdfPath, pageNo):
    pdfDoc = PdfDocument()
    pdfDoc.LoadFromFile(pdfPath)
    page = pdfDoc.Pages.get_Item(pageNo)
    
    imageHelper = PdfImageHelper()
    imagesInfo = imageHelper.GetImagesInfo(page)
    
    if not imagesInfo:
        print(f"There is no image on {pageNo + 1}.")
    else:
        os.makedirs("Extracted_images", exist_ok=True)
        for index, imgInfo in enumerate(imagesInfo):
            image = imgInfo.Image
            outputImg = f"Extracted_images/Page-{pageNo + 1}-Image-{index + 1}.png"
            image.Save(outputImg)
        print(f"Extracted {len(imagesInfo)} images from page {pageNo + 1}.")
    
    pdfDoc.Dispose()

def extractText(pdfPath, pageNo):
    pdfReader = PdfReader(pdfPath)
    page = pdfReader.pages[pageNo]
    extractedTxt = page.extract_text()
    
    if extractedTxt:
        print("Extracted Text:")
        print(extractedTxt)
    else:
        print(f"There is no text on {pageNo + 1}.")

def extractTable(pdfPath, pageNo):
    with pdfplumber.open(pdfPath) as pdf:
        page = pdf.pages[pageNo]
        tables = page.extract_tables()
        
        if not tables:
            print(f"There is no table on {pageNo + 1}.")
        else:
            print("Extracted Table:")
            for table in tables:
                print(tabulate(table, tablefmt="grid"))

pdfFile = "sample.pdf"  
    
try:
    pageNo = int(input("Enter a Page Number: ")) - 1  
    if pageNo < 0:
        print("Invalid Page Number.")
        exit()
        
    extType = input("Enter extraction type (Text/Image/Table): ").strip().lower()
        
    if extType == "text":
        extractText(pdfFile, pageNo)
    elif extType == "image":
        extractImages(pdfFile, pageNo)
    elif extType == "table":
        extractTable(pdfFile, pageNo)
    else:
        print("Invalid selection. Choose from 'Text', 'Image', or 'Table'.")
except Exception as error:
    print(f"An error occurred: {error}")
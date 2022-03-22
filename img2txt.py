import os
from PIL import Image
from pytesseract import pytesseract
from colorama import Fore
from docx import Document

path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
imagesDir = r"images/"
outDir = r'out/'
document = Document()

if not os.path.exists(outDir):
    os.mkdir(outDir)

pytesseract.tesseract_cmd = path_to_tesseract
for index, filename in enumerate(os.listdir(imagesDir)):
    try:
        img = Image.open(imagesDir + filename)
        text = pytesseract.image_to_string(img)
        with open(outDir+filename.split('.')[0]+'.txt', 'w') as f:
            f.write(text)
        print(f"[{index}]\t" + Fore.LIGHTGREEN_EX +
              filename + ": SUCCESS" + Fore.RESET)
    except:
        print(f"[{index}]\t" + Fore.LIGHTRED_EX +
              filename + ": ERROR" + Fore.RESET)

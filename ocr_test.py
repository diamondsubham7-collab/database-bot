import pytesseract
from PIL import Image
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Folder mein jo photo hai usko directly read karo
folder = "received_files"
for file in os.listdir(folder):
    if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
        image_path = os.path.join(folder, file)
        print(f"Reading: {image_path}")
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        print("===== EXTRACTED TEXT =====")
        print(text)
        print("==========================")
        if text.strip():
            print("✅ OCR SUCCESSFUL!")
        else:
            print("❌ No text detected")
        break
else:
    print("❌ No image found in received_files folder")
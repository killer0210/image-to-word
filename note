import sys
import io
import os
# Windows терминал дээр UTF-8 кодчлолыг тохируулж байна
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Шаардлагатай сангуудыг import хийж байна
import pytesseract  # Зургаас текст таних
from PIL import Image  # Зураг боловсруулах
from docx import Document  # Word файл үүсгэх

# Tesseract OCR програмын байршлыг зааж өгч байна
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Tesseract-д зориулсан хэлний файлууд байрлах замыг зааж өгч байна
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# Зургийг нээх
image_path = '44.jpg'  # Зургийн файлын нэр
img = Image.open(image_path)

# Алдааг барих
try:
    # Зургийн текстийг OCR ашиглан таних
    text = pytesseract.image_to_string(img, lang='mon')
except pytesseract.TesseractNotFoundError:
    print("Tesseract OCR олдсонгүй. Зам зөв эсэхийг шалгана уу")
    exit()
except pytesseract.TesseractError as e:
    print(f"Алдаа гарлаа: {str(e)}")
    print("Монгол хэлний файлууд суусан эсэхийг шалгана уу")
    exit()

# Word баримт үүсгэх
doc = Document()
doc.add_paragraph(text)  # OCR-р олдсон текстийг Word файлд нэмэх

# Word файлыг хадгалах
output_path = 'output.docx'
doc.save(output_path)

print(f"Текстийг {output_path} файлд амжилттай хадгаллаа.")

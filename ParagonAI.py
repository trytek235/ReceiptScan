import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import openpyxl

# Ścieżka do Tesseract OCR na twoim komputerze
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Funkcja do odczytywania tekstu z obrazu
def extract_text(image_path):
    # Otwórz obraz
    image = Image.open(image_path)
    image = image.convert('L')
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)  # Zwiększ kontrast
    image = image.filter(ImageFilter.MedianFilter())
    image = image.point(lambda x: 0 if x < 128 else 255, '1')

    # Użyj Tesseract OCR do wyodrębnienia tekstu
    text = pytesseract.image_to_string(image, lang='pol')
    return text

# Funkcja do zapisywania danych do pliku Excel
def save_to_excel(data, excel_path):
    # Otwórz lub stwórz plik Excel
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # Dodaj nagłówki
        sheet.append(["Data", "Pozycja", "Cena"])

    # Przetwarzaj dane i wstawiaj do Excela (tu można rozbudować parsing tekstu)
    for line in data.split("\n"):
        if any(char in line for char in ["A", "B", "C"]):  # Przykład znalezienia pozycji z ceną
            parts = line.split()
            item = "C".join(parts[:-1])
            price = parts[-1]
            # Zakładając, że paragon zawiera datę na początku (tu trzeba dostosować do formatu)
            sheet.append([None, item, price])

    workbook.save(excel_path)
    print(f"Dane zapisane w {excel_path}")

# Główna funkcja
if __name__ == "__main__":
    image_path = r"C:\Users\GRZESIU\Documents\Scanned Documents\Paragony\Biedronka23092024.jpeg"  # Ścieżka do obrazu paragonu
    excel_path = "paragony.xlsx"  # Plik Excel do zapisu danych
    
    # Wyodrębnij tekst z obrazu
    extracted_text = extract_text(image_path)
    print("Znaleziony tekst:", extracted_text)
    
    # Zapisz wyodrębnione dane do pliku Excel
    save_to_excel(extracted_text, excel_path)

from PIL import Image
import pytesseract
import re
import os
from dotenv import load_dotenv

# Tesseract windows configuration
load_dotenv()
pytesseract.pytesseract.tesseract_cmd = os.path.join(os.getenv('TESSERACT_DIR_LOCATION'), 'tesseract.exe')


def get_text_from_captcha(filename):
    text: str = pytesseract.image_to_string(Image.open(filename), lang='rus')
    letters = re.findall(r'\w', text)
    capcha_text = ''.join(letters)
    return capcha_text[:5].lower()

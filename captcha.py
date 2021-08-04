"""This module is responsible for interaction with tesseract-ocr"""
import re
import os
import pytesseract

from PIL import Image
from dotenv import load_dotenv

# Tesseract windows configuration
load_dotenv()
pytesseract.pytesseract.tesseract_cmd = os.path.join(os.getenv('TESSERACT_DIR_LOCATION'), 'tesseract.exe')


def get_text_from_captcha(filename: str) -> str:
    """
    Using tesseract-ocr extracts text from captcha image.

    Note: Fssp uses a standard captcha with a length of 5 characters.
    """
    text: str = pytesseract.image_to_string(Image.open(filename), lang='rus')
    letters = re.findall(r'\w', text)
    captcha_text: str = ''.join(letters)

    return captcha_text[:5].lower()

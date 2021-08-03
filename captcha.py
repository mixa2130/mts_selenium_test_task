try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
import re


def get_text_from_captcha(filename):
    text: str = pytesseract.image_to_string(Image.open(filename), lang='rus')
    letters = re.findall(r'\w', text)
    capcha_text = ''.join(letters)
    return capcha_text[:5].lower()


from typing import NamedTuple
import urllib.request
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from captcha import get_text_from_captcha
from typing import List


class InputArgs(NamedTuple):
    first_name: str
    last_name: str
    patronymic: str
    date: str


class SessionFssp:
    def __init__(self):
        self.browser = webdriver.Firefox()
        self.wait = WebDriverWait(self.browser, 6)

        self.browser.get('https://fssp.gov.ru/')
        info_close_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.tingle-modal__close')))
        info_close_button.click()

    def get_debts(self, args: InputArgs) -> List[tuple]:
        def get_text(search_el):
            return search_el.text

        start_page: bool = True
        try:
            # Advanced searching available only at the start page
            big_search_btn = self.browser.find_element_by_css_selector('a.btn.btn-light')
            big_search_btn.click()
        except NoSuchElementException:
            start_page = False

        args_dict = args._asdict()
        last_form = None
        for key in args_dict.keys():
            input_form_arg = self.browser.find_element_by_name(f"is[{key}]")
            input_form_arg.clear()
            input_form_arg.send_keys(args_dict.get(key))
            last_form = input_form_arg

        last_form.send_keys(keys.Keys.ENTER)
        if not start_page:
            last_form.send_keys(keys.Keys.ENTER)

        self._solve_captcha()

        res_table = self.browser.find_element_by_css_selector('div.results')
        try:
            res_body = res_table.find_element_by_css_selector('div.results-frame').find_element_by_css_selector('tbody')
        except NoSuchElementException:
            return [(' '.join(args_dict.values()), 'Нет задолженностей')]

        # Results exists
        debtors = list()
        debt_info = res_body.find_elements_by_tag_name('td')
        counter: int = 0

        while counter < len(debt_info):
            if debt_info[counter].get_attribute('class') == 'first':
                tmp_lst: list = debt_info[counter:counter + 4] + debt_info[counter + 5:counter + 8]
                debtors.append(tuple(map(get_text, tmp_lst)))

                counter += 7
            counter += 1

        return debtors

    def _solve_captcha(self):
        time.sleep(3)
        while True:
            try:
                time.sleep(2)
                capcha = self.browser.find_element_by_id('capchaVisual')
                src = capcha.get_attribute('src')
            except NoSuchElementException:
                # We've logged in
                return 0

            try:
                captcha_btn = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.input-submit-capcha')))
            except TimeoutException:
                return 0
            img = urllib.request.urlretrieve(src)
            captcha_text = get_text_from_captcha(img[0])
            print(captcha_text)
            input_form = self.browser.find_element_by_id('captcha-popup-code')
            input_form.send_keys(captcha_text)

            captcha_btn.click()

    def __del__(self):
        self.browser.quit()


if __name__ == '__main__':
    smth = SessionFssp()
    args = InputArgs('Антон', 'Мартынов', 'Валерьевич', '')
    args2 = InputArgs('Михаил', 'Ступак', 'Викторович', '21.11.2000')
    print(smth.get_debts(args))
    print(smth.get_debts(args2))


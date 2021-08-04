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
from typing import List, Deque
from excel_interaction import write_excel_file, InputArgs, read_excel_file


class SessionFssp:
    def __init__(self):
        self.browser = webdriver.Firefox()
        self.wait = WebDriverWait(self.browser, 6)
        self.browser.implicitly_wait(10)

        self.browser.get('https://fssp.gov.ru/')
        info_close_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.tingle-modal__close')))
        info_close_button.click()

    def get_debts(self, args: InputArgs, debts: list) -> bool:
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

        if self._solve_captcha() == -1:
            return False

        try:
            res_table = self.browser.find_element_by_css_selector('div.results')
        except NoSuchElementException:
            return False

        try:
            res_body = res_table.find_element_by_css_selector('div.results-frame').find_element_by_css_selector('tbody')
        except NoSuchElementException:
            debts.append((' '.join(args_dict.values()), 'Нет задолженностей'))
            return True

        # Results exists
        debt_info = res_body.find_elements_by_tag_name('td')
        counter: int = 0

        while counter < len(debt_info):
            if debt_info[counter].get_attribute('class') == 'first':
                tmp_lst: list = debt_info[counter:counter + 4] + debt_info[counter + 5:counter + 8]
                debts.append(tuple(map(get_text, tmp_lst)))

                counter += 7
            counter += 1

        return True

    def _solve_captcha(self) -> int:
        """0 - captcha solved/-1 - restart"""
        while True:
            try:
                captcha = self.browser.find_element_by_id('capchaVisual')
                src = captcha.get_attribute('src')
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
            # if captcha_text == '':
            #     # обновить страницу
            #     self.browser.refresh()
            #     return -1

            input_form = self.browser.find_element_by_id('captcha-popup-code')
            input_form.send_keys(captcha_text)

            captcha_btn.click()

    def __del__(self):
        self.browser.quit()


if __name__ == '__main__':
    potential_debtors: Deque = read_excel_file()
    session = SessionFssp()
    debts = []
    while len(potential_debtors) > 0:
        el = potential_debtors.popleft()
        debt_found: bool = session.get_debts(args=el, debts=debts)
        if not debt_found:
            print('failed')
            potential_debtors.append(el)

    write_excel_file(debts)

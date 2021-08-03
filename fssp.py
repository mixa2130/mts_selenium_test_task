from typing import NamedTuple
import urllib.request
import requests
import time
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from captcha import get_text_from_captcha

browser = webdriver.Firefox()
wait = WebDriverWait(browser, 12)


class InputArgs(NamedTuple):
    first_name: str
    last_name: str
    patronymic: str
    date: str


class SessionFssp:
    def __init__(self):
        browser.implicitly_wait(7)
        browser.get('https://fssp.gov.ru/')
        info_close_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.tingle-modal__close')))
        info_close_button.click()

    def get_debts(self, args: InputArgs):
        try:
            # Advanced searching available only at the start page
            big_search_btn = browser.find_element_by_css_selector('a.btn.btn-light')
            big_search_btn.click()
        except NoSuchElementException:
            pass

        args_dict = args._asdict()
        last_form = None
        for key in args_dict.keys():
            input_form_arg = browser.find_element_by_name(f"is[{key}]")
            input_form_arg.send_keys(args_dict.get(key))
            last_form = input_form_arg

        last_form.send_keys(keys.Keys.ENTER)

        self._solve_captcha()

        res_table = browser.find_element_by_css_selector('div.results')
        try:
            res_body = res_table.find_element_by_css_selector('div.results-frame').find_element_by_css_selector('tbody')
        except NoSuchElementException:
            return {'name': args_dict, 'message': 'Нет задолженностей'}

        # Results exists
        # names = res_body.find_elements_by_css_selector('td.first')
        debtors = []
        # {name, Enforcement Requisites date subject department bailiff}

        debt_info = res_body.find_elements_by_tag_name('td')
        counter: int = 0
        # while counter < len(debt_info):
        #     if debt_info[counter].get_attribute('first') is not None:
        #         # Debtor starts
        #         debtor_tmp: list = []
        #         for ind in range(counter, counter+8):
        #             debtor_tmp.append(debt_info[ind].text)
        #         debtor_tmp.pop(4)
        #         debtors.append(tuple(debtor_tmp))
        #         counter += 7
        #     counter += 1
        # print(debtors)
        for el in debt_info:
            if el.get_attribute('colspan') is None:
                # Not republic header
                print(f"{counter}: {el.text.strip()}")
                counter += 1
                if counter % 8 == 0:
                    counter = 0

    def _solve_captcha(self):
        time.sleep(3)
        btn_wait = WebDriverWait(browser, 6)
        while True:
            try:
                time.sleep(2)
                capcha = browser.find_element_by_id('capchaVisual')
                src = capcha.get_attribute('src')
            except NoSuchElementException:
                # We've logged in
                return 0

            try:
                captcha_btn = btn_wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.input-submit-capcha')))
            except TimeoutException:
                return 0
            img = urllib.request.urlretrieve(src)
            captcha_text = get_text_from_captcha(img[0])
            print(captcha_text)
            input_form = browser.find_element_by_id('captcha-popup-code')
            input_form.send_keys(captcha_text)

            captcha_btn.click()

    def __del__(self):
        browser.quit()


if __name__ == '__main__':
    smth = SessionFssp()
    args = InputArgs('Мартынов', 'Антон', 'Валерьевич', '')
    args2 = InputArgs('Мартынов', 'Антон', '', '')
    smth.get_debts(args)

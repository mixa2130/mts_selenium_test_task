"""This module is responsible for processing sudrf(https://sudrf.ru/index.php?id=300#sp) website"""

import time
from typing import Deque

from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from excel_interaction import FILES, read_excel_file, InputArgs, write_excel_file


default_court_value: str = "Второй кассационный суд общей юрисдикции"


class SessionSudrf:
    """Main class for interacting  with https://sudrf.ru/index.php?id=300#sp website"""

    def __init__(self):
        self.browser = webdriver.Firefox()
        self.wait = WebDriverWait(self.browser, 6)

        self.browser.implicitly_wait(6)
        self.browser.get('https://sudrf.ru/index.php?id=300#sp')

    def get_lawsuits(self, args: InputArgs, lawsuits: list):
        """
        Gets all lawsuits from sudrf website for the specified arguments in args param.

        How it works:
        Adds to the lawsuits - tuple of debt values.
        If there are no debts - (name, "Нет задолженностей").


        :param args: search arguments
        :param lawsuits: the list, to which the result will be sent

        :return: 0 - Ok
        """
        def get_text(lawsuit):
            return lawsuit.text

        fio: str = ' '.join((args.last_name, args.first_name, args.patronymic))

        court_subj = Select(self.browser.find_element_by_xpath('(//select[@id="court_subj"])[2]'))
        court_subj.select_by_visible_text('Город Москва')

        suds_subj = Select(self.browser.find_element_by_xpath('(//select[@id="suds_subj"])[1]'))
        suds_subj.select_by_visible_text(default_court_value)

        f_name = self.browser.find_element_by_id('f_name')
        f_name.clear()
        f_name.send_keys(fio)

        form_btn = self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[4]/div[6]/form/table/tbody/tr[8]/td[2]/input[1]')))
        form_btn.click()
        time.sleep(5)

        try:
            result_table = self.browser.find_element_by_xpath(
                '(//div[@id="resultTable"])').find_element_by_css_selector('tbody')
        except NoSuchElementException:
            lawsuits.append((fio, 'Нет дел'))
            return 0

        lawsuits_info = result_table.find_elements_by_tag_name('td')
        counter: int = 0

        while counter < len(lawsuits_info):
            tmp_lst: list = lawsuits_info[counter:counter+9]
            lawsuits.append(tuple(map(get_text, tmp_lst)))
            counter += 9

        return 0

    def __del__(self):
        self.browser.quit()


if __name__ == '__main__':
    data: Deque = read_excel_file(FILES[1])
    session = SessionSudrf()
    suits = []

    while len(data) > 0:
        el = data.popleft()
        session.get_lawsuits(el, suits)

    write_excel_file(suits, FILES[1])

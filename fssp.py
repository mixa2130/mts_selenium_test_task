"""
This module is responsible for processing fssp(https://fssp.gov.ru/) website.

Important!
Long waits are associated with unstable website operation.
These time intervals were clearly verified experimentally.
"""

import urllib.request
from typing import Deque
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from captcha import get_text_from_captcha
from excel_interaction import write_excel_file, InputArgs, read_excel_file, FILES


class SessionFssp:
    """Main class for interacting  with https://fssp.gov.ru website"""

    def __init__(self):
        self.browser = webdriver.Firefox()
        self.wait = WebDriverWait(self.browser, 6)

        self.browser.implicitly_wait(10)
        self._restart_sesion()

    def _restart_sesion(self):
        self.browser.get('https://fssp.gov.ru/')

        info_close_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.tingle-modal__close')))
        info_close_button.click()

        # Физическое лицо
        # tmp_btn = self.browser.find_element_by_css_selector('label.field__toggle-wrap')
        # tmp_btn.click()

    def _solve_captcha(self):
        """Solves Fssp page captcha using tesseract-ocr"""
        prev_src: str = ''
        while True:
            # Проверка наличия капчи на странице.
            try:
                captcha = self.browser.find_element_by_id('capchaVisual')
                src = captcha.get_attribute('src')
            except NoSuchElementException:
                # Капчи нет - мы уже авторизованы.
                break

            if src[-1:-10:-1] == prev_src:
                # Баг - зациклились
                return -1

            img: tuple = urllib.request.urlretrieve(src)  # (tmp_filename, )
            captcha_text: str = get_text_from_captcha(img[0])

            input_form = self.browser.find_element_by_id('captcha-popup-code')
            input_form.send_keys(captcha_text)

            # Так как к концу каждой итерации мы ждём заверщения процессинга нашего ответа на капчу,
            # а в начале мы проверяем наличие капчи - к данному этапу кнопка должна существовать,
            # и быть доступна.
            captcha_btn = self.browser.find_element_by_css_selector('input.input-submit-capcha')
            captcha_btn.click()

            prev_src = src[-1:-10:-1]
            # Сайт слабенький, поэтому иногда приходится ждать до 8 секунд ответа проверки капчи.
            # Если наш ответ неверен - появится новое src и новая кнопка.
            # Иначе - мы получим результат запроса(случай истечения времени ожидания).
            try:
                self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.input-submit-capcha')))
            except TimeoutException:
                break

        return 0

    def _pagination(self, pages_cnt: int) -> list:
        """
        Retrieves debts from pages.

        :return: [] if something bad has happened/
            else will return List[tuple] with debts values
        """

        def get_text(search_el):
            """Retrieves text from chosen html block"""
            return search_el.text

        tmp_debts = list()
        for page_ind in range(1, pages_cnt + 1):
            try:
                res_body = self.browser.find_element_by_css_selector('div.results').find_element_by_css_selector(
                    'div.results-frame').find_element_by_css_selector('tbody')
            except NoSuchElementException:
                # Долгий процессинг каптч, не уложились во времени - стоит начать заново со стартовой страницы
                return []

            debt_info: list = res_body.find_elements_by_tag_name('td')
            # debt_info содержит:
            # ячейки с названием республики, к которой относится данный долг
            # ячейки с информацией о задолженностях
            counter: int = 0

            while counter < len(debt_info):
                if debt_info[counter].get_attribute('class') == 'first':
                    # class first хранит информацию о имени должника.
                    # => Следующие 7 ячеек информация о задолженности.
                    # Одна из них - "сервис", возможность немедленной оплаты, её игнорируем
                    tmp_lst: list = debt_info[counter:counter + 4] + debt_info[counter + 5:counter + 8]
                    tmp_debts.append(tuple(map(get_text, tmp_lst)))

                    counter += 7
                counter += 1

            if page_ind != pages_cnt:
                # Ещё есть страницы c долгами
                pages = self.browser.find_element_by_css_selector('div.pagination').find_elements_by_tag_name('a')
                next_page = pages[len(pages) - 1]
                next_page.click()

                if self._solve_captcha() == -1:
                    return []

        return tmp_debts

    def get_debts(self, args: InputArgs, all_debts: list) -> bool:
        """
        Gets all debts from fssp website for the specified arguments in args param.

        How it works:
        Adds to the all_debts - tuple of debt values. If there are no debts -
        (name, "Нет задолженностей").

        :param args: search arguments
        :param all_debts: the list, to which the result will be sent

        :return: True - debts found/
        False - failed, we didn't meet the deadline,
        perhaps the site is lagging(in this case site will be reloaded)
        """

        def get_pages_cnt(href: str) -> int:
            _tmp: list = href.split('&page=')
            return int(_tmp[1])

        start_page: bool = True

        try:
            # Кнопка расширенного поиска доступна только на стартовой странице
            big_search_btn = self.browser.find_element_by_css_selector('a.btn.btn-light')
            big_search_btn.click()
        except NoSuchElementException:
            start_page = False

        args_dict: dict = args._asdict()
        last_form = None
        for key in args_dict.keys():
            input_form_arg = self.browser.find_element_by_name(f"is[{key}]")
            input_form_arg.clear()
            input_form_arg.send_keys(args_dict.get(key))
            last_form = input_form_arg

        last_form.send_keys(keys.Keys.ENTER)
        if not start_page:
            # Форма выбора даты мешает сделать дальнейшие действия.
            # Первый enter скроет её. Повторный запустит поиск.
            last_form.send_keys(keys.Keys.ENTER)

        if self._solve_captcha() == -1:
            self._restart_sesion()
            return False

        try:
            res_table = self.browser.find_element_by_css_selector('div.results')
        except NoSuchElementException:
            # Долгий процессинг каптч, не уложились во времени - стоит начать заново со стартовой страницы
            self._restart_sesion()
            return False

        try:
            res_table.find_element_by_css_selector('div.results-frame').find_element_by_css_selector('tbody')
        except NoSuchElementException:
            all_debts.append((' '.join(args_dict.values()), 'Нет задолженностей'))
            return True

        # Задолженности точно есть. Возможно они не умещаются в одну страницу.
        try:
            pagination = self.browser.find_element_by_css_selector('div.pagination')
        except NoSuchElementException:
            pages_cnt: int = 1
        else:
            pages = pagination.find_elements_by_tag_name('a')
            pages_cnt: int = get_pages_cnt(pages[len(pages) - 2].get_attribute('href'))

        _debts_res: list = self._pagination(pages_cnt)
        if len(_debts_res) == 0:
            self._restart_sesion()
            return False

        all_debts.extend(_debts_res)
        return True

    def __del__(self):
        self.browser.quit()


if __name__ == '__main__':
    potential_debtors: Deque = read_excel_file(FILES[0])

    session = SessionFssp()
    debts = []

    while len(potential_debtors) > 0:
        el = potential_debtors.popleft()

        debt_found: bool = session.get_debts(args=el, all_debts=debts)
        if not debt_found:
            potential_debtors.append(el)

    write_excel_file(data=debts, file_desc=FILES[0])

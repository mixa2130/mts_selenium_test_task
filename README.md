# mts_selenium_test_task

## Installation

(Установка [virtual environment](https://pypi.python.org/pypi/virtualenv) всегда рекомендуема.)

```bash
pip3 install -r requirements.txt
```

Разгадывание капчи производится средствами [tesseract-ocr](https://github.com/tesseract-ocr/tesseract), поэтому для
корректной работы программы необходимо установить соответствущую
[версию ocr](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.0.0-alpha.20210506.exe).

Для корректной работы библиотеки Selenium потребуется
установить [geckodriver](https://github.com/mozilla/geckodriver/releases/download/v0.29.1/geckodriver-v0.29.1-win64.zip)
в рабочую директорию.

### Configuration

Проект конфигурируется с помощью файла переменных окружений `.env` в рабочей директории.

Пример `.env`:

```env
# Абсолютный путь до директории, в которой расположен tesseract.exe
TESSERACT_DIR_LOCATION="C:\Program Files\Tesseract-OCR"
```

### Тестовые файлы
[Тестовая выборка для задания 1](https://disk.yandex.ru/i/e0TLwD33XS23bA) 
[Результат обработки тестовой выборки задания 1](https://disk.yandex.ru/i/O2Y3WVCmlXUXzw)

## Задание 1

Дано:

* Язык разработки Python 3.6+
* Модуль Pywin32 (работа с Excel)
* Модуль Selenium (работа с Web страницей)

Получение данных должно осуществляться через selenium, через сайт. НЕ через API

Нужно:
Получить на вход таблицу Excel (колонки Фамилия, Имя, Отчество, Дата рождения) и построчно получить информацию по
исполнительным производствам из сайта ФССП (http://fssprus.ru/). 
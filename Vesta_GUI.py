"""
Графический интерфейс
"""
from diff_tables import find_diffrence  # Функции для нахождения разницы 2 таблиц
from decl_case import declension_fio_by_case  # Функция для склонения ФИО по падежам
from comparsion_two_tables import merging_two_tables  # Функция для сравнения, слияния 2 таблиц
from table_stat import counting_by_category  # Функция для подсчета категориальных переменныъ
from table_stat import counting_quantitative_stat  # функция для подсчета количественных статистик
from processing_date import proccessing_date  # Функция для объединения двух таблиц
from union_tables import union_tables  # Функция для объедения множества таблиц
from extract_data_from_xlsx import \
    extract_data_from_hard_xlsx  # Функция для извлечения данных из таблиц со сложной структурой
from generate_docs import generate_docs_from_template  # Функция для создания документов Word из шаблона
from split_table import split_table  # Функция для разделения таблицы по отдельным листам и файлам
from preparation_list import prepare_list  # Функция для очистки и обработки списка
from create_svod import generate_svod_for_columns  # Функция для создания сводных таблиц по заданным колонкам
import pandas as pd
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import sys
import locale
import logging

logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def select_file_template_doc():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_doc
    name_file_template_doc = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_doc():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_doc
    # Получаем путь к файлу
    name_file_data_doc = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_doc():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_doc
    path_to_end_folder_doc = filedialog.askdirectory()


# Функции для вкладки извлечение данных
def select_file_params_calculate_data():
    """
    Функция для выбора файла c ячейками которые нужно подсчитать
    :return: Путь к файлу
    """
    global name_file_params_calculate_data
    name_file_params_calculate_data = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_files_data_calculate_data():
    """
    Функция для выбора файлов с данными параметры из которых нужно подсчитать
    :return: Путь к файлам с данными
    """
    global names_files_calculate_data
    # Получаем путь к файлу
    names_files_calculate_data = filedialog.askdirectory()


def select_end_folder_calculate_data():
    """
    Функция для выбора папки куда будут генерироваться файл  с результатом подсчета и файл с проверочной инфомрацией
    :return:
    """
    global path_to_end_folder_calculate_data
    path_to_end_folder_calculate_data = filedialog.askdirectory()


def calculate_data():
    """
    Функция для подсчета данных из файлов
    :return:
    """
    try:
        mode_text = mode_text_value.get()
        extract_data_from_hard_xlsx(mode_text, name_file_params_calculate_data, names_files_calculate_data,
                                    path_to_end_folder_calculate_data)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите параметры ,папку или файл с данными и папку куда будут генерироваться файлы')


# Функции для слияния таблиц

def select_end_folder_merger():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_merger
    path_to_end_folder_merger = filedialog.askdirectory()


def select_folder_data_merger():
    """
    Функция для выбора папки где хранятся нужные файлы
    :return:
    """
    global dir_name
    dir_name = filedialog.askdirectory()


def select_params_file_merger():
    """
    Функция для выбора файла c ячейками которые нужно подсчитать
    :return: Путь к файлу
    """
    if group_rb_type_harvest.get() == 2:
        global params_harvest
        params_harvest = filedialog.askopenfilename(
            filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
    else:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             'Выберите вариант слияния В и попробуйте снова ')


def select_standard_file_merger():
    """
    Функция для выбора файла c ячейками которые нужно подсчитать
    :return: Путь к файлу
    """
    global file_standard_merger
    file_standard_merger = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def merge_tables():
    """
    Функция для слияния таблиц с одинаковой структурой в одну большую таблицу
    """
    # Звпускаем функцию обработки
    try:
        checkbox_harvest = group_rb_type_harvest.get()  # получаем значение чекбокса
        merger_entry_skip_rows = merger_number_skip_rows.get()  # получаем сколько строк занимает заголовок
        # проверяем значение чекбокса на случай
        if checkbox_harvest != 2:
            file_params = None
        else:
            file_params = params_harvest

        union_tables(checkbox_harvest, merger_entry_skip_rows, file_standard_merger, dir_name,
                     path_to_end_folder_merger, file_params)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите папку с файлами,эталонный файл и папку куда будут генерироваться файлы')


def generate_docs_other():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных Веста Обработка таблиц и создание документов ver 1.35)
    :return:
    """
    try:
        name_column = entry_name_column_data.get()
        name_type_file = entry_type_file.get()
        name_value_column = entry_value_column.get()

        # получаем состояние чекбокса создания pdf
        mode_pdf = mode_pdf_value.get()
        # Получаем состояние  чекбокса объединения файлов в один
        mode_combine = mode_combine_value.get()
        # Получаем состояние чекбокса создания индвидуального файла
        mode_group = mode_group_doc.get()

        generate_docs_from_template(name_column, name_type_file, name_value_column, mode_pdf, name_file_template_doc,
                                    name_file_data_doc, path_to_end_folder_doc,
                                    mode_combine, mode_group)


    except NameError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')


def set_rus_locale():
    """
    Функция чтобы можно было извлечь русские названия месяцев
    """
    locale.setlocale(
        locale.LC_ALL,
        'rus_rus' if sys.platform == 'win32' else 'ru_RU.UTF-8')


def convert_date(cell):
    """
    Функция для конвертации даты в формате 1957-05-10 в формат 10.05.1957(строковый)
    """

    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date

    except TypeError:
        print(cell)
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             'Проверьте правильность заполнения ячеек с датой!!!')
        logging.exception('AN ERROR HAS OCCURRED')
        quit()


"""
Функции для получения параметров обработки даты рождения
"""


def select_file_data_date():
    """
    Функция для выбора файла с данными для которого нужно разбить по категориям
    :return: Путь к файлу с данными
    """
    global name_file_data_date
    # Получаем путь к файлу
    name_file_data_date = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_date():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_date
    path_to_end_folder_date = filedialog.askdirectory()


def calculate_date():
    """
    Функция для разбиения по категориям, подсчета текущего возраста и выделения месяца,года
    :return:
    """
    try:
        raw_selected_date = entry_date.get()
        name_column = entry_name_column.get()
        # Устанавливаем русскую локаль
        set_rus_locale()
        proccessing_date(raw_selected_date, name_column, name_file_data_date, path_to_end_folder_date)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')


"""
Функции для подсчета статистик по таблице
"""


def select_file_data_groupby():
    """
    Функция для выбора файла с данными
    :return:
    """
    global name_file_data_groupby
    # Получаем путь к файлу
    name_file_data_groupby = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_groupby():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_groupby
    path_to_end_folder_groupby = filedialog.askdirectory()


def groupby_category():
    """
    Подсчет категорий по всем колонкам таблицы
    """
    try:
        counting_by_category(name_file_data_groupby, path_to_end_folder_groupby)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')


def groupby_stat():
    """
    Подсчет категорий по всем колонкам таблицы
    """
    try:
        counting_quantitative_stat(name_file_data_groupby, path_to_end_folder_groupby)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')


# Функциия для слияния 2 таблиц
def select_file_params_comparsion():
    """
    Функция для выбора файла с параметрами колонок т.е. кокие колонки нужно обрабатывать
    :return:
    """
    global file_params
    file_params = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_first_comparison():
    """
    Функция для выбора  первого файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_first_file_comparison
    # Получаем путь к файлу
    name_first_file_comparison = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_comparison():
    """
    Функция для выбора  второго файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_second_file_comparison
    # Получаем путь к файлу
    name_second_file_comparison = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_comparison():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_comparison
    path_to_end_folder_comparison = filedialog.askdirectory()


def processing_comparison():
    """
    Функция для сравнения,слияния 2 таблиц
    :return:
    """
    # получаем названия листов
    try:
        first_sheet = entry_first_sheet_name.get()
        second_sheet = entry_second_sheet_name.get()

        merging_two_tables(file_params, first_sheet, second_sheet, name_first_file_comparison,
                           name_second_file_comparison, path_to_end_folder_comparison)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для склонения ФИО по падежам
"""


def select_data_decl_case():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_decl_case
    # Получаем путь к файлу
    data_decl_case = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_decl_case():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_decl_case
    path_to_end_folder_decl_case = filedialog.askdirectory()


def process_decl_case():
    """
    Функция для проведения склонения ФИО по падежам
    :return:
    """
    try:
        fio_column = decl_case_entry_fio.get()
        declension_fio_by_case(fio_column, data_decl_case, path_to_end_folder_decl_case)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Нахождения разницы 2 таблиц
Функции  получения параметров для find_diffrenece 
"""


def select_first_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_first_diffrence
    # Получаем путь к файлу
    data_first_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_second_diffrence
    # Получаем путь к файлу
    data_second_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_diffrence():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_diffrence
    path_to_end_folder_diffrence = filedialog.askdirectory()


def processing_diffrence():
    """
    Функция для получения названий листов и путей к файлам которые нужно сравнить
    :return:
    """
    # названия листов в таблицах
    try:
        first_sheet = entry_first_sheet_name_diffrence.get()
        second_sheet = entry_second_sheet_name_diffrence.get()
        # находим разницу
        find_diffrence(first_sheet, second_sheet, data_first_diffrence, data_second_diffrence,
                       path_to_end_folder_diffrence)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для разделения таблицы
"""


def select_file_split():
    """
    Функция для выбора файла с таблицей которую нужно разделить
    :return: Путь к файлу с данными
    """
    global file_data_split
    # Получаем путь к файлу
    file_data_split = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_split():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_split
    path_to_end_folder_split = filedialog.askdirectory()


def processing_split_table():
    """
    Функция для получения разделения таблицы по значениям
    :return:
    """
    # названия листов в таблицах
    try:
        # name_sheet = str(entry_sheet_name_split.get()) # получаем имя листа
        number_column = entry_number_column_split.get()  # получаем порядковый номер колонки
        number_column = int(number_column)  # конвертируем в инт

        checkbox_split = group_rb_type_split.get()  # получаем значения переключиталея

        # находим разницу
        split_table(file_data_split, number_column, checkbox_split, path_to_end_folder_split)
    except ValueError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Введите целое числа начиная с 1 !!!')
        logging.exception('AN ERROR HAS OCCURRED')
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для вкладки подготовка файлов
"""


def select_prep_file():
    """
    Функция для выбора файла который нужно преобразовать
    :return:
    """
    global glob_prep_file
    # Получаем путь к файлу
    glob_prep_file = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_prep():
    """
    Функция для выбора папки куда будет сохранен преобразованный файл
    :return:
    """
    global glob_path_to_end_folder_prep
    glob_path_to_end_folder_prep = filedialog.askdirectory()


def processing_preparation_file():
    """
    Функция для генерации документов
    """
    try:
        # name_sheet = var_name_sheet_prep.get() # получаем название листа
        checkbox_dupl = mode_dupl_value.get()
        prepare_list(glob_prep_file, glob_path_to_end_folder_prep, checkbox_dupl)

    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для сводных таблиц
"""


def select_file_svod():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_svod
    # Получаем путь к файлу
    data_svod = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_svod():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_svod
    path_to_end_folder_svod = filedialog.askdirectory()


def processing_svod():
    """
    Функция для получения названий листов и путей к файлам которые нужно сравнить
    :return:
    """
    # названия листов в таблицах
    try:
        sheet_name = entry_sheet_name_svod.get()  # название листа
        svod_columns = entry_columns_svod.get()  # колонки свода
        target_columns = entry_target_svod.get()
        # находим разницу
        generate_svod_for_columns(data_svod, sheet_name, path_to_end_folder_svod, svod_columns, target_columns)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для создания контекстного меню(Копировать,вставить,вырезать)
"""


def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")


def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))


def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)


def on_scroll(*args):
    canvas.yview(*args)


def set_window_size(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Устанавливаем размер окна в 80% от ширины и высоты экрана
    if screen_width >= 3840:
        width = int(screen_width * 0.28)
    elif screen_width >= 2560:
        width = int(screen_width * 0.39)
    elif screen_width >= 1920:
        width = int(screen_width * 0.48)
    elif screen_width >= 1600:
        width = int(screen_width * 0.58)
    elif screen_width >= 1280:
        width = int(screen_width * 0.70)
    elif screen_width >= 1024:
        width = int(screen_width * 0.85)
    else:
        width = int(screen_width * 1)

    height = int(screen_height * 0.8)

    # Рассчитываем координаты для центрирования окна
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размер и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")

"""
Создание нового окна
"""


def open_license():
    # Создание нового окна
    new_window = Toplevel(window)

    # Настройка нового окна
    new_window.title("Лицензия")
    text = Text(new_window,width=90, height=50)
    text.insert(INSERT, """                    GNU GENERAL PUBLIC LICENSE
                       Version 3, 29 June 2007

 Copyright (C) 2007 Free Software Foundation, Inc. <http://fsf.org/>
 Everyone is permitted to copy and distribute verbatim copies
 of this license document, but changing it is not allowed.

                            Preamble

  The GNU General Public License is a free, copyleft license for
software and other kinds of works.

  The licenses for most software and other practical works are designed
to take away your freedom to share and change the works.  By contrast,
the GNU General Public License is intended to guarantee your freedom to
share and change all versions of a program--to make sure it remains free
software for all its users.  We, the Free Software Foundation, use the
GNU General Public License for most of our software; it applies also to
any other work released this way by its authors.  You can apply it to
your programs, too.

  When we speak of free software, we are referring to freedom, not
price.  Our General Public Licenses are designed to make sure that you
have the freedom to distribute copies of free software (and charge for
them if you wish), that you receive source code or can get it if you
want it, that you can change the software or use pieces of it in new
free programs, and that you know you can do these things.

  To protect your rights, we need to prevent others from denying you
these rights or asking you to surrender the rights.  Therefore, you have
certain responsibilities if you distribute copies of the software, or if
you modify it: responsibilities to respect the freedom of others.

  For example, if you distribute copies of such a program, whether
gratis or for a fee, you must pass on to the recipients the same
freedoms that you received.  You must make sure that they, too, receive
or can get the source code.  And you must show them these terms so they
know their rights.

  Developers that use the GNU GPL protect your rights with two steps:
(1) assert copyright on the software, and (2) offer you this License
giving you legal permission to copy, distribute and/or modify it.

  For the developers' and authors' protection, the GPL clearly explains
that there is no warranty for this free software.  For both users' and
authors' sake, the GPL requires that modified versions be marked as
changed, so that their problems will not be attributed erroneously to
authors of previous versions.

  Some devices are designed to deny users access to install or run
modified versions of the software inside them, although the manufacturer
can do so.  This is fundamentally incompatible with the aim of
protecting users' freedom to change the software.  The systematic
pattern of such abuse occurs in the area of products for individuals to
use, which is precisely where it is most unacceptable.  Therefore, we
have designed this version of the GPL to prohibit the practice for those
products.  If such problems arise substantially in other domains, we
stand ready to extend this provision to those domains in future versions
of the GPL, as needed to protect the freedom of users.

  Finally, every program is threatened constantly by software patents.
States should not allow patents to restrict development and use of
software on general-purpose computers, but in those that do, we wish to
avoid the special danger that patents applied to a free program could
make it effectively proprietary.  To prevent this, the GPL assures that
patents cannot be used to render the program non-free.

  The precise terms and conditions for copying, distribution and
modification follow.

                       TERMS AND CONDITIONS

  0. Definitions.

  "This License" refers to version 3 of the GNU General Public License.

  "Copyright" also means copyright-like laws that apply to other kinds of
works, such as semiconductor masks.

  "The Program" refers to any copyrightable work licensed under this
License.  Each licensee is addressed as "you".  "Licensees" and
"recipients" may be individuals or organizations.

  To "modify" a work means to copy from or adapt all or part of the work
in a fashion requiring copyright permission, other than the making of an
exact copy.  The resulting work is called a "modified version" of the
earlier work or a work "based on" the earlier work.

  A "covered work" means either the unmodified Program or a work based
on the Program.

  To "propagate" a work means to do anything with it that, without
permission, would make you directly or secondarily liable for
infringement under applicable copyright law, except executing it on a
computer or modifying a private copy.  Propagation includes copying,
distribution (with or without modification), making available to the
public, and in some countries other activities as well.

  To "convey" a work means any kind of propagation that enables other
parties to make or receive copies.  Mere interaction with a user through
a computer network, with no transfer of a copy, is not conveying.

  An interactive user interface displays "Appropriate Legal Notices"
to the extent that it includes a convenient and prominently visible
feature that (1) displays an appropriate copyright notice, and (2)
tells the user that there is no warranty for the work (except to the
extent that warranties are provided), that licensees may convey the
work under this License, and how to view a copy of this License.  If
the interface presents a list of user commands or options, such as a
menu, a prominent item in the list meets this criterion.

  1. Source Code.

  The "source code" for a work means the preferred form of the work
for making modifications to it.  "Object code" means any non-source
form of a work.

  A "Standard Interface" means an interface that either is an official
standard defined by a recognized standards body, or, in the case of
interfaces specified for a particular programming language, one that
is widely used among developers working in that language.

  The "System Libraries" of an executable work include anything, other
than the work as a whole, that (a) is included in the normal form of
packaging a Major Component, but which is not part of that Major
Component, and (b) serves only to enable use of the work with that
Major Component, or to implement a Standard Interface for which an
implementation is available to the public in source code form.  A
"Major Component", in this context, means a major essential component
(kernel, window system, and so on) of the specific operating system
(if any) on which the executable work runs, or a compiler used to
produce the work, or an object code interpreter used to run it.

  The "Corresponding Source" for a work in object code form means all
the source code needed to generate, install, and (for an executable
work) run the object code and to modify the work, including scripts to
control those activities.  However, it does not include the work's
System Libraries, or general-purpose tools or generally available free
programs which are used unmodified in performing those activities but
which are not part of the work.  For example, Corresponding Source
includes interface definition files associated with source files for
the work, and the source code for shared libraries and dynamically
linked subprograms that the work is specifically designed to require,
such as by intimate data communication or control flow between those
subprograms and other parts of the work.

  The Corresponding Source need not include anything that users
can regenerate automatically from other parts of the Corresponding
Source.

  The Corresponding Source for a work in source code form is that
same work.

  2. Basic Permissions.

  All rights granted under this License are granted for the term of
copyright on the Program, and are irrevocable provided the stated
conditions are met.  This License explicitly affirms your unlimited
permission to run the unmodified Program.  The output from running a
covered work is covered by this License only if the output, given its
content, constitutes a covered work.  This License acknowledges your
rights of fair use or other equivalent, as provided by copyright law.

  You may make, run and propagate covered works that you do not
convey, without conditions so long as your license otherwise remains
in force.  You may convey covered works to others for the sole purpose
of having them make modifications exclusively for you, or provide you
with facilities for running those works, provided that you comply with
the terms of this License in conveying all material for which you do
not control copyright.  Those thus making or running the covered works
for you must do so exclusively on your behalf, under your direction
and control, on terms that prohibit them from making any copies of
your copyrighted material outside their relationship with you.

  Conveying under any other circumstances is permitted solely under
the conditions stated below.  Sublicensing is not allowed; section 10
makes it unnecessary.

  3. Protecting Users' Legal Rights From Anti-Circumvention Law.

  No covered work shall be deemed part of an effective technological
measure under any applicable law fulfilling obligations under article
11 of the WIPO copyright treaty adopted on 20 December 1996, or
similar laws prohibiting or restricting circumvention of such
measures.

  When you convey a covered work, you waive any legal power to forbid
circumvention of technological measures to the extent such circumvention
is effected by exercising rights under this License with respect to
the covered work, and you disclaim any intention to limit operation or
modification of the work as a means of enforcing, against the work's
users, your or third parties' legal rights to forbid circumvention of
technological measures.

  4. Conveying Verbatim Copies.

  You may convey verbatim copies of the Program's source code as you
receive it, in any medium, provided that you conspicuously and
appropriately publish on each copy an appropriate copyright notice;
keep intact all notices stating that this License and any
non-permissive terms added in accord with section 7 apply to the code;
keep intact all notices of the absence of any warranty; and give all
recipients a copy of this License along with the Program.

  You may charge any price or no price for each copy that you convey,
and you may offer support or warranty protection for a fee.

  5. Conveying Modified Source Versions.

  You may convey a work based on the Program, or the modifications to
produce it from the Program, in the form of source code under the
terms of section 4, provided that you also meet all of these conditions:

    a) The work must carry prominent notices stating that you modified
    it, and giving a relevant date.

    b) The work must carry prominent notices stating that it is
    released under this License and any conditions added under section
    7.  This requirement modifies the requirement in section 4 to
    "keep intact all notices".

    c) You must license the entire work, as a whole, under this
    License to anyone who comes into possession of a copy.  This
    License will therefore apply, along with any applicable section 7
    additional terms, to the whole of the work, and all its parts,
    regardless of how they are packaged.  This License gives no
    permission to license the work in any other way, but it does not
    invalidate such permission if you have separately received it.

    d) If the work has interactive user interfaces, each must display
    Appropriate Legal Notices; however, if the Program has interactive
    interfaces that do not display Appropriate Legal Notices, your
    work need not make them do so.

  A compilation of a covered work with other separate and independent
works, which are not by their nature extensions of the covered work,
and which are not combined with it such as to form a larger program,
in or on a volume of a storage or distribution medium, is called an
"aggregate" if the compilation and its resulting copyright are not
used to limit the access or legal rights of the compilation's users
beyond what the individual works permit.  Inclusion of a covered work
in an aggregate does not cause this License to apply to the other
parts of the aggregate.

  6. Conveying Non-Source Forms.

  You may convey a covered work in object code form under the terms
of sections 4 and 5, provided that you also convey the
machine-readable Corresponding Source under the terms of this License,
in one of these ways:

    a) Convey the object code in, or embodied in, a physical product
    (including a physical distribution medium), accompanied by the
    Corresponding Source fixed on a durable physical medium
    customarily used for software interchange.

    b) Convey the object code in, or embodied in, a physical product
    (including a physical distribution medium), accompanied by a
    written offer, valid for at least three years and valid for as
    long as you offer spare parts or customer support for that product
    model, to give anyone who possesses the object code either (1) a
    copy of the Corresponding Source for all the software in the
    product that is covered by this License, on a durable physical
    medium customarily used for software interchange, for a price no
    more than your reasonable cost of physically performing this
    conveying of source, or (2) access to copy the
    Corresponding Source from a network server at no charge.

    c) Convey individual copies of the object code with a copy of the
    written offer to provide the Corresponding Source.  This
    alternative is allowed only occasionally and noncommercially, and
    only if you received the object code with such an offer, in accord
    with subsection 6b.

    d) Convey the object code by offering access from a designated
    place (gratis or for a charge), and offer equivalent access to the
    Corresponding Source in the same way through the same place at no
    further charge.  You need not require recipients to copy the
    Corresponding Source along with the object code.  If the place to
    copy the object code is a network server, the Corresponding Source
    may be on a different server (operated by you or a third party)
    that supports equivalent copying facilities, provided you maintain
    clear directions next to the object code saying where to find the
    Corresponding Source.  Regardless of what server hosts the
    Corresponding Source, you remain obligated to ensure that it is
    available for as long as needed to satisfy these requirements.

    e) Convey the object code using peer-to-peer transmission, provided
    you inform other peers where the object code and Corresponding
    Source of the work are being offered to the general public at no
    charge under subsection 6d.

  A separable portion of the object code, whose source code is excluded
from the Corresponding Source as a System Library, need not be
included in conveying the object code work.

  A "User Product" is either (1) a "consumer product", which means any
tangible personal property which is normally used for personal, family,
or household purposes, or (2) anything designed or sold for incorporation
into a dwelling.  In determining whether a product is a consumer product,
doubtful cases shall be resolved in favor of coverage.  For a particular
product received by a particular user, "normally used" refers to a
typical or common use of that class of product, regardless of the status
of the particular user or of the way in which the particular user
actually uses, or expects or is expected to use, the product.  A product
is a consumer product regardless of whether the product has substantial
commercial, industrial or non-consumer uses, unless such uses represent
the only significant mode of use of the product.

  "Installation Information" for a User Product means any methods,
procedures, authorization keys, or other information required to install
and execute modified versions of a covered work in that User Product from
a modified version of its Corresponding Source.  The information must
suffice to ensure that the continued functioning of the modified object
code is in no case prevented or interfered with solely because
modification has been made.

  If you convey an object code work under this section in, or with, or
specifically for use in, a User Product, and the conveying occurs as
part of a transaction in which the right of possession and use of the
User Product is transferred to the recipient in perpetuity or for a
fixed term (regardless of how the transaction is characterized), the
Corresponding Source conveyed under this section must be accompanied
by the Installation Information.  But this requirement does not apply
if neither you nor any third party retains the ability to install
modified object code on the User Product (for example, the work has
been installed in ROM).

  The requirement to provide Installation Information does not include a
requirement to continue to provide support service, warranty, or updates
for a work that has been modified or installed by the recipient, or for
the User Product in which it has been modified or installed.  Access to a
network may be denied when the modification itself materially and
adversely affects the operation of the network or violates the rules and
protocols for communication across the network.

  Corresponding Source conveyed, and Installation Information provided,
in accord with this section must be in a format that is publicly
documented (and with an implementation available to the public in
source code form), and must require no special password or key for
unpacking, reading or copying.

  7. Additional Terms.

  "Additional permissions" are terms that supplement the terms of this
License by making exceptions from one or more of its conditions.
Additional permissions that are applicable to the entire Program shall
be treated as though they were included in this License, to the extent
that they are valid under applicable law.  If additional permissions
apply only to part of the Program, that part may be used separately
under those permissions, but the entire Program remains governed by
this License without regard to the additional permissions.

  When you convey a copy of a covered work, you may at your option
remove any additional permissions from that copy, or from any part of
it.  (Additional permissions may be written to require their own
removal in certain cases when you modify the work.)  You may place
additional permissions on material, added by you to a covered work,
for which you have or can give appropriate copyright permission.

  Notwithstanding any other provision of this License, for material you
add to a covered work, you may (if authorized by the copyright holders of
that material) supplement the terms of this License with terms:

    a) Disclaiming warranty or limiting liability differently from the
    terms of sections 15 and 16 of this License; or

    b) Requiring preservation of specified reasonable legal notices or
    author attributions in that material or in the Appropriate Legal
    Notices displayed by works containing it; or

    c) Prohibiting misrepresentation of the origin of that material, or
    requiring that modified versions of such material be marked in
    reasonable ways as different from the original version; or

    d) Limiting the use for publicity purposes of names of licensors or
    authors of the material; or

    e) Declining to grant rights under trademark law for use of some
    trade names, trademarks, or service marks; or

    f) Requiring indemnification of licensors and authors of that
    material by anyone who conveys the material (or modified versions of
    it) with contractual assumptions of liability to the recipient, for
    any liability that these contractual assumptions directly impose on
    those licensors and authors.

  All other non-permissive additional terms are considered "further
restrictions" within the meaning of section 10.  If the Program as you
received it, or any part of it, contains a notice stating that it is
governed by this License along with a term that is a further
restriction, you may remove that term.  If a license document contains
a further restriction but permits relicensing or conveying under this
License, you may add to a covered work material governed by the terms
of that license document, provided that the further restriction does
not survive such relicensing or conveying.

  If you add terms to a covered work in accord with this section, you
must place, in the relevant source files, a statement of the
additional terms that apply to those files, or a notice indicating
where to find the applicable terms.

  Additional terms, permissive or non-permissive, may be stated in the
form of a separately written license, or stated as exceptions;
the above requirements apply either way.

  8. Termination.

  You may not propagate or modify a covered work except as expressly
provided under this License.  Any attempt otherwise to propagate or
modify it is void, and will automatically terminate your rights under
this License (including any patent licenses granted under the third
paragraph of section 11).

  However, if you cease all violation of this License, then your
license from a particular copyright holder is reinstated (a)
provisionally, unless and until the copyright holder explicitly and
finally terminates your license, and (b) permanently, if the copyright
holder fails to notify you of the violation by some reasonable means
prior to 60 days after the cessation.

  Moreover, your license from a particular copyright holder is
reinstated permanently if the copyright holder notifies you of the
violation by some reasonable means, this is the first time you have
received notice of violation of this License (for any work) from that
copyright holder, and you cure the violation prior to 30 days after
your receipt of the notice.

  Termination of your rights under this section does not terminate the
licenses of parties who have received copies or rights from you under
this License.  If your rights have been terminated and not permanently
reinstated, you do not qualify to receive new licenses for the same
material under section 10.

  9. Acceptance Not Required for Having Copies.

  You are not required to accept this License in order to receive or
run a copy of the Program.  Ancillary propagation of a covered work
occurring solely as a consequence of using peer-to-peer transmission
to receive a copy likewise does not require acceptance.  However,
nothing other than this License grants you permission to propagate or
modify any covered work.  These actions infringe copyright if you do
not accept this License.  Therefore, by modifying or propagating a
covered work, you indicate your acceptance of this License to do so.

  10. Automatic Licensing of Downstream Recipients.

  Each time you convey a covered work, the recipient automatically
receives a license from the original licensors, to run, modify and
propagate that work, subject to this License.  You are not responsible
for enforcing compliance by third parties with this License.

  An "entity transaction" is a transaction transferring control of an
organization, or substantially all assets of one, or subdividing an
organization, or merging organizations.  If propagation of a covered
work results from an entity transaction, each party to that
transaction who receives a copy of the work also receives whatever
licenses to the work the party's predecessor in interest had or could
give under the previous paragraph, plus a right to possession of the
Corresponding Source of the work from the predecessor in interest, if
the predecessor has it or can get it with reasonable efforts.

  You may not impose any further restrictions on the exercise of the
rights granted or affirmed under this License.  For example, you may
not impose a license fee, royalty, or other charge for exercise of
rights granted under this License, and you may not initiate litigation
(including a cross-claim or counterclaim in a lawsuit) alleging that
any patent claim is infringed by making, using, selling, offering for
sale, or importing the Program or any portion of it.

  11. Patents.

  A "contributor" is a copyright holder who authorizes use under this
License of the Program or a work on which the Program is based.  The
work thus licensed is called the contributor's "contributor version".

  A contributor's "essential patent claims" are all patent claims
owned or controlled by the contributor, whether already acquired or
hereafter acquired, that would be infringed by some manner, permitted
by this License, of making, using, or selling its contributor version,
but do not include claims that would be infringed only as a
consequence of further modification of the contributor version.  For
purposes of this definition, "control" includes the right to grant
patent sublicenses in a manner consistent with the requirements of
this License.

  Each contributor grants you a non-exclusive, worldwide, royalty-free
patent license under the contributor's essential patent claims, to
make, use, sell, offer for sale, import and otherwise run, modify and
propagate the contents of its contributor version.

  In the following three paragraphs, a "patent license" is any express
agreement or commitment, however denominated, not to enforce a patent
(such as an express permission to practice a patent or covenant not to
sue for patent infringement).  To "grant" such a patent license to a
party means to make such an agreement or commitment not to enforce a
patent against the party.

  If you convey a covered work, knowingly relying on a patent license,
and the Corresponding Source of the work is not available for anyone
to copy, free of charge and under the terms of this License, through a
publicly available network server or other readily accessible means,
then you must either (1) cause the Corresponding Source to be so
available, or (2) arrange to deprive yourself of the benefit of the
patent license for this particular work, or (3) arrange, in a manner
consistent with the requirements of this License, to extend the patent
license to downstream recipients.  "Knowingly relying" means you have
actual knowledge that, but for the patent license, your conveying the
covered work in a country, or your recipient's use of the covered work
in a country, would infringe one or more identifiable patents in that
country that you have reason to believe are valid.

  If, pursuant to or in connection with a single transaction or
arrangement, you convey, or propagate by procuring conveyance of, a
covered work, and grant a patent license to some of the parties
receiving the covered work authorizing them to use, propagate, modify
or convey a specific copy of the covered work, then the patent license
you grant is automatically extended to all recipients of the covered
work and works based on it.

  A patent license is "discriminatory" if it does not include within
the scope of its coverage, prohibits the exercise of, or is
conditioned on the non-exercise of one or more of the rights that are
specifically granted under this License.  You may not convey a covered
work if you are a party to an arrangement with a third party that is
in the business of distributing software, under which you make payment
to the third party based on the extent of your activity of conveying
the work, and under which the third party grants, to any of the
parties who would receive the covered work from you, a discriminatory
patent license (a) in connection with copies of the covered work
conveyed by you (or copies made from those copies), or (b) primarily
for and in connection with specific products or compilations that
contain the covered work, unless you entered into that arrangement,
or that patent license was granted, prior to 28 March 2007.

  Nothing in this License shall be construed as excluding or limiting
any implied license or other defenses to infringement that may
otherwise be available to you under applicable patent law.

  12. No Surrender of Others' Freedom.

  If conditions are imposed on you (whether by court order, agreement or
otherwise) that contradict the conditions of this License, they do not
excuse you from the conditions of this License.  If you cannot convey a
covered work so as to satisfy simultaneously your obligations under this
License and any other pertinent obligations, then as a consequence you may
not convey it at all.  For example, if you agree to terms that obligate you
to collect a royalty for further conveying from those to whom you convey
the Program, the only way you could satisfy both those terms and this
License would be to refrain entirely from conveying the Program.

  13. Use with the GNU Affero General Public License.

  Notwithstanding any other provision of this License, you have
permission to link or combine any covered work with a work licensed
under version 3 of the GNU Affero General Public License into a single
combined work, and to convey the resulting work.  The terms of this
License will continue to apply to the part which is the covered work,
but the special requirements of the GNU Affero General Public License,
section 13, concerning interaction through a network will apply to the
combination as such.

  14. Revised Versions of this License.

  The Free Software Foundation may publish revised and/or new versions of
the GNU General Public License from time to time.  Such new versions will
be similar in spirit to the present version, but may differ in detail to
address new problems or concerns.

  Each version is given a distinguishing version number.  If the
Program specifies that a certain numbered version of the GNU General
Public License "or any later version" applies to it, you have the
option of following the terms and conditions either of that numbered
version or of any later version published by the Free Software
Foundation.  If the Program does not specify a version number of the
GNU General Public License, you may choose any version ever published
by the Free Software Foundation.

  If the Program specifies that a proxy can decide which future
versions of the GNU General Public License can be used, that proxy's
public statement of acceptance of a version permanently authorizes you
to choose that version for the Program.

  Later license versions may give you additional or different
permissions.  However, no additional obligations are imposed on any
author or copyright holder as a result of your choosing to follow a
later version.

  15. Disclaimer of Warranty.

  THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY
APPLICABLE LAW.  EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT
HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM "AS IS" WITHOUT WARRANTY
OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO,
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
PURPOSE.  THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM
IS WITH YOU.  SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF
ALL NECESSARY SERVICING, REPAIR OR CORRECTION.

  16. Limitation of Liability.

  IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING
WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MODIFIES AND/OR CONVEYS
THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES, INCLUDING ANY
GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE
USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF
DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY YOU OR THIRD
PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER PROGRAMS),
EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE POSSIBILITY OF
SUCH DAMAGES.

  17. Interpretation of Sections 15 and 16.

  If the disclaimer of warranty and limitation of liability provided
above cannot be given local legal effect according to their terms,
reviewing courts shall apply local law that most closely approximates
an absolute waiver of all civil liability in connection with the
Program, unless a warranty or assumption of liability accompanies a
copy of the Program in return for a fee.

                     END OF TERMS AND CONDITIONS
    """)
    text.configure(state='disabled')
    text.pack(side=LEFT)

    scroll = Scrollbar(new_window,command=text.yview)
    scroll.pack(side=LEFT, fill=Y)

    text.config(yscrollcommand=scroll.set)



if __name__ == '__main__':
    window = Tk()
    window.title('Веста Обработка таблиц и создание документов ver 1.50')
    # Устанавливаем размер и положение окна
    set_window_size(window)
    # window.geometry('774x760')
    # window.geometry('980x910+700+100')
    window.resizable(True, True)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)

    # Создаем вертикальный скроллбар
    scrollbar = Scrollbar(window, orient="vertical")

    # Создаем холст
    canvas = Canvas(window, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    # Привязываем скроллбар к холсту
    scrollbar.config(command=canvas.yview)

    # Создаем ноутбук (вкладки)
    tab_control = ttk.Notebook(canvas)

    """
    Создаем вкладку для предварительной обработки списка
    """
    tab_preparation = ttk.Frame(tab_control)
    tab_control.add(tab_preparation, text='Обработка\nсписка')

    preparation_frame_description = LabelFrame(tab_preparation)
    preparation_frame_description.pack()

    lbl_hello_preparation = Label(preparation_frame_description,
                                  text='Очистка от лишних пробелов и символов; поиск пропущенных значений\n в колонках с персональными данными,'
                                       '(ФИО,паспортные данные,\nтелефон,e-mail,дата рождения,ИНН)\n преобразование СНИЛС в формат ХХХ-ХХХ-ХХХ ХХ.\n'
                                       'Создание списка дубликатов по каждой колонке\n'
                                       'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                       'Для корректной работы программы уберите из таблицы\nобъединенные ячейки',
                                  width=60)
    lbl_hello_preparation.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_preparation = resource_path('logo.png')
    img_preparation = PhotoImage(file=path_to_img_preparation)
    Label(preparation_frame_description,
          image=img_preparation, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем кнопку вывода сообщения о программе
    btn_about = Button(preparation_frame_description, text='О программе', font=('Arial Bold', 14),
                                  command=select_prep_file)
    btn_about.pack(side=TOP, anchor=N,padx=10, pady=10)

    # Создаем кнопку вывода сообщения о программе
    btn_help = Button(preparation_frame_description, text='Руководство', font=('Arial Bold', 14),
                                  command=select_prep_file)
    btn_help.pack(side=TOP, anchor=N,padx=10, pady=10)




    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_prep = LabelFrame(tab_preparation, text='Подготовка')
    frame_data_prep.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_prep_file = Button(frame_data_prep, text='1) Выберите файл', font=('Arial Bold', 14),
                                  command=select_prep_file)
    btn_choose_prep_file.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_prep = Button(frame_data_prep, text='2) Выберите конечную папку', font=('Arial Bold', 14),
                                        command=select_end_folder_prep)
    btn_choose_end_folder_prep.pack(padx=10, pady=10)

    # Создаем переменную для хранения результа переключения чекбокса
    mode_dupl_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_dupl_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_dupl = Checkbutton(frame_data_prep,
                                  text='Проверить каждую колонку таблицы на дубликаты',
                                  variable=mode_dupl_value,
                                  offvalue='No',
                                  onvalue='Yes')
    chbox_mode_dupl.pack(padx=10, pady=10)

    # Создаем кнопку очистки
    btn_choose_processing_prep = Button(tab_preparation, text='3) Выполнить обработку', font=('Arial Bold', 20),
                                        command=processing_preparation_file)
    btn_choose_processing_prep.pack(padx=10, pady=10)

    """
    Создание вкладки для разбиения таблицы на несколько штук по значениям в определенной колонке
    """
    # Создаем вкладку для подсчета данных по категориям
    tab_split_tables = ttk.Frame(tab_control)
    tab_control.add(tab_split_tables, text='Разделение\n таблицы')

    split_tables_frame_description = LabelFrame(tab_split_tables)
    split_tables_frame_description.pack()

    lbl_hello_split_tables = Label(split_tables_frame_description,
                                   text='Разделение таблицы Excel по листам и файлам\n'
                                        'Для корректной работы программы уберите из таблицы\nобъединенные ячейки\n'
                                        'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                        'Заголовок таблицы должен занимать ОДНУ СТРОКУ\n и в нем не должно быть объединенных ячеек!',
                                   width=60)
    lbl_hello_split_tables.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_split_tables = resource_path('logo.png')
    img_split_tables = PhotoImage(file=path_to_img_split_tables)
    Label(split_tables_frame_description,
          image=img_split_tables, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_split = LabelFrame(tab_split_tables, text='Подготовка')
    frame_data_for_split.pack(padx=10, pady=10)
    # Переключатель:вариант слияния файлов
    # Создаем переключатель
    group_rb_type_split = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_split = LabelFrame(frame_data_for_split, text='1) Выберите вариант разделения')
    frame_rb_type_split.pack(padx=10, pady=10)
    #
    Radiobutton(frame_rb_type_split, text='А) По листам в одном файле', variable=group_rb_type_split,
                value=0).pack()
    Radiobutton(frame_rb_type_split, text='Б) По отдельным файлам', variable=group_rb_type_split,
                value=1).pack()

    # Создаем кнопку Выбрать файл

    btn_example_split = Button(frame_data_for_split, text='2) Выберите файл с таблицей', font=('Arial Bold', 14),
                               command=select_file_split)
    btn_example_split.pack(padx=10, pady=10)

    # Определяем числовую переменную для порядкового номера
    entry_number_column_split = IntVar()
    # Описание поля
    label_number_column_split = Label(frame_data_for_split,
                                      text='3) Введите порядковый номер колонки начиная с 1\nпо значениям которой нужно разделить таблицу')
    label_number_column_split.pack(padx=10, pady=10)
    # поле ввода имени листа
    entry_number_column_split = Entry(frame_data_for_split, textvariable=entry_number_column_split,
                                      width=30)
    entry_number_column_split.pack(ipady=5)

    btn_choose_end_folder_split = Button(frame_data_for_split, text='4) Выберите конечную папку',
                                         font=('Arial Bold', 14),
                                         command=select_end_folder_split
                                         )
    btn_choose_end_folder_split.pack(padx=10, pady=10)

    # Создаем кнопку слияния

    btn_split_process = Button(tab_split_tables, text='5) Разделить таблицу',
                               font=('Arial Bold', 20),
                               command=processing_split_table)
    btn_split_process.pack(padx=10, pady=10)

    """
    Создаем вкладку создания документов
    """
    tab_create_doc = Frame(tab_control)
    tab_control.add(tab_create_doc, text='Создание\nдокументов')

    create_doc_frame_description = LabelFrame(tab_create_doc)
    create_doc_frame_description.pack()

    lbl_hello = Label(create_doc_frame_description,
                      text='Генерация документов по шаблону'
                           '\nДля корректной работы программы уберите из таблицы\nобъединенные ячейки'
                           '\nДанные обрабатываются только с первого листа файла Excel!!!', width=60)
    lbl_hello.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # #
    # #
    # Картинка
    path_to_img = resource_path('logo.png')
    img = PhotoImage(file=path_to_img)
    Label(create_doc_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем фрейм для действий
    create_doc_frame_action = LabelFrame(tab_create_doc, text='Подготовка')
    create_doc_frame_action.pack()

    # Создаем кнопку Выбрать шаблон
    btn_template_doc = Button(create_doc_frame_action, text='1) Выберите шаблон документа', font=('Arial Bold', 14),
                              command=select_file_template_doc
                              )
    btn_template_doc.pack(padx=10, pady=10)

    btn_data_doc = Button(create_doc_frame_action, text='2) Выберите файл с данными', font=('Arial Bold', 14),
                          command=select_file_data_doc
                          )
    btn_data_doc.pack(padx=10, pady=10)
    #
    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    # Определяем текстовую переменную
    entry_name_column_data = StringVar()
    # Описание поля
    label_name_column_data = Label(create_doc_frame_action,
                                   text='3) Введите название колонки в таблице\n по которой будут создаваться имена файлов')
    label_name_column_data.pack(padx=10, pady=10)
    # поле ввода
    data_column_entry = Entry(create_doc_frame_action, textvariable=entry_name_column_data, width=30)
    data_column_entry.pack(ipady=5)

    # Поле для ввода названия генериуемых документов
    # Определяем текстовую переменную
    entry_type_file = StringVar()
    # Описание поля
    label_name_column_type_file = Label(create_doc_frame_action, text='4) Введите название создаваемых документов')
    label_name_column_type_file.pack(padx=10, pady=10)
    # поле ввода
    type_file_column_entry = Entry(create_doc_frame_action, textvariable=entry_type_file, width=30)
    type_file_column_entry.pack(ipady=5)

    btn_choose_end_folder_doc = Button(create_doc_frame_action, text='5) Выберите конечную папку',
                                       font=('Arial Bold', 14),
                                       command=select_end_folder_doc
                                       )
    btn_choose_end_folder_doc.pack(padx=10, pady=10)

    # Создаем область для того чтобы поместить туда опции
    frame_data_for_options = LabelFrame(tab_create_doc, text='Дополнительные опции')
    frame_data_for_options.pack(padx=10, pady=10)

    # Создаем переменную для хранения результа переключения чекбокса
    mode_combine_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_combine_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_calculate = Checkbutton(frame_data_for_options,
                                       text='Поставьте галочку, если вам нужно чтобы все файлы были объединены в один',
                                       variable=mode_combine_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_calculate.pack()

    # Создаем чекбокс для режима создания pdf
    # Создаем переменную для хранения результа переключения чекбокса
    mode_pdf_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_pdf_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_pdf = Checkbutton(frame_data_for_options,
                                 text='Поставьте галочку, если вам нужно чтобы \n'
                                      'дополнительно создавались pdf версии документов',
                                 variable=mode_pdf_value,
                                 offvalue='No',
                                 onvalue='Yes')
    chbox_mode_pdf.pack()

    # создаем чекбокс для единичного документа

    # Создаем переменную для хранения результа переключения чекбокса
    mode_group_doc = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_group_doc.set('No')
    # Создаем чекбокс для выбора режима подсчета
    chbox_mode_group = Checkbutton(frame_data_for_options,
                                   text='Поставьте галочку, если вам нужно создать один документ\nдля конкретного значения (например для определенного ФИО)',
                                   variable=mode_group_doc,
                                   offvalue='No',
                                   onvalue='Yes')
    chbox_mode_group.pack(padx=10, pady=10)
    # Создаем поле для ввода значения по которому будет создаваться единичный документ
    # Определяем текстовую переменную
    entry_value_column = StringVar()
    # Описание поля
    label_name_column_group = Label(frame_data_for_options,
                                    text='Введите значение из колонки\nуказанной на шаге 3 для которого нужно создать один документ,\nнапример конкретное ФИО')
    label_name_column_group.pack()
    # поле ввода
    type_file_group_entry = Entry(frame_data_for_options, textvariable=entry_value_column, width=30)
    type_file_group_entry.pack(ipady=5)

    # Создаем кнопку для создания документов из таблиц с произвольной структурой
    btn_create_files_other = Button(tab_create_doc, text='6) Создать документ(ы)',
                                    font=('Arial Bold', 20),
                                    command=generate_docs_other
                                    )
    btn_create_files_other.pack(padx=10, pady=10)

    """
    Создаем вкладку для обработки дат рождения
    """

    tab_calculate_date = Frame(tab_control)
    tab_control.add(tab_calculate_date, text='Обработка\nдат рождения')

    calculate_date_frame_description = LabelFrame(tab_calculate_date)
    calculate_date_frame_description.pack()

    lbl_hello_calculate_date = Label(calculate_date_frame_description,
                                     text='Подсчет по категориям,выделение месяца,года\nподсчет текущего возраста'
                                          '\nДля корректной работы программы уберите из таблицы\n объединенные ячейки'
                                          '\nДанные обрабатываются только с первого листа файла Excel!!!', width=60)
    lbl_hello_calculate_date.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # #
    # #
    # Картинка
    path_to_img_calculate_date = resource_path('logo.png')
    img_calculate_date = PhotoImage(file=path_to_img_calculate_date)
    Label(calculate_date_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем фрейм для действий
    calculate_date_frame_action = LabelFrame(tab_calculate_date, text='Подготовка')
    calculate_date_frame_action.pack()

    # Определяем текстовую переменную которая будет хранить дату
    entry_date = StringVar()
    # Описание поля
    label_name_date_field = Label(calculate_date_frame_action,
                                  text='Введите  дату в формате XX.XX.XXXX\n относительно, которой нужно подсчитать текущий возраст')
    label_name_date_field.pack(padx=10, pady=10)
    # поле ввода
    date_field = Entry(calculate_date_frame_action, textvariable=entry_date, width=30)
    date_field.pack(ipady=5)

    # Создаем кнопку Выбрать файл с данными
    btn_data_date = Button(calculate_date_frame_action, text='1) Выберите файл с данными', font=('Arial Bold', 14),
                           command=select_file_data_date)
    btn_data_date.pack(padx=10, pady=10)

    btn_choose_end_folder_date = Button(calculate_date_frame_action, text='2) Выберите конечную папку',
                                        font=('Arial Bold', 14),
                                        command=select_end_folder_date
                                        )
    btn_choose_end_folder_date.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_name_column = StringVar()
    # Описание поля
    label_name_column = Label(calculate_date_frame_action,
                              text='3) Введите название колонки с датами рождения,\nкоторые нужно обработать ')
    label_name_column.pack(padx=10, pady=10)
    # поле ввода
    column_entry = Entry(calculate_date_frame_action, textvariable=entry_name_column, width=30)
    column_entry.pack(ipady=5, pady=10)

    btn_calculate_date = Button(tab_calculate_date, text='4) Обработать', font=('Arial Bold', 20),
                                command=calculate_date)
    btn_calculate_date.pack(padx=10, pady=10)

    """
    Создаем вкладку для подсчета данных по категориям
    """
    tab_groupby_data = Frame(tab_control)
    tab_control.add(tab_groupby_data, text='Подсчет\nданных')

    groupby_data_frame_description = LabelFrame(tab_groupby_data)
    groupby_data_frame_description.pack()

    lbl_hello_groupby_data = Label(groupby_data_frame_description,
                                   text='Для корректной работы программы уберите из таблицы\n объединенные ячейки'
                                        '\nДанные обрабатываются только с первого листа файла Excel!!!', width=60)
    lbl_hello_groupby_data.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_groupby_data = resource_path('logo.png')
    img_groupby_data = PhotoImage(file=path_to_img_groupby_data)
    Label(groupby_data_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_groupby = LabelFrame(tab_groupby_data, text='Подготовка')
    frame_data_for_groupby.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_groupby = Button(frame_data_for_groupby, text='1) Выберите файл с данными', font=('Arial Bold', 14),
                              command=select_file_data_groupby
                              )
    btn_data_groupby.pack(padx=10, pady=10)

    btn_choose_end_folder_groupby = Button(frame_data_for_groupby, text='2) Выберите конечную папку',
                                           font=('Arial Bold', 14),
                                           command=select_end_folder_groupby
                                           )
    btn_choose_end_folder_groupby.pack(padx=10, pady=10)

    # Создаем кнопки подсчета

    btn_groupby_category = Button(tab_groupby_data, text='Подсчитать количество по категориям\nдля всех колонок',
                                  font=('Arial Bold', 20),
                                  command=groupby_category)
    btn_groupby_category.pack(padx=10, pady=10)

    btn_groupby_stat = Button(tab_groupby_data, text='Подсчитать базовую статистику\nдля всех колонок',
                              font=('Arial Bold', 20),
                              command=groupby_stat)
    btn_groupby_stat.pack(padx=10, pady=10)

    """
    Создаем вкладку для сравнения 2 столбцов
    """
    tab_comparison = Frame(tab_control)
    tab_control.add(tab_comparison, text='Слияние\n2 таблиц')

    comparison_frame_description = LabelFrame(tab_comparison)
    comparison_frame_description.pack()

    lbl_hello_comparison = Label(comparison_frame_description,
                                 text='Для корректной работы программы уберите из таблицы\nобъединенные ячейки',
                                 width=60)
    lbl_hello_comparison.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_comparison = resource_path('logo.png')
    img_comparison = PhotoImage(file=path_to_img_comparison)
    Label(comparison_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_comparison = LabelFrame(tab_comparison, text='Подготовка')
    frame_data_for_comparison.pack(padx=10, pady=10)

    # Создаем кнопку выбрать файл с параметрами
    btn_columns_params = Button(frame_data_for_comparison, text='1) Выберите файл с параметрами слияния',
                                font=('Arial Bold', 14),
                                command=select_file_params_comparsion)
    btn_columns_params.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_comparison = Button(frame_data_for_comparison, text='2) Выберите первый файл с данными',
                                       font=('Arial Bold', 14),
                                       command=select_first_comparison
                                       )
    btn_data_first_comparison.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_first_sheet_name = StringVar()
    # Описание поля
    label_first_sheet_name = Label(frame_data_for_comparison,
                                   text='3) Введите название листа в первом файле')
    label_first_sheet_name.pack(padx=10, pady=10)
    # поле ввода имени листа
    first_sheet_name_entry = Entry(frame_data_for_comparison, textvariable=entry_first_sheet_name, width=30)
    first_sheet_name_entry.pack(ipady=5)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_comparison = Button(frame_data_for_comparison, text='4) Выберите второй файл с данными',
                                        font=('Arial Bold', 14),
                                        command=select_second_comparison
                                        )
    btn_data_second_comparison.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_sheet_name = StringVar()
    # Описание поля
    label_second_sheet_name = Label(frame_data_for_comparison,
                                    text='5) Введите название листа во втором файле')
    label_second_sheet_name.pack(padx=10, pady=10)
    # поле ввода
    second__sheet_name_entry = Entry(frame_data_for_comparison, textvariable=entry_second_sheet_name, width=30)
    second__sheet_name_entry.pack(ipady=5)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_comparison = Button(frame_data_for_comparison, text='6) Выберите конечную папку',
                                       font=('Arial Bold', 14),
                                       command=select_end_folder_comparison
                                       )
    btn_select_end_comparison.pack(padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_comparison = Button(tab_comparison, text='7) Произвести слияние\nтаблиц', font=('Arial Bold', 20),
                                    command=processing_comparison
                                    )
    btn_data_do_comparison.pack(padx=10, pady=10)

    """
    Извлечение данных из таблиц со сложной структурой
    """
    tab_extract_data = Frame(tab_control)
    tab_control.add(tab_extract_data, text='Извлечение\nданных')

    extract_data_frame_description = LabelFrame(tab_extract_data)
    extract_data_frame_description.pack()

    lbl_hello_extract_data = Label(extract_data_frame_description,
                                   text='Извлечение данных из файлов Excel с сложной структурой', width=60)

    lbl_hello_extract_data.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_extract_data = resource_path('logo.png')
    img_extract_data = PhotoImage(file=path_to_img_extract_data)
    Label(extract_data_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_extract_data = LabelFrame(tab_extract_data, text='Подготовка')
    frame_data_extract_data.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать файл с параметрами
    btn_select_file_params = Button(frame_data_extract_data, text='1) Выбрать файл с параметрами',
                                    font=('Arial Bold', 14),
                                    command=select_file_params_calculate_data
                                    )
    btn_select_file_params.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_select_files_data = Button(frame_data_extract_data, text='2) Выбрать папку с данными', font=('Arial Bold', 14),
                                   command=select_files_data_calculate_data
                                   )
    btn_select_files_data.pack(padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(frame_data_extract_data, text='3) Выбрать конечную папку', font=('Arial Bold', 14),
                                   command=select_end_folder_calculate_data
                                   )
    btn_choose_end_folder.pack(padx=10, pady=10)

    # Создаем переменную для хранения результа переключения чекбокса
    mode_text_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_text_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_calculate = Checkbutton(frame_data_extract_data,
                                       text='Поставьте галочку, если вам нужно подсчитать текстовые данные ',
                                       variable=mode_text_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_calculate.pack(padx=10, pady=10)

    # Создаем кнопку для запуска подсчета файлов

    btn_calculate = Button(tab_extract_data, text='4) Подсчитать', font=('Arial Bold', 20),
                           command=calculate_data
                           )
    btn_calculate.pack(padx=10, pady=10)

    """
    Создание вкладки для объединения таблиц в одну большую
    """
    tab_merger_tables = Frame(tab_control)
    tab_control.add(tab_merger_tables, text='Слияние\nфайлов')

    merger_tables_frame_description = LabelFrame(tab_merger_tables)
    merger_tables_frame_description.pack()

    lbl_hello_merger_tables = Label(merger_tables_frame_description,
                                    text='Слияние файлов Excel с одинаковой структурой'
                                         '\nДля корректной работы программы уберите из таблицы\n объединенные ячейки',
                                    width=60)

    lbl_hello_merger_tables.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_merger_tables = resource_path('logo.png')
    img_merger_tables = PhotoImage(file=path_to_img_merger_tables)
    Label(merger_tables_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_merger_tables = LabelFrame(tab_merger_tables, text='Подготовка')
    frame_data_merger_tables.pack(padx=10, pady=10)

    # Переключатель:вариант слияния файлов
    # Создаем переключатель
    group_rb_type_harvest = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_harvest = LabelFrame(frame_data_merger_tables, text='1) Выберите вариант слияния')
    frame_rb_type_harvest.pack(padx=10, pady=10)
    #
    Radiobutton(frame_rb_type_harvest, text='А) Простое слияние по названию листов', variable=group_rb_type_harvest,
                value=0).pack()
    Radiobutton(frame_rb_type_harvest, text='Б) Слияние по порядку листов', variable=group_rb_type_harvest,
                value=1).pack()
    Radiobutton(frame_rb_type_harvest, text='В) Сложное слияние по названию листов', variable=group_rb_type_harvest,
                value=2).pack()

    # Создаем кнопку Выбрать папку с данными

    btn_data_merger = Button(frame_data_merger_tables, text='2) Выберите папку с данными', font=('Arial Bold', 14),
                             command=select_folder_data_merger
                             )
    btn_data_merger.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать эталонный файл

    btn_example_merger = Button(frame_data_merger_tables, text='3) Выберите эталонный файл', font=('Arial Bold', 14),
                                command=select_standard_file_merger)
    btn_example_merger.pack(padx=10, pady=10)

    btn_choose_end_folder_merger = Button(frame_data_merger_tables, text='4) Выберите конечную папку',
                                          font=('Arial Bold', 14),
                                          command=select_end_folder_merger
                                          )
    btn_choose_end_folder_merger.pack(padx=10, pady=10)

    # Определяем переменную в которой будем хранить количество пропускаемых строк
    merger_entry_skip_rows = StringVar()
    # Описание поля
    merger_label_skip_rows = Label(frame_data_merger_tables,
                                   text='5) Введите количество строк\nв листах,чтобы пропустить\nзаголовок\n'
                                        'ТОЛЬКО для вариантов слияния А и Б ')
    merger_label_skip_rows.pack(padx=10, pady=10)
    # поле ввода
    merger_number_skip_rows = Entry(frame_data_merger_tables, textvariable=merger_entry_skip_rows, width=5)
    merger_number_skip_rows.pack(ipady=5)

    # Создаем кнопку выбора файла с параметрами
    btn_params_merger = Button(frame_data_merger_tables, text='Выберите файл с параметрами слияния\n'
                                                              'ТОЛЬКО для варианта В', font=('Arial Bold', 14),
                               command=select_params_file_merger)
    btn_params_merger.pack(padx=10, pady=10)
    # Создаем кнопку слияния

    btn_merger_process = Button(tab_merger_tables, text='6) Произвести слияние \nфайлов',
                                font=('Arial Bold', 20),
                                command=merge_tables)
    btn_merger_process.pack(padx=10, pady=10)

    """
    Создание вкладки для склонения ФИО по падежам
    """
    # Создаем вкладку для подсчета данных по категориям
    tab_decl_by_cases = Frame(tab_control)
    tab_control.add(tab_decl_by_cases, text='Склонение ФИО\nпо падежам')

    decl_by_cases_frame_description = LabelFrame(tab_decl_by_cases)
    decl_by_cases_frame_description.pack()

    lbl_hello_decl_by_cases = Label(decl_by_cases_frame_description,
                                    text='Для корректной работы программы уберите из таблицы\n объединенные ячейки',
                                    width=60)

    lbl_hello_decl_by_cases.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_decl_by_cases = resource_path('logo.png')
    img_decl_by_cases = PhotoImage(file=path_to_img_decl_by_cases)
    Label(decl_by_cases_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_decl_by_cases = LabelFrame(tab_decl_by_cases, text='Подготовка')
    frame_data_decl_by_cases.pack(padx=10, pady=10)

    # выбрать файл с данными
    btn_data_decl_case = Button(frame_data_decl_by_cases, text='1) Выберите файл с данными', font=('Arial Bold', 14),
                                command=select_data_decl_case)
    btn_data_decl_case.pack(padx=10, pady=10)

    # Ввести название колонки с ФИО
    # # Определяем переменную
    decl_case_fio_col = StringVar()
    # Описание поля ввода
    decl_case_label_fio = Label(frame_data_decl_by_cases,
                                text='2) Введите название колонки\n с ФИО в им.падеже')
    decl_case_label_fio.pack(padx=10, pady=10)
    # поле ввода
    decl_case_entry_fio = Entry(frame_data_decl_by_cases, textvariable=decl_case_fio_col, width=25)
    decl_case_entry_fio.pack(ipady=5)
    #
    btn_choose_end_folder_decl_case = Button(frame_data_decl_by_cases, text='3) Выберите конечную папку',
                                             font=('Arial Bold', 14),
                                             command=select_end_folder_decl_case
                                             )
    btn_choose_end_folder_decl_case.pack(padx=10, pady=10)

    # Создаем кнопку склонения по падежам

    btn_decl_case_process = Button(tab_decl_by_cases, text='4) Произвести склонение \nпо падежам',
                                   font=('Arial Bold', 20),
                                   command=process_decl_case)
    btn_decl_case_process.pack(padx=10, pady=10)

    """
    Разница двух таблиц
    """
    tab_diffrence = Frame(tab_control)
    tab_control.add(tab_diffrence, text='Разница\n2 таблиц')

    diffrence_frame_description = LabelFrame(tab_diffrence)
    diffrence_frame_description.pack()

    lbl_hello_diffrence = Label(diffrence_frame_description,
                                text='Количество строк и колонок в таблицах должно совпадать\n'
                                     'Названия колонок в таблицах должны совпадать'
                                     '\nДля корректной работы программы уберите из таблицы\n объединенные ячейки',
                                width=60)

    lbl_hello_diffrence.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_diffrence = resource_path('logo.png')
    img_diffrence = PhotoImage(file=path_to_img_diffrence)
    Label(diffrence_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_diffrence = LabelFrame(tab_diffrence, text='Подготовка')
    frame_data_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_diffrence = Button(frame_data_diffrence, text='1) Выберите файл с первой таблицей',
                                      font=('Arial Bold', 14),
                                      command=select_first_diffrence
                                      )
    btn_data_first_diffrence.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_first_sheet_name_diffrence = StringVar()
    # Описание поля
    label_first_sheet_name_diffrence = Label(frame_data_diffrence,
                                             text='2) Введите название листа, где находится первая таблица')
    label_first_sheet_name_diffrence.pack(padx=10, pady=10)
    # поле ввода имени листа
    first_sheet_name_entry_diffrence = Entry(frame_data_diffrence, textvariable=entry_first_sheet_name_diffrence,
                                             width=30)
    first_sheet_name_entry_diffrence.pack(ipady=5)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_diffrence = Button(frame_data_diffrence, text='3) Выберите файл со второй таблицей',
                                       font=('Arial Bold', 14),
                                       command=select_second_diffrence
                                       )
    btn_data_second_diffrence.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_sheet_name_diffrence = StringVar()
    # Описание поля
    label_second_sheet_name_diffrence = Label(frame_data_diffrence,
                                              text='4) Введите название листа, где находится вторая таблица')
    label_second_sheet_name_diffrence.pack(padx=10, pady=10)
    # поле ввода
    second__sheet_name_entry_diffrence = Entry(frame_data_diffrence, textvariable=entry_second_sheet_name_diffrence,
                                               width=30)
    second__sheet_name_entry_diffrence.pack(ipady=5)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_diffrence = Button(frame_data_diffrence, text='5) Выберите конечную папку',
                                      font=('Arial Bold', 14),
                                      command=select_end_folder_diffrence
                                      )
    btn_select_end_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_diffrence = Button(tab_diffrence, text='6) Обработать таблицы', font=('Arial Bold', 20),
                                   command=processing_diffrence
                                   )
    btn_data_do_diffrence.pack(padx=10, pady=10)

    """
    Создание сводных таблиц
    """
    tab_svod = Frame(tab_control)
    tab_control.add(tab_svod, text='Сводные\nтаблицы')

    svod_frame_description = LabelFrame(tab_svod)
    svod_frame_description.pack()

    lbl_hello_svod = Label(svod_frame_description,
                           text='Создание сводных таблиц: Сумма,Среднее,Медиана,Минимум\n'
                                'Максимум,Количество,Количество уникальных,Самое частое\n'
                                'Количество самых частых, Количество дубликатов'
                                '\nДля корректной работы программы уберите из таблицы\n объединенные ячейки\n'
                                'Заголовок таблицы должен быть на первой строке', width=60)

    lbl_hello_svod.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_svod = resource_path('logo.png')
    img_svod = PhotoImage(file=path_to_img_svod)
    Label(svod_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_svod = LabelFrame(tab_svod, text='Подготовка')
    frame_data_svod.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  файл с данными
    btn_data_svod = Button(frame_data_svod, text='1) Выберите файл с данными',
                           font=('Arial Bold', 14),
                           command=select_file_svod
                           )
    btn_data_svod.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_sheet_name_svod = StringVar()
    # Описание поля
    label_sheet_name_svod = Label(frame_data_svod,
                                  text='2) Введите название листа в файле где находятся данные')
    label_sheet_name_svod.pack(padx=10, pady=10)
    # поле ввода имени листа
    sheet_name_entry_svod = Entry(frame_data_svod, textvariable=entry_sheet_name_svod,
                                  width=30)
    sheet_name_entry_svod.pack(ipady=5)

    # Определяем текстовую переменную
    entry_columns_svod = StringVar()
    # Описание поля
    label_columns_svod = Label(frame_data_svod,
                               text='3) Введите порядковые номера колонок по которым \n'
                                    'будут группироваться данные\n'
                                    'Например: 2,4,8')
    label_columns_svod.pack(padx=10, pady=10)
    # поле ввода имени листа
    entry_svod_columns = Entry(frame_data_svod, textvariable=entry_columns_svod,
                               width=30)
    entry_svod_columns.pack(ipady=5)

    # Определяем текстовую переменную
    entry_target_svod = StringVar()
    # Описание поля
    label_target_svod = Label(frame_data_svod,
                              text='4) Введите порядковые номера колонок по которым\n'
                                   'будет вестить подсчет\n'
                                   'Например: 7,8')
    label_target_svod.pack(padx=10, pady=10)
    # поле ввода
    entry_target_columns = Entry(frame_data_svod, textvariable=entry_target_svod,
                                 width=30)
    entry_target_columns.pack(ipady=5)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_svod = Button(frame_data_svod, text='5) Выберите конечную папку',
                                 font=('Arial Bold', 14),
                                 command=select_end_folder_svod
                                 )
    btn_select_end_svod.pack(padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_svod = Button(tab_svod, text='6) Обработать данные', font=('Arial Bold', 20),
                              command=processing_svod
                              )
    btn_data_do_svod.pack(padx=10, pady=10)

    """
    Создаем вкладку для размещения описания программы, руководства пользователя,лицензии.
    """

    tab_about = ttk.Frame(tab_control)
    tab_control.add(tab_about, text='О ПРОГРАММЕ')

    # Кнопка, которая вызывает функцию open_new_window
    button = Button(tab_about, text="Лицензия", command=open_license)
    button.pack()




    # Создаем виджет для управления полосой прокрутки
    canvas.create_window((0, 0), window=tab_control, anchor="nw")

    # Конфигурируем холст для обработки скроллинга
    canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
    scrollbar.pack(side="right", fill="y")

    # Вешаем событие скроллинга
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)
    window.mainloop()

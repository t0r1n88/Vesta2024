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
        checkbox_mix_alphabets = mode_mix_alphabets.get()
        prepare_list(glob_prep_file, glob_path_to_end_folder_prep, checkbox_dupl,checkbox_mix_alphabets)

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
    text_area = Text(new_window,width=90, height=50)

    with open(license_path, 'r', encoding='utf-8') as file:
        text = file.read()
        text_area.insert(END, text)
    text_area.configure(state='disabled')
    text_area.pack(side=LEFT)

    scroll = Scrollbar(new_window,command=text_area.yview)
    scroll.pack(side=LEFT, fill=Y)

    text_area.config(yscrollcommand=scroll.set)



if __name__ == '__main__':
    window = Tk()
    window.title('Веста Обработка таблиц и создание документов ver 1.51')
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
                                       'Создание списка дубликатов по каждой колонке.\n'
                                       'Поиск со смешаным написанием русских и английских букв.\n'
                                       'ПРИМЕЧАНИЯ\n'
                                       'Заголовок таблицы должен занимать только первую строку!\n'
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

    # Создаем переменную для хранения результа переключения чекбокса поиска смешения
    mode_mix_alphabets = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_mix_alphabets.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_mix_alphabets = Checkbutton(frame_data_prep,
                                  text='Проверить каждую ячейку таблицы на смешение русских и английских букв',
                                  variable=mode_mix_alphabets,
                                  offvalue='No',
                                  onvalue='Yes')
    chbox_mode_mix_alphabets.pack(padx=10, pady=10)


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

    about_frame_description = LabelFrame(tab_about,text='О программе')
    about_frame_description.pack()

    lbl_about = Label(about_frame_description,
                           text="""Веста - Утилита для обработки таблиц и создания однотипных документов
                           Версия 1.51
                           Язык программирования - Python 3\n
                           Используемая лицензия BSD-2-Clause\n
                           Copyright (c) <2024> <Будаев Олег Тимурович>
                           Адрес сайта программы: https://itdarhan.ru/vesta.html
                           """, width=60)

    lbl_about.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_about = resource_path('logo.png')
    img_about = PhotoImage(file=path_to_img_about)
    Label(about_frame_description,
          image=img_about, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем поле для лицензий библиотек
    dop_library_license_frame_description = LabelFrame(tab_about,text='Дополнительные библиотеки Python')
    dop_library_license_frame_description.pack()

    text_area = Text(dop_library_license_frame_description, width=84, height=20)
    lib_license_path = resource_path('LibraryLicense.txt')  # путь к файлу лицензии
    with open(lib_license_path, 'r', encoding='utf-8') as file:
        text = file.read()
        text_area.insert(END, text)
    text_area.configure(state='disabled')
    text_area.pack(side=LEFT)

    scroll = Scrollbar(dop_library_license_frame_description,command=text_area.yview)
    scroll.pack(side=LEFT, fill=Y)

    text_area.config(yscrollcommand=scroll.set)

    text_area.configure(state='disabled')
    text_area.pack(side=LEFT)





    # Кнопка, которая вызывает функцию open_new_window
    license_path = resource_path('License.txt') # путь к файлу лицензии
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

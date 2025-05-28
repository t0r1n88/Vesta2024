"""
Скрипт для подсчета табличных отчетов
"""
import copy

from support_functions import convert_to_int, convert_to_float,write_df_big_dct_to_excel, del_sheet

import pandas as pd
pd.options.display.width= None
pd.options.display.max_columns= None
import openpyxl
import time
import re
import os
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action="ignore", category=pd.errors.PerformanceWarning)
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None

from tkinter import messagebox

class NotCorrectParams(Exception):
    """
    Исключение для случаев когда нет ни одного корретного параметра
    """
    pass

def check_range(df:pd.DataFrame):
    """
    Функция для проверки корректности записи диапазона
    :param датафрейм: параметры
    :return: Правильно или Неправильно
    """
    error_df = pd.DataFrame(
        columns=['Название файла','Название листа', 'Описание ошибки'])  # датафрейм для ошибок

    pattern = r'^([A-Z]{1,2})(\d+):([A-Z]{1,2})(\d+)$'

    out_dct = {} # словарь для хранения параметров
    # перебираем строки
    for row in df.itertuples():
        name_sheet = row[1] # название листа
        name_range = row[2] # диапазон
        number_cols = row[3] # количество колонок
        # Проверяем на корректность диапазон
        result = re.search(pattern,name_range)
        if result:
            prep_name_range = name_range.replace(':','-') # заменяем двоеточие на тире чтобы потом не было проблем с сохранением листов
            out_dct[f'{name_sheet}_{prep_name_range}'] = {'Название листа':name_sheet,'Диапазон':name_range,'Количество колонок':number_cols}

        else:
            temp_error_df = pd.DataFrame(columns=['Название файла','Название листа', 'Описание ошибки'],
                data=[['Ошибка в файле параметров', f'Строка {name_sheet}', f'Ошибка в написании диапазона- {name_range} . Правильный формат написания 1 или 2 буквы латинского алфавита потом число потом двоеточие потом одна или 2 буквы латинского алфавита потом число. Например F2:AZ10']])
            error_df = pd.concat([error_df,temp_error_df])
    return out_dct,error_df


def counting_table_report(file_params:str, report_dir:str,path_end_folder:str):
    """
    Функция для подсчета данных из определенного диапазона таблицы
    :param file_params: файл с параметрами, где указаны названия листов, диапазон с данными, количество обрабатываемых колонок
    :param report_dir: папка с отчетами
    :param path_end_folder: конечная папка
    """
    current_time = time.strftime('%H_%M_%S')

    error_df = pd.DataFrame(
        columns=['Название файла','Название листа', 'Описание ошибки'])  # датафрейм для ошибок

    try:
        try:
            params_df = pd.read_excel(file_params,usecols='A:C')
        except:
            messagebox.showerror('Веста Обработка таблиц и создание документов', 'Не удалось обработать файл с параметрами обработки отчетов!\n'
                                                                'Проверьте файл на повреждения. Пересохраните в новом файле.')
        params_df.dropna(how='any',inplace=True) # очищаем от неполных строк
        params_df.columns = ['Название листа','Диапазон','Количество колонок'] # переименовываем
        # Приводим к нужным типам данных
        params_df[['Название листа','Диапазон']] = params_df[['Название листа','Диапазон']].astype(str)
        params_df['Количество колонок'] = params_df['Количество колонок'].apply(convert_to_int)
        params_df = params_df[params_df['Количество колонок'] >= 1] # отбираем только те колонки, что больше нуля

        dct_params,error_params_df = check_range(params_df) # создаем словарь с параметрами и датафрейм с ошибками
        # добавляем в датафрейм ошибок
        error_df = pd.concat([error_df,error_params_df],ignore_index=True)
        if len(dct_params) == 0:
            raise NotCorrectParams

        # получаем листы которые нужно обработать
        lst_required_sheets = [value['Название листа'] for key, value in dct_params.items()]


        dct_work = copy.deepcopy(dct_params) # создаем словарь в котором будет производить всю работу
        # Создаем датафреймы для каждого листа для хранения данных
        for key,value in dct_work.items():
            # Датафрейм для суммирования результатов
            dct_work[key]['Данные'] = pd.DataFrame(columns=range(value['Количество колонок']))
            dct_work[key]['Список для проверки'] = pd.DataFrame(columns=[f'Колонка {i}' for i in range(1,value['Количество колонок'] + 1)])
            dct_work[key]['Список для проверки'].insert(0,'Номер строки','')
            dct_work[key]['Список для проверки'].insert(1,'Название файла','')

        for dirpath, dirnames, filenames in os.walk(report_dir):
            for file in filenames:
                if file.endswith('.xls'):
                    temp_error_df = pd.DataFrame(
                        data=[[f'{file}',
                               f'Программа обрабатывает файлы с разрешением xlsx. XLS файлы не обрабатываются !'
                               ]],
                        columns=['Название файла',
                                 'Описание ошибки'])
                    error_df = pd.concat([error_df, temp_error_df], axis=0,
                                         ignore_index=True)
                    continue

                if not file.startswith('~$') and file.endswith('.xlsx'):
                    name_file = file.split('.xlsx')[0]
                    print(name_file) # обрабатываемый файл
                    # Проверяем чтобы файл не был резервной копией или файлом с другим расширением.
                    if file.startswith('~$'):
                        continue
                    try:
                        wb = openpyxl.load_workbook(f'{dirpath}/{file}',data_only=True) # открываем файл
                    except:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{file}',
                                   f'Не удалось обработать файл. Возможно файл поврежден'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        continue
                    diff_req_sheet = set(lst_required_sheets).difference(set(wb.sheetnames)) # получаем разницу
                    if len(diff_req_sheet) != 0:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{file}',
                                   f'В файле отсутствуют указанные листы :{";".join(diff_req_sheet)}'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        continue

                    # Извлекаем данные
                    for key,value in dct_work.items():
                        name_sheet = value['Название листа']
                        try:
                            cells_range = wb[name_sheet][value['Диапазон']]
                        except ValueError:
                            # Проверяем на длину получившихся данных
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}',
                                       f'{name_sheet}',
                                       f'Указанный для листа {value["Название листа"]} диапазон {value["Диапазон"]} превышает допустимые для листа xlsx значения. Допустимые значения от 1 до 1048576'
                                       ]],
                                columns=['Название файла', 'Название листа',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0,
                                                 ignore_index=True)
                            continue

                        # Создаем пустой список для хранения данных
                        data = []
                        check_data = [] # список для хранения проверочных данных

                        # Перебираем строки и извлекаем данные
                        for idx,row in enumerate(cells_range,1):
                            row_to_append = [cell.value for cell in row] # строка с данными
                            data.append(row_to_append) # добавляем в будущий датафрейм для суммирования

                            # добавляем вспомогательные колонки для последующего добавления в проверочный датафрейм
                            copy_row = row_to_append.copy()
                            copy_row.insert(0,f'Строка {idx}')
                            copy_row.insert(1,name_file)
                            check_data.append(copy_row)

                        # Проверяем на длину получившихся данных
                        if len(data[0]) != value['Количество колонок']:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}',
                                       f'{name_sheet}',
                                       f'Указанный для листа {value["Название листа"]} диапазон {value["Диапазон"]} не совпадает с количество обрабатываемых колонок. В параметрах отчета указано количество обрабатываемых колонок равное {value["Количество колонок"]}, а указанный диапазон занимает {len(data[0])}'
                                       ]],
                                columns=['Название файла','Название листа',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0,
                                                 ignore_index=True)
                            continue

                        # Преобразуем список в DataFrame
                        temp_df = pd.DataFrame(data,columns=range(value['Количество колонок']))
                        temp_df = temp_df.applymap(convert_to_float)
                        if len(value['Данные']) == 0:
                            value['Данные'] = pd.concat([value['Данные'],temp_df])
                        else:
                            value['Данные'] = value['Данные'].add(temp_df,fill_value= 0)
                        # Создаем проверочный датафрейм
                        check_lst = [f'Колонка {i}' for i in range(1,value['Количество колонок'] + 1)]
                        int_cols = check_lst.copy() # колонки с числовыми значениями которые нужно привести к флоату
                        check_lst.insert(0,'Номер строки')
                        check_lst.insert(1,'Название файла')
                        check_df = pd.DataFrame(check_data,columns=check_lst)
                        check_df[int_cols] = check_df[int_cols].applymap(convert_to_float)
                        # Добавляем в проверочный датафрейм
                        value['Список для проверки'] = pd.concat([value['Список для проверки'],check_df])





        # Переименовываем индексы и колонки для удобства
        for key, value in dct_work.items():
            value['Данные'].columns = [f'Колонка {i}' for i in range(1,value['Данные'].shape[1] + 1)]
            value['Данные'].index = [f'Строка {i}' for i in range(1,value['Данные'].shape[0] + 1)]


        # Записываем итоговый файл
        itog_dct = {key:value['Данные'] for key,value in dct_work.items()}
        itog_wb = write_df_big_dct_to_excel(itog_dct, write_index=True)
        itog_wb = del_sheet(itog_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        itog_wb.save(f'{path_end_folder}/Итог подсчета {current_time}.xlsx')

        # Записываем датафрейм с данными для проверки правильности подсчета
        # Записываем итоговый файл
        data_dct = {key:value['Список для проверки'] for key,value in dct_work.items()}
        data_wb = write_df_big_dct_to_excel(data_dct, write_index=False)
        data_wb = del_sheet(data_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        data_wb.save(f'{path_end_folder}/Данные для проверки подсчета {current_time}.xlsx')



        error_wb = write_df_big_dct_to_excel({'Ошибки':error_df}, write_index=False)
        error_wb = del_sheet(error_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        error_wb.save(f'{path_end_folder}/Ошибки при подсчете {current_time}.xlsx')


    except ZeroDivisionError:
        print('заглушка')
    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов',
                            'Подсчет данных завершен!!!')

if __name__ == '__main__':
    main_file_params = 'data/Подсчет отчетов/Параметры отчета Пример 1.xlsx'
    main_data_folder = 'data/Подсчет отчетов/Табличные отчеты Пример 1'
    main_path_end_folder = 'data/result'

    counting_table_report(main_file_params,main_data_folder,main_path_end_folder)

    print('Lindy Booth!')





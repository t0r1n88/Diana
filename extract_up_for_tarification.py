"""
Скрипт для извлечения данных о дисциплинах находящихся в учебном плане скачанного с сетевого города для использования при подсчете тарификации
"""

import pandas as pd
pd.options.display.width= None
pd.options.display.max_columns= None
import openpyxl
import time
import re
import copy
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
    Исключение для случаев когда нет ни одного корректного параметра
    """
    pass

class NotFile(Exception):
    """
    Обработка случаев когда нет файлов в папке
    """
    pass


def load_file_params(file_params:str):
    """
    Функция для обработки файла с параметрами
    :param file_params:файл с параметрами
    :return:словарь с параметрами
    """
    try:
        params_df = pd.read_excel(file_params, usecols='A:B')
    except:
        messagebox.showerror('Диана',
                             'Не удалось обработать файл с параметрами обработки учебных планов!\n'
                             'Проверьте файл на повреждения. Пересохраните в новом файле.')
    params_df.dropna(how='any', inplace=True)  # очищаем от неполных строк
    params_df.columns = ['Параметр', 'Значение']  # переименовываем
    # Приводим к нужным типам данных
    # params_df[['Название листа', 'Диапазон']] = params_df[['Название листа', 'Диапазон']].astype(str)
    # params_df['Количество колонок'] = params_df['Количество колонок'].apply(convert_to_int)


def processing_data_up_for_tarification(data_folder:str, file_params:str, result_folder:str):
    """
    Функция для извлечения данных по учебным планам
    :param data_folder: папка с данными
    :param file_params: файл с параметрами
    :param result_folder: конечная папка
    """
    error_df = pd.DataFrame(
        columns=['Название файла','Описание ошибки'])  # датафрейм для ошибок

    dct_params = load_file_params(file_params) # получаем параметры # TODO Доделать функцию
    name_sheet = 'Учебный план' # название листа
    quantity_header = 6 # количество строк заголовка
    number_main_column = 2 - 1 # порядковый номер колонки с наименованиями
    quantity_cols = 15 # количество колонок с данными которые нужно собрать

    dct_file = dict() # основной словарь со структорой Название файла:{Дисциплина:список значений}

    for dirpath, dirnames, filenames in os.walk(data_folder):
        for file in filenames:
            if file.endswith('.xls') or file.endswith('.ods'):
                temp_error_df = pd.DataFrame(
                    data=[[f'{file}',
                           f'Программа обрабатывает файлы с разрешением xlsx. XLS и ODS файлы не обрабатываются !'
                           ]],
                    columns=['Название файла',
                             'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0,
                                     ignore_index=True)
                continue
            if not file.startswith('~$') and file.endswith('.xlsx'):
                name_file = file.split('.xlsx')[0]
                print(name_file)  # обрабатываемый файл
                try:
                    temp_df = pd.read_excel(f'{dirpath}/{file}',sheet_name=name_sheet,skiprows=quantity_header,header=None,usecols=list(range(0,quantity_cols)))  # открываем файл
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

                lst_name_subject= [spec for spec in
                             temp_df.iloc[:,number_main_column].unique()]  # получаем список предметов и дисциплин

                # Названия колонок
                column_cat = [f'Колонка {i}' for i in range(1, quantity_cols+1)]
                spec_dct = {key: None for key in column_cat} #
                dct_file[name_file] = {code: copy.deepcopy(spec_dct) for code in lst_name_subject}

                # Создание словаря для хранения данных с основного листа
                for row in temp_df.itertuples():
                    data_row = row[1:]  # получаем срез с нужными данными колонки в которых есть числа
                    for idx_col, value in enumerate(data_row, start=1):
                        dct_file[name_file][row[number_main_column+1]][f'Колонка {idx_col}'] = value

                print(dct_file)









if __name__ == '__main__':
    main_data_folder = 'data/Учебные планы'
    main_file_params = 'data/Параметры сбора.xlsx'
    main_result_folder = 'data/Результат'
    processing_data_up_for_tarification(main_data_folder,main_file_params,main_result_folder)
    print('Lindy Booth')
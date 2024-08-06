"""
Скрипт для создания отчетности по преподавателям
"""
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from docxtpl import DocxTemplate
import string
import time
import re
import os
from tkinter import messagebox
from jinja2 import exceptions

pd.options.mode.chained_assignment = None  # default='warn'
pd.set_option('display.max_columns', None)  # Отображать все столбцы
pd.set_option('display.expand_frame_repr', False)  # Не переносить строки
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=FutureWarning, module='openpyxl')
warnings.filterwarnings("ignore", category=DeprecationWarning)


def create_report_teacher(data_folder: str, result_folder: str):
    """
    Функция для создания отчетности по преподавателям
    :param data_folder: папка где хранятся личные дела преподавателей
    :param result_folder: папка в которую будут сохраняться итоговые файлы
    """
    # обязательные листы
    required_sheets_columns = {'Общие сведения': ['ФИО', 'Дата рождения', 'Дата начала работы в ПОО',
                                                  'Преподаваемая дисциплина', 'Общий стаж работы',
                                                  'Педагогический стаж',
                                                  'Сведения об образовании (образовательное учреждение, квалификация, год окончания)',
                                                  'Категория', '№ приказа на аттестацию, дата',
                                                  'Наличие личного сайта, блога (ссылка)'],
                               'Повышение квалификации': ['Название программы повышения квалификации',
                                                          'Вид повышения квалификации',
                                                          'Место прохождения программы',
                                                          'Дата прохождения программы (с какого по какое число, месяц, год)',
                                                          'Количество академических часов',
                                                          'Наименование подтверждающего документа, его номер и дата выдачи'],
                               'Стажировка': ['Место стажировки', 'Кол-во часов', 'Дата '],
                               'Методические разработки': ['Вид методического издания', 'Название издания',
                                                           'Профессия/специальность ',
                                                           'Дата разработки', 'Кем утверждена'],
                               'Мероприятия, пров. ППС': ['Название мероприятия', 'Дата', 'Уровень мероприятия'],
                               'Личное выступление ППС': ['Дата', 'Название мероприятия', 'Тема', 'Вид мероприятия',
                                                          'Уровень мероприятия',
                                                          'Способ участия', 'Результат участия'],
                               'Публикации': ['Полное название статьи', 'Издание', 'Дата выпуска'],
                               'Открытые уроки': ['Дисциплина', 'Группа', 'Тема', 'Дата проведения'],
                               'Взаимопосещение': ['ФИО посещенного педагога', 'Дата посещения', 'Группа', 'Тема'],
                               'УИРС': ['ФИО обучающегося', 'Профессия/специальность', 'Группа', 'Вид мероприятия',
                                        'Название мероприятия',
                                        'Тема', 'Способ участия', 'Уровень мероприятия', 'Дата проведения',
                                        'Результат участия'],
                               'Работа по НМР': ['Тема НИРП', 'Проведено ли обобщение опыта', 'Форма обобщения опыта',
                                                 'Место обобщения опыта', 'Дата обобщения опыта'],
                               'Данные для списков': []
                               }
    error_df = pd.DataFrame(
        columns=['Название файла', 'Название листа', 'Описание ошибки'])  # датафрейм для ошибок

    for idx, file in enumerate(os.listdir(data_folder)):
        if not file.startswith('~$') and file.endswith('.xlsx'):
            name_file = file.split('.xlsx')[0]  # Получаем название файла
            # загружаем файл, чтобы получить названия листов
            check_sheets_wb = openpyxl.load_workbook(f'{data_folder}/{file}', read_only=True)
            file_sheets = check_sheets_wb.sheetnames  # получаем названия листов
            check_sheets_wb.close()  # закрываем файл
            # проверяем наличие нужных листов
            diff_sheets = required_sheets_columns.difference(set(file_sheets))
            if len(diff_sheets) != 0:
                # Записываем ошибку
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', ';'.join(diff_sheets),
                                                    'Не найдены указанные обязательные колонки']],
                                             columns=['Название файла', 'Название листа',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            # Проверка листов


            # генерируем текущее время
            t = time.localtime()
            current_time = time.strftime('%H_%M_%S', t)
            error_df.to_excel(f'{result_folder}/Ошибки {current_time}.xlsx', index=False)


if __name__ == '__main__':
    main_data_folder = 'data/Данные'
    main_result_folder = 'data/Результат'

    create_report_teacher(main_data_folder, main_result_folder)

    print('Lindy Booth')

"""
Функция для создания документов Word
"""

import datetime
from collections import Counter

import pandas as pd
import openpyxl
from docxtpl import DocxTemplate
import time
import re
import os

pd.options.mode.chained_assignment = None  # default='warn'
pd.set_option('display.max_columns', None)  # Отображать все столбцы
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=FutureWarning, module='openpyxl')
warnings.filterwarnings("ignore", category=DeprecationWarning)


# def check_required_sheet_in_file(path_to_checked_file: str, func_required_sheets_columns: dict, func_name_file: str):
#     """
#     Функция для проверки наличия обязательных листов в файле
#     :param path_to_checked_file: путь к проверяемому файлу
#     :param func_required_sheets_columns: словарь с данными для проверки формата {Лист:[Список обязательных колонок}
#     :param func_name_file:  имя проверяемого файла
#     :return: датафрейм с найденными ошибками
#     """
#     # загружаем файл, чтобы получить названия листов
#     check_sheets_wb = openpyxl.load_workbook(path_to_checked_file, read_only=True)
#     file_sheets = check_sheets_wb.sheetnames  # получаем названия листов
#     check_sheets_wb.close()  # закрываем файл
#     # проверяем наличие нужных листов
#     diff_sheets = set(func_required_sheets_columns.keys()).difference(set(file_sheets))
#     if len(diff_sheets) != 0:
#         # Записываем ошибку
#         temp_error_df = pd.DataFrame(data=[[f'{func_name_file}', ';'.join(diff_sheets),
#                                             'Не найдены указанные обязательные листы']],
#                                      columns=['Название файла', 'Название листа',
#                                               'Описание ошибки'])
#         return temp_error_df
#
#
# def check_required_columns_in_sheet(path_to_checked_file: str, func_required_sheets_columns: dict, func_name_file: str):
#     """
#     Функция для проверки наличия обязательных колонок на каждом листе в файле
#     :param path_to_checked_file: путь к проверяемому файлу
#     :param func_required_sheets_columns: словарь с данными для проверки формата {Лист:[Список обязательных колонок}
#     :param func_name_file:  имя проверяемого файла
#     :return: датафрейм с найденными ошибками
#     """
#     # датафрейм для сбора ошибок
#     check_error_req_columns_df = pd.DataFrame(columns=['Название файла', 'Название листа', 'Описание ошибки'])
#     for name_sheet, lst_req_cols in func_required_sheets_columns.items():
#         check_cols_df = pd.read_excel(path_to_checked_file, sheet_name=name_sheet)  # открываем файл
#         diff_cols = set(lst_req_cols).difference(set(check_cols_df.columns))  # ищем разницу в колонках
#         if len(diff_cols) != 0:
#             # Записываем ошибку
#             temp_error_df = pd.DataFrame(data=[[f'{func_name_file}', name_sheet,
#                                                 f'На листе не найдены указанные обязательные колонки: {";".join(diff_cols)}']],
#                                          columns=['Название файла', 'Название листа',
#                                                   'Описание ошибки'])
#             check_error_req_columns_df = pd.concat([check_error_req_columns_df, temp_error_df], axis=0,
#                                                    ignore_index=True)
#
#     if len(check_error_req_columns_df) != 0:
#         return check_error_req_columns_df
#
#
# def prepare_sheet_general_inf(path_to_file: str, name_sheet: str, lst_columns: list):
#     """
#     Функция для извлечения данных из листа Общие сведения
#     :param path_to_file: путь к файлу
#     :param name_sheet: название листа
#     :param lst_columns: список колонок из которых будут извлекаться данные
#     :return: подготовленный датафрейм
#     """
#     func_df = pd.read_excel(path_to_file, sheet_name=name_sheet, usecols=lst_columns)
#     func_df.dropna(how='all', inplace=True)  # удаляем пустые строки
#     # очищаем от пробельных символов в начале и конце каждой ячейки
#     func_df = func_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
#     person_df = func_df.copy()  # создаем датафрейм для сохранения всех строк
#     person_df['ФИО'] = person_df.iloc[0, 0]
#     person_df['Дата рождения'] = person_df.iloc[0, 1]
#     person_df['Дата начала работы в ПОО'] = person_df.iloc[0, 2]
#     person_df['Общий стаж работы'] = person_df.iloc[0, 4]
#     person_df['Педагогический стаж'] = person_df.iloc[0, 5]
#     person_df['Категория'] = person_df.iloc[0, 9]
#     person_df['№ приказа на аттестацию, дата'] = person_df.iloc[0, 10]
#     person_df['Наличие личного сайта, блога (ссылка)'] = person_df.iloc[0, 11]
#     dis_lst = func_df['Преподаваемая дисциплина'].tolist()  # список дисциплин преподавателя
#     # список полученных дипломов об образовании
#     educ_lst = func_df['Сведения об образовании (образовательное учреждение, квалификация, год окончания)'].tolist()
#     func_df = func_df.iloc[0, :].to_frame().transpose()
#     func_df.loc[0, 'Преподаваемая дисциплина'] = ';'.join(dis_lst)  # превращаем в строку список дисциплин
#     func_df.loc[0, 'Сведения об образовании (образовательное учреждение, квалификация, год окончания)'] = ';'.join(
#         educ_lst)  # превращаем в строку список дисциплин
#     return func_df, person_df
#
#
# def prepare_sheet_standart(path_to_file: str, name_sheet: str, lst_columns: list):
#     """
#     Функция для извлечения данных из листа Повышение квалификации
#     :param path_to_file: путь к файлу
#     :param name_sheet: название листа
#     :param lst_columns: список колонок из которых будут извлекаться данные
#     :return: подготовленный датафрейм
#     """
#     func_df = pd.read_excel(path_to_file, sheet_name=name_sheet, usecols=lst_columns)
#     func_df.dropna(how='all', inplace=True)  # удаляем пустые строки
#     # очищаем от пробельных символов в начале и конце каждой ячейки
#     func_df = func_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
#
#     return func_df
#
#
# def extract_last_date(value: str):
#     """
#     Функция для извлечения последней даты из ячейки
#     :param value:значение ячейки
#     :return: дата в виде строки или None
#     """
#     result_lst = re.findall(r'\d{2}.\d{2}.\d{4}', value)
#     if result_lst:
#         return result_lst[-1]
#     else:
#         return None
#
#
# def prepare_date(value):
#     """
#     Функция для конвертации дат с учетом возможных вариантов вида 01.01.2023-02.03.2023
#     :param value: значение
#     :return: дата в формате timestamp
#     """
#     try:
#         prep_date = pd.to_datetime(value, dayfirst=True, format='mixed')  # конвертируем в дату
#         return prep_date
#     except:
#         value_str = str(value)  # конвертируем в строку
#         prep_date = extract_last_date(value_str)  # ищем последнюю дату
#         if prep_date:
#             prep_date = datetime.datetime.strptime(prep_date, '%d.%m.%Y')  # конвертируем в дату
#             return prep_date
#         else:
#             return None
#
#
# def selection_by_date(df: pd.DataFrame, start_date: str, end_date: str, name_date_column, name_file: str,
#                       name_sheet: str,
#                       union_df: pd.DataFrame, error_df: pd.DataFrame, name_teacher: str):
#     """
#     Функция для отбора тех строк датафрейма что подходят под выбранные даты
#     :param df: датафрейм с данными
#     :param start_date: стартовая дата в формате строки 01.01.2024
#     :param end_date: конечная дата в формате строки 17.12.2024
#     :param name_date_column: список названий колонок с датами
#     :param name_file:  название файла
#     :param name_sheet: название листа с ошибкой
#     :param union_df: датафрейм для консолидации
#     :param error_df: общий датафрейм с ошибками
#     :param name_teacher: ФИО педагога
#     :return: отфильтрованный датафрейм
#     """
#     if name_sheet != 'Общие сведения':
#         # конвертируем даты в формат дат
#         start_date = pd.to_datetime(start_date, dayfirst=True)
#         end_date = pd.to_datetime(end_date, dayfirst=True)
#
#         df['_Отбор даты'] = df[name_date_column].apply(prepare_date)
#         date_error_df = df[df['_Отбор даты'].isnull()]  # отбираем строки с ошибками в датах
#         if len(date_error_df) != 0:
#             number_row_error = list(
#                 map(lambda x: str(x + 2), date_error_df.index))  # получаем номера строк с ошибками прибавляя 2
#             temp_error_df = pd.DataFrame(data=[[f'{name_file}', name_sheet,
#                                                 f'В колонке {name_date_column} в указанных строках неправильно записаны даты: {";".join(number_row_error)}. Требуемый формат: 21.05.2024'
#                                                 f' или 05.06.2024-15.08.2024']],
#                                          columns=['Название файла', 'Название листа',
#                                                   'Описание ошибки'])
#             error_df = pd.concat([error_df, temp_error_df], axis=0,
#                                  ignore_index=True)
#
#         df = df[df['_Отбор даты'].notnull()]  # отбираем строки с правильной датой
#
#         df = df[df['_Отбор даты'].between(start_date, end_date)]  # получаем значения в указанном диапазоне
#         df[name_date_column] = df[name_date_column].apply(
#             lambda x: x.strftime('%d.%m.%Y') if isinstance(x, (pd.Timestamp, datetime.datetime)) else x)
#         df.drop(columns=['_Отбор даты'], inplace=True)  # удаляем служебную колонку
#         df.insert(0, 'ФИО', name_teacher)
#
#     # Соединяем
#     union_df = pd.concat([union_df, df])
#
#     return union_df, df, error_df
#
#
# def create_report_teacher(data_folder: str, result_folder: str, start_date: str, end_date: str):
#     """
#     Функция для создания отчетности по преподавателям
#     :param data_folder: папка где хранятся личные дела преподавателей
#     :param result_folder: папка в которую будут сохраняться итоговые файлы
#     :param start_date: начальная дата, если ничего не указано то 01.01.1900
#     :param end_date: конечная дата, если ничего не указано то 01.01.2100
#     """
#     # обязательные листы
#     required_sheets_columns = {'Общие сведения': ['ФИО', 'Дата рождения', 'Дата начала работы в ПОО',
#                                                   'Преподаваемая дисциплина', 'Общий стаж работы',
#                                                   'Педагогический стаж',
#                                                   'Сведения об образовании (образовательное учреждение, квалификация, год окончания)',
#                                                   'Квалификация', 'Год окончания',
#                                                   'Категория', '№ приказа на аттестацию, дата',
#                                                   'Наличие личного сайта, блога (ссылка)'],
#                                'Повышение квалификации': ['Название программы повышения квалификации',
#                                                           'Вид повышения квалификации',
#                                                           'Место прохождения программы',
#                                                           'Дата прохождения программы (с какого по какое число, месяц, год)',
#                                                           'Количество академических часов',
#                                                           'Наименование подтверждающего документа, его номер и дата выдачи'],
#                                'Стажировка': ['Место стажировки', 'Кол-во часов', 'Дата'],
#                                'Методические разработки': ['Вид методического издания', 'Название издания',
#                                                            'Профессия/специальность ',
#                                                            'Дата разработки', 'Кем утверждена'],
#                                'Мероприятия, пров. ППС': ['Название мероприятия', 'Дата', 'Уровень мероприятия'],
#                                'Личное выступление ППС': ['Дата', 'Название мероприятия', 'Тема', 'Вид мероприятия',
#                                                           'Уровень мероприятия',
#                                                           'Способ участия', 'Результат участия'],
#                                'Публикации': ['Полное название статьи', 'Издание', 'Дата выпуска'],
#                                'Открытые уроки': ['Вид занятия','Дисциплина', 'Группа', 'Тема', 'Дата проведения'],
#                                'Взаимопосещение': ['ФИО посещенного педагога', 'Дата посещения', 'Группа', 'Тема'],
#                                'УИРС': ['ФИО обучающегося', 'Профессия/специальность', 'Группа', 'Вид мероприятия',
#                                         'Название мероприятия',
#                                         'Тема', 'Способ участия', 'Уровень мероприятия', 'Дата проведения',
#                                         'Результат участия'],
#                                'Работа по НМР': ['Тема НИРП', 'Проведено ли обобщение опыта', 'Форма обобщения опыта',
#                                                  'Место обобщения опыта', 'Дата обобщения опыта'],
#                                'Данные для списков': []
#                                }
#     error_df = pd.DataFrame(
#         columns=['Название файла', 'Название листа', 'Описание ошибки'])  # датафрейм для ошибок
#
#     teachers_dct = dict()  # словарь в котором будут храниться словари с данными листов для каждого преподавателя
#
#     # Создаем датафреймы для консолидации данных
#     general_inf_df = pd.DataFrame(columns=required_sheets_columns['Общие сведения'])
#     skills_dev_df = pd.DataFrame(columns=required_sheets_columns['Повышение квалификации'])
#     skills_dev_df.insert(0, 'ФИО', '')
#     internship_df = pd.DataFrame(columns=required_sheets_columns['Стажировка'])
#     internship_df.insert(0, 'ФИО', '')
#     method_dev_df = pd.DataFrame(columns=required_sheets_columns['Методические разработки'])
#     method_dev_df.insert(0, 'ФИО', '')
#     events_teacher_df = pd.DataFrame(columns=required_sheets_columns['Мероприятия, пров. ППС'])
#     events_teacher_df.insert(0, 'ФИО', '')
#     personal_perf_df = pd.DataFrame(columns=required_sheets_columns['Личное выступление ППС'])
#     personal_perf_df.insert(0, 'ФИО', '')
#     publications_df = pd.DataFrame(columns=required_sheets_columns['Публикации'])
#     publications_df.insert(0, 'ФИО', '')
#     open_lessons_df = pd.DataFrame(columns=required_sheets_columns['Открытые уроки'])
#     open_lessons_df.insert(0, 'ФИО', '')
#     mutual_visits_df = pd.DataFrame(columns=required_sheets_columns['Взаимопосещение'])
#     mutual_visits_df.insert(0, 'ФИО', '')
#     student_perf_df = pd.DataFrame(columns=required_sheets_columns['УИРС'])
#     student_perf_df.insert(0, 'ФИО', '')
#     nmr_df = pd.DataFrame(columns=required_sheets_columns['Работа по НМР'])
#     nmr_df.insert(0, 'ФИО', '')
#
#     for idx, file in enumerate(os.listdir(data_folder)):
#         if not file.startswith('~$') and file.endswith('.xlsx'):
#             name_file = file.split('.xlsx')[0]  # Получаем название файла
#             path_to_file = f'{data_folder}/{file}'
#             # Проверяем наличие нужных листов в файле
#             error_req_sheet_df = check_required_sheet_in_file(path_to_file, required_sheets_columns, name_file)
#             if error_req_sheet_df is not None:
#                 error_df = pd.concat([error_df, error_req_sheet_df], axis=0, ignore_index=True)
#                 continue
#             # Проверка наличия нужных колонок в листах
#             error_req_columns_sheet_df = check_required_columns_in_sheet(path_to_file, required_sheets_columns,
#                                                                          name_file)
#             if error_req_columns_sheet_df is not None:
#                 error_df = pd.concat([error_df, error_req_columns_sheet_df], axis=0, ignore_index=True)
#                 continue
#             print(name_file)
#             # Обрабатываем лист Общие сведения
#             teachers_dct[name_file] = {key: pd.DataFrame() for key in
#                                        required_sheets_columns.keys()}  # Создаем ключи для конкретного преподавателя
#             prep_general_inf_df, teachers_dct[name_file]['Общие сведения'] = prepare_sheet_general_inf(path_to_file,
#                                                                                                        'Общие сведения',
#                                                                                                        required_sheets_columns[
#                                                                                                            'Общие сведения'])
#             fio_teacher = teachers_dct[name_file]['Общие сведения'].iloc[
#                 0, 0]  # получаем ФИО преподавателя для добавления в
#             # Отбираем данные по датам
#             general_inf_df, _, error_df = selection_by_date(prep_general_inf_df, start_date, end_date,
#                                                             'Дата начала работы в ПОО',
#                                                             name_file, 'Общие сведения', general_inf_df, error_df,
#                                                             fio_teacher)
#             # Обрабатываем лист Повышение квалификации
#             prep_skills_dev_df = prepare_sheet_standart(path_to_file, 'Повышение квалификации',
#                                                         required_sheets_columns['Повышение квалификации'])
#             # сохраняем в датафрейм
#             skills_dev_df, teachers_dct[name_file]['Повышение квалификации'], error_df = selection_by_date(
#                 prep_skills_dev_df, start_date, end_date,
#                 'Дата прохождения программы (с какого по какое число, месяц, год)',
#                 name_file, 'Повышение квалификации', skills_dev_df, error_df, fio_teacher)
#             # Обрабатываем лист Стажировка
#             prep_internship_df = prepare_sheet_standart(path_to_file, 'Стажировка',
#                                                         required_sheets_columns['Стажировка'])
#             internship_df, teachers_dct[name_file]['Стажировка'], error_df = selection_by_date(prep_internship_df,
#                                                                                                start_date, end_date,
#                                                                                                'Дата',
#                                                                                                name_file, 'Стажировка',
#                                                                                                internship_df, error_df,
#                                                                                                fio_teacher)
#
#             # Обрабатываем лист Методические разработки
#             prep_method_dev_df = prepare_sheet_standart(path_to_file, 'Методические разработки',
#                                                         required_sheets_columns['Методические разработки'])
#             method_dev_df, teachers_dct[name_file]['Методические разработки'], error_df = selection_by_date(
#                 prep_method_dev_df, start_date, end_date, 'Дата разработки',
#                 name_file, 'Методические разработки', method_dev_df, error_df, fio_teacher)
#
#             # Обрабатываем лист Мероприятия, пров. ППС
#             prep_events_teacher_df = prepare_sheet_standart(path_to_file, 'Мероприятия, пров. ППС',
#                                                             required_sheets_columns['Мероприятия, пров. ППС'])
#             events_teacher_df, teachers_dct[name_file]['Мероприятия, пров. ППС'], error_df = selection_by_date(
#                 prep_events_teacher_df, start_date, end_date, 'Дата',
#                 name_file, 'Мероприятия, пров. ППС', events_teacher_df, error_df, fio_teacher)
#
#             # Обрабатываем лист Личное выступление ППС
#             prep_personal_perf_df = prepare_sheet_standart(path_to_file, 'Личное выступление ППС',
#                                                            required_sheets_columns['Личное выступление ППС'])
#             personal_perf_df, teachers_dct[name_file]['Личное выступление ППС'], error_df = selection_by_date(
#                 prep_personal_perf_df, start_date, end_date, 'Дата',
#                 name_file, 'Личное выступление ППС', personal_perf_df,
#                 error_df, fio_teacher)
#
#             # Обрабатываем лист Публикации
#             prep_publications_df = prepare_sheet_standart(path_to_file, 'Публикации',
#                                                           required_sheets_columns['Публикации'])
#             publications_df, teachers_dct[name_file]['Публикации'], error_df = selection_by_date(prep_publications_df,
#                                                                                                  start_date, end_date,
#                                                                                                  'Дата выпуска',
#                                                                                                  name_file,
#                                                                                                  'Публикации',
#                                                                                                  publications_df,
#                                                                                                  error_df, fio_teacher)
#             # Обрабатываем лис Открытые уроки
#             prep_open_lessons_df = prepare_sheet_standart(path_to_file, 'Открытые уроки',
#                                                           required_sheets_columns['Открытые уроки'])
#             open_lessons_df, teachers_dct[name_file]['Открытые уроки'], error_df = selection_by_date(
#                 prep_open_lessons_df, start_date, end_date, 'Дата проведения',
#                 name_file, 'Открытые уроки', open_lessons_df,
#                 error_df, fio_teacher)
#             # Обрабатываем лист Взаимопосещение
#             prep_mutual_visits_df = prepare_sheet_standart(path_to_file, 'Взаимопосещение',
#                                                            required_sheets_columns['Взаимопосещение'])
#             mutual_visits_df, teachers_dct[name_file]['Взаимопосещение'], error_df = selection_by_date(
#                 prep_mutual_visits_df, start_date, end_date, 'Дата посещения',
#                 name_file, 'Взаимопосещение', mutual_visits_df,
#                 error_df, fio_teacher)
#             # Обрабатываем лист УИРС
#             prep_student_perf_df = prepare_sheet_standart(path_to_file, 'УИРС', required_sheets_columns['УИРС'])
#             student_perf_df, teachers_dct[name_file]['УИРС'], error_df = selection_by_date(prep_student_perf_df,
#                                                                                            start_date, end_date,
#                                                                                            'Дата проведения',
#                                                                                            name_file, 'УИРС',
#                                                                                            student_perf_df,
#                                                                                            error_df, fio_teacher)
#             # Обрабатываем лист Работа по НМР
#             prep_nmr_df = prepare_sheet_standart(path_to_file, 'Работа по НМР',
#                                                  required_sheets_columns['Работа по НМР'])
#             nmr_df, teachers_dct[name_file]['Работа по НМР'], error_df = selection_by_date(prep_nmr_df, start_date,
#                                                                                            end_date,
#                                                                                            'Дата обобщения опыта',
#                                                                                            name_file, 'Работа по НМР',
#                                                                                            nmr_df,
#                                                                                            error_df, fio_teacher)
#
#     dct_df = {'Общие сведения': general_inf_df, 'Повышение квалификации': skills_dev_df,
#               'Стажировка': internship_df, 'Методические разработки': method_dev_df,
#               'Мероприятия, пров. ППС': events_teacher_df,
#               'Личное выступление ППС': personal_perf_df, 'Публикации': publications_df,
#               'Открытые уроки': open_lessons_df,
#               'Взаимопосещение': mutual_visits_df, 'УИРС': student_perf_df, 'Работа по НМР': nmr_df}
#     return {'Общий отчет': dct_df}


def generate_table_method_dev(df: pd.DataFrame):
    """
    Функция для генерации сложной таблицы методических разработок
    :param df: датафрейм
    :return: датафрейм
    """
    main_df = pd.DataFrame(columns=df.columns)  # создаем датафрейм куда будут добавляться данные
    main_df.insert(0, 'Номер', '')
    df.insert(0, 'Номер', 0)
    count = 1  # счетчик строк
    lst_type = ['методические рекомендации', 'методические разработки', 'учебное пособие', 'электронный курс',
                'рабочая тетрадь',
                'тест', 'видеоурок', 'профессиональная проба', 'иное']
    for type in lst_type:
        name_table = type.capitalize()  # получаем название промежуточной таблицы
        row_header = pd.DataFrame(columns=main_df.columns,
                                  data=[[name_table, '', '', '', '', '', '']])
        main_df = pd.concat([main_df, row_header], axis=0, ignore_index=True)
        temp_df = df[df['Вид'] == type]
        if len(temp_df) != 0:
            temp_df['Номер'] = range(count, count + len(temp_df))  # присваеваем номера строк
            main_df = pd.concat([main_df, temp_df], axis=0, ignore_index=True)
            count += len(temp_df)

    quantity_met = df.shape[0]  # количество метод изданий
    quantity_teacher = len(df['ФИО'].unique())  # количество преподавателей
    result_teacher = f'ИТОГО преподавателей-{quantity_teacher}'
    count_type = Counter(df['Вид'].tolist())
    result_str_met = f'ИТОГО изданий-{quantity_met}:\nметодические рекомендации-{count_type["методические рекомендации"]}\nметодические разработки-{count_type["методические разработки"]}\nучебное пособие-{count_type["учебное пособие"]}\nэлектронный курс-{count_type["электронный курс"]}\nрабочая тетрадь-{count_type["рабочая тетрадь"]}\nтест-{count_type["тест"]}\nвидеоурок-{count_type["видеоурок"]}\nпрофессиональная проба-{count_type["профессиональная проба"]}\nиное-{count_type["иное"]}\n'

    main_df.loc[-1] = ['', result_teacher, '', result_str_met, '', '', '']

    return main_df


def generate_table_events_teacher(df: pd.DataFrame):
    """
    Функция для генерации сложной таблицы общего отчета по мероприятияем проведенным ППС
    :param df:
    :return:
    """
    main_df = pd.DataFrame(columns=['ФИО', 'Внутр', 'Мун', 'Рег', 'Межрег'])
    df['Название'] = df['Название'] + ', ' + df['Дата']
    lst_teacher = df['ФИО'].unique()  # получаем список преподавателей
    for teacher in lst_teacher:
        temp_df = df[df['ФИО'] == teacher]
        # Внутренние мероприятия
        local_df = temp_df[temp_df['Уровень'] == 'внутренний']
        local_lst = local_df['Название'].tolist()
        local_str = ';\n'.join(local_lst)
        # Муниципальный уровень
        mun_df = temp_df[temp_df['Уровень'] == 'муниципальный']
        mun_lst = mun_df['Название'].tolist()
        mun_str = ';\n'.join(mun_lst)
        # Региональный уровень
        reg_df = temp_df[temp_df['Уровень'] == 'региональный']
        reg_lst = reg_df['Название'].tolist()
        reg_str = ';\n'.join(reg_lst)
        # Межрегиональный
        meg_reg_df = temp_df[temp_df['Уровень'] == 'межрегиональный']
        meg_reg_lst = meg_reg_df['Название'].tolist()
        meg_reg_str = ';\n'.join(meg_reg_lst)
        # Создаем строку для добавления в главный датафрейм
        row_df = pd.DataFrame(columns=main_df.columns,
                              data=[[teacher, local_str, mun_str, reg_str, meg_reg_str]])
        main_df = pd.concat([main_df, row_df])
    # Результирующая строка
    quantity_teacher = len(df['ФИО'].unique()) # количество педагогов
    count_type_event = Counter(df['Уровень'].tolist())
    result_str_teacher = f'ИТОГО преподавателей-{quantity_teacher}\nфедеральных-{count_type_event["федеральный"]}\nмеждународных-{count_type_event["международный"]}'
    main_df.loc[-1] = [result_str_teacher,f'{count_type_event["внутренний"]} внутренних',f'{count_type_event["муниципальный"]} муниципальных'
                                   ,f'{count_type_event["региональный"]} региональных',f'{count_type_event["межрегиональный"]} межрегиональных']

    return main_df

def generate_table_personal_perf(df:pd.DataFrame):
    """
    Функция для генерации сложной таблицы по личным выступлениям педагога
    :param df: датафрейм
    :return: датафрейм
    """
    main_df = pd.DataFrame(columns=df.columns)  # создаем датафрейм куда будут добавляться данные
    main_df.insert(0, 'Номер', '')
    df.insert(0, 'Номер', 0)
    df['Название'] = df['Название'] + '\nТема '+ df['Тема']
    count = 1  # счетчик строк
    lst_type = ['конкурс', 'научно-практическая конференция', 'олимпиада', 'иное']
    for type in lst_type:
        name_table = type.capitalize()  # получаем название промежуточной таблицы
        row_header = pd.DataFrame(columns=main_df.columns,
                                  data=[[name_table, '', '', '', '', '', '','','']])
        main_df = pd.concat([main_df, row_header], axis=0, ignore_index=True)
        temp_df = df[df['Вид'] == type]
        if len(temp_df) != 0:
            temp_df['Номер'] = range(count, count + len(temp_df))  # присваеваем номера строк
            main_df = pd.concat([main_df, temp_df], axis=0, ignore_index=True)
            count += len(temp_df)
        # Результирующая строка
        quantity_event = len(temp_df) #  количество конкурсов
        quantity_teacher = len(temp_df['ФИО'].unique()) # количество педагогов
        count_type_event = Counter(temp_df['Уровень']) # количество по уровням
        lst_type_event = [f'{key}-{value}' for key,value in count_type_event.items()]

        count_way_event = Counter(temp_df['Способ']) # количество по способам
        lst_way_event = [f'{key}-{value}' for key,value in count_way_event.items()]



        count_result = Counter(temp_df['Результат']) # количество по результатам
        result_str_result = f'1 место-{count_result["1 место"]}\n2 место-{count_result["2 место"]}\n3 место-{count_result["3 место"]}\nпобедитель номинации-{count_result["победитель номинации"]}\nноминация-{count_result["номинация"]}'
        row_itog = pd.DataFrame(columns=main_df.columns,
                                data=[['Итого','',f'мероприятий-{quantity_event}',
                                       f'преподавателей-{quantity_teacher}','','','\n'.join(lst_type_event),'\n'.join(lst_way_event),result_str_result]])
        main_df = pd.concat([main_df,row_itog])

    return main_df

def generate_table_student_perf(df:pd.DataFrame):
    """
    Функция для создания сложной таблицы по результатам обучающихся
    :param df: датафрейм
    :return: датафрейм
    """
    main_df = pd.DataFrame(columns=df.columns)  # создаем датафрейм куда будут добавляться данные
    main_df.insert(0, 'Номер', '')
    df.insert(0, 'Номер', 0)

    count = 1  # счетчик строк
    lst_type = ['конкурс', 'научно-практическая конференция', 'олимпиада', 'иное']
    for type in lst_type:
        name_table = type.capitalize()  # получаем название промежуточной таблицы
        row_header = pd.DataFrame(columns=main_df.columns,
                                  data=[[name_table, '', '', '', '', '', '','','','','','']])
        main_df = pd.concat([main_df, row_header], axis=0, ignore_index=True)
        temp_df = df[df['Вид'] == type]
        if len(temp_df) != 0:
            temp_df['Номер'] = range(count, count + len(temp_df))  # присваеваем номера строк
            temp_df['Профессия'] = temp_df['Профессия'] + ', ' + temp_df['Группа']
            temp_df['Название'] = temp_df['Название'] + '\nТема ' + temp_df['Тема']
            main_df = pd.concat([main_df, temp_df], axis=0, ignore_index=True)
            count += len(temp_df)
        # Результирующая строка
        all_quantity_student = len(temp_df) # количество участий студентов
        result_student = f'участий обучающихся-{all_quantity_student}\n'

        # Считаем преподавателей
        quantity_teacher = len(temp_df['ФИО'].unique())

        # считаем по уровням
        count_type = Counter(temp_df['Уровень'])
        result_str_type = f'Внутренние-{count_type["внутренний"]}\nМуниципальные-{count_type["муниципальный"]}\nРегиональные-{count_type["региональный"]}\nМежрегиональные-{count_type["межрегиональный"]}\nФедеральные-{count_type["федеральный"]}\nМеждународные-{count_type["международный"]}'
        # Считаем результат
        count_result = Counter(temp_df['Результат'])  # количество по результатам
        result_str_result = f'1 место-{count_result["1 место"]}\n2 место-{count_result["2 место"]}\n3 место-{count_result["3 место"]}\nпобедитель номинации-{count_result["победитель номинации"]}\nноминация-{count_result["номинация"]}'



        row_itog = pd.DataFrame(columns=main_df.columns,
                                data=[['ИТОГО',f'руководителей-{quantity_teacher}', result_student,'', '','','',
                                       '', '',result_str_type,'',result_str_result]])

        main_df = pd.concat([main_df, row_itog])
    return main_df

def generate_table_pmutual_visits(df:pd.DataFrame):
    """
    Функция для создания таблицы взаимопосещений уроков
    :param df: датафрейм
    :return: датафрейм
    """
    main_df = pd.DataFrame(columns=['ФИО','ФИО_посещенного']) # собирающий датафрейм
    lst_teacher = df['ФИО'].unique() # список преподавателей
    for teacher in lst_teacher:
        temp_df = df[df['ФИО'] == teacher]
        lst_visitors = temp_df['ФИО_посещенного'].tolist() # получаем список всех посещенных преподавателей
        row_df = pd.DataFrame(columns=['ФИО','ФИО_посещенного'],
                              data=[[teacher,'\n'.join(lst_visitors)]])
        main_df = pd.concat([main_df,row_df])

    quantity_teacher = len(df['ФИО'].unique())
    quantity_lesson = len(df)
    row_itog = pd.DataFrame(columns=['ФИО', 'ФИО_посещенного'],
                          data=[[f'Преподавателей-{quantity_teacher}', f'Занятий-{quantity_lesson}']])
    main_df = pd.concat([main_df,row_itog])
    return main_df



def generate_context(dct_value:dict):
    """
    Функция для создания словаря контекста для записи в файл docx
    :param dct_value: словарь с данными
    :return: подготовленный словарь для записи в файл
    """
    context = dict()  # словарь
    # Создаем переменные для листов
    first_sheet_df = dct_value['Общие сведения']
    # Переименовываем колонки для удобства
    first_sheet_df.columns = ['ФИО', 'Дата_рождения', 'Дата_ПОО', 'Дисциплина', 'Стаж', 'Педстаж',
                              'Организация', 'Квалификация', 'Год_окончания', 'Категория', 'Приказ', 'Сайт']
    first_sheet_df[['Стаж', 'Педстаж']] = first_sheet_df[['Стаж', 'Педстаж']].fillna(0)  # заменяем нан нулями
    first_sheet_df[['Стаж', 'Педстаж']] = first_sheet_df[['Стаж', 'Педстаж']].applymap(
        lambda x: int(x) if isinstance(x, (int, float)) else x)
    first_sheet_df.fillna('', inplace=True)

    context['Преподаватель'] = first_sheet_df.iloc[0, 0]  # ФИО преподавателя
    context['Общая_информация'] = first_sheet_df[['Дисциплина', 'Дата_рождения', 'Дата_ПОО', 'Стаж', 'Педстаж',
                                                  'Категория', 'Приказ', 'Сайт']].to_dict('records')
    context['Образование'] = first_sheet_df[['Организация', 'Квалификация', 'Год_окончания']].to_dict('records')

    # данные с листа Повышение квалификации
    skills_dev_df = dct_value['Повышение квалификации']
    skills_dev_df.columns = ['ФИО', 'Название', 'Вид', 'Место', 'Дата', 'Часов', 'Документ']
    # Приводим колонку
    skills_dev_df['Часов'] = skills_dev_df['Часов'].fillna(0)
    skills_dev_df['Часов'] = skills_dev_df['Часов'].apply(lambda x: int(x) if isinstance(x, (int, float)) else x)
    skills_dev_df.fillna('', inplace=True)
    context['Доп_образование'] = skills_dev_df.to_dict('records')
    # Создаем датафрейм для общего отчета с результирующей строкой
    itog_skills_dev_df = skills_dev_df.copy()  # создаем новый датафрейм
    quantity_teacher = len(
        itog_skills_dev_df['ФИО'].unique())  # получаем количество уникальных преподавателей прошедших ПК
    quantity_course = itog_skills_dev_df.shape[0]  # общее количество курсов
    count_type_course = Counter(itog_skills_dev_df['Вид'].tolist())
    # Результирующая строка по преподавателям и количеству курсов
    result_str_teacher = f'ИТОГО:\nпреподавателей-{quantity_teacher}\nповышений квалификаций-{quantity_course} '
    quantity_kpk = count_type_course['курс повышения квалификации']
    quantity_pp = count_type_course['профессиональная переподготовка']
    quantity_other = count_type_course['иное']
    result_str_course = f'ИТОГО:\nКПК-{quantity_kpk}\nкурсов переподготовки-{quantity_pp}\nИное-{quantity_other}'

    itog_skills_dev_df.loc[-1] = [result_str_teacher, result_str_course, '', '', '', '',
                                                           '']
    context['Доп_образование_итог'] = itog_skills_dev_df.to_dict('records')

    internship_df = dct_value['Стажировка']
    internship_df.columns = ['ФИО', 'Место', 'Часов', 'Дата']
    internship_df['Часов'] = internship_df['Часов'].fillna(0)
    internship_df['Часов'] = internship_df['Часов'].apply(lambda x: int(x) if isinstance(x, (int, float)) else x)
    internship_df.fillna('', inplace=True)
    context['Стажировка'] = internship_df.to_dict('records')
    itog_internship_df = internship_df.copy()
    quantity_teacher = len(
        itog_internship_df['ФИО'].unique())  # получаем количество уникальных преподавателей прошедших стажировку
    quantity_internship = itog_internship_df.shape[0]  # общее количество стажировок
    result_str_internship = (f'ИТОГО:\n'
                             f'стажировок-{quantity_internship}\nпедагогов-{quantity_teacher}')
    itog_internship_df.loc[-1] = [result_str_internship, '', '', '']
    context['Стажировка_итог'] = itog_internship_df.to_dict('records')

    method_dev_df = dct_value['Методические разработки']
    method_dev_df.columns = ['ФИО', 'Вид', 'Название', 'Профессия', 'Дата', 'Утверждено']
    method_dev_df.fillna('', inplace=True)
    context['Метод_разработки'] = method_dev_df.to_dict('records')

    itog_method_dev_df = method_dev_df.copy()
    context['Метод_разработки_итог'] = generate_table_method_dev(itog_method_dev_df).to_dict(
        'records')  # генерируем таблицу

    events_teacher_df = dct_value['Мероприятия, пров. ППС']
    events_teacher_df.columns = ['ФИО', 'Название', 'Дата', 'Уровень']
    events_teacher_df.fillna('', inplace=True)
    context['Пров_мероприятия'] = events_teacher_df.to_dict('records')

    itog_events_teacher_df = events_teacher_df.copy()
    context['Пров_мероприятия_итог'] = generate_table_events_teacher(itog_events_teacher_df).to_dict(
        'records')  # генерируем таблицу

    personal_perf_df = dct_value['Личное выступление ППС']
    personal_perf_df.columns = ['ФИО', 'Дата', 'Название', 'Тема', 'Вид', 'Уровень', 'Способ', 'Результат']
    personal_perf_df.fillna('', inplace=True)
    context['Выступления'] = personal_perf_df.to_dict('records')

    itog_personal_perf_df = personal_perf_df.copy()
    context['Выступления_итог'] = generate_table_personal_perf(itog_personal_perf_df).to_dict(
        'records')  # генерируем таблицу

    publications_df = dct_value['Публикации']
    publications_df.columns = ['ФИО', 'Название', 'Издание', 'Дата']
    publications_df.fillna('', inplace=True)
    context['Публикации'] = publications_df.to_dict('records')

    itog_publications_df = publications_df.copy()
    quantity_teacher = len(
        itog_publications_df['ФИО'].unique())  # получаем количество уникальных преподавателей прошедших стажировку
    quantity_publications = itog_publications_df.shape[0]  # общее количество стажировок
    result_str_publications = (f'ИТОГО:\n'
                               f'публикаций-{quantity_publications}\nпедагогов-{quantity_teacher}')

    itog_publications_df.loc[-1] = ['', result_str_publications, '', '']
    context['Публикации_итог'] = itog_publications_df.to_dict('records')

    open_lessons_df = dct_value['Открытые уроки']
    open_lessons_df.columns = ['ФИО', 'Вид', 'Дисциплина', 'Группа', 'Тема', 'Дата']
    open_lessons_df.fillna('', inplace=True)
    context['Открытые_уроки'] = open_lessons_df.to_dict('records')

    itog_open_lessons_df = open_lessons_df.copy()
    quantity_course = len(
        itog_open_lessons_df['Дисциплина'].unique())  # получаем количество уникальных дисциплин

    quantity_open_lessons = itog_open_lessons_df.shape[0]  # общее количество открытых уроков

    count_type_open_lessons = Counter(itog_open_lessons_df['Вид'].tolist())
    lst_type_open = [f'{key}-{value}' for key, value in count_type_open_lessons.items()]
    result_str_open_lessons = f'ИТОГО:\nоткрытых уроков-{quantity_open_lessons}\nдисциплин-{quantity_course}\nПо типам занятий:\n'
    type_str = '\n'.join(lst_type_open)
    itog_open_lessons_df.loc[-1] = ['', '', result_str_open_lessons + type_str, '', '',
                                                               '']
    context['Открытые_уроки_итог'] = itog_open_lessons_df.to_dict('records')

    mutual_visits_df = dct_value['Взаимопосещение']
    mutual_visits_df.columns = ['ФИО', 'ФИО_посещенного', 'Дата', 'Группа', 'Тема']
    mutual_visits_df.fillna('', inplace=True)
    context['Взаимопосещение'] = mutual_visits_df.to_dict('records')

    itog_mutual_visits_df = mutual_visits_df.copy()
    context['Взаимопосещение_итог'] = generate_table_pmutual_visits(itog_mutual_visits_df).to_dict(
        'records')  # генерируем таблицу

    student_perf_df = dct_value['УИРС']
    student_perf_df.columns = ['ФИО', 'ФИО_студента', 'Профессия', 'Группа', 'Вид', 'Название', 'Тема', 'Способ',
                               'Уровень', 'Дата', 'Результат']
    student_perf_df.fillna('', inplace=True)
    context['УИРС'] = student_perf_df.to_dict('records')

    itog_student_perf_df = student_perf_df.copy()
    context['УИРС_итог'] = generate_table_student_perf(itog_student_perf_df).to_dict(
        'records')  # генерируем таблицу

    nmr_df = dct_value['Работа по НМР']
    nmr_df.columns = ['ФИО', 'Тема', 'Обобщение', 'Форма', 'Место', 'Дата']
    nmr_df.fillna('', inplace=True)
    context['НМР'] = nmr_df.to_dict('records')

    itog_nmr_df = nmr_df.copy()
    quantity_teacher = len(itog_nmr_df['ФИО'].unique())
    itog_nmr_df.loc[-1] = [f'ИТОГО преподавателей-{quantity_teacher}', '', '', '', '', '']
    context['НМР_итог'] = itog_nmr_df.to_dict('records')

    return context



def generate_docs(master_dct: dict, template_folder: str, result_folder: str):
    """
    Функция для генерации документов
    :param master_dct: словарь содержащий данные в формате {Ключ:{Ключ:датафрейм}}
    :param template_folder: папка с шаблонами документов docx
    :param result_folder:куда сохранять результат
    """
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    for name_dct, dct_value in master_dct.items():
        if name_dct == 'Личные дела':
            for teacher,dct_data in dct_value.items():
                context = generate_context(dct_data)
                for file in os.listdir(template_folder):
                    if file.endswith('.docx') and not file.startswith('~$'):  # получаем только файлы docx и не временные
                        name_file = file.split('.docx')[0]
                        name_file = re.sub(r'[Шш]аблон\s*','',name_file)
                        doc = DocxTemplate(f'{template_folder}/{file}')

                        doc.render(context)
                        teacher = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', teacher)  # очищаем от некорректных символов

                        path_folder_teacher = f'{result_folder}/{teacher}'
                        # Создаем папки
                        if not os.path.exists(path_folder_teacher):
                            os.makedirs(path_folder_teacher)


                        if 'отчет' in file.lower():
                            doc.save(f'{path_folder_teacher}/{name_file}.docx')
                        else:
                            doc.save(f'{path_folder_teacher}/Личное дело {teacher[:40]}.docx')
        else:
            context = generate_context(dct_value)
            for file in os.listdir(template_folder):
                if file.endswith('.docx') and not file.startswith('~$'):  # получаем только файлы docx и не временные
                    if 'отчет' in file.lower():
                        name_file = file.split('.docx')[0]
                        name_file = re.sub(r'[Шш]аблон\s*','',name_file)
                        doc = DocxTemplate(f'{template_folder}/{file}')
                        doc.render(context)
                        doc.save(f'{result_folder}/{name_file} {current_time}.docx')



if __name__ == '__main__':
    main_data_folder = 'data/Данные'
    main_result_folder = 'data/Результат'
    main_start_date = '06.06.1900'
    main_end_date = '01.05.2100'
    main_template = 'data/Шаблоны'

    main_dct = create_report_teacher(main_data_folder, main_result_folder, main_start_date, main_end_date)
    generate_docs(main_dct, main_template, main_result_folder)

    print('Lindy Booth')

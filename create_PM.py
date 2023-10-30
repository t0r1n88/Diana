"""
скрипт для создания программ профессиональныъ модулей
"""
import pandas as pd
import numpy as np
import openpyxl
from docxtpl import DocxTemplate
import string
import time
import re
from tkinter import messagebox

pd.options.mode.chained_assignment = None  # default='warn'
pd.set_option('display.max_columns', None)  # Отображать все столбцы
pd.set_option('display.expand_frame_repr', False)  # Не переносить строки
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=FutureWarning, module='openpyxl')


def convert_to_int(cell):
    """
    Метод для проверки значения ячейки
    :param cell: значение ячейки
    :return: число в формате int
    """
    if cell is np.nan:
        return ''
    try:
        value = float(cell)
        return int(value)
    except:
        return cell


def extract_lr(cell):
    """
    Функция для создания 2 списков из данных одной колонки
    первая колонка это код личностного результата а вторая колонка  это описание
    извлечение будет с помощью регулярок
    :param df: датафрем из одной колонки
    :return:
    """
    value = str(cell)  # делаем строковй
    result = re.split(r'(\d+\.\s*?)', value)
    result = [value for value in result if value]  # убираем пустые значения
    if len(result) >= 3:
        end_lr = result[1].strip()
        end_lr = end_lr.rstrip(string.punctuation)  # очищаем от точки в конце
        return f'{result[0]}{end_lr}'
    else:
        lr = result[0].strip()
        lr = lr.rstrip(string.punctuation)
        return lr


def extract_descr_lr(cell):
    """
    Функция для создания 2 списков из данных одной колонки
    первая колонка это код личностного результата а вторая колонка  это описание
    извлечение будет с помощью регулярок
    :param df: датафрем из одной колонки
    :return:
    """
    value = str(cell)  # делаем строковй
    result = re.split(r'(\d+\.\s*?)', value)
    result = [value for value in result if value]
    if len(result) >= 3:
        descr_lr = result[2].strip()
        descr_lr = descr_lr.rstrip(string.punctuation)  # очищаем от точки в конце
        return f'{descr_lr}.'
    else:
        return ''


def processing_punctuation_end_string(lst_phrase: list, sep_string: str, sep_begin: str, sep_end: str) -> str:
    """
    Очистка каждого элемента списка от знаков пунктации в конце,
    добавление разделитея, добавление точки в конце
    :param lst_phrase: список элементов
    :param sep_string: разделитель между элементами списка
    :param sep_begin: начальный разделитель
    :param sep_end: знак пунктуации в конце
    :return: строку с разделителями и переносами строки
    """
    if len(lst_phrase) == 0:
        return ''
    lst_phrase = list(map(str, lst_phrase))
    temp_lst = list(map(lambda x: sep_begin + x, lst_phrase))
    temp_lst = list(map(lambda x: x.strip(), temp_lst))  # очищаем от прбельных символов в начале и конце
    temp_lst = list(map(lambda x: x.rstrip(string.punctuation), temp_lst))  # очищаем от знаков пунктуации
    temp_lst[-1] = temp_lst[-1] + sep_end  # добавляем конечный знак пунктуации
    temp_str = f'{sep_string}'.join(temp_lst)  # создаем строку с разделителями
    return temp_str


def insert_type_source(df: pd.DataFrame) -> list:
    """
    Вставка в строку слов [Электронный ресурс] Форма доступа:
    :param lst_phrase:датафрейм
    :return: список измененных строк
    """
    out_lst = []  # список для хранения строк
    for row in df.itertuples():
        name = str(row[1])
        url_ii = str(row[2])
        name = name.strip()  # очищаем от пробельных символов
        name = name.rstrip(string.punctuation)  # очищаем от знаков препинания
        name = name.strip()
        url_ii = url_ii.strip()  # очищаем от пробельных символов
        url_ii = url_ii.rstrip(string.punctuation)  # очищаем от знаков препинания
        url_ii = url_ii.strip()
        temp_str = f'{name} [Электронный ресурс] Форма доступа:{url_ii}'
        out_lst.append(temp_str)

    return out_lst


def processing_publ(row):
    """
    Функция для генерации строки с литературой в нужном формате
    :param row: строка датафрейма
    :return: сумма строк датафрейма в нужном формате
    """
    author = row[0]  # автор(ы)
    name_book = row[1]  # название
    full_city = row[2]  # полное название города
    short_city = row[3]  # краткое название города
    publ_house = row[4]  # издательство
    year = row[5]  # год издания
    quan_pages = row[6]  # число страниц
    author = author.rstrip(string.punctuation)  # очищаем от символа пунктуации в конце
    name_book = name_book.rstrip(string.punctuation)  # очищаем от символа пунктуации в конце
    short_city = short_city.rstrip(string.punctuation)  # очищаем от символа пунктуации в конце
    publ_house = publ_house.rstrip(string.punctuation)  # очищаем от символа пунктуации в конце
    # извлекаем год
    result = re.search(r'\d{4}', year)
    if result:
        clean_year = result.group()
    else:
        clean_year = 'Неправильно заполнен год издания, введите год в формате 4 цифры без букв'

    # извлекаем количество страниц
    result = re.search(r'\d+', quan_pages)
    if result:
        clean_quan_pages = result.group()
    else:
        clean_quan_pages = 'Неправильно заполнено количество страниц, введите количество в виде числа без букв'
    # Формируем итоговую строку
    out_str = f'{author}. {name_book}.- {short_city}.: {publ_house}, {clean_year}.- {clean_quan_pages} c.'

    return out_str


def extract_data_mdk(data_pm,sheet_name)->pd.DataFrame:
    """
    Функция для получения датафрейма из листа файла
    :param data_pm: путь к файлу
    :param sheet_name: имя листа
    :return: датафрейм
    """

    df_plan_pm = pd.read_excel(data_pm, sheet_name=sheet_name, usecols='A:H')
    df_plan_pm.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    df_plan_pm.columns = ['Курс_семестр', 'Раздел', 'Тема', 'Содержание', 'Количество_часов', 'Практика', 'Вид_занятия',
                          'СРС']
    df_plan_pm['Курс_семестр'].fillna('Пусто', inplace=True)
    df_plan_pm['Раздел'].fillna('Пусто', inplace=True)
    df_plan_pm['Тема'].fillna('Пусто', inplace=True)

    borders = df_plan_pm[
        df_plan_pm['Курс_семестр'].str.contains('семестр')].index  # получаем индексы строк где есть слово семестр

    part_df = []  # список для хранения кусков датафрейма
    previos_border = -1
    # делим датафрем по границам
    for value_border in borders:
        part = df_plan_pm.iloc[previos_border:value_border]
        part_df.append(part)
        previos_border = value_border

    # добавляем последнюю часть
    last_part = df_plan_pm.iloc[borders[-1]:]
    part_df.append(last_part)

    part_df.pop(0)  # удаляем нулевой элемент так как он пустой

    main_df = pd.DataFrame(
        columns=['Курс_семестр', 'Раздел', 'Тема', 'Содержание', 'Количество_часов', 'Практика', 'Вид_занятия',
                 'СРС'])  # создаем базовый датафрейм

    lst_type_lesson = ['урок', 'практическое занятие', 'лабораторное занятие',
                       'курсовая работа (КП)']  # список типов занятий
    for df in part_df:
        dct_sum_result = {key: 0 for key in lst_type_lesson}  # создаем словарь для подсчета значений
        for type_lesson in lst_type_lesson:
            _df = df[df['Вид_занятия'] == type_lesson]  # фильтруем датафрейм
            _df['Количество_часов'].fillna(0, inplace=True)
            _df['Количество_часов'] = _df['Количество_часов'].astype(int)
            dct_sum_result[type_lesson] = _df['Количество_часов'].sum()
        # создаем строку с описанием
        margint_text = 'Итого часов за семестр:\nиз них\nтеория\nпрактические занятия\nлабораторные занятия\nкурсовая работа (КП)'

        all_hours = sum(dct_sum_result.values())  # общая сумма часов

        theory_hours = dct_sum_result['урок']  # часы теории
        praktice_hours = dct_sum_result['практическое занятие']  # часы практики
        lab_hours = dct_sum_result['лабораторное занятие']  # часы лабораторных
        kurs_hours = dct_sum_result['курсовая работа (КП)']  # часы курсовых

        value_text = f'{all_hours}\n \n{theory_hours}\n{praktice_hours}\n{lab_hours}\n{kurs_hours}'  # строка со значениями
        temp_df = pd.DataFrame([{'Тема': margint_text, 'Количество_часов': value_text}])
        df = pd.concat([df, temp_df], ignore_index=True)  # добаляем итоговую строку
        main_df = pd.concat([main_df, df], ignore_index=True)  # добавляем в основной датафрейм

    main_df.insert(0, 'Номер', np.nan)  # добавляем колонку с номерами занятий

    main_df['Содержание'] = main_df['Содержание'].fillna('Пусто')  # заменяем наны на пусто

    count = 0  # счетчик
    for idx, row in enumerate(main_df.itertuples()):
        if (row[5] == 'Пусто') | ('Итого часов' in row[5]):
            main_df.iloc[idx, 0] = ''
        else:
            count += 1
            main_df.iloc[idx, 0] = count

    # очищаем от пустых символов и строки Пусто
    main_df['Курс_семестр'] = main_df['Курс_семестр'].fillna('Пусто')
    main_df['Раздел'] = main_df['Раздел'].fillna('Пусто')

    main_df['Курс_семестр'] = main_df['Курс_семестр'].replace('Пусто', '')
    main_df['Тема'] = main_df['Тема'].replace('Пусто', '')
    main_df['Раздел'] = main_df['Раздел'].replace('Пусто', '')
    main_df['Содержание'] = main_df['Содержание'].replace('Пусто', '')

    main_df['Вид_занятия'] = main_df['Вид_занятия'].fillna('')

    main_df['Количество_часов'] = main_df['Количество_часов'].apply(convert_to_int)
    main_df['Количество_часов'] = main_df['Количество_часов'].fillna('')
    main_df['Практика'] = main_df['Практика'].fillna(0)
    main_df['Практика'] = main_df['Практика'].astype(int, errors='ignore')
    main_df['Практика'] = main_df['Практика'].apply(lambda x: '' if x == 0 else x)

    main_df['СРС'] = main_df['СРС'].fillna(0)
    main_df['СРС'] = main_df['СРС'].astype(int, errors='ignore')
    main_df['СРС'] = main_df['СРС'].apply(lambda x: '' if x == 0 else x)
    main_df['Содержание'] = main_df['Курс_семестр'] + main_df['Раздел'] + main_df['Тема'] + main_df['Содержание']
    main_df.drop(columns=['Курс_семестр', 'Раздел', 'Тема'], inplace=True)

    return main_df




def processing_mdk(data_pm) -> list:
    """
    Функция для обработки листов с названием МДК
    :param data_pm: путь к файлу
    :return: список датафреймов
    """
    # Получаем список листов

    dct_mdk = dict()
    wb = openpyxl.load_workbook(data_pm,read_only=True)
    for sheet_name in wb.sheetnames:
        if 'МДК' in sheet_name:
            name_mdk = str(wb[sheet_name]['D1'].value) # получаем название МДК, делаем строковыми на случай нан
            if 'МДК' in name_mdk:
                temp_mdk_df = extract_data_mdk(data_pm,sheet_name) # извлекаем данные из датафрейма
                dct_mdk[name_mdk] = temp_mdk_df
    wb.close()
    print(dct_mdk['МДК 02.03 Организация и контроль безопасности на железнодорожном транспорте и в пунктах прибытия (отправления) поездов'])






def create_pm(template_pm: str, data_pm: str, end_folder: str):
    """
    Функция для создания программ профессиональных модулей
    :param template_pm: шаблон профмодуля
    :param data_pm: таблица Excel
    :param end_folder: конечная папка
    :return:
    """
    lst_mdk_df = processing_mdk(data_pm)


if __name__ == '__main__':
    template_pm_main = 'data/Шаблон автозаполнения ПМ.docx'
    data_pm_main = 'data/Пример заполнения ПМ.xlsx'
    end_folder_main = 'data'

    create_pm(template_pm_main, data_pm_main, end_folder_main)

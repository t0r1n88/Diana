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

class ControlWord_PM(Exception):
    """
    Исключение для случая когда на листе Контроль шаблона ПМ пропущены слова Умения: ,Знания:, Практический опыт:
    """
    pass



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


def sum_column_any_value(df:pd.DataFrame,name_column:str):
    """
    Суммирование колонки с разными типами значений в том числе строковыми.
    """
    lst_value = df[name_column].dropna().tolist()
    sum_value = [value for value in lst_value if isinstance(value,(int,float))] # отбираем только числа
    return sum(sum_value) # возвращаем сумму

def extract_data_mdk(data_pm,sheet_name):
    """
    Функция для получения датафрейма из листа файла
    :param data_pm: путь к файлу
    :param sheet_name: имя листа
    :return: датафрейм
    """
    lst_type_lesson = ['урок', 'практическое занятие', 'лабораторное занятие',
                       'курсовая работа (КП)']  # список типов занятий
    dct_all_sum_result = {key: 0 for key in lst_type_lesson}  # создаем словарь для подсчета значений



    df_plan_pm = pd.read_excel(data_pm,sheet_name=sheet_name,skiprows=1, usecols='A:H')
    df_plan_pm.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    df_plan_pm.columns = ['Курс_семестр', 'Раздел', 'Тема', 'Содержание', 'Количество_часов', 'Практика', 'Вид_занятия',
                          'СРС']
    df_plan_pm['Курс_семестр'].fillna('Пусто', inplace=True)
    df_plan_pm['Раздел'].fillna('Пусто', inplace=True)
    df_plan_pm['Тема'].fillna('Пусто', inplace=True)

    # Считаем общие суммы
    mdk_all_sum = int(sum_column_any_value(df_plan_pm, 'Количество_часов'))  # получаем сумму общие часы
    mdk_all_prac_sum = int(sum_column_any_value(df_plan_pm, 'Практика'))  # получаем сумму общие часы
    mdk_all_srs_sum = int(sum_column_any_value(df_plan_pm, 'СРС'))  # сумма срс
    for type_lesson in lst_type_lesson:
        _df = df_plan_pm[df_plan_pm['Вид_занятия'] == type_lesson]  # фильтруем датафрейм
        dct_all_sum_result[type_lesson] = int(sum_column_any_value(_df, 'Количество_часов'))  # получаем значение

    dct_all_sum_result['Всего часов'] = mdk_all_sum
    dct_all_sum_result['Всего практики'] = mdk_all_prac_sum
    dct_all_sum_result['Всего СРС'] = mdk_all_srs_sum


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

    return (main_df,dct_all_sum_result) # возвращаем кортеж




def processing_mdk(data_pm) -> dict:
    """
    Функция для обработки листов с названием МДК
    :param data_pm: путь к файлу
    :return: список датафреймов
    """
    # Получаем список листов

    dct_mdk = dict()
    wb = openpyxl.load_workbook(data_pm,read_only=True) # получаем названия листов содержащих МДК
    for sheet_name in wb.sheetnames:
        if 'МДК' in sheet_name:
            name_mdk = str(wb[sheet_name]['D1'].value) # получаем название МДК, делаем строковыми на случай нан
            if 'МДК' in name_mdk:
                temp_mdk_df,temp_dct_result = extract_data_mdk(data_pm,sheet_name) # извлекаем данные из датафрейма
                dct_mdk[name_mdk] = {'Итог':temp_dct_result,'Данные':temp_mdk_df}
    wb.close()
    return dct_mdk





def create_pm(template_pm: str, data_pm: str, end_folder: str):
    """
    Функция для создания программ профессиональных модулей
    :param template_pm: шаблон профмодуля
    :param data_pm: таблица Excel
    :param end_folder: конечная папка
    :return:
    """
    # названия листов
    desc_rp = 'Описание ПМ'
    pers_result = 'Лич_результаты'
    volume_pm = 'Объем ПМ'
    volume_all_mdk = 'Объем МДК'
    kp = 'Тематика КП(КР)'
    up = 'УП'
    pp = 'ПП'
    mto = 'МТО'
    main_publ = 'ОИ'
    second_publ = 'ДИ'
    ii_publ = 'ИИ'
    control = 'Контроль'
    pk_ok_vpd = 'ПК,ОК,ВПД'
    fgos = 'Данные ФГОС'

    # Обрабатываем лист Описание ПМ
    df_desc_rp = pd.read_excel(data_pm, sheet_name=desc_rp, nrows=1, usecols='A:J')  # загружаем датафрейм
    df_desc_rp.fillna('НЕ ЗАПОЛНЕНО !!!', inplace=True)  # заполняем не заполненные разделы
    df_desc_rp.columns = ['Тип_программы', 'Название_модуля', 'Цикл', 'Перечень', 'Код_специальность_профессия',
                          'Год', 'Разработчик', 'Методист', 'Название_ПЦК', 'Пред_ПЦК']

    # Создаем переменные для ФИО утверждающих
    df_accept_fio = pd.read_excel(data_pm, sheet_name=desc_rp, nrows=2, usecols='K:L')  # загружаем датафрейм
    accept_UR = df_accept_fio.iloc[0,1]
    accept_PR = df_accept_fio.iloc[1,1]


    # Обрабатываем лист Лич_результаты

    df_pers_result = pd.read_excel(data_pm, sheet_name=pers_result, usecols='A')
    df_pers_result.dropna(inplace=True)  # удаляем пустые строки
    df_pers_result.columns = ['Описание']
    df_pers_result['Код'] = df_pers_result['Описание'].apply(extract_lr)
    df_pers_result['Результат'] = df_pers_result['Описание'].apply(extract_descr_lr)

    # # Обрабатываем лист Объем ПМ
    df_volume_pm = pd.read_excel(data_pm,sheet_name=volume_pm,usecols='A:B')
    df_volume_pm.dropna(inplace=True) # удаляем пустые строки
    df_volume_pm.columns = ['Наименование', 'Объем']


    """
    Обрабатываем лист Объем МДК
    """
    df_volume_all_mdk = pd.read_excel(data_pm,sheet_name=volume_all_mdk,usecols='A:J')
    df_volume_all_mdk.columns = ['Наименование','Всего','Прак_под','Обяз','Прак_зан','КР','СРС','УП','ПП','КА']
    df_volume_all_mdk.dropna(inplace=True,thresh=1) # удаляем пустые строки
    df_volume_all_mdk.fillna(0,inplace=True)  # заполняем наны
    df_volume_all_mdk.iloc[:,1:] = df_volume_all_mdk.iloc[:,1:].applymap(lambda x: int(x) if isinstance(x,(int,float)) else 0) # приводим к инту
    sum_row = df_volume_all_mdk.sum() # получаем строку общей суммы
    df_volume_all_mdk.loc['Сумма'] = sum_row # добавляем строку в датафрейм
    df_volume_all_mdk.at['Сумма','Наименование'] = 'Итого'
    df_volume_all_mdk = df_volume_all_mdk.astype(int,errors='ignore') # делай интовыми
    df_volume_all_mdk = df_volume_all_mdk.applymap(lambda x:'' if x ==0 else x)








    # # Открываем файл
    # wb = openpyxl.load_workbook(data_pm, read_only=True)
    # target_value = 'итог'
    #
    # # Поиск значения в выбранном столбце
    # column_number = 1  # Номер столбца, в котором ищем значение (например, столбец A)
    # target_row = None  # Номер строки с искомым значением
    #
    # for row in wb['Объем УД'].iter_rows(min_row=1, min_col=column_number, max_col=column_number):
    #     cell_value = row[0].value
    #     if target_value in str(cell_value).lower():
    #         target_row = row[0].row
    #         break
    #
    # if not target_row:
    #     # если не находим строку в которой есть слово итог то выдаем исключение
    #     messagebox.showerror('Диана Создание рабочих программ',
    #                          'Не найдена строка с названием Итоговая аттестация в листе Объем УД')
    #
    # wb.close()  # закрываем файл
    #
    # # если значение найдено то считываем нужное количество строк и  7 колонок
    # df_structure = pd.read_excel(data_pm, sheet_name=structure, nrows=target_row,
    #                              usecols='A:C', dtype=str)
    #
    # df_structure.iloc[:-1, 1:3] = df_structure.iloc[:-1, 1:3].applymap(convert_to_int)
    # df_structure.columns = ['Вид', 'Всего', 'Практика']
    # df_structure.fillna('', inplace=True)
    # max_load = df_structure.loc[0, 'Всего']  # максимальная учебная нагрузка
    # mand_load = df_structure.loc[1, 'Всего']  # обязательная нагрузка
    # practice_load = df_structure.loc[1, 'Практика']  # практическая нагрузка
    #
    # sam_df = df_structure[
    #     df_structure['Вид'] == 'Самостоятельная работа обучающегося (всего)']  # получаем часы самостоятельной работы
    # if len(sam_df) == 0:
    #     messagebox.showerror('Диана Создание рабочих программ',
    #                          'Проверьте наличие строки Самостоятельная работа обучающегося (всего) в листе Объем УД')
    # sam_load = sam_df.iloc[0, 1]


    """
    Обрабатываем листы с МДК
    """

    _dct_mdk_df = processing_mdk(data_pm) # получам словарь где ключ это название МДК а значение это словарь с данными и итогами по подсчету этих данных
    dct_mdk_df ={mdk:value['Данные'] for mdk,value in _dct_mdk_df.items()} # создаем словарь извлекая датафрейм
    _dct_mdk_data ={mdk:value['Итог'] for mdk,value in _dct_mdk_df.items()} # создаем словарь извлекая словарь с данными
    dct_mdk_data = dict() # считаем общую сумму
    for name,dct in _dct_mdk_data.items():
        for key,value in dct.items():
            if key not in dct_mdk_data:
                dct_mdk_data[key] = value
            else:
                dct_mdk_data[key] += value

    """Обрабатываем лист ПК и ОК

       """
    df_pk_ok = pd.read_excel(data_pm, sheet_name=pk_ok_vpd, usecols='A:C')
    df_pk_ok.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    # Обработка ПК
    lst_pk = df_pk_ok['Наименование ПК'].dropna().tolist()
    lst_pk = processing_punctuation_end_string(lst_pk, ';\n', '- ', '.')

    # Обработка ОК
    lst_ok = df_pk_ok['Наименование ОК'].dropna().tolist()
    lst_ok = processing_punctuation_end_string(lst_ok, ';\n', '- ', '.')

    # Разворачиваем ОК в две колонки
    df_flat_ok = df_pk_ok['Наименование ОК'].to_frame()
    df_flat_ok.dropna(inplace=True)

    df_flat_ok.columns = ['Описание']
    df_flat_ok['Код'] = df_flat_ok['Описание'].apply(extract_lr)
    df_flat_ok['Результат'] = df_flat_ok['Описание'].apply(extract_descr_lr)

    # Обработка ВПД
    lst_vpd = df_pk_ok['Виды профессиональной деятельности'].dropna().tolist()
    lst_vpd = processing_punctuation_end_string(lst_vpd, ';\n', '- ', '.')

    """
    Обрабатываем лист Контроль и Оценка
    """
    df_control = pd.read_excel(data_pm, sheet_name=control, usecols='A:B')
    df_control.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_control.columns = ['Результаты_обучения', 'Контроль_обучения']
    _lst_result_educ = df_control['Результаты_обучения'].dropna().tolist()  # создаем список
    control_set = {'Знания:','Умения:','Практический опыт:'}

    if not control_set.issubset(set(_lst_result_educ)): # проверяем наличие нужных слов
        raise ControlWord_PM

    border_divide = _lst_result_educ.index('Знания:')  # граница разделения знания
    border_divide_second = _lst_result_educ.index('Практический опыт:') # граница разделения Опыта
    lst_skill = _lst_result_educ[1:border_divide]  # получаем список умений
    lst_knowledge = _lst_result_educ[border_divide + 1:border_divide_second]  # получаем список знаний
    lst_prac_exp = _lst_result_educ[border_divide_second+1:]



    lst_skill = processing_punctuation_end_string(lst_skill, ';\n', '- ', '.')  # форматируем выходную строку
    lst_knowledge = processing_punctuation_end_string(lst_knowledge, ';\n', '- ', '.')  # форматируем выходную строку
    lst_prac_exp = processing_punctuation_end_string(lst_prac_exp, ';\n', '- ', '.')  # форматируем выходную строку

    df_control.fillna('', inplace=True)


    """
    Обрабатываем лист темы курсовых работ
        """
    df_kp = pd.read_excel(data_pm,sheet_name=kp,usecols='A')
    lst_kp = df_kp.iloc[:,0].dropna().tolist()
    lst_kp = processing_punctuation_end_string(lst_kp, ';\n', '- ', '.')  # форматируем выходную строку



    """
    Обрабатываем лист УП (Учебная практика)
    """
    df_up = pd.read_excel(data_pm, sheet_name=up, usecols='A:C')
    df_up.columns = ['Вид','Содержание','Объем']
    df_up.fillna(0,inplace=True)
    df_up = df_up.applymap(lambda x:int(x) if isinstance(x,float) else x)
    df_up = df_up.applymap(lambda x: '' if x ==0 else x)
    sum_up = sum_column_any_value(df_up,'Объем') # сумма учебной практики

    """
    Обрабатываем лист ПП (Производственная практика)
    """
    df_pp = pd.read_excel(data_pm, sheet_name=pp, usecols='A:C')
    df_pp.columns = ['Вид','Содержание','Объем']
    df_pp.fillna(0,inplace=True)
    df_pp = df_pp.applymap(lambda x:int(x) if isinstance(x,float) else x)
    df_pp = df_pp.applymap(lambda x: '' if x ==0 else x)
    sum_pp = sum_column_any_value(df_pp, 'Объем') # сумма производственной практики

    """
    Обрабатываем лист МТО
    """
    df_mto = pd.read_excel(data_pm, sheet_name=mto, usecols='A:E')
    df_mto.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    name_kab = df_mto['Наименование_учебного_кабинета'].dropna().tolist()  #
    name_kab = ','.join(name_kab)  # получаем все названия которые есть в колонке Наименование учебного кабинета

    name_lab = df_mto['Наименование_лаборатории'].dropna().tolist()
    if len(name_lab) == 0:
        name_lab = []
    else:
        name_lab = ','.join(name_lab)

    # Списки кабинета и средств обучения
    lst_obor_cab = df_mto[
        'Оборудование_учебного_кабинета'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_obor_cab) == 0:
        lst_obor_cab = ['На листе МТО НЕ заполнено оборудование учебного кабинета !!!']
    obor_cab = processing_punctuation_end_string(lst_obor_cab, ';\n', '- ',
                                                 '.')  # обрабатываем знаки пунктуации для каждой строки

    lst_tecn_educ = df_mto[
        'Технические_средства_обучения'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_tecn_educ) == 0:
        lst_tecn_educ = ['На листе МТО НЕ заполнены технические средства обучения !!!']
    tecn_educ = processing_punctuation_end_string(lst_tecn_educ, ';\n', '- ', '.')

    # Оборудование лаборатории
    lst_obor_labor = df_mto[
        'Оборудование_лаборатории'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_obor_labor) == 0 and name_lab:
        obor_labor = ['На листе МТО НЕ заполнено оборудование лаборатории !!!']
    elif len(lst_obor_labor) != 0 and not name_lab:
        obor_labor = ['На листе МТО заполнено оборудование лаборатории но не заполнено наименование лаборатории !!!']
    else:
        obor_labor = processing_punctuation_end_string(lst_obor_labor, ';\n', '- ',
                                                       '.')  # обрабатываем знаки пунктуации для каждой строки
    """
    Обрабатываем лист Основные источники
    """

    df_main_publ = pd.read_excel(data_pm, sheet_name=main_publ, usecols='A:G')
    if df_main_publ.shape[0] != 0:
        df_main_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
        df_main_publ.fillna('Не заполнено !!!', inplace=True)
        df_main_publ = df_main_publ.applymap(str)  # приводим к строковому виду
        df_main_publ = df_main_publ.applymap(str.strip)  # очищаем от пробелов в начале и конце

        df_main_publ['Основной_источник'] = df_main_publ.apply(processing_publ, axis=1)  # формируем строку
        df_main_publ.sort_values(by='Основной_источник', inplace=True)
        lst_main_source = df_main_publ['Основной_источник'].tolist()
    else:
        lst_main_source = 'Не заполнено'

    """
    Обрабатываем лист дополнительные источники
    """
    df_second_publ = pd.read_excel(data_pm, sheet_name=second_publ, usecols='A:G')
    if df_second_publ.shape[0] != 0:
        df_second_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
        df_second_publ.fillna('Не заполнено !!!', inplace=True)
        df_second_publ = df_second_publ.applymap(str)  # приводим к строковому виду
        df_second_publ = df_second_publ.applymap(str.strip)  # очищаем от пробелов в начале и конце

        df_second_publ['Основной_источник'] = df_second_publ.apply(processing_publ, axis=1)  # формируем строку
        df_second_publ.sort_values(by='Основной_источник', inplace=True)
        lst_slave_source = df_second_publ['Основной_источник'].tolist()
    else:
        lst_slave_source = 'Не заполнено'

    """
    Обрабатываем лист интернет источники
    """
    df_ii_publ = pd.read_excel(data_pm, sheet_name=ii_publ, usecols='A:B')
    if df_ii_publ.shape[0] != 0:
        df_ii_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки

        df_ii_publ.sort_values(by='Название', inplace=True)

        lst_inet_source = insert_type_source(df_ii_publ.copy())
    else:
        lst_inet_source = 'Не заполнено'

    """
    Обрабатываем лист ФГОС
    """
    df_fgos = pd.read_excel(data_pm,sheet_name=fgos,usecols='A:C')
    df_fgos.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_fgos.fillna('Не заполнено',inplace=True)






    doc = DocxTemplate(template_pm)



    # Конвертируем датафрейм с описанием программы в список словарей и добавляем туда нужные элементы
    data_program = df_desc_rp.to_dict('records')
    context = data_program[0]
    context['ФИО_УР'] = accept_UR # Переменные для ФИО утверждающих УР и ПР
    context['ФИО_ПР'] = accept_PR

    context['Лич_результаты'] = df_pers_result.to_dict('records')  # добаввляем лист личностных результатов
    context['Объем_ПМ'] = df_volume_pm.to_dict('records')  # объем ПМ
    # context['Учебная_работа'] = df_structure.to_dict('records')
    context['Контроль_оценка'] = df_control.to_dict('records')
    context['Знать'] = lst_knowledge
    context['Уметь'] = lst_skill
    context['Прак_опыт'] = lst_prac_exp

    # Объем МДК
    context['Объем_МДК'] = df_volume_all_mdk.to_dict('records')

    # добавляем единичные переменные
    context['Всего'] = dct_mdk_data['Всего часов']
    context['Всего_прак_под'] = dct_mdk_data['Всего практики']
    context['СРС'] = dct_mdk_data['Всего СРС']
    context['КР'] = dct_mdk_data['курсовая работа (КП)']

    context['КП'] = lst_kp # список тем курсовых работ

    context['Объем_УП'] = sum_up
    context['Объем_ПП'] = sum_pp


    # context['Макс_нагрузка'] = max_load
    # context['Обяз_нагрузка'] = mand_load
    # context['Практ_подготовка'] = practice_load
    # context['Сам_работа'] = sam_load
    #
    context['УП'] = df_up.to_dict('records')
    context['ПП'] = df_pp.to_dict('records')

    # #лист МТО
    context['Учебный_кабинет'] = name_kab
    context['Лаборатория'] = name_lab
    context['Список_оборудования'] = obor_cab
    context['Средства_обучения'] = tecn_educ
    context['Оборудование_лаборатории'] = obor_labor
    # лист Учебные издания
    context['Основные_источники'] = lst_main_source
    context['Дополнительные_источники'] = lst_slave_source
    context['Интернет_источники'] = lst_inet_source
    # Листы данные ОК и  данные ПК,ВПД
    context['ОК'] = lst_ok
    context['ПК'] = lst_pk
    context['ВПД'] = lst_vpd
    context['Общ_компетенции'] = df_flat_ok.to_dict('records')

    # Переменные ФГОС
    # даты
    context['ФГОС_у_дата'] = df_fgos.at[0,'Дата']
    context['ФГОС_з_дата'] = df_fgos.at[1,'Дата']
    context['Прог_з_дата'] = df_fgos.at[2,'Дата']

    # номера
    context['ФГОС_у_номер'] = df_fgos.at[0,'Номер документа']
    context['ФГОС_з_номер'] = df_fgos.at[1,'Номер документа']
    context['Прог_з_номер'] = df_fgos.at[2,'Номер документа']

    for idx,tpl_dct in enumerate(dct_mdk_df.items(),start=1):
        key,value = tpl_dct # распаковываем кортеж
        # делаем переменные для названий МДК
        name_var = f'МДК_{idx}'
        context[name_var] = key
        # создаем переменные датафреймов
        name_mdk_table = f'План_МДК_{idx}'
        context[name_mdk_table] = value.to_dict('records')
    # Создаем документ
    doc.render(context)
    # сохраняем документ
    # название программы
    # name_rp = df_desc_rp['Название_дисциплины'].tolist()[0]
    name_rp = 'Тест'
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    doc.save(f'{end_folder}/РП {name_rp[:40]} {current_time}.docx')


# messagebox.showerror('Диана Создание рабочих программ',
#                      'На листе Контроль в первой колонке должно быть указаны слова\n'
#                      'Умения: , Знания: , Практический опыт:\n'
#                      'Посмотрите пример в исходном шаблоне')

# jinja2.exceptions.UndefinedError: 'Интернет' is undefined неправильная запись в шаблоне

if __name__ == '__main__':
    template_pm_main = 'data/Шаблон автозаполнения ПМ.docx'
    data_pm_main = 'data/Пример заполнения ПМ.xlsx'
    end_folder_main = 'data'

    create_pm(template_pm_main, data_pm_main, end_folder_main)
    print('Lindy Booth')


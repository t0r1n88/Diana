"""
Скрипт для отработки генерации рабочих программ учебных дисциплин с помощью шаблонов docxtemplate
"""

import pandas as pd
import numpy as np
import openpyxl
from docxtpl import DocxTemplate
import string
import time
import re

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
    temp_lst = list(map(lambda x: sep_begin + x, lst_phrase))
    temp_lst = list(map(lambda x: x.strip(), temp_lst)) # очищаем от прбельных символов в начале и конце
    temp_lst = list(map(lambda x: x.rstrip(string.punctuation), temp_lst))  # очищаем от знаков пунктуации
    temp_lst[-1] = temp_lst[-1] + sep_end  # добавляем конечный знак пунктуации
    temp_str = f'{sep_string}'.join(temp_lst)  # создаем строку с разделителями
    return temp_str

def insert_type_source(lst_phrase:list)->list:
    """
    Вставка в строку слов [Электронный ресурс] Форма доступа:
    :param lst_phrase:список строк
    :return: список измененных строк
    """
    out_lst = []
    pattern = r'(?=[A-Za-z])' # регулярка для разделения
    for row in lst_phrase:
        temp_lst = re.split(pattern,row,maxsplit=1) # делим по первой английской букве
        temp_lst.insert(1,' [Электронный ресурс] Форма доступа:') # вставляем в середину списка нужную строку
        temp_str = ' '.join(temp_lst)
        out_lst.append(temp_str)
    return out_lst

def create_RP_for_UD(template_work_program:str,data_work_program:str,end_folder:str):
    """
    Скрипт для генерации рабочей програамы для учебной дисциплины с помощью DocxTemplate
    :param template_work_program: путь к шаблону рабочей программы в формате docx
    :type template_work_program: str
    :param data_work_program: путь к файлу с данными для рабочей программы в формате xlsx
    :type data_work_program: str
    :param end_folder: путь к папке куда будет сохранен созданный файл рабочей программы
    :type end_folder: str
    """
    # названия листов
    desc_rp = 'Описание РП'
    pers_result = 'Лич_результаты'
    structure = 'Объем УД'
    plan = 'План УД'
    mto = 'МТО'
    educ_publ = 'Учебные издания'
    ok = 'Данные для ОК'
    pk = 'Данные для ПК'

    # Обрабатываем лист Описание РП
    df_desc_rp = pd.read_excel(data_work_program, sheet_name=desc_rp, nrows=1, usecols='A:E')  # загружаем датафрейм
    df_desc_rp.fillna('НЕ ЗАПОЛНЕНО !!!', inplace=True)  # заполняем не заполненные разделы

    # Обрабатываем лист Лич_результаты

    df_pers_result = pd.read_excel(data_work_program, sheet_name=pers_result, usecols='A:B')
    df_pers_result.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_pers_result.columns = ['Код', 'Результат']

    # Обрабатываем лист Структура
    # Открываем файл
    wb = openpyxl.load_workbook(data_work_program, read_only=True)
    target_value = 'итог'

    # Поиск значения в выбранном столбце
    column_number = 1  # Номер столбца, в котором ищем значение (например, столбец A)
    target_row = None  # Номер строки с искомым значением

    for row in wb['Объем УД'].iter_rows(min_row=1, min_col=column_number, max_col=column_number):
        cell_value = row[0].value
        if target_value in str(cell_value).lower():
            target_row = row[0].row
            break

    if not target_row:
        # если не находим строку в которой есть слово итог то выдаем исключение
        print('Не найдена строка с названием Итоговая аттестация в форме ...')

    wb.close()  # закрываем файл

    # если значение найдено то считываем нужное количество строк и  7 колонок
    df_structure = pd.read_excel(data_work_program, sheet_name=structure, nrows=target_row,
                                 usecols='A:C', dtype=str)

    df_structure.iloc[:-1, 1:3] = df_structure.iloc[:-1, 1:3].applymap(convert_to_int)
    df_structure.columns = ['Вид', 'Всего', 'Практика']
    df_structure.fillna('', inplace=True)
    max_load = df_structure.loc[0, 'Всего']  # максимальная учебная нагрузка
    mand_load = df_structure.loc[1, 'Всего']  # обязательная нагрузка
    sam_df = df_structure[
        df_structure['Вид'] == 'Самостоятельная работа обучающегося (всего)']  # получаем часы самостоятельной работы
    if len(sam_df) == 0:
        print('Проверьте наличие строки Самостоятельная работа обучающегося (всего)')
    sam_load = sam_df.iloc[0, 1]

    """
    Обрабатываем лист План УД
    """
    df_plan_ud = pd.read_excel(data_work_program, sheet_name='План УД', usecols='A:F')
    df_plan_ud.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    df_plan_ud.columns = ['Раздел', 'Тема', 'Количество_часов', 'Практика', 'Вид_занятия', 'СРС']

    df_plan_ud['Раздел'].fillna('Пусто', inplace=True)

    borders = df_plan_ud[
        df_plan_ud['Раздел'].str.contains('семестр')].index  # получаем индексы строк где есть слово семестр

    part_df = []  # список для хранения кусков датафрейма
    previos_border = -1
    # делим датафрем по границам
    for value_border in borders:
        part = df_plan_ud.iloc[previos_border:value_border]
        part_df.append(part)
        previos_border = value_border

    # добавляем последнюю часть
    last_part = df_plan_ud.iloc[borders[-1]:]
    part_df.append(last_part)

    part_df.pop(0)  # удаляем нулевой элемент так как он пустой

    main_df = pd.DataFrame(
        columns=['Раздел', 'Тема', 'Количество_часов', 'Практика', 'Вид_занятия', 'СРС'])  # создаем базовый датафрейм

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
        # itog_row = {'Тема':margint_text,'Количество_часов':value_text} # создаем строку
        temp_df = pd.DataFrame([{'Тема': margint_text, 'Количество_часов': value_text}])
        df = pd.concat([df, temp_df], ignore_index=True)  # добаляем итоговую строку
        main_df = pd.concat([main_df, df], ignore_index=True)  # добавляем в основной датафрейм

    main_df.insert(0, 'Номер', np.nan)  # добавляем колонку с номерами занятий

    main_df['Тема'] = main_df['Тема'].fillna('Пусто')  # заменяем наны на пусто

    count = 0  # счетчик
    for idx, row in enumerate(main_df.itertuples()):
        if (row[3] == 'Пусто') | ('Итого часов' in row[3]):
            main_df.iloc[idx, 0] = ''
        else:
            count += 1
            main_df.iloc[idx, 0] = count

    # очищаем от пустых символов и строки Пусто
    main_df['Раздел'] = main_df['Раздел'].fillna('Пусто')
    main_df['Тема'] = main_df['Тема'].replace('Пусто', '')
    main_df['Раздел'] = main_df['Раздел'].replace('Пусто', '')
    main_df['Вид_занятия'] = main_df['Вид_занятия'].fillna('')

    main_df['Количество_часов'] = main_df['Количество_часов'].apply(convert_to_int)
    main_df['Количество_часов'] = main_df['Количество_часов'].fillna('')

    main_df['Практика'] = main_df['Практика'].apply(convert_to_int)
    main_df['СРС'] = main_df['СРС'].apply(convert_to_int)
    main_df['Тема'] = main_df['Раздел'] + main_df['Тема']
    main_df.drop(columns=['Раздел'], inplace=True)

    """
    Обрабатываем лист МТО
    """
    df_mto = pd.read_excel(data_work_program, sheet_name=mto, usecols='A:E')
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
    Обрабатываем лист Учебные издания
    """
    df_educ_publ = pd.read_excel(data_work_program, sheet_name=educ_publ, usecols='A:C')
    df_educ_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    lst_main_source = df_educ_publ[
        'Основные_источники'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_main_source) == 0:
        lst_main_source = ['На листе Учебные издания НЕ заполнены основные источники !!!']

    # Дополнительные источники
    lst_slave_source = df_educ_publ[
        'Дополнительные_источники'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_slave_source) == 0:
        lst_slave_source = ['На листе Учебные издания НЕ заполнены дополнительные источники !!!']

    # Интернет -источники
    lst_inet_source = df_educ_publ[
        'Интернет_ресурсы'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_inet_source) == 0:
        lst_inet_source = ['На листе Учебные издания НЕ заполнены интернет источники !!!']
    lst_inet_source = insert_type_source(lst_inet_source)

    """
    Обрабатываем лист Контроль и Оценка
    """
    df_control = pd.read_excel(data_work_program, sheet_name='Контроль и оценка', usecols='A:B')
    df_control.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_control.columns = ['Результаты_обучения', 'Контроль_обучения']
    _lst_result_educ = df_control['Результаты_обучения'].dropna().tolist()  # создаем список
    border_divide = _lst_result_educ.index('Знания:')  # граница разделения
    lst_skill = _lst_result_educ[1:border_divide]  # получаем список умений
    lst_knowledge = _lst_result_educ[border_divide + 1:]  # получаем список знаний

    lst_skill = processing_punctuation_end_string(lst_skill, ';\n', '- ', '.')  # форматируем выходную строку
    lst_knowledge = processing_punctuation_end_string(lst_knowledge, ';\n', '- ', '.')  # форматируем выходную строку
    df_control.fillna('', inplace=True)

    """
    Обрабатываем лист Данные для ОК
    """
    df_ok = pd.read_excel(data_work_program, sheet_name=ok, usecols='A')
    lst_ok = df_ok['Наименование'].dropna().tolist()
    lst_ok = processing_punctuation_end_string(lst_ok, ';\n', '- ', '.')

    """
    Обрабатываем лист Данные для ПК
    """
    df_pk = pd.read_excel(data_work_program, sheet_name=pk, usecols='A')
    lst_pk = df_pk['Наименование'].dropna().tolist()
    lst_pk = processing_punctuation_end_string(lst_pk, ';\n', '- ', '.')

    # Конвертируем датафрейм с описанием программы в список словарей и добавляем туда нужные элементы
    data_program = df_desc_rp.to_dict('records')
    context = data_program[0]
    context['Лич_результаты'] = df_pers_result.to_dict('records')  # добаввляем лист личностных результатов
    context['План_УД'] = main_df.to_dict('records')  # содержание учебной дисциплины
    context['Учебная_работа'] = df_structure.to_dict('records')
    context['Контроль_оценка'] = df_control.to_dict('records')
    context['Знать'] = lst_knowledge
    context['Уметь'] = lst_skill

    # добавляем единичные переменные
    context['Макс_нагрузка'] = max_load
    context['Обяз_нагрузка'] = mand_load
    context['Сам_работа'] = sam_load
    # лист МТО
    context['Учебный_кабинет'] = name_kab
    context['Лаборатория'] = name_lab
    context['Список_оборудования'] = obor_cab
    context['Средства_обучения'] = tecn_educ
    context['Оборудование_лаборатории'] = obor_labor
    # лист Учебные издания
    context['Основные_источники'] = lst_main_source
    context['Дополнительные_источники'] = lst_slave_source
    context['Интернет_источники'] = lst_inet_source
    # Листы данные ОК и  данные ПК
    context['ОК'] = lst_ok
    context['ПК'] = lst_pk

    doc = DocxTemplate(template_work_program)
    # Получаем ключи используемые в шаблоне
    set_of_variables = doc.get_undeclared_template_variables()

    # Создаем документ
    doc.render(context)
    # сохраняем документ
    # название программы
    name_rp = df_desc_rp['Название_дисциплины'].tolist()[0]
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    doc.save(f'{end_folder}/РП {name_rp[:40]} {current_time}.docx')


if __name__=='__main__':
    template_work_program = 'data/Шаблон автозаполнения РП.docx'
    data_work_program = 'data/Автозаполнение РП.xlsx'
    end_folder = 'data'

    # названия листов
    desc_rp = 'Описание РП'
    pers_result = 'Лич_результаты'
    structure = 'Объем УД'
    plan = 'План УД'
    mto = 'МТО'
    educ_publ = 'Учебные издания'
    ok = 'Данные для ОК'
    pk = 'Данные для ПК'

    # Обрабатываем лист Описание РП
    df_desc_rp = pd.read_excel(data_work_program, sheet_name=desc_rp, nrows=1, usecols='A:G')  # загружаем датафрейм
    df_desc_rp.fillna('НЕ ЗАПОЛНЕНО !!!', inplace=True)  # заполняем не заполненные разделы

    # Обрабатываем лист Лич_результаты

    df_pers_result = pd.read_excel(data_work_program, sheet_name=pers_result, usecols='A:B')
    df_pers_result.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_pers_result.columns = ['Код', 'Результат']

    # Обрабатываем лист Структура
    # Открываем файл
    wb = openpyxl.load_workbook(data_work_program, read_only=True)
    target_value = 'итог'

    # Поиск значения в выбранном столбце
    column_number = 1  # Номер столбца, в котором ищем значение (например, столбец A)
    target_row = None  # Номер строки с искомым значением

    for row in wb['Объем УД'].iter_rows(min_row=1, min_col=column_number, max_col=column_number):
        cell_value = row[0].value
        if target_value in str(cell_value).lower():
            target_row = row[0].row
            break

    if not target_row:
        # если не находим строку в которой есть слово итог то выдаем исключение
        print('Не найдена строка с названием Итоговая аттестация в форме ...')

    wb.close()  # закрываем файл

    # если значение найдено то считываем нужное количество строк и  7 колонок
    df_structure = pd.read_excel(data_work_program, sheet_name=structure, nrows=target_row,
                                 usecols='A:C', dtype=str)

    df_structure.iloc[:-1, 1:3] = df_structure.iloc[:-1, 1:3].applymap(convert_to_int)
    df_structure.columns = ['Вид', 'Всего', 'Практика']
    df_structure.fillna('', inplace=True)
    max_load = df_structure.loc[0, 'Всего']  # максимальная учебная нагрузка
    mand_load = df_structure.loc[1, 'Всего']  # обязательная нагрузка
    sam_df = df_structure[
        df_structure['Вид'] == 'Самостоятельная работа обучающегося (всего)']  # получаем часы самостоятельной работы
    if len(sam_df) == 0:
        print('Проверьте наличие строки Самостоятельная работа обучающегося (всего)')
    sam_load = sam_df.iloc[0, 1]

    """
    Обрабатываем лист План УД
    """
    df_plan_ud = pd.read_excel('data/Автозаполнение РП.xlsx', sheet_name='План УД', usecols='A:F')
    df_plan_ud.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    df_plan_ud.columns = ['Раздел', 'Тема', 'Количество_часов', 'Практика', 'Вид_занятия', 'СРС']

    df_plan_ud['Раздел'].fillna('Пусто', inplace=True)

    borders = df_plan_ud[
        df_plan_ud['Раздел'].str.contains('семестр')].index  # получаем индексы строк где есть слово семестр

    part_df = []  # список для хранения кусков датафрейма
    previos_border = -1
    # делим датафрем по границам
    for value_border in borders:
        part = df_plan_ud.iloc[previos_border:value_border]
        part_df.append(part)
        previos_border = value_border

    # добавляем последнюю часть
    last_part = df_plan_ud.iloc[borders[-1]:]
    part_df.append(last_part)

    part_df.pop(0)  # удаляем нулевой элемент так как он пустой

    main_df = pd.DataFrame(
        columns=['Раздел', 'Тема', 'Количество_часов', 'Практика', 'Вид_занятия', 'СРС'])  # создаем базовый датафрейм

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
        # itog_row = {'Тема':margint_text,'Количество_часов':value_text} # создаем строку
        temp_df = pd.DataFrame([{'Тема': margint_text, 'Количество_часов': value_text}])
        df = pd.concat([df, temp_df], ignore_index=True)  # добаляем итоговую строку
        main_df = pd.concat([main_df, df], ignore_index=True)  # добавляем в основной датафрейм

    main_df.insert(0, 'Номер', np.nan)  # добавляем колонку с номерами занятий

    main_df['Тема'] = main_df['Тема'].fillna('Пусто')  # заменяем наны на пусто

    count = 0  # счетчик
    for idx, row in enumerate(main_df.itertuples()):
        if (row[3] == 'Пусто') | ('Итого часов' in row[3]):
            main_df.iloc[idx, 0] = ''
        else:
            count += 1
            main_df.iloc[idx, 0] = count

    # очищаем от пустых символов и строки Пусто
    main_df['Раздел'] = main_df['Раздел'].fillna('Пусто')
    main_df['Тема'] = main_df['Тема'].replace('Пусто', '')
    main_df['Раздел'] = main_df['Раздел'].replace('Пусто', '')
    main_df['Вид_занятия'] = main_df['Вид_занятия'].fillna('')

    main_df['Количество_часов'] = main_df['Количество_часов'].apply(convert_to_int)
    main_df['Количество_часов'] = main_df['Количество_часов'].fillna('')

    main_df['Практика'] = main_df['Практика'].apply(convert_to_int)
    main_df['СРС'] = main_df['СРС'].apply(convert_to_int)
    main_df['Тема'] = main_df['Раздел'] + main_df['Тема']
    main_df.drop(columns=['Раздел'],inplace=True)



    """
    Обрабатываем лист МТО
    """
    df_mto = pd.read_excel(data_work_program, sheet_name=mto, usecols='A:E')
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
    obor_cab = processing_punctuation_end_string(lst_obor_cab, ';\n', '- ','.')  # обрабатываем знаки пунктуации для каждой строки

    lst_tecn_educ = df_mto['Технические_средства_обучения'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
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
        obor_labor = processing_punctuation_end_string(lst_obor_labor, ';\n', '- ','.')  # обрабатываем знаки пунктуации для каждой строки



    """
    Обрабатываем лист Учебные издания
    """
    df_educ_publ = pd.read_excel(data_work_program, sheet_name=educ_publ, usecols='A:C')
    df_educ_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    lst_main_source = df_educ_publ['Основные_источники'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_main_source) == 0:
        lst_main_source = ['На листе Учебные издания НЕ заполнены основные источники !!!']

    # Дополнительные источники
    lst_slave_source = df_educ_publ[
        'Дополнительные_источники'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_slave_source) == 0:
        lst_slave_source = ['На листе Учебные издания НЕ заполнены дополнительные источники !!!']

    # Интернет -источники
    lst_inet_source = df_educ_publ['Интернет_ресурсы'].dropna().tolist()  # создаем список удаляя все незаполненные ячейки
    if len(lst_inet_source) == 0:
        lst_inet_source = ['На листе Учебные издания НЕ заполнены интернет источники !!!']
    lst_inet_source = insert_type_source(lst_inet_source)

    """
    Обрабатываем лист Контроль и Оценка
    """
    df_control = pd.read_excel(data_work_program,sheet_name='Контроль и оценка',usecols='A:B')
    df_control.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_control.columns = ['Результаты_обучения','Контроль_обучения']
    _lst_result_educ = df_control['Результаты_обучения'].dropna().tolist() # создаем список
    border_divide = _lst_result_educ.index('Знания:') # граница разделения
    lst_skill = _lst_result_educ[1:border_divide] # получаем список умений
    lst_knowledge = _lst_result_educ[border_divide+1:] # получаем список знаний

    lst_skill = processing_punctuation_end_string(lst_skill, ';\n', '- ','.') # форматируем выходную строку
    lst_knowledge = processing_punctuation_end_string(lst_knowledge, ';\n', '- ','.') # форматируем выходную строку
    df_control.fillna('',inplace=True)

    """
    Обрабатываем лист Данные для ОК
    """
    df_ok = pd.read_excel(data_work_program,sheet_name=ok,usecols='A')
    lst_ok = df_ok['Наименование'].dropna().tolist()
    lst_ok = processing_punctuation_end_string(lst_ok, ';\n', '- ','.')

    """
    Обрабатываем лист Данные для ПК
    """
    df_pk = pd.read_excel(data_work_program,sheet_name=pk,usecols='A')
    lst_pk = df_pk['Наименование'].dropna().tolist()
    lst_pk = processing_punctuation_end_string(lst_pk, ';\n', '- ','.')


    # Конвертируем датафрейм с описанием программы в список словарей и добавляем туда нужные элементы
    data_program = df_desc_rp.to_dict('records')
    context = data_program[0]
    context['Лич_результаты'] = df_pers_result.to_dict('records')  # добаввляем лист личностных результатов
    context['План_УД'] = main_df.to_dict('records') # содержание учебной дисциплины
    context['Учебная_работа'] = df_structure.to_dict('records')
    context['Контроль_оценка'] = df_control.to_dict('records')
    context['Знать'] = lst_knowledge
    context['Уметь'] = lst_skill


    # добавляем единичные переменные
    context['Макс_нагрузка'] = max_load
    context['Обяз_нагрузка'] = mand_load
    context['Сам_работа'] = sam_load
    # лист МТО
    context['Учебный_кабинет'] = name_kab
    context['Лаборатория'] = name_lab
    context['Список_оборудования'] = obor_cab
    context['Средства_обучения'] = tecn_educ
    context['Оборудование_лаборатории'] = obor_labor
    # лист Учебные издания
    context['Основные_источники'] = lst_main_source
    context['Дополнительные_источники'] = lst_slave_source
    context['Интернет_источники'] = lst_inet_source
    # Листы данные ОК и  данные ПК
    context['ОК'] = lst_ok
    context['ПК'] = lst_pk



    doc = DocxTemplate(template_work_program)
    # Получаем ключи используемые в шаблоне
    set_of_variables = doc.get_undeclared_template_variables()

    # Создаем документ
    doc.render(context)
    # сохраняем документ
    # название программы
    name_rp = df_desc_rp['Название_дисциплины'].tolist()[0]
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    doc.save(f'{end_folder}/РП {name_rp[:40]} {current_time}.docx')
    print('Lindy Booth!')

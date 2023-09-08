"""
Скрипт для отработки генерации рабочих программ учебных дисциплин общеобразовательного типа с помощью шаблонов docxtemplate
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
    if len(lst_phrase) == 0:
        return ''
    lst_phrase = list(map(str,lst_phrase)) # делаем строковыми элементы
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

def processing_publ(row):
    """
    Функция для генерации строки с литературой в нужном формате
    :param row: строка датафрейма
    :return: сумма строк датафрейма в нужном формате
    """
    author = row[0] # автор(ы)
    name_book = row[1] # название
    full_city = row[2] # полное название города
    short_city = row[3] # краткое название города
    publ_house = row[4] # издательство
    year = row[5] # год издания
    quan_pages = row[6] # число страниц
    author = author.rstrip(string.punctuation) # очищаем от символа пунктуации в конце
    name_book = name_book.rstrip(string.punctuation) # очищаем от символа пунктуации в конце
    short_city = short_city.rstrip(string.punctuation)  # очищаем от символа пунктуации в конце
    publ_house = publ_house.rstrip(string.punctuation)  # очищаем от символа пунктуации в конце
    # извлекаем год
    result = re.search(r'\d{4}',year)
    if result:
        clean_year = result.group()
    else:
        clean_year = 'Неправильно заполнен год издания, введите год в формате 4 цифры без букв'

    # извлекаем количество страниц
    result = re.search(r'\d+',quan_pages)
    if result:
        clean_quan_pages = result.group()
    else:
        clean_quan_pages = 'Неправильно заполнено количество страниц, введите количество в виде числа без букв'
    # Формируем итоговую строку
    out_str = f'{author}. {name_book}.- {short_city}.: {publ_house}, {clean_year}.- {clean_quan_pages} c.'

    return out_str

def create_RP_for_UD_OOD(template_work_program:str,data_work_program:str,end_folder:str):
    """
    Скрипт для генерации рабочей програамы для учебной дисциплины ООД (общеобразовательной) с помощью DocxTemplate
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
    content = 'Содержание'
    target = 'Цели'
    result = 'Результаты'
    uupd = 'УУПД'
    main_publ = 'ОИ'
    second_publ = 'ДИ'
    ii_publ = 'ИИ'
    control = 'Контроль'
    # Обрабатываем лист Описание РП
    df_desc_rp = pd.read_excel(data_work_program, sheet_name=desc_rp, nrows=1, usecols='A:L')  # загружаем датафрейм
    df_desc_rp.fillna('НЕ ЗАПОЛНЕНО !!!', inplace=True)  # заполняем не заполненные разделы
    df_desc_rp.columns = ['Тип_программы', 'Название_дисциплины', 'Цикл', 'Перечень', 'Код_специальность_профессия',
                          'Год', 'Разработчик', 'Методист', 'ПЦК', 'Пред_ПЦК', 'Должность', 'Утверждающий']

    # Обрабатываем лист Объем УД
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
    prof_load = df_structure.loc[1, 'Практика']  # практическая нагрузка

    sam_df = df_structure[
        df_structure['Вид'] == 'Самостоятельная работа обучающегося (всего)']  # получаем часы самостоятельной работы
    if len(sam_df) == 0:
        print('Проверьте наличие строки Самостоятельная работа обучающегося (всего)')
    sam_load = sam_df.iloc[0, 1]

    """
    Обрабатываем лист План УД
    """
    df_plan_ud = pd.read_excel(data_work_program, sheet_name=plan, usecols='A:H')
    df_plan_ud.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    df_plan_ud.columns = ['Курс_семестр', 'Раздел', 'Тема', 'Содержание', 'Количество_часов', 'Практика', 'Вид_занятия',
                          'СРС']
    df_plan_ud['Курс_семестр'].fillna('Пусто', inplace=True)
    df_plan_ud['Раздел'].fillna('Пусто', inplace=True)
    df_plan_ud['Тема'].fillna('Пусто', inplace=True)

    borders = df_plan_ud[
        df_plan_ud['Курс_семестр'].str.contains('семестр')].index  # получаем индексы строк где есть слово семестр

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

    main_df['Практика'] = main_df['Практика'].apply(convert_to_int)
    main_df['СРС'] = main_df['СРС'].apply(convert_to_int)
    main_df['Содержание'] = main_df['Курс_семестр'] + main_df['Раздел'] + main_df['Тема'] + main_df['Содержание']
    main_df.drop(columns=['Курс_семестр', 'Раздел', 'Тема'], inplace=True)

    """
    Обрабатываем лист Содержание
    """
    df_content = pd.read_excel(data_work_program, sheet_name=content, usecols='A',nrows=1,dtype=str)
    df_content = df_content.applymap(str.strip) # очищаем от пробелов

    """
    Обрабатываем листы Цели,Результаты, УУПД
    """
    #Цели
    df_target = pd.read_excel(data_work_program, sheet_name=target, usecols='A')
    lst_target = df_target['Цели'].dropna().tolist()
    lst_target = processing_punctuation_end_string(lst_target, ';\n', '- ', '.')
    # Результаты
    # Личностные результаты
    df_result = pd.read_excel(data_work_program, sheet_name=result, usecols='A:C')
    lst_perconal_result = df_result['Личностные_результаты'].dropna().tolist()
    lst_perconal_result = processing_punctuation_end_string(lst_perconal_result, ';\n', '- ', '.')
    # Метапредметные результаты
    lst_meta_result = df_result['Метапредметные_результаты'].dropna().tolist()
    lst_meta_result = processing_punctuation_end_string(lst_meta_result, ';\n', '- ', '.')
    # Предметные результаты
    lst_predmet_result = df_result['Предметные_результаты'].dropna().tolist()
    lst_predmet_result = processing_punctuation_end_string(lst_predmet_result, ';\n', '- ', '.')
    # УУПД
    df_uupd = pd.read_excel(data_work_program, sheet_name=uupd, usecols='A')
    lst_uupd = df_target.iloc[:,0].dropna().tolist()
    lst_uupd = processing_punctuation_end_string(lst_uupd, ';\n', '- ', '.')







    """
    Обрабатываем лист Основные источники
    """


    df_main_publ = pd.read_excel(data_work_program, sheet_name=main_publ, usecols='A:G')
    df_main_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_main_publ.fillna('Не заполнено !!!', inplace=True)
    df_main_publ = df_main_publ.applymap(str)  # приводим к строковому виду
    df_main_publ = df_main_publ.applymap(str.strip)  # очищаем от пробелов в начале и конце

    df_main_publ['Основной_источник'] = df_main_publ.apply(processing_publ, axis=1)  # формируем строку
    df_main_publ.sort_values(by='Основной_источник', inplace=True)
    lst_main_source = df_main_publ['Основной_источник'].tolist()

    """
    Обрабатываем лист дополнительные источники
    """
    df_second_publ = pd.read_excel(data_work_program, sheet_name=second_publ, usecols='A:G')
    df_second_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_second_publ.fillna('Не заполнено !!!', inplace=True)
    df_second_publ = df_second_publ.applymap(str)  # приводим к строковому виду
    df_second_publ = df_second_publ.applymap(str.strip)  # очищаем от пробелов в начале и конце

    df_second_publ['Основной_источник'] = df_second_publ.apply(processing_publ, axis=1)  # формируем строку
    df_second_publ.sort_values(by='Основной_источник', inplace=True)
    lst_slave_source = df_second_publ['Основной_источник'].tolist()

    """
    Обрабатываем лист интернет источники
    """
    df_ii_publ = pd.read_excel(data_work_program, sheet_name=ii_publ, usecols='A')
    df_ii_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_ii_publ.sort_values(by='Интернет_ресурсы', inplace=True)
    lst_inet_source = df_ii_publ['Интернет_ресурсы'].tolist()
    lst_inet_source = insert_type_source(lst_inet_source)

    """
    Обрабатываем лист Контроль и Оценка
    """
    df_control = pd.read_excel(data_work_program, sheet_name=control, usecols='A:B')
    df_control.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_control.columns = ['Результаты_обучения', 'Контроль_обучения']
    _lst_result_educ = df_control['Результаты_обучения'].dropna().tolist()  # создаем список
    border_divide = _lst_result_educ.index('Знания:')  # граница разделения
    lst_skill = _lst_result_educ[1:border_divide]  # получаем список умений
    lst_knowledge = _lst_result_educ[border_divide + 1:]  # получаем список знаний

    lst_skill = processing_punctuation_end_string(lst_skill, ';\n', '- ', '.')  # форматируем выходную строку
    lst_knowledge = processing_punctuation_end_string(lst_knowledge, ';\n', '- ', '.')  # форматируем выходную строку
    df_control.fillna('', inplace=True)


    # Конвертируем датафрейм с описанием программы в список словарей и добавляем туда нужные элементы
    data_program = df_desc_rp.to_dict('records')
    context = data_program[0]
    context['План_УД'] = main_df.to_dict('records')  # содержание учебной дисциплины
    context['Учебная_работа'] = df_structure.to_dict('records')
    context['Контроль_оценка'] = df_control.to_dict('records')
    context['Знать'] = lst_knowledge
    context['Уметь'] = lst_skill

    # добавляем единичные переменные
    context['Макс_нагрузка'] = max_load
    context['Обяз_нагрузка'] = mand_load
    context['Проф_направленность'] = prof_load
    context['Сам_работа'] = sam_load

    # Листы Цели,результаты,УУПД
    context['Цели'] = lst_target
    context['Личностные_результаты'] =lst_perconal_result
    context['Метапредметные_результаты'] =lst_meta_result
    context['Предметные_результаты'] =lst_predmet_result
    context['УУПД'] =lst_uupd



    # Лист Содержание
    context['Содержание_дисциплины'] = df_content.iloc[0,0]
    # лист Учебные издания
    context['Основные_источники'] = lst_main_source
    context['Дополнительные_источники'] = lst_slave_source
    context['Интернет_источники'] = lst_inet_source

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
    doc.save(f'{end_folder}/РП ООД {name_rp[:40]} {current_time}.docx')

if __name__ == '__main__':

    template_work_program = 'data/Шаблон автозаполнения ООД.docx'
    data_work_program = 'data/Автозаполнение ООД.xlsx'
    end_folder = 'data'

    # названия листов
    desc_rp = 'Описание РП'
    pers_result = 'Лич_результаты'
    structure = 'Объем УД'
    plan = 'План УД'
    content = 'Содержание'
    target = 'Цели'
    result = 'Результаты'
    uupd = 'УУПД'
    main_publ = 'ОИ'
    second_publ = 'ДИ'
    ii_publ = 'ИИ'
    control = 'Контроль'
    # Обрабатываем лист Описание РП
    df_desc_rp = pd.read_excel(data_work_program, sheet_name=desc_rp, nrows=1, usecols='A:L')  # загружаем датафрейм
    df_desc_rp.fillna('НЕ ЗАПОЛНЕНО !!!', inplace=True)  # заполняем не заполненные разделы
    df_desc_rp.columns = ['Тип_программы', 'Название_дисциплины', 'Цикл', 'Перечень', 'Код_специальность_профессия',
                          'Год', 'Разработчик', 'Методист', 'ПЦК', 'Пред_ПЦК', 'Должность', 'Утверждающий']

    # Обрабатываем лист Объем УД
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
    prof_load = df_structure.loc[1, 'Практика']  # практическая нагрузка

    sam_df = df_structure[
        df_structure['Вид'] == 'Самостоятельная работа обучающегося (всего)']  # получаем часы самостоятельной работы
    if len(sam_df) == 0:
        print('Проверьте наличие строки Самостоятельная работа обучающегося (всего)')
    sam_load = sam_df.iloc[0, 1]

    """
    Обрабатываем лист План УД
    """
    df_plan_ud = pd.read_excel(data_work_program, sheet_name=plan, usecols='A:H')
    df_plan_ud.dropna(inplace=True, thresh=1)  # удаляем пустые строки

    df_plan_ud.columns = ['Курс_семестр', 'Раздел', 'Тема', 'Содержание', 'Количество_часов', 'Практика', 'Вид_занятия',
                          'СРС']
    df_plan_ud['Курс_семестр'].fillna('Пусто', inplace=True)
    df_plan_ud['Раздел'].fillna('Пусто', inplace=True)
    df_plan_ud['Тема'].fillna('Пусто', inplace=True)

    borders = df_plan_ud[
        df_plan_ud['Курс_семестр'].str.contains('семестр')].index  # получаем индексы строк где есть слово семестр

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

    main_df['Практика'] = main_df['Практика'].apply(convert_to_int)
    main_df['СРС'] = main_df['СРС'].apply(convert_to_int)
    main_df['Содержание'] = main_df['Курс_семестр'] + main_df['Раздел'] + main_df['Тема'] + main_df['Содержание']
    main_df.drop(columns=['Курс_семестр', 'Раздел', 'Тема'], inplace=True)

    """
    Обрабатываем лист Содержание
    """
    df_content = pd.read_excel(data_work_program, sheet_name=content, usecols='A',nrows=1,dtype=str)
    df_content = df_content.applymap(str.strip) # очищаем от пробелов

    """
    Обрабатываем листы Цели,Результаты, УУПД
    """
    #Цели
    df_target = pd.read_excel(data_work_program, sheet_name=target, usecols='A')
    lst_target = df_target['Цели'].dropna().tolist()
    lst_target = processing_punctuation_end_string(lst_target, ';\n', '- ', '.')
    # Результаты
    # Личностные результаты
    df_result = pd.read_excel(data_work_program, sheet_name=result, usecols='A:C')
    lst_perconal_result = df_result['Личностные_результаты'].dropna().tolist()
    lst_perconal_result = processing_punctuation_end_string(lst_perconal_result, ';\n', '- ', '.')
    # Метапредметные результаты
    lst_meta_result = df_result['Метапредметные_результаты'].dropna().tolist()
    lst_meta_result = processing_punctuation_end_string(lst_meta_result, ';\n', '- ', '.')
    # Предметные результаты
    lst_predmet_result = df_result['Предметные_результаты'].dropna().tolist()
    lst_predmet_result = processing_punctuation_end_string(lst_predmet_result, ';\n', '- ', '.')
    # УУПД
    df_uupd = pd.read_excel(data_work_program, sheet_name=uupd, usecols='A')
    lst_uupd = df_target.iloc[:,0].dropna().tolist()
    lst_uupd = processing_punctuation_end_string(lst_uupd, ';\n', '- ', '.')







    """
    Обрабатываем лист Основные источники
    """


    df_main_publ = pd.read_excel(data_work_program, sheet_name=main_publ, usecols='A:G')
    df_main_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_main_publ.fillna('Не заполнено !!!', inplace=True)
    df_main_publ = df_main_publ.applymap(str)  # приводим к строковому виду
    df_main_publ = df_main_publ.applymap(str.strip)  # очищаем от пробелов в начале и конце

    df_main_publ['Основной_источник'] = df_main_publ.apply(processing_publ, axis=1)  # формируем строку
    df_main_publ.sort_values(by='Основной_источник', inplace=True)
    lst_main_source = df_main_publ['Основной_источник'].tolist()

    """
    Обрабатываем лист дополнительные источники
    """
    df_second_publ = pd.read_excel(data_work_program, sheet_name=second_publ, usecols='A:G')
    df_second_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_second_publ.fillna('Не заполнено !!!', inplace=True)
    df_second_publ = df_second_publ.applymap(str)  # приводим к строковому виду
    df_second_publ = df_second_publ.applymap(str.strip)  # очищаем от пробелов в начале и конце

    df_second_publ['Основной_источник'] = df_second_publ.apply(processing_publ, axis=1)  # формируем строку
    df_second_publ.sort_values(by='Основной_источник', inplace=True)
    lst_slave_source = df_second_publ['Основной_источник'].tolist()

    """
    Обрабатываем лист интернет источники
    """
    df_ii_publ = pd.read_excel(data_work_program, sheet_name=ii_publ, usecols='A')
    df_ii_publ.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_ii_publ.sort_values(by='Интернет_ресурсы', inplace=True)
    lst_inet_source = df_ii_publ['Интернет_ресурсы'].tolist()
    lst_inet_source = insert_type_source(lst_inet_source)

    """
    Обрабатываем лист Контроль и Оценка
    """
    df_control = pd.read_excel(data_work_program, sheet_name=control, usecols='A:B')
    df_control.dropna(inplace=True, thresh=1)  # удаляем пустые строки
    df_control.columns = ['Результаты_обучения', 'Контроль_обучения']
    _lst_result_educ = df_control['Результаты_обучения'].dropna().tolist()  # создаем список
    border_divide = _lst_result_educ.index('Знания:')  # граница разделения
    lst_skill = _lst_result_educ[1:border_divide]  # получаем список умений
    lst_knowledge = _lst_result_educ[border_divide + 1:]  # получаем список знаний

    lst_skill = processing_punctuation_end_string(lst_skill, ';\n', '- ', '.')  # форматируем выходную строку
    lst_knowledge = processing_punctuation_end_string(lst_knowledge, ';\n', '- ', '.')  # форматируем выходную строку
    df_control.fillna('', inplace=True)


    # Конвертируем датафрейм с описанием программы в список словарей и добавляем туда нужные элементы
    data_program = df_desc_rp.to_dict('records')
    context = data_program[0]
    context['План_УД'] = main_df.to_dict('records')  # содержание учебной дисциплины
    context['Учебная_работа'] = df_structure.to_dict('records')
    context['Контроль_оценка'] = df_control.to_dict('records')
    context['Знать'] = lst_knowledge
    context['Уметь'] = lst_skill

    # добавляем единичные переменные
    context['Макс_нагрузка'] = max_load
    context['Обяз_нагрузка'] = mand_load
    context['Проф_направленность'] = prof_load
    context['Сам_работа'] = sam_load

    # Листы Цели,результаты,УУПД
    context['Цели'] = lst_target
    context['Личностные_результаты'] =lst_perconal_result
    context['Метапредметные_результаты'] =lst_meta_result
    context['Предметные_результаты'] =lst_predmet_result
    context['УУПД'] =lst_uupd



    # Лист Содержание
    context['Содержание_дисциплины'] = df_content.iloc[0,0]
    # лист Учебные издания
    context['Основные_источники'] = lst_main_source
    context['Дополнительные_источники'] = lst_slave_source
    context['Интернет_источники'] = lst_inet_source

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
    doc.save(f'{end_folder}/РП ООД {name_rp[:40]} {current_time}.docx')
    print('Lindy Booth')


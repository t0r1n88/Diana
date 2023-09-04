"""
Скрипт для отработки генерации рабочих программ дисциплин с помощью шаблонов docxtemplate
"""

import pandas as pd
import numpy as np
import openpyxl
from docxtpl import DocxTemplate
import time
pd.options.mode.chained_assignment = None  # default='warn'
pd.set_option('display.max_columns', None)  # Отображать все столбцы
pd.set_option('display.expand_frame_repr', False)  # Не переносить строки
import warnings
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
    if cell.isdigit():
        return int(cell)
    else:
        return ''



template_work_program = 'data/Шаблон автозаполнения РП.docx'
data_work_program = 'data/Автозаполнение РП.xlsx'

# названия листов
desc_rp = 'Описание РП'
pers_result = 'Лич_результаты'
structure = 'Структура'
mto = 'МТО'
educ_publ = 'Учебные издания'


# Обрабатываем лист Описание РП
df_desc_rp = pd.read_excel(data_work_program, sheet_name=desc_rp, nrows=1, usecols='A:E') # загружаем датафрейм
df_desc_rp.fillna('НЕ ЗАПОЛНЕНО !!!', inplace=True) # заполняем не заполненные разделы

# Обрабатываем лист Лич_результаты

df_pers_result = pd.read_excel(data_work_program,sheet_name=pers_result,usecols='A:B')
df_pers_result.dropna(inplace=True,thresh=1) # удаляем пустые строки
df_pers_result.columns = ['Код','Результат']

# Обрабатываем лист Структура
# Открываем файл
wb = openpyxl.load_workbook(data_work_program,read_only=True)
target_value = 'итог'

# Поиск значения в выбранном столбце
column_number = 1  # Номер столбца, в котором ищем значение (например, столбец A)
target_row = None  # Номер строки с искомым значением

for row in wb['Структура'].iter_rows(min_row=1, min_col=column_number, max_col=column_number):
    cell_value = row[0].value
    if target_value in str(cell_value).lower():
        target_row = row[0].row
        break


if not target_row:
    # если не находим строку в которой есть слово итог то выдаем исключение
    print('Не найдена строка с названием Итоговая аттестация в форме ...')

wb.close() # закрываем файл

# если значение найдено то считываем нужное количество строк и  7 колонок
df_structure = pd.read_excel(data_work_program, sheet_name=structure, nrows=target_row,
                      usecols='A:C', dtype=str)

df_structure.iloc[:-1, 1:3] = df_structure.iloc[:-1, 1:3].applymap(convert_to_int)
df_structure.columns = ['Вид','Всего','Практика']
df_structure.fillna('',inplace=True)
max_load = df_structure.loc[0,'Всего'] # максимальная учебная нагрузка
mand_load = df_structure.loc[1,'Всего'] # обязательная нагрузка
sam_df = df_structure[df_structure['Вид'] == 'Самостоятельная работа обучающегося (всего)'] # получаем часы самостоятельной работы
if len(sam_df) == 0:
    print('Проверьте наличие строки Самостоятельнgая работа обучающегося (всего)')
sam_load = sam_df.iloc[0,1]

"""
Обрабатываем лист МТО
"""
df_mto = pd.read_excel(data_work_program,sheet_name=mto,usecols='A:E')
df_mto.dropna(inplace=True,thresh=1) # удаляем пустые строки
name_kab = df_mto['Наименование_учебного_кабинета'].dropna().tolist() #
name_kab = ','.join(name_kab) # получаем все названия которые есть в колонке Наименование учебного кабинета

name_lab = df_mto['Наименование_лаборатории'].dropna().tolist()
if len(name_lab) == 0:
    name_lab = ['']
name_lab = ','.join(name_lab)

# Списки кабинета и средств обучения
lst_obor_cab = df_mto['Оборудование_учебного_кабинета'].dropna().tolist() # создаем список удаляя все незаполненные ячейки
if len(lst_obor_cab) == 0:
    lst_obor_cab = ['На листе МТО НЕ заполнено оборудование учебного кабинета !!!']

lst_obor_cab = list(map(lambda x:'- '+ x,lst_obor_cab))
obor_cab = ';\n'.join(lst_obor_cab)

lst_tecn_educ = df_mto['Технические_средства_обучения'].dropna().tolist() # создаем список удаляя все незаполненные ячейки
if len(lst_tecn_educ) == 0:
    lst_tecn_educ = ['На листе МТО НЕ заполнены технические средства обучения !!!']

lst_tecn_educ = list(map(lambda x:'- '+ x,lst_tecn_educ))
tecn_educ = ';\n'.join(lst_tecn_educ)


"""
Обрабатываем лист Учебные издания
"""
df_educ_publ = pd.read_excel(data_work_program,sheet_name=educ_publ,usecols='A:C')
df_educ_publ.dropna(inplace=True,thresh=1) # удаляем пустые строки

lst_main_source = df_educ_publ['Основные_источники'].dropna().tolist() # создаем список удаляя все незаполненные ячейки
if len(lst_main_source) == 0:
    lst_main_source = ['На листе Учебные издания НЕ заполнены основные источники !!!']

# lst_main_source = list(map(lambda x:'- '+ x,lst_main_source))
# main_source = ';\n'.join(lst_main_source)








# Конвертируем датафрейм с описанием программы в список словарей и добавляем туда нужные элементы
data_program = df_desc_rp.to_dict('records')
context = data_program[0]
context['lst_pers_result'] = df_pers_result.to_dict('records') # добаввляем лист личностных результатов
context['lst_up'] = df_structure.to_dict('records')

# добавляем единичные переменные
context['Макс_нагрузка'] = max_load
context['Обяз_нагрузка'] = mand_load
context['Сам_работа'] = sam_load
# лист МТО
context['Учебный_кабинет'] = name_kab
context['Лаборатория'] = name_lab
context['Список_оборудования'] = obor_cab
context['Средства_обучения'] = tecn_educ
# лист Учебные издания

context['lst_main_source'] = lst_main_source




doc = DocxTemplate(template_work_program)
# Создаем документ
doc.render(context)
# сохраняем документ
# название программы
name_rp = df_desc_rp['Название_дисциплины'].tolist()[0]
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
doc.save(f'data/РП {name_rp[:40]} {current_time}.docx')

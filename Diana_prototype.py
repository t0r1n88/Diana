"""
Скрипт для отработки генерации рабочих программ дисциплин с помощью шаблонов docxtemplate
"""

import pandas as pd
import openpyxl
from docxtpl import DocxTemplate
import time
pd.options.mode.chained_assignment = None  # default='warn'
pd.set_option('display.max_columns', None)  # Отображать все столбцы
pd.set_option('display.expand_frame_repr', False)  # Не переносить строки
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=FutureWarning, module='openpyxl')





template_work_program = 'data/Шаблон автозаполнения РП.docx'
data_work_program = 'data/Автозаполнение РП.xlsx'

# названия листов
desc_rp = 'Описание РП'
pers_result = 'Лич_результаты'
mto = 'МТО'
educ_publ = 'Учебные издания'


# Обрабатываем лист Описание РП
desc_rp_df = pd.read_excel(data_work_program,sheet_name=desc_rp,nrows=1,usecols='A:E') # загружаем датафрейм
desc_rp_df.fillna('НЕ ЗАПОЛНЕНО !!!',inplace=True) # заполняем не заполненные разделы

# Конвертируем датафрейм с описанием программы в список словарей
data_program = desc_rp_df.to_dict('records')
context = data_program[0]
print(context)

doc = DocxTemplate(template_work_program)
# Создаем документ
doc.render(context)
# сохраняем документ
# название программы
name_rp = desc_rp_df['Название_дисциплины'].tolist()[0]
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
doc.save(f'data/РП {name_rp[:40]}{current_time}.docx')

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


def generate_table_method_dev(df: pd.DataFrame):
    """
    Функция для генерации сложной таблицы методических разработок
    :param df: датафрейм
    :return: датафрейм
    """
    main_df = pd.DataFrame(columns=df.columns)  # создаем датафрейм куда будут добавляться данные
    main_df.insert(0, 'Номер', '')
    df.insert(0, 'Номер', 0)
    quantity_met = df.shape[0]  # количество метод изданий
    count = 1  # счетчик строк
    lst_type = sorted(df['Вид'].unique())
    result_str_met = f'ИТОГО изданий-{quantity_met}:'
    for type in lst_type:
        name_table = type.capitalize()  # получаем название промежуточной таблицы
        row_header = pd.DataFrame(columns=main_df.columns,
                                  data=[[name_table, '', '', '', '', '', '']])
        main_df = pd.concat([main_df, row_header], axis=0, ignore_index=True)
        temp_df = df[df['Вид'] == type]
        if len(temp_df) != 0:
            temp_df['Номер'] = range(count, count + len(temp_df))  # присваеваем номера строк
            main_df = pd.concat([main_df, temp_df], axis=0, ignore_index=True)
            result_str_met +=f'\n{type}-{str(len(temp_df))}'
            count += len(temp_df)


    quantity_teacher = len(df['ФИО'].unique())  # количество преподавателей
    result_teacher = f'ИТОГО преподавателей-{quantity_teacher}'

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
    result_str_teacher = f'ИТОГО преподавателей-{quantity_teacher}\nМероприятий:\nфедеральных-{count_type_event["федеральный"]}\nмеждународных-{count_type_event["международный"]}'
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
    result_df = df.copy()  # копируем датафрейм чтобы потом посчитать общую результирующую строку
    df['Название'] = df['Название'] + '\nТема '+ df['Тема']
    count = 1  # счетчик строк
    lst_type = sorted(df['Вид'].unique())
    for type in lst_type:
        name_table = type.capitalize()  # получаем название промежуточной таблицы
        row_header = pd.DataFrame(columns=main_df.columns,
                                  data=[[name_table, '', '', '', '', '', '','','','']])
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

        count_way_event = Counter(temp_df['Форма']) # количество по формам участия
        lst_way_event = [f'{key}-{value}' for key,value in count_way_event.items()]



        count_result = Counter(temp_df['Результат']) # количество по результатам
        result_str_result = f'1 место-{count_result["1 место"]}\n2 место-{count_result["2 место"]}\n3 место-{count_result["3 место"]}\nноминация-{count_result["номинация"]}'
        row_itog = pd.DataFrame(columns=main_df.columns,
                                data=[['Итого','',f'выступлений-{quantity_event}',
                                       f'преподавателей-{quantity_teacher}','','','\n'.join(lst_type_event),'\n'.join(lst_way_event),result_str_result,'']])
        main_df = pd.concat([main_df,row_itog])

    # Результирующая строка для всего датафрейма
    # Результирующая строка
    quantity_event = len(result_df)  # количество конкурсов
    quantity_teacher = len(result_df['ФИО'].unique())  # количество педагогов
    count_type_event = Counter(result_df['Уровень'])  # количество по уровням
    lst_type_event = [f'{key}-{value}' for key, value in count_type_event.items()]

    count_way_event = Counter(result_df['Форма'])  # количество по формам участия
    lst_way_event = [f'{key}-{value}' for key, value in count_way_event.items()]

    count_result = Counter(result_df['Результат'])  # количество по результатам
    result_str_result = f'1 место-{count_result["1 место"]}\n2 место-{count_result["2 место"]}\n3 место-{count_result["3 место"]}\nноминация-{count_result["номинация"]}'
    result_type_event_str = 'По уровню мероприятий '+'\n'.join(lst_type_event)
    row_itog = pd.DataFrame(columns=result_df.columns,
                            data=[['Итого по всем выступлениям', '', f'выступлений-{quantity_event}',
                                   f'преподавателей выступило-{quantity_teacher}', '', '', result_type_event_str,
                                   '\n'.join(lst_way_event), result_str_result, '']])
    main_df = pd.concat([main_df, row_itog])

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

    lst_type = sorted(df['Вид'].unique())
    for type in lst_type:
        name_table = type.capitalize()  # получаем название промежуточной таблицы
        row_header = pd.DataFrame(columns=main_df.columns,
                                  data=[[name_table, '', '', '', '', '', '','','','','','','','']])
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
        result_str_result = f'1 место-{count_result["1 место"]}\n2 место-{count_result["2 место"]}\n3 место-{count_result["3 место"]}\nноминация-{count_result["номинация"]}'



        row_itog = pd.DataFrame(columns=main_df.columns,
                                data=[['ИТОГО',f'руководителей-{quantity_teacher}', result_student,'', '','','',
                                       '', '',result_str_type,'',result_str_result,'','']])

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
    first_sheet_df.columns = ['ФИО', 'Дата_рождения', 'Дата_ПОО', 'Дисциплина', 'Стаж', 'Педстаж','Стаж_в_ПОО',
                              'Организация', 'Квалификация', 'Год_окончания', 'Категория', 'Приказ', 'Сайт']
    first_sheet_df[['Стаж', 'Педстаж']] = first_sheet_df[['Стаж', 'Педстаж']].fillna(0)  # заменяем нан нулями
    first_sheet_df[['Стаж', 'Педстаж']] = first_sheet_df[['Стаж', 'Педстаж']].applymap(
        lambda x: int(x) if isinstance(x, (int, float)) else x)
    first_sheet_df.fillna('', inplace=True)

    context['Преподаватель'] = first_sheet_df.iloc[0, 0]  # ФИО преподавателя
    context['Общая_информация'] = first_sheet_df[['Дисциплина', 'Дата_рождения', 'Дата_ПОО', 'Стаж', 'Педстаж','Стаж_в_ПОО',
                                                  'Категория', 'Приказ', 'Сайт']].to_dict('records')
    context['Образование'] = first_sheet_df[['Организация', 'Квалификация', 'Год_окончания']].to_dict('records')

    # данные с листа Повышение квалификации
    skills_dev_df = dct_value['Повышение квалификации']
    skills_dev_df.columns = ['ФИО', 'Вид','Название',  'Место', 'Дата', 'Часов', 'Документ']
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
    internship_df.columns = ['ФИО', 'Место','Наименование', 'Часов', 'Дата']
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
    itog_internship_df.loc[-1] = [result_str_internship, '','', '', '']
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
    personal_perf_df.columns = ['ФИО', 'Дата', 'Форма','Уровень','Вид','Название', 'Тема',  'Результат','Номинация']
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

    open_lessons_df = dct_value['Открытые уроки_мастер_классы']
    open_lessons_df.columns = ['ФИО', 'Вид', 'Дисциплина', 'Курс','Профессия','Группа', 'Тема', 'Дата']
    open_lessons_df.fillna('', inplace=True)
    context['Открытые_уроки'] = open_lessons_df.to_dict('records') # оставим такое название

    itog_open_lessons_df = open_lessons_df.copy()
    quantity_course = len(
        itog_open_lessons_df['Дисциплина'].unique())  # получаем количество уникальных дисциплин

    quantity_open_lessons = itog_open_lessons_df.shape[0]  # общее количество открытых уроков

    count_type_open_lessons = Counter(itog_open_lessons_df['Вид'].tolist())
    lst_type_open = [f'{key}-{value}' for key, value in count_type_open_lessons.items()]
    result_str_open_lessons = f'ИТОГО:\nоткрытых уроков-{quantity_open_lessons}\nдисциплин-{quantity_course}\nПо типам занятий:\n'
    type_str = '\n'.join(lst_type_open)
    itog_open_lessons_df.loc[-1] = ['', '', result_str_open_lessons + type_str, '', '',
                                                               '','','']
    context['Открытые_уроки_итог'] = itog_open_lessons_df.to_dict('records')

    mutual_visits_df = dct_value['Взаимопосещение']
    mutual_visits_df.columns = ['ФИО', 'ФИО_посещенного', 'Дата','Дисциплина' ,'Курс','Профессия','Группа', 'Тема']
    mutual_visits_df.fillna('', inplace=True)
    context['Взаимопосещение'] = mutual_visits_df.to_dict('records')

    itog_mutual_visits_df = mutual_visits_df.copy()
    context['Взаимопосещение_итог'] = generate_table_pmutual_visits(itog_mutual_visits_df).to_dict(
        'records')  # генерируем таблицу

    student_perf_df = dct_value['УИРС']
    student_perf_df.columns = ['ФИО','Дата','Форма','Уровень','Вид','Название','Тема', 'ФИО_студента','Курс', 'Профессия', 'Группа',
                                 'Результат','Номинация']
    student_perf_df.fillna('', inplace=True)
    context['УИРС'] = student_perf_df.to_dict('records')

    itog_student_perf_df = student_perf_df.copy()
    context['УИРС_итог'] = generate_table_student_perf(itog_student_perf_df).to_dict(
        'records')  # генерируем таблицу

    nmr_df = dct_value['Работа по НМР']
    nmr_df.columns = ['ФИО', 'Тема', 'Обобщение', 'Форма','Наименование', 'Место', 'Дата']
    nmr_df.fillna('', inplace=True)
    context['НМР'] = nmr_df.to_dict('records')

    itog_nmr_df = nmr_df.copy()
    quantity_teacher = len(itog_nmr_df['ФИО'].unique())
    itog_nmr_df.loc[-1] = [f'ИТОГО преподавателей-{quantity_teacher}', '', '', '', '', '','']
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

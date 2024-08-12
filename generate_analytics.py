"""
Скрипт для создания аналитической справки в формате xlsx по преподавателям
"""
import pandas as pd
import openpyxl



def create_analytics_report(dct_data:dict,result_folder:str):
    """
    Функция для создания аналитической отчетности в формате xlsx
    :param dct_data: словарь где ключ это название листа а значение это датафрейм с данными из этого листа
    :param result_folder: папка для сохранения результа
    :return: файл xlsx с несколькими листами
    """
    print(dct_data.keys())
    dct_df = dict() # словарь для хранения созданных аналитических таблиц
    # Повышение квалификации
    skills_dev_df = dct_data['Повышение квалификации']
    # В разрезе преподавателей
    teacher_skills_dev_df = pd.pivot_table(skills_dev_df,index=['ФИО','Вид повышения квалификации'],
                                           values=['Название программы повышения квалификации'],
                                           aggfunc='count')
    dct_df['ПК_преподаватели'] = teacher_skills_dev_df
    # В разрезе видов повышения квалификации
    course_skills_dev_df = pd.pivot_table(skills_dev_df,index=['Вид повышения квалификации','ФИО'],
                                           values=['Название программы повышения квалификации'],
                                           aggfunc='count')
    dct_df['ПК_виды'] = course_skills_dev_df

    # Стажировка
    internship_df = dct_data['Стажировка']
    # В разрезе преподавателей
    teacher_internship_df = pd.pivot_table(internship_df,index=['ФИО','Место стажировки'],
                                           values=['Дата'],
                                           aggfunc='count')
    dct_df['Стажировка_преп'] = teacher_internship_df
    # В разрезе мест стажировки
    course_internship_df = pd.pivot_table(internship_df,index=['Место стажировки','ФИО'],
                                           values=['Дата'],
                                           aggfunc='count')
    dct_df['Стажировка_места'] = course_internship_df

    # Методические разработки
    method_dev_df = dct_data['Методические разработки']
    # В разрезе преподавателей
    teacher_method_dev_df = pd.pivot_table(method_dev_df, index=['ФИО', 'Вид методического издания'],
                                           values=['Дата разработки'],
                                           aggfunc='count')
    dct_df['Метод_разр_преп'] = teacher_method_dev_df
    # В разрезе видов
    type_method_df = pd.pivot_table(method_dev_df, index=['Вид методического издания', 'ФИО'],
                                          values=['Дата разработки'],
                                          aggfunc='count')
    dct_df['Метод_разр_вид'] = type_method_df
    # В разрезе профессий
    prof_method_df = pd.pivot_table(method_dev_df, index=['Профессия/специальность ','Вид методического издания', 'ФИО'],
                                          values=['Дата разработки'],
                                          aggfunc='count')
    dct_df['Метод_разр_проф'] = prof_method_df
    print(prof_method_df)
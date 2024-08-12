"""
Скрипт для создания аналитической справки в формате xlsx по преподавателям
"""
import pandas as pd
import openpyxl
import time



def create_analytics_report(dct_data:dict,result_folder:str):
    """
    Функция для создания аналитической отчетности в формате xlsx
    :param dct_data: словарь где ключ это название листа а значение это датафрейм с данными из этого листа
    :param result_folder: папка для сохранения результа
    :return: файл xlsx с несколькими листами
    """
    # Повышение квалификации
    skills_dev_df = dct_data['Повышение квалификации']
    # В разрезе преподавателей
    teacher_skills_dev_df_one_col = pd.pivot_table(skills_dev_df,index=['ФИО'],
                                           values=['Название программы повышения квалификации'],
                                           aggfunc='count').rename(columns={'Название программы повышения квалификации':'Количество'})
    # В разрезе видов повышения квалификации
    course_skills_dev_df_one_col = pd.pivot_table(skills_dev_df,index=['Вид повышения квалификации'],
                                           values=['Название программы повышения квалификации'],
                                           aggfunc='count').rename(columns={'Название программы повышения квалификации':'Количество'})

    # В разрезе преподавателей
    teacher_skills_dev_df = pd.pivot_table(skills_dev_df,index=['ФИО','Вид повышения квалификации'],
                                           values=['Название программы повышения квалификации'],
                                           aggfunc='count').rename(columns={'Название программы повышения квалификации':'Количество'})
    # В разрезе видов повышения квалификации
    course_skills_dev_df = pd.pivot_table(skills_dev_df,index=['Вид повышения квалификации','ФИО'],
                                           values=['Название программы повышения квалификации'],
                                           aggfunc='count').rename(columns={'Название программы повышения квалификации':'Количество'})



    # Стажировка
    internship_df = dct_data['Стажировка']
    # В разрезе преподавателей
    teacher_internship_df_one_col = pd.pivot_table(internship_df,index=['ФИО'],
                                           values=['Дата'],
                                           aggfunc='count').rename(columns={'Дата':'Количество'})
    teacher_internship_df = pd.pivot_table(internship_df,index=['ФИО','Место стажировки'],
                                           values=['Дата'],
                                           aggfunc='count').rename(columns={'Дата':'Количество'})
    # В разрезе мест стажировки
    course_internship_df_one_col = pd.pivot_table(internship_df,index=['Место стажировки'],
                                           values=['Дата'],
                                           aggfunc='count').rename(columns={'Дата':'Количество'})
    course_internship_df = pd.pivot_table(internship_df,index=['Место стажировки','ФИО'],
                                           values=['Дата'],
                                           aggfunc='count').rename(columns={'Дата':'Количество'})

    # Методические разработки
    method_dev_df = dct_data['Методические разработки']
    # В разрезе преподавателей
    teacher_method_dev_df_one_col = pd.pivot_table(method_dev_df, index=['ФИО'],
                                           values=['Дата разработки'],
                                           aggfunc='count').rename(columns={'Дата разработки':'Количество'})
    teacher_method_dev_df = pd.pivot_table(method_dev_df, index=['ФИО', 'Вид методического издания'],
                                           values=['Дата разработки'],
                                           aggfunc='count').rename(columns={'Дата разработки':'Количество'})
    # В разрезе видов
    type_method_df_one_col = pd.pivot_table(method_dev_df, index=['Вид методического издания'],
                                          values=['Дата разработки'],
                                          aggfunc='count').rename(columns={'Дата разработки':'Количество'})
    type_method_df = pd.pivot_table(method_dev_df, index=['Вид методического издания', 'ФИО'],
                                          values=['Дата разработки'],
                                          aggfunc='count').rename(columns={'Дата разработки':'Количество'})


    # В разрезе профессий
    prof_method_df_one_col = pd.pivot_table(method_dev_df, index=['Профессия/специальность '],
                                          values=['Дата разработки'],
                                          aggfunc='count').rename(columns={'Дата разработки':'Количество'})

    prof_method_df = pd.pivot_table(method_dev_df, index=['Профессия/специальность ','Вид методического издания', 'ФИО'],
                                          values=['Дата разработки'],
                                          aggfunc='count').rename(columns={'Дата разработки':'Количество'})
    # Мероприятия проведенные ППС
    events_teacher_df = dct_data['Мероприятия, пров. ППС']
    # В разрезе преподавателей
    teacher_events_df = pd.pivot_table(events_teacher_df, index=['ФИО', 'Уровень мероприятия'],
                                      values=['Дата'],
                                      aggfunc='count')
    # В разрезе мест стажировки
    level_events_df = pd.pivot_table(events_teacher_df, index=['Уровень мероприятия', 'ФИО'],
                                     values=['Дата'],
                                     aggfunc='count')
    # Личное выступление ППС
    personal_perf_df = dct_data['Личное выступление ППС']
    # В разрезе преподавателей
    teacher_personal_perf_df = pd.pivot_table(personal_perf_df, index=['ФИО', 'Уровень мероприятия','Вид мероприятия','Способ участия'],
                                              columns=['Результат участия'],
                                      values=['Дата'],
                                      aggfunc='count')
    # # В разрезе видов мероприятий
    # course_personal_perf_df = pd.pivot_table(personal_perf_df, index=['Вид мероприятия','Уровень мероприятия', 'ФИО','Результат участия'],
    #                                  values=['Дата'],
    #                                  aggfunc='count')
    # dct_df['Стажировка_места'] = course_blank_df
    # print(course_blank_df)

    # генерируем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    with pd.ExcelWriter(f'{result_folder}/Статистика {current_time}.xlsx') as writer:
        teacher_skills_dev_df_one_col.to_excel(writer,sheet_name='Повышение квалификации')
        course_skills_dev_df_one_col.to_excel(writer,sheet_name='Повышение квалификации',startrow=len(teacher_skills_dev_df_one_col)+3)
        teacher_skills_dev_df.to_excel(writer,sheet_name='Повышение квалификации',startrow=len(teacher_skills_dev_df_one_col)+len(course_skills_dev_df_one_col)+5)
        course_skills_dev_df.to_excel(writer,sheet_name='Повышение квалификации',startrow=len(teacher_skills_dev_df_one_col)+len(course_skills_dev_df_one_col)+len(teacher_skills_dev_df)+10)
        # Стажировка
        teacher_internship_df_one_col.to_excel(writer, sheet_name='Стажировка')
        course_internship_df_one_col.to_excel(writer, sheet_name='Стажировка',startrow=len(teacher_internship_df_one_col)+3)
        teacher_internship_df.to_excel(writer, sheet_name='Стажировка',startrow=len(teacher_internship_df_one_col)+len(course_internship_df_one_col)+5)
        course_internship_df.to_excel(writer, sheet_name='Стажировка',startrow=len(teacher_internship_df_one_col)+len(course_internship_df_one_col)+len(teacher_internship_df)+10)
        # Метод разработки
        teacher_method_dev_df_one_col.to_excel(writer, sheet_name='Методические разработки')
        type_method_df_one_col.to_excel(writer, sheet_name='Методические разработки',startcol=teacher_method_dev_df_one_col.shape[1] + 5)
        prof_method_df_one_col.to_excel(writer, sheet_name='Методические разработки',startcol=teacher_method_dev_df_one_col.shape[1] + teacher_method_dev_df_one_col.shape[1]+10)
        max_row = max(len(teacher_method_dev_df_one_col),len(type_method_df_one_col),len(prof_method_df_one_col)) # получаем строку с которой надо начинать запись второго ряда
        teacher_method_dev_df.to_excel(writer, sheet_name='Методические разработки',startrow=max_row + 5)
        type_method_df.to_excel(writer, sheet_name='Методические разработки',startrow=max_row + 5,startcol=teacher_method_dev_df.shape[1] + 5)
        prof_method_df.to_excel(writer, sheet_name='Методические разработки',startcol=teacher_method_dev_df.shape[1]+type_method_df.shape[1]+10,startrow=max_row + 5)

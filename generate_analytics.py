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
    teacher_events_df_one_col = pd.pivot_table(events_teacher_df, index=['ФИО'],
                                      values=['Дата'],
                                      aggfunc='count').rename(columns={'Дата':'Количество'})
    teacher_events_df = pd.pivot_table(events_teacher_df, index=['ФИО', 'Уровень мероприятия'],
                                      values=['Дата'],
                                      aggfunc='count').rename(columns={'Дата':'Количество'})
    # В разрезе уровней
    level_events_df_one_col = pd.pivot_table(events_teacher_df, index=['Уровень мероприятия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    level_events_df = pd.pivot_table(events_teacher_df, index=['Уровень мероприятия', 'ФИО'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    # Личное выступление ППС
    personal_perf_df = dct_data['Личное выступление ППС']
    # В разрезе преподавателей
    teacher_personal_perf_df_one_col = pd.pivot_table(personal_perf_df, index=['ФИО'],
                                      values=['Дата'],
                                      aggfunc='count').rename(columns={'Дата':'Количество'})
    teacher_personal_perf_df = pd.pivot_table(personal_perf_df, index=['ФИО','Вид мероприятия','Уровень мероприятия','Результат участия'],
                                      values=['Дата'],
                                      aggfunc='count').rename(columns={'Дата':'Количество'})

    # # В разрезе видов мероприятий
    course_personal_perf_df_one_col = pd.pivot_table(personal_perf_df, index=['Вид мероприятия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    course_personal_perf_df = pd.pivot_table(personal_perf_df, index=['Вид мероприятия','Уровень мероприятия', 'ФИО','Результат участия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    # В разрезе уровней
    level_personal_perf_df_one_col = pd.pivot_table(personal_perf_df, index=['Уровень мероприятия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    level_personal_perf_df = pd.pivot_table(personal_perf_df, index=['Уровень мероприятия','Вид мероприятия', 'ФИО','Результат участия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    # В разрезе результатов
    result_personal_perf_df_one_col = pd.pivot_table(personal_perf_df, index=['Результат участия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    result_personal_perf_df = pd.pivot_table(personal_perf_df, index=['Результат участия','Уровень мероприятия','Вид мероприятия','ФИО'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    # В разрезе способов
    way_personal_perf_df_one_col = pd.pivot_table(personal_perf_df, index=['Способ участия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    way_personal_perf_df = pd.pivot_table(personal_perf_df, index=['Способ участия','Результат участия','Уровень мероприятия','Вид мероприятия','ФИО'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})

    # Публикации
    publications_df = dct_data['Публикации']
    # В разрезе преподавателей
    teacher_publications_df_one_col = pd.pivot_table(publications_df, index=['ФИО'],
                                      values=['Дата выпуска'],
                                      aggfunc='count').rename(columns={'Дата выпуска':'Количество'})
    teacher_publications_df = pd.pivot_table(publications_df, index=['ФИО','Издание'],
                                      values=['Дата выпуска'],
                                      aggfunc='count').rename(columns={'Дата выпуска':'Количество'})
    # В разрезе изданий
    publ_publications_df_one_col = pd.pivot_table(publications_df, index=['Издание'],
                                      values=['Дата выпуска'],
                                      aggfunc='count').rename(columns={'Дата выпуска':'Количество'})
    publ_publications_df = pd.pivot_table(publications_df, index=['Издание','ФИО'],
                                      values=['Дата выпуска'],
                                      aggfunc='count').rename(columns={'Дата выпуска':'Количество'})


    # Открытые уроки
    open_lessons_df = dct_data['Открытые уроки']
    # В разрезе преподавателей
    teacher_open_lessons_df_one_col = pd.pivot_table(open_lessons_df, index=['ФИО'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    teacher_open_lessons_df = pd.pivot_table(open_lessons_df, index=['ФИО','Вид занятия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    # В разрезе видов занятий
    type_open_lessons_df_one_col = pd.pivot_table(open_lessons_df, index=['Вид занятия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    type_open_lessons_df = pd.pivot_table(open_lessons_df, index=['Вид занятия','ФИО'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})

    # В разрезе дисциплин
    lesson_open_lessons_df_one_col = pd.pivot_table(open_lessons_df, index=['Дисциплина'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    lesson_open_lessons_df = pd.pivot_table(open_lessons_df, index=['Дисциплина','Группа'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    # В разрезе групп
    group_open_lessons_df_one_col = pd.pivot_table(open_lessons_df, index=['Группа'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    group_open_lessons_df = pd.pivot_table(open_lessons_df, index=['Группа','Дисциплина'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})

    # Взаимопосещения
    mutual_visits_df = dct_data['Взаимопосещение']
    teacher_mutual_visits_df_one_col = pd.pivot_table(mutual_visits_df, index=['ФИО'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})
    teacher_mutual_visits_df = pd.pivot_table(mutual_visits_df, index=['ФИО','ФИО посещенного педагога'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})

    teacher_visited_visits_df_one_col = pd.pivot_table(mutual_visits_df, index=['ФИО посещенного педагога'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})
    teacher_visited_mutual_visits_df = pd.pivot_table(mutual_visits_df, index=['ФИО посещенного педагога','ФИО'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})

    group_mutual_visits_df_one_col = pd.pivot_table(mutual_visits_df, index=['Группа'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})
    group_mutual_visits_df = pd.pivot_table(mutual_visits_df, index=['Группа','ФИО посещенного педагога'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})
    theme_mutual_visits_df_one_col = pd.pivot_table(mutual_visits_df, index=['Тема'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})
    theme_mutual_visits_df = pd.pivot_table(mutual_visits_df, index=['Тема','ФИО посещенного педагога'],
                                      values=['Дата посещения'],
                                      aggfunc='count').rename(columns={'Дата посещения':'Количество'})





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
        # Мероприятия проведенные ППС
        teacher_events_df_one_col.to_excel(writer, sheet_name='Мероприятия, пров. ППС')
        level_events_df_one_col.to_excel(writer, sheet_name='Мероприятия, пров. ППС',startcol=teacher_events_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_events_df_one_col),len(level_events_df_one_col))
        teacher_events_df.to_excel(writer, sheet_name='Мероприятия, пров. ППС',startrow=max_row+5)
        level_events_df.to_excel(writer, sheet_name='Мероприятия, пров. ППС',startrow=max_row+5,startcol=teacher_events_df_one_col.shape[1] + 5)
        # Личное выступление ППС
        teacher_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС')
        course_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row = max(len(teacher_personal_perf_df_one_col),len(course_personal_perf_df_one_col))
        teacher_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        course_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row = max(len(teacher_personal_perf_df)+max_row+5,len(course_personal_perf_df)+max_row+5)
        level_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        result_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row = max(len(level_personal_perf_df_one_col)+max_row+3, len(result_personal_perf_df_one_col)+max_row+3)
        level_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        result_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row =  max(len(level_personal_perf_df)+max_row+3, len(result_personal_perf_df)+max_row+3)
        way_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        way_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)

        # Публикации
        teacher_publications_df_one_col.to_excel(writer, sheet_name='Публикации')
        publ_publications_df_one_col.to_excel(writer, sheet_name='Публикации',startcol=teacher_publications_df_one_col.shape[1]+5)
        max_row = max(len(teacher_publications_df_one_col),len(publ_publications_df_one_col))
        teacher_publications_df.to_excel(writer, sheet_name='Публикации',startrow=max_row+3)
        publ_publications_df.to_excel(writer, sheet_name='Публикации',startrow=max_row+3,startcol=teacher_publications_df_one_col.shape[1]+5)
        # Открытые уроки
        teacher_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки')
        type_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки',
                                              startcol=teacher_open_lessons_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_open_lessons_df_one_col), len(type_open_lessons_df_one_col))
        teacher_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3)
        type_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3,
                                      startcol=teacher_open_lessons_df_one_col.shape[1] + 5)

        max_row = max(len(teacher_open_lessons_df)+max_row+5, len(teacher_open_lessons_df)+ max_row+5)
        lesson_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки',startrow=max_row+3)
        group_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки',startrow=max_row+3,
                                              startcol=teacher_open_lessons_df_one_col.shape[1] + 5)
        max_row = max(len(lesson_open_lessons_df_one_col)+max_row+5, len(group_open_lessons_df_one_col)+max_row+5)
        lesson_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3)
        group_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3,
                                      startcol=teacher_open_lessons_df_one_col.shape[1] + 5)
        # Взаимопосещения
        teacher_mutual_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение')
        teacher_visited_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение',
                                              startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_mutual_visits_df_one_col), len(teacher_visited_visits_df_one_col))
        teacher_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3)
        teacher_visited_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3,
                                      startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_mutual_visits_df) + max_row + 5, len(teacher_visited_mutual_visits_df) + max_row + 5)
        group_mutual_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение',startrow=max_row+3)
        theme_mutual_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение',startrow=max_row+3,startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)
        max_row = max(len(group_mutual_visits_df_one_col) + max_row + 5, len(theme_mutual_visits_df_one_col) + max_row + 5)
        group_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3)
        theme_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3,
                                      startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)






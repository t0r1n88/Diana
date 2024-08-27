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
                                           values=['Дата утверждения'],
                                           aggfunc='count').rename(columns={'Дата утверждения':'Количество'})
    teacher_method_dev_df = pd.pivot_table(method_dev_df, index=['ФИО', 'Вид методического издания'],
                                           values=['Дата утверждения'],
                                           aggfunc='count').rename(columns={'Дата утверждения':'Количество'})
    # В разрезе видов
    type_method_df_one_col = pd.pivot_table(method_dev_df, index=['Вид методического издания'],
                                          values=['Дата утверждения'],
                                          aggfunc='count').rename(columns={'Дата утверждения':'Количество'})
    type_method_df = pd.pivot_table(method_dev_df, index=['Вид методического издания', 'ФИО'],
                                          values=['Дата утверждения'],
                                          aggfunc='count').rename(columns={'Дата утверждения':'Количество'})


    # В разрезе профессий
    prof_method_df_one_col = pd.pivot_table(method_dev_df, index=['Профессия/специальность'],
                                          values=['Дата утверждения'],
                                          aggfunc='count').rename(columns={'Дата утверждения':'Количество'})

    prof_method_df = pd.pivot_table(method_dev_df, index=['Профессия/специальность','Вид методического издания', 'ФИО'],
                                          values=['Дата утверждения'],
                                          aggfunc='count').rename(columns={'Дата утверждения':'Количество'})
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
    # В разрезе форм участия
    way_personal_perf_df_one_col = pd.pivot_table(personal_perf_df, index=['Форма участия'],
                                     values=['Дата'],
                                     aggfunc='count').rename(columns={'Дата':'Количество'})
    way_personal_perf_df = pd.pivot_table(personal_perf_df, index=['Форма участия','Результат участия','Уровень мероприятия','Вид мероприятия','ФИО'],
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
    # УИРС
    student_perf_df = dct_data['УИРС']
    teacher_student_perf_df_one_col = pd.pivot_table(student_perf_df, index=['ФИО обучающегося','Результат участия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    teacher_student_perf_df = pd.pivot_table(student_perf_df, index=['ФИО','ФИО обучающегося','Результат участия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})

    prof_student_perf_df_one_col = pd.pivot_table(student_perf_df, index=['Профессия/специальность','Группа'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    prof_student_perf_df = pd.pivot_table(student_perf_df, index=['Профессия/специальность','Группа', 'ФИО обучающегося', 'Результат участия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})

    type_student_perf_df_one_col = pd.pivot_table(student_perf_df, index=['Вид мероприятия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    type_student_perf_df = pd.pivot_table(student_perf_df, index=['Вид мероприятия','Уровень мероприятия', 'ФИО обучающегося', 'Результат участия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})

    level_student_perf_df_one_col = pd.pivot_table(student_perf_df, index=['Уровень мероприятия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})
    level_student_perf_df = pd.pivot_table(student_perf_df, index=['Уровень мероприятия','Вид мероприятия','ФИО обучающегося', 'Результат участия'],
                                      values=['Дата проведения'],
                                      aggfunc='count').rename(columns={'Дата проведения':'Количество'})

    # Работа по НМР
    nmr_df = dct_data['Работа по НМР']
    teacher_nmr_df_one_col = pd.pivot_table(nmr_df, index=['ФИО'],
                                      values=['Дата обобщения опыта'],
                                      aggfunc='count').rename(columns={'Дата обобщения опыта':'Количество'})
    teacher_nmr_df = pd.pivot_table(nmr_df, index=['ФИО','Форма обобщения опыта','Проведено ли обобщение опыта'],
                                      values=['Дата обобщения опыта'],
                                      aggfunc='count').rename(columns={'Дата обобщения опыта':'Количество'})

    form_nmr_df_one_col = pd.pivot_table(nmr_df, index=['Форма обобщения опыта'],
                                      values=['Дата обобщения опыта'],
                                      aggfunc='count').rename(columns={'Дата обобщения опыта':'Количество'})
    form_nmr_df = pd.pivot_table(nmr_df, index=['Форма обобщения опыта','ФИО','Проведено ли обобщение опыта'],
                                      values=['Дата обобщения опыта'],
                                      aggfunc='count').rename(columns={'Дата обобщения опыта':'Количество'})


    bool_nmr_df_one_col = pd.pivot_table(nmr_df, index=['Проведено ли обобщение опыта'],
                                      values=['Дата обобщения опыта'],
                                      aggfunc='count').rename(columns={'Дата обобщения опыта':'Количество'})
    bool_nmr_df = pd.pivot_table(nmr_df, index=['Проведено ли обобщение опыта','Форма обобщения опыта','ФИО'],
                                      values=['Дата обобщения опыта'],
                                      aggfunc='count').rename(columns={'Дата обобщения опыта':'Количество'})

    place_nmr_df_one_col = pd.pivot_table(nmr_df, index=['Место обобщения опыта'],
                                         values=['Дата обобщения опыта'],
                                         aggfunc='count').rename(columns={'Дата обобщения опыта': 'Количество'})
    place_nmr_df = pd.pivot_table(nmr_df, index=['Место обобщения опыта','Форма обобщения опыта', 'ФИО','Проведено ли обобщение опыта'],
                                 values=['Дата обобщения опыта'],
                                 aggfunc='count').rename(columns={'Дата обобщения опыта': 'Количество'})



    # генерируем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    with pd.ExcelWriter(f'{result_folder}/Статистика {current_time}.xlsx') as writer:
        if len(teacher_skills_dev_df_one_col) != 0:
            teacher_skills_dev_df_one_col.to_excel(writer,sheet_name='Повышение квалификации')
        if len(course_skills_dev_df_one_col) != 0:
            course_skills_dev_df_one_col.to_excel(writer,sheet_name='Повышение квалификации',startrow=len(teacher_skills_dev_df_one_col)+3)
        if len(teacher_skills_dev_df) !=0:
            teacher_skills_dev_df.to_excel(writer,sheet_name='Повышение квалификации',startrow=len(teacher_skills_dev_df_one_col)+len(course_skills_dev_df_one_col)+5)
        if len(course_skills_dev_df) != 0:
            course_skills_dev_df.to_excel(writer,sheet_name='Повышение квалификации',startrow=len(teacher_skills_dev_df_one_col)+len(course_skills_dev_df_one_col)+len(teacher_skills_dev_df)+10)
        # Стажировка
        if len(teacher_internship_df_one_col) != 0:
            teacher_internship_df_one_col.to_excel(writer, sheet_name='Стажировка')
        if len(course_internship_df_one_col) != 0:
            course_internship_df_one_col.to_excel(writer, sheet_name='Стажировка',startrow=len(teacher_internship_df_one_col)+3)
        if len(teacher_internship_df) != 0:
            teacher_internship_df.to_excel(writer, sheet_name='Стажировка',startrow=len(teacher_internship_df_one_col)+len(course_internship_df_one_col)+5)
        if len(course_internship_df) != 0:
            course_internship_df.to_excel(writer, sheet_name='Стажировка',startrow=len(teacher_internship_df_one_col)+len(course_internship_df_one_col)+len(teacher_internship_df)+10)
        # Метод разработки
        if len(teacher_method_dev_df_one_col) != 0:
            teacher_method_dev_df_one_col.to_excel(writer, sheet_name='Методические разработки')
        if len(type_method_df_one_col) != 0:
            type_method_df_one_col.to_excel(writer, sheet_name='Методические разработки',startcol=teacher_method_dev_df_one_col.shape[1] + 5)
        if len(prof_method_df_one_col) != 0:
            prof_method_df_one_col.to_excel(writer, sheet_name='Методические разработки',startcol=teacher_method_dev_df_one_col.shape[1] + teacher_method_dev_df_one_col.shape[1]+10)
        max_row = max(len(teacher_method_dev_df_one_col),len(type_method_df_one_col),len(prof_method_df_one_col)) # получаем строку с которой надо начинать запись второго ряда
        if len(teacher_method_dev_df) != 0:
            teacher_method_dev_df.to_excel(writer, sheet_name='Методические разработки',startrow=max_row + 5)
        if len(type_method_df) != 0:
            type_method_df.to_excel(writer, sheet_name='Методические разработки',startrow=max_row + 5,startcol=teacher_method_dev_df.shape[1] + 5)
        if len(prof_method_df) != 0:
            prof_method_df.to_excel(writer, sheet_name='Методические разработки',startcol=teacher_method_dev_df.shape[1]+type_method_df.shape[1]+10,startrow=max_row + 5)
        # Мероприятия проведенные ППС
        if len(teacher_events_df_one_col) != 0:
            teacher_events_df_one_col.to_excel(writer, sheet_name='Мероприятия, пров. ППС')
        if len(level_events_df_one_col) != 0:
            level_events_df_one_col.to_excel(writer, sheet_name='Мероприятия, пров. ППС',startcol=teacher_events_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_events_df_one_col),len(level_events_df_one_col))
        if len(teacher_events_df) != 0:
            teacher_events_df.to_excel(writer, sheet_name='Мероприятия, пров. ППС',startrow=max_row+5)
        if len(level_events_df) != 0:
            level_events_df.to_excel(writer, sheet_name='Мероприятия, пров. ППС',startrow=max_row+5,startcol=teacher_events_df_one_col.shape[1] + 5)
        # Личное выступление ППС
        if len(teacher_personal_perf_df_one_col) != 0:
            teacher_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС')
        if len(course_personal_perf_df_one_col) != 0:
            course_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row = max(len(teacher_personal_perf_df_one_col),len(course_personal_perf_df_one_col))
        if len(teacher_personal_perf_df) != 0:
            teacher_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        if len(course_personal_perf_df) != 0:
            course_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row = max(len(teacher_personal_perf_df)+max_row+5,len(course_personal_perf_df)+max_row+5)
        if len(level_personal_perf_df_one_col) != 0:
            level_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        if len(result_personal_perf_df_one_col) != 0:
            result_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row = max(len(level_personal_perf_df_one_col)+max_row+3, len(result_personal_perf_df_one_col)+max_row+3)
        if len(level_personal_perf_df) != 0:
            level_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        if len(result_personal_perf_df) != 0:
            result_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)
        max_row =  max(len(level_personal_perf_df)+max_row+3, len(result_personal_perf_df)+max_row+3)
        if len(way_personal_perf_df_one_col) != 0:
            way_personal_perf_df_one_col.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5)
        if len(way_personal_perf_df) != 0:
            way_personal_perf_df.to_excel(writer, sheet_name='Личное выступление ППС',startrow=max_row+5,startcol=teacher_personal_perf_df_one_col.shape[1]+5)

        # Публикации
        if len(teacher_publications_df_one_col) != 0:
            teacher_publications_df_one_col.to_excel(writer, sheet_name='Публикации')
        if len(publ_publications_df_one_col) != 0:
            publ_publications_df_one_col.to_excel(writer, sheet_name='Публикации',startcol=teacher_publications_df_one_col.shape[1]+5)
        max_row = max(len(teacher_publications_df_one_col),len(publ_publications_df_one_col))
        if len(teacher_publications_df) != 0:
            teacher_publications_df.to_excel(writer, sheet_name='Публикации',startrow=max_row+3)
        if len(publ_publications_df) != 0:
            publ_publications_df.to_excel(writer, sheet_name='Публикации',startrow=max_row+3,startcol=teacher_publications_df_one_col.shape[1]+5)
        # Открытые уроки
        if len(teacher_open_lessons_df_one_col) != 0:
            teacher_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки')
        if len(type_open_lessons_df_one_col) != 0:
            type_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки',
                                              startcol=teacher_open_lessons_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_open_lessons_df_one_col), len(type_open_lessons_df_one_col))
        if len(teacher_open_lessons_df) != 0:
            teacher_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3)
        if len(type_open_lessons_df) != 0:
            type_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3,
                                      startcol=teacher_open_lessons_df_one_col.shape[1] + 5)

        max_row = max(len(teacher_open_lessons_df)+max_row+5, len(teacher_open_lessons_df)+ max_row+5)
        if len(lesson_open_lessons_df_one_col) != 0:
            lesson_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки',startrow=max_row+3)
        if len(group_open_lessons_df_one_col) != 0:
            group_open_lessons_df_one_col.to_excel(writer, sheet_name='Открытые уроки',startrow=max_row+3,
                                              startcol=teacher_open_lessons_df_one_col.shape[1] + 5)
        max_row = max(len(lesson_open_lessons_df_one_col)+max_row+5, len(group_open_lessons_df_one_col)+max_row+5)
        if len(lesson_open_lessons_df) != 0:
            lesson_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3)
        if len(group_open_lessons_df) != 0:
            group_open_lessons_df.to_excel(writer, sheet_name='Открытые уроки', startrow=max_row + 3,
                                      startcol=teacher_open_lessons_df_one_col.shape[1] + 5)
        # Взаимопосещения
        if len(teacher_mutual_visits_df_one_col) != 0:
            teacher_mutual_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение')
        if len(teacher_visited_visits_df_one_col) != 0:
            teacher_visited_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение',
                                              startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_mutual_visits_df_one_col), len(teacher_visited_visits_df_one_col))
        if len(teacher_mutual_visits_df) != 0:
            teacher_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3)
        if len(teacher_visited_mutual_visits_df) != 0:
            teacher_visited_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3,
                                      startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_mutual_visits_df) + max_row + 5, len(teacher_visited_mutual_visits_df) + max_row + 5)
        if len(group_mutual_visits_df_one_col) != 0:
            group_mutual_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение',startrow=max_row+3)
        if len(theme_mutual_visits_df_one_col) != 0:
            theme_mutual_visits_df_one_col.to_excel(writer, sheet_name='Взаимопосещение',startrow=max_row+3,startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)
        max_row = max(len(group_mutual_visits_df_one_col) + max_row + 5, len(theme_mutual_visits_df_one_col) + max_row + 5)
        if len(group_mutual_visits_df) != 0:
            group_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3)
        if len(theme_mutual_visits_df) != 0:
            theme_mutual_visits_df.to_excel(writer, sheet_name='Взаимопосещение', startrow=max_row + 3,
                                      startcol=teacher_mutual_visits_df_one_col.shape[1] + 5)
        #УИРС
        if len(teacher_student_perf_df_one_col) != 0:
            teacher_student_perf_df_one_col.to_excel(writer, sheet_name='УИРС')
        if len(prof_student_perf_df_one_col) != 0:
            prof_student_perf_df_one_col.to_excel(writer, sheet_name='УИРС',
                                              startcol=teacher_student_perf_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_student_perf_df_one_col), len(prof_student_perf_df_one_col))
        if len(teacher_student_perf_df) != 0:
            teacher_student_perf_df.to_excel(writer, sheet_name='УИРС', startrow=max_row + 3)
        if len(prof_student_perf_df) != 0:
            prof_student_perf_df.to_excel(writer, sheet_name='УИРС', startrow=max_row + 3,
                                      startcol=teacher_student_perf_df_one_col.shape[1] + 5)

        max_row = max(len(teacher_student_perf_df) + max_row + 5, len(prof_student_perf_df) + max_row + 5)
        if len(type_student_perf_df_one_col) != 0:
            type_student_perf_df_one_col.to_excel(writer, sheet_name='УИРС',startrow=max_row+3)
        if len(level_student_perf_df_one_col) != 0:
            level_student_perf_df_one_col.to_excel(writer, sheet_name='УИРС',startrow=max_row+3,startcol=teacher_student_perf_df_one_col.shape[1] + 5)
        max_row = max(len(type_student_perf_df_one_col) + max_row + 5,len(level_student_perf_df_one_col) + max_row + 5)
        if len(type_student_perf_df) != 0:
            type_student_perf_df.to_excel(writer, sheet_name='УИРС', startrow=max_row + 3)
        if len(level_student_perf_df) != 0:
            level_student_perf_df.to_excel(writer, sheet_name='УИРС', startrow=max_row + 3,
                                        startcol=teacher_student_perf_df_one_col.shape[1] + 5)
        # Работа по НМР
        if len(teacher_nmr_df_one_col) != 0:
            teacher_nmr_df_one_col.to_excel(writer, sheet_name='Работа по НМР')
        if len(form_nmr_df_one_col) != 0:
            form_nmr_df_one_col.to_excel(writer, sheet_name='Работа по НМР',
                                              startcol=teacher_nmr_df_one_col.shape[1] + 5)
        max_row = max(len(teacher_nmr_df_one_col), len(form_nmr_df_one_col))
        if len(teacher_nmr_df) != 0:
            teacher_nmr_df.to_excel(writer, sheet_name='Работа по НМР', startrow=max_row + 3)
        if len(form_nmr_df) != 0:
            form_nmr_df.to_excel(writer, sheet_name='Работа по НМР', startrow=max_row + 3,
                                      startcol=teacher_nmr_df_one_col.shape[1] + 5)

        max_row = max(len(teacher_nmr_df) + max_row + 5, len(form_nmr_df) + max_row + 5)
        if len(bool_nmr_df_one_col) != 0:
            bool_nmr_df_one_col.to_excel(writer, sheet_name='Работа по НМР',startrow=max_row+3)
        if len(place_nmr_df_one_col) != 0:
            place_nmr_df_one_col.to_excel(writer, sheet_name='Работа по НМР',startrow=max_row+3,startcol=teacher_nmr_df_one_col.shape[1] + 5)
        max_row = max(len(bool_nmr_df_one_col) + max_row + 5,len(place_nmr_df_one_col) + max_row + 5)
        if len(bool_nmr_df) != 0:
            bool_nmr_df.to_excel(writer, sheet_name='Работа по НМР', startrow=max_row + 3)
        if len(place_nmr_df) != 0:
            place_nmr_df.to_excel(writer, sheet_name='Работа по НМР', startrow=max_row + 3,
                                        startcol=teacher_nmr_df_one_col.shape[1] + 5)
        end_df = pd.DataFrame(columns=['Информация'],
                              data=[[None]])
        end_df.to_excel(writer,sheet_name='1')






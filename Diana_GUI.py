"""
Графический интерфейс для генерации документов методотдела БРИТ
"""
from create_RP_UD import create_RP_for_UD # Функция для генерации РП для УД
from create_RP_UD_OOD import create_RP_for_UD_OOD # функция для генерации РП для ООД
from create_PM import create_pm # функция для генерации программы профессионального модуля
from create_UP_PM import create_rp_up #  функция для генерации РП для УП (учебных практик)
from create_PP_PM import create_rp_pp #  функция для генерации РП для ПП (учебных практик)
from create_PRED_DIP_PRAC import create_pred_dip_prac # функция для генерации рабочей программы преддипломной практики
from create_report_teacher import create_report_teacher # функция для генерации отчетов по преподавателям
import pandas as pd
import sys
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from pandas._libs.tslibs.parsing import DateParseError
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

from jinja2 import exceptions


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
"""
Функции для создания контекстного меню(Копировать,вставить,вырезать)
"""


def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")


def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))


def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)


def on_scroll(*args):
    canvas.yview(*args)


def set_window_size(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Устанавливаем размер окна в 80% от ширины и высоты экрана
    if screen_width >= 3840:
        width = int(screen_width * 0.28)
    elif screen_width >= 2560:
        width = int(screen_width * 0.39)
    elif screen_width >= 1920:
        width = int(screen_width * 0.48)
    elif screen_width >= 1600:
        width = int(screen_width * 0.58)
    elif screen_width >= 1280:
        width = int(screen_width * 0.70)
    elif screen_width >= 1024:
        width = int(screen_width * 0.85)
    else:
        width = int(screen_width * 1)

    height = int(screen_height * 0.8)

    # Рассчитываем координаты для центрирования окна
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размер и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")

def select_folder_data():
    """
    Функция для выбора папки c данными
    :return:
    """
    global path_folder_data
    path_folder_data = filedialog.askdirectory()

def select_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()

def select_file_docx():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_template
    file_template = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_file_data_xlsx():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx
    # Получаем путь к файлу
    file_data_xlsx = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


"""
Функции для ООД
"""
def select_end_folder_ood():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_ood
    path_to_end_folder_ood = filedialog.askdirectory()

def select_file_docx_ood():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_template_ood
    file_template_ood = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_file_data_xlsx_ood():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx_ood
    # Получаем путь к файлу
    file_data_xlsx_ood = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_files_data_xlsx():
    """
    Функция для выбора нескоьких файлов с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global files_data_xlsx
    # Получаем путь файлы
    files_data_xlsx = filedialog.askopenfilenames(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

"""
Функции для создания проф модуля
"""
def select_end_folder_pm():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_pm
    path_to_end_folder_pm = filedialog.askdirectory()

def select_file_docx_pm():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_template_pm
    file_template_pm = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_file_data_xlsx_pm():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx_pm
    # Получаем путь к файлу
    file_data_xlsx_pm = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

"""
Функции для создания рабочей программы учебной практики профмодуля
"""
def select_end_folder_up_pm():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_up_pm
    path_to_end_folder_up_pm = filedialog.askdirectory()

def select_file_docx_up_pm():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_template_up_pm
    file_template_up_pm = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_file_data_xlsx_up_pm():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx_up_pm
    # Получаем путь к файлу
    file_data_xlsx_up_pm = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

"""
Функции для создания рабочей программы производственной практики профмодуля
"""
def select_end_folder_pp_pm():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_pp_pm
    path_to_end_folder_pp_pm = filedialog.askdirectory()

def select_file_docx_pp_pm():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_template_pp_pm
    file_template_pp_pm = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_file_data_xlsx_pp_pm():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx_pp_pm
    # Получаем путь к файлу
    file_data_xlsx_pp_pm = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


"""
Функции для создания рабочей программы преддипломной практики 
"""
def select_end_folder_pp_prac():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_pp_prac
    path_to_end_folder_pp_prac = filedialog.askdirectory()

def select_file_docx_pp_prac():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_template_pp_prac
    file_template_pp_prac = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_file_data_xlsx_pp_prac():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx_pp_prac
    # Получаем путь к файлу
    file_data_xlsx_pp_prac = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


"""
Функции для генерации отчетов по преподавателям
"""
def select_templates_folder_report_teacher():
    """
    Функция для выбора папки с шаблонами
    :return:
    """
    global path_to_templates_folder_report_teacher
    path_to_templates_folder_report_teacher = filedialog.askdirectory()


def select_data_folder_report_teacher():
    """
    Функция для выбора папки с шаблонами
    :return:
    """
    global path_to_data_folder_report_teacher
    path_to_data_folder_report_teacher = filedialog.askdirectory()


def select_end_folder_report_teacher():
    """
    Функция для выбора папки с шаблонами
    :return:
    """
    global path_to_end_folder_report_teacher
    path_to_end_folder_report_teacher = filedialog.askdirectory()






def processing_create_RP_for_UD():
    """
    Фугкция для создания рабочей программы для учебной дисциплины
    :return:
    """
    try:
        create_RP_for_UD(file_template,file_data_xlsx,path_to_end_folder)
    except NameError:
        messagebox.showerror('Диана',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

def processing_create_RP_for_OOD():
    """
    Фугкция для создания рабочей программы для учебной  общеобразовательной дисциплины
    :return:
    """
    try:
        create_RP_for_UD_OOD(file_template_ood,file_data_xlsx_ood,path_to_end_folder_ood)
    except NameError:
        messagebox.showerror('Диана',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_create_RP_for_PM():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_pm(file_template_pm,file_data_xlsx_pm,path_to_end_folder_pm)
    except NameError:
        messagebox.showerror('Диана',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

def processing_create_RP_for_UP():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_rp_up(file_template_up_pm,file_data_xlsx_up_pm,path_to_end_folder_up_pm)
    except NameError:
        messagebox.showerror('Диана',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

def processing_create_RP_for_PP():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_rp_pp(file_template_pp_pm,file_data_xlsx_pp_pm,path_to_end_folder_pp_pm)
    except NameError:
        messagebox.showerror('Диана',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_create_RP_for_PP_prac():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_pred_dip_prac(file_template_pp_prac,file_data_xlsx_pp_prac,path_to_end_folder_pp_prac)
    except NameError:
        messagebox.showerror('Диана',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_create_report_teacher():
    """
    Фугкция для создания отчетов по преподавателям
    :return:
    """
    try:
        start_date = var_start_date.get() # начальная дата
        end_date = var_end_date.get() # конечная дата

        # Обрабатываем даты диапазона
        # Если ничего
        if not start_date:
            start_date = '01.01.1900'
        if not end_date:
            end_date = '01.01.2100'
        start_date = pd.to_datetime(start_date, dayfirst=True, errors='raise')
        end_date = pd.to_datetime(end_date, dayfirst=True, errors='raise')


        create_report_teacher(path_to_templates_folder_report_teacher,path_to_data_folder_report_teacher,path_to_end_folder_report_teacher,start_date,end_date)
    except NameError:
        messagebox.showerror('Диана',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except DateParseError:
        messagebox.showerror('Диана',
                             f'Введено некорректное значение начальной или конечной даты.\n Если вам нужен конкретный диапазон '
                             f'Вводите даты в формате 14.06.2024')





if __name__ == '__main__':
    window = Tk()
    window.title('Диана Автоматизация работы методиста СПО ver 3.0')
    # Устанавливаем размер и положение окна
    set_window_size(window)
    window.resizable(True, True)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)

    # Создаем вертикальный скроллбар
    scrollbar = Scrollbar(window, orient="vertical")

    # Создаем холст
    canvas = Canvas(window, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    # Привязываем скроллбар к холсту
    scrollbar.config(command=canvas.yview)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_rp_for_ud = ttk.Frame(tab_control)
    tab_control.add(tab_rp_for_ud, text='Создание РП\n для УД')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание РП для УД
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_rp_for_ud,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание рабочей программы для учебной дисциплины\n с помощью шаблона')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_rp_for_ud,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон
    btn_choose_data = Button(tab_rp_for_ud, text='1) Выберите шаблон РП УД', font=('Arial Bold', 20),
                             command=select_file_docx
                             )
    btn_choose_data.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными для шаблона
    btn_choose_data = Button(tab_rp_for_ud, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                             command=select_file_data_xlsx
                             )
    btn_choose_data.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_rp_for_ud, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=4, padx=10, pady=10)

    #Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_rp_for_ud, text='4) Обработать данные', font=('Arial Bold', 20),
                                  command=processing_create_RP_for_UD
                                  )
    btn_proccessing_data.grid(column=0, row=5, padx=10, pady=10)

    """
    Интерфейс для УД ООД
    """
    tab_rp_for_ood = ttk.Frame(tab_control)
    tab_control.add(tab_rp_for_ood, text='Создание РП\n для ООД')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание РП для ООД
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_rp_for_ood,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание рабочей программы для ООД с помощью шаблона')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_ood = resource_path('logo.png')

    img_ood = PhotoImage(file=path_to_img_ood)
    Label(tab_rp_for_ood,
          image=img_ood
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон
    btn_choose_data = Button(tab_rp_for_ood, text='1) Выберите шаблон РП ООД', font=('Arial Bold', 20),
                             command=select_file_docx_ood
                             )
    btn_choose_data.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными для шаблона
    btn_choose_data = Button(tab_rp_for_ood, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                             command=select_file_data_xlsx_ood
                             )
    btn_choose_data.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_rp_for_ood, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder_ood
                                   )
    btn_choose_end_folder.grid(column=0, row=4, padx=10, pady=10)

    #Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_rp_for_ood, text='4) Обработать данные', font=('Arial Bold', 20),
                                  command=processing_create_RP_for_OOD
                                  )
    btn_proccessing_data.grid(column=0, row=5, padx=10, pady=10)

    """
    Интерфейс для РП профессионального модуля
    """
    tab_rp_for_pm = ttk.Frame(tab_control)
    tab_control.add(tab_rp_for_pm, text='Создание РП\n для ПМ')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание РП для ООД
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_rp_for_pm,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание рабочей программы для профессионального модуля\n с помощью шаблона')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_pm = resource_path('logo.png')

    img_pm = PhotoImage(file=path_to_img_pm)
    Label(tab_rp_for_pm,
          image=img_pm
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон
    btn_choose_template_pm = Button(tab_rp_for_pm, text='1) Выберите шаблон РП ПМ', font=('Arial Bold', 20),
                                    command=select_file_docx_pm
                                    )
    btn_choose_template_pm.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными для шаблона
    btn_choose_data_pm = Button(tab_rp_for_pm, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                                command=select_file_data_xlsx_pm
                                )
    btn_choose_data_pm.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_pm = Button(tab_rp_for_pm, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                      command=select_end_folder_pm
                                      )
    btn_choose_end_folder_pm.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_pm = Button(tab_rp_for_pm, text='4) Обработать данные', font=('Arial Bold', 20),
                                command=processing_create_RP_for_PM
                                )
    btn_proccessing_pm.grid(column=0, row=5, padx=10, pady=10)

    """
    Интерфейс для РП учебной практики профессионального модуля
    """
    tab_rp_for_up_pm = ttk.Frame(tab_control)
    tab_control.add(tab_rp_for_up_pm, text='Создание РП\n для УП')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание РП для УП
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_rp_for_up_pm,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание рабочей программы для учебной практики ПМ\n с помощью шаблона')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_up_pm = resource_path('logo.png')

    img_up_pm = PhotoImage(file=path_to_img_up_pm)
    Label(tab_rp_for_up_pm,
          image=img_up_pm
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон
    btn_choose_template_up_pm = Button(tab_rp_for_up_pm, text='1) Выберите шаблон РП УП', font=('Arial Bold', 20),
                                       command=select_file_docx_up_pm
                                       )
    btn_choose_template_up_pm.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными для шаблона
    btn_choose_data_up_pm = Button(tab_rp_for_up_pm, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                                   command=select_file_data_xlsx_up_pm
                                   )
    btn_choose_data_up_pm.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_up_pm = Button(tab_rp_for_up_pm, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                         command=select_end_folder_up_pm
                                         )
    btn_choose_end_folder_up_pm.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_up_pm = Button(tab_rp_for_up_pm, text='4) Обработать данные', font=('Arial Bold', 20),
                                   command=processing_create_RP_for_UP
                                   )
    btn_proccessing_up_pm.grid(column=0, row=5, padx=10, pady=10)

    """
    Интерфейс для РП производственной практики профессионального модуля
    """
    tab_rp_for_pp_pm = ttk.Frame(tab_control)
    tab_control.add(tab_rp_for_pp_pm, text='Создание РП\n для ПП')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание РП для УП
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_rp_for_pp_pm,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание рабочей программы для производственной практики ПМ\n с помощью шаблона')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_pp_pm = resource_path('logo.png')

    img_pp_pm = PhotoImage(file=path_to_img_pp_pm)
    Label(tab_rp_for_pp_pm,
          image=img_pp_pm
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон
    btn_choose_template_pp_pm = Button(tab_rp_for_pp_pm, text='1) Выберите шаблон РП ПП', font=('Arial Bold', 20),
                                       command=select_file_docx_pp_pm
                                       )
    btn_choose_template_pp_pm.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными для шаблона
    btn_choose_data_pp_pm = Button(tab_rp_for_pp_pm, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                                   command=select_file_data_xlsx_pp_pm
                                   )
    btn_choose_data_pp_pm.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_pp_pm = Button(tab_rp_for_pp_pm, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                         command=select_end_folder_pp_pm
                                         )
    btn_choose_end_folder_pp_pm.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_pp_pm = Button(tab_rp_for_pp_pm, text='4) Обработать данные', font=('Arial Bold', 20),
                                   command=processing_create_RP_for_PP
                                   )
    btn_proccessing_pp_pm.grid(column=0, row=5, padx=10, pady=10)

    """
    Интерфейс для РП преддипломной практики
    """
    tab_rp_for_pp_prac = ttk.Frame(tab_control)
    tab_control.add(tab_rp_for_pp_prac, text='Создание РП\n для ПдП')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание РП для УП
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_rp_for_pp_prac,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание рабочей программы для преддипломной практики\n с помощью шаблона')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_pp_prac = resource_path('logo.png')

    img_pp_prac = PhotoImage(file=path_to_img_pp_prac)
    Label(tab_rp_for_pp_prac,
          image=img_pp_prac
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон
    btn_choose_template_pp_prac = Button(tab_rp_for_pp_prac, text='1) Выберите шаблон РП ПдП', font=('Arial Bold', 20),
                                         command=select_file_docx_pp_prac
                                         )
    btn_choose_template_pp_prac.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными для шаблона
    btn_choose_data_pp_prac = Button(tab_rp_for_pp_prac, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                                     command=select_file_data_xlsx_pp_prac
                                     )
    btn_choose_data_pp_prac.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_pp_prac = Button(tab_rp_for_pp_prac, text='3) Выберите конечную папку',
                                           font=('Arial Bold', 20),
                                           command=select_end_folder_pp_prac
                                           )
    btn_choose_end_folder_pp_prac.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_pp_prac = Button(tab_rp_for_pp_prac, text='4) Обработать данные', font=('Arial Bold', 20),
                                     command=processing_create_RP_for_PP_prac
                                     )
    btn_proccessing_pp_prac.grid(column=0, row=5, padx=10, pady=10)

    tab_for_report_teacher = ttk.Frame(tab_control)
    tab_control.add(tab_for_report_teacher, text='Отчеты по\n преподавателям')
    tab_control.pack(expand=1, fill='both')
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_for_report_teacher,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание отчетов и личных дел для преподавателей')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_report_teacher = resource_path('logo.png')

    img_report_teacher = PhotoImage(file=path_to_img_report_teacher)
    Label(tab_for_report_teacher,
          image=img_report_teacher
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными для шаблона
    btn_choose_template_report_teacher = Button(tab_for_report_teacher, text='1) Выберите папку с шаблонами',
                                            font=('Arial Bold', 20),
                                            command=select_templates_folder_report_teacher
                                            )
    btn_choose_template_report_teacher.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки c данными в формате xlsx

    btn_choose_data_folder_report_teacher = Button(tab_for_report_teacher, text='2) Выберите папку с данными',
                                                  font=('Arial Bold', 20),
                                                  command=select_data_folder_report_teacher
                                                  )
    btn_choose_data_folder_report_teacher.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_report_teacher = Button(tab_for_report_teacher, text='3) Выберите конечную папку',
                                                  font=('Arial Bold', 20),
                                                  command=select_end_folder_report_teacher
                                                  )
    btn_choose_end_folder_report_teacher.grid(column=0, row=5, padx=10, pady=10)

    # Определяем текстовую переменную для стартовой даты
    var_start_date = StringVar()
    # Описание поля
    label_start_date = Label(tab_for_report_teacher,
                                   text='Введите начальную дату диапазона за который вы хотите получить данные в формате: 10.05.2024\n'
                                        'Если вы ничего не введете то начальной датой будет считаться 01.01.1900')
    label_start_date.grid(column=0, row=6, padx=10, pady=10)
    # поле ввода
    entry_start_date = Entry(tab_for_report_teacher, textvariable=var_start_date, width=30)
    entry_start_date.grid(column=0, row=7, padx=10, pady=10)

    # Определяем текстовую переменную для конечной даты
    var_end_date = StringVar()
    # Описание поля
    label_end_date = Label(tab_for_report_teacher,
                                   text='Введите конечную дату диапазона за который вы хотите получить данные в формате: 10.05.2024\n'
                                        'Если вы ничего не введете то конечной датой будет считаться 01.01.2100')
    label_end_date.grid(column=0, row=8, padx=10, pady=10)
    # поле ввода
    entry_end_date = Entry(tab_for_report_teacher, textvariable=var_end_date, width=30)
    entry_end_date.grid(column=0, row=9, padx=10, pady=10)


    # Создаем кнопку обработки данных

    btn_proccessing_report_teacher = Button(tab_for_report_teacher, text='4) Обработать данные',
                                            font=('Arial Bold', 20),
                                            command=processing_create_report_teacher
                                            )
    btn_proccessing_report_teacher.grid(column=0, row=10, padx=10, pady=10)





    # Создаем виджет для управления полосой прокрутки
    canvas.create_window((0, 0), window=tab_control, anchor="nw")

    # Конфигурируем холст для обработки скроллинга
    canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
    scrollbar.pack(side="right", fill="y")

    # Вешаем событие скроллинга
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)
    window.mainloop()
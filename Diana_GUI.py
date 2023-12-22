"""
Графический интерфейс для генерации документов методотдела БРИТ
"""
from create_RP_UD import create_RP_for_UD # Функция для генерации РП для УД
from create_RP_UD_OOD import create_RP_for_UD_OOD # функция для генерации РП для ООД
from create_PM import create_pm # функция для генерации программы профессионального модуля
from create_UP_PM import create_rp_up #  функция для генерации РП для УП (учебных практик)
from create_PP_PM import create_rp_pp #  функция для генерации РП для ПП (учебных практик)
from create_PRED_DIP_PRAC import create_pred_dip_prac # функция для генерации рабочей программы преддипломной практики
import tkinter
import sys
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time
# pd.options.mode.chained_assignment = None  # default='warn'
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




def processing_create_RP_for_UD():
    """
    Фугкция для создания рабочей программы для учебной дисциплины
    :return:
    """
    try:
        create_RP_for_UD(file_template,file_data_xlsx,path_to_end_folder)
    except NameError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

def processing_create_RP_for_OOD():
    """
    Фугкция для создания рабочей программы для учебной  общеобразовательной дисциплины
    :return:
    """
    try:
        create_RP_for_UD_OOD(file_template_ood,file_data_xlsx_ood,path_to_end_folder_ood)
    except NameError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_create_RP_for_PM():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_pm(file_template_pm,file_data_xlsx_pm,path_to_end_folder_pm)
    except NameError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

def processing_create_RP_for_UP():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_rp_up(file_template_up_pm,file_data_xlsx_up_pm,path_to_end_folder_up_pm)
    except NameError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

def processing_create_RP_for_PP():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_rp_pp(file_template_pp_pm,file_data_xlsx_pp_pm,path_to_end_folder_pp_pm)
    except NameError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_create_RP_for_PP_prac():
    """
    Фугкция для создания рабочей программы для профессионального модуля
    :return:
    """
    try:
        create_pred_dip_prac(file_template_pp_prac,file_data_xlsx_pp_prac,path_to_end_folder_pp_prac)
    except NameError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')



if __name__ == '__main__':
    window = Tk()
    window.title('Диана Создание рабочих программ ver 2.7')
    window.geometry('800x760')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_rp_for_ud = ttk.Frame(tab_control)
    tab_control.add(tab_rp_for_ud, text='Создание РП для УД')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание РП для УД
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_rp_for_ud,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Создание рабочей программы для учебной дисциплины с помощью шаблона')
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
    tab_control.add(tab_rp_for_ood, text='Создание РП для ООД')
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
    tab_control.add(tab_rp_for_pm, text='Создание РП для ПМ')
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
    tab_control.add(tab_rp_for_up_pm, text='Создание РП для УП')
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
    tab_control.add(tab_rp_for_pp_pm, text='Создание РП для ПП')
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
    tab_control.add(tab_rp_for_pp_prac, text='Создание РП для ПдП')
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



    window.mainloop()
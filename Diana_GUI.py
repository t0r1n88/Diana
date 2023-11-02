"""
Графический интерфейс для генерации документов методотдела БРИТ
"""
from create_RP_UD import create_RP_for_UD # Функция для генерации РП для УД
from create_RP_UD_OOD import create_RP_for_UD_OOD # функция для генерации РП для ООД
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
    except KeyError as e:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'В таблице не найдена колонка с названием {e.args}!\nПроверьте написание названия колонки')
    except ValueError as e:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'В таблице не найден лист с названием {e.args}!\nПроверьте написание названия листа')

    except FileNotFoundError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except exceptions.TemplateSyntaxError:
        messagebox.showerror('Диана Создание рабочих программ',
                             f'Ошибка в оформлении вставляемых значений в шаблоне\n'
                             f'Проверьте свой шаблон на наличие следующих ошибок:\n'
                             f'1) Вставляемые значения должны быть оформлены двойными фигурными скобками\n'
                             f'{{{{Вставляемое_значение}}}}\n'
                             f'2) В названии колонки в таблице откуда берутся данные - есть пробелы,цифры,знаки пунктуации и т.п.\n'
                             f'в названии колонки должны быть только буквы и нижнее подчеркивание.\n'
                             f'{{{{Дата_рождения}}}}')


    else:
        messagebox.showinfo('Диана Создание рабочих программ', 'Данные успешно обработаны')

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



if __name__ == '__main__':
    window = Tk()
    window.title('Диана Создание рабочих программ ver 1.6')
    window.geometry('700x860')
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






    window.mainloop()
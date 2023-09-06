"""
Графический интерфейс для генерации документов методотдела БРИТ
"""
from create_RP_UD import create_RP_for_UD
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
    else:
        messagebox.showinfo('Диана Создание рабочих программ', 'Данные успешно обработаны')


if __name__ == '__main__':
    window = Tk()
    window.title('Диана Создание рабочих программ ver 1.0')
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
    btn_choose_data = Button(tab_rp_for_ud, text='1) Выберите шаблон', font=('Arial Bold', 20),
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

    window.mainloop()
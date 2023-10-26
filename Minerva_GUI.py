"""
Графический интерфейс
"""
from create_program_prob import generate_program_prob # Скрипт генерации программ профпроб
from base_report_bvb import main_create_report # Скрипт для генерации некоторых отчетов по Билету в будущее
import numpy as np
import os
import pandas as pd
from dateutil.parser import ParserError
from docxtpl import DocxTemplate
from docx2pdf import convert
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import tkinter
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl import load_workbook
import time
import datetime
import warnings
from collections import Counter
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.options.mode.chained_assignment = None
import sys
import locale
import logging
logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)

import re

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def select_file_template_doc():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_doc
    name_file_template_doc = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_doc():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_doc
    # Получаем путь к файлу
    name_file_data_doc = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_doc():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_doc
    path_to_end_folder_doc = filedialog.askdirectory()

def generate_file_profprobs():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных Веста Обработка таблиц и создание документов ver 1.29)
    :return:
    """
    # получаем состояние чекбокса создания pdf
    mode_pdf = mode_pdf_value.get()

    try:
        generate_program_prob(name_file_template_doc,name_file_data_doc,path_to_end_folder_doc,mode_pdf)

    except NameError as e:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')

"""
Функции для создания отчетов по профпробам
"""
def select_file_data_students():
    """
    Функция для выбора файла с данными учеников на основе которых будут генеририроваться отчеты
    :return: Путь к файлу с данными
    """
    global name_file_data_students
    # Получаем путь к файлу
    name_file_data_students = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_students():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_students
    path_to_end_folder_students = filedialog.askdirectory()

def processing_report_students():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных Веста Обработка таблиц и создание документов ver 1.29)
    :return:
    """

    try:
        main_create_report(name_file_data_students,path_to_end_folder_students)

    except NameError as e:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'Выберите файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError as e:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'В таблице не найдена указанная колонка {e.args}')
    except PermissionError:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'Закройте все файлы созданные Минервой')
    except FileNotFoundError:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    else:
        messagebox.showinfo('Минерва Отчеты и пробы Билет в будущее',
                            'Создание отчетов завершено!')



if __name__ == '__main__':
    window = Tk()
    window.title('Минерва Отчеты и пробы Билет в будущее ver 1.2')
    window.geometry('750x860')
    window.resizable(False, False)

    tab_control = ttk.Notebook(window)

    """
    Создаем вкладку для создания отчетов по профпробам
    """
    tab_create_report_bvb = ttk.Frame(tab_control)
    tab_control.add(tab_create_report_bvb, text='Отчеты по БВБ')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Создание документов
    # Создаем метку для описания назначения программы
    lbl_hello_report = Label(tab_create_report_bvb,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nСоздание отчетов по БВБ')
    lbl_hello_report.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_report = resource_path('logo.png')
    img_report = PhotoImage(file=path_to_img_report)
    Label(tab_create_report_bvb,
          image=img_report
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_report = LabelFrame(tab_create_report_bvb, text='Подготовка')
    frame_data_for_report.grid(column=0, row=2, padx=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_report = Button(frame_data_for_report, text='1) Выберите файл с учениками', font=('Arial Bold', 15),
                          command=select_file_data_students
                          )
    btn_data_report.grid(column=0, row=4, padx=10, pady=10)
    #
    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_report = Button(frame_data_for_report, text='2) Выберите конечную папку', font=('Arial Bold', 15),
                                       command=select_end_folder_students
                                       )
    btn_choose_end_folder_report.grid(column=0, row=5, padx=10, pady=10)

    # Создаем кнопку генерации

    btn_processing_report = Button(frame_data_for_report, text='3) Создать отчеты', font=('Arial Bold', 15),
                             command=processing_report_students
                             )
    btn_processing_report.grid(column=0, row=6, padx=10, pady=10)


    """
    Создание программ профпроб
    """

    tab_create_program_prob = ttk.Frame(tab_control)
    tab_control.add(tab_create_program_prob, text='Создание программ профпроб')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Создание документов
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_create_program_prob,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nГенерация программ профпроб')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img = PhotoImage(file=path_to_img)
    Label(tab_create_program_prob,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_doc = LabelFrame(tab_create_program_prob, text='Подготовка')
    frame_data_for_doc.grid(column=0, row=2, padx=10)

    # Создаем кнопку Выбрать шаблон
    btn_template_doc = Button(frame_data_for_doc, text='1) Выберите шаблон документа', font=('Arial Bold', 15),
                              command=select_file_template_doc
                              )
    btn_template_doc.grid(column=0, row=3, padx=10, pady=10)
    #
    # Создаем кнопку Выбрать файл с данными
    btn_data_doc = Button(frame_data_for_doc, text='2) Выберите файл с данными', font=('Arial Bold', 15),
                          command=select_file_data_doc
                          )
    btn_data_doc.grid(column=0, row=4, padx=10, pady=10)
    #
    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_doc = Button(frame_data_for_doc, text='3) Выберите конечную папку', font=('Arial Bold', 15),
                                       command=select_end_folder_doc
                                       )
    btn_choose_end_folder_doc.grid(column=0, row=9, padx=10, pady=10)

    # Создаем область для того чтобы поместить туда опции
    frame_data_for_options = LabelFrame(tab_create_program_prob, text='Дополнительные опции')
    frame_data_for_options.grid(column=0, row=10, padx=10)

    # Создаем чекбокс для режима создания pdf
    # Создаем переменную для хранения результа переключения чекбокса
    mode_pdf_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_pdf_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_pdf = Checkbutton(frame_data_for_options,
                                       text='Поставьте галочку, если вам нужно чтобы \n'
                                            'дополнительно создавались pdf версии программ',
                                       variable=mode_pdf_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_pdf.grid(column=0, row=12, padx=1, pady=1)


    # Создаем кнопку для создания документов из таблиц с произвольной структурой
    btn_create_files_other = Button(tab_create_program_prob, text='4) Создать программы профпроб',
                                    font=('Arial Bold', 15),
                                    command=generate_file_profprobs
                                    )
    btn_create_files_other.grid(column=0, row=16, padx=10, pady=10)

    window.mainloop()
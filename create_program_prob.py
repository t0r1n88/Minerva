"""
Скрипт для генерации программ профпроб
"""
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
from collections import Counter
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
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




def generate_program_prob(name_file_template_doc:str,name_file_data_doc:str,path_to_end_folder_doc,mode_pdf):
    """
    Скрипт генерирующей из шаблона Word name_file_template_doc и файла с данными по профпробам name_file_data_doc  программы профпроб
    с указанным форматированием
    :param name_file_template_doc:
    :param name_file_data_doc:
    :param path_to_end_folder_doc:
    :param mode_pdf:
    :return:
    """
    try:
        df = pd.read_excel(name_file_data_doc, dtype=str)

        df.fillna('error', inplace=True) # заполняем Nan
        # соединяем значения из колонок возрастных категорий
        # df['Возрастная_категория'] = df.iloc[:, 13:16].apply(lambda row: ';'.join(row), axis=1)
        df['Возрастная_категория'] = df.iloc[:, 13:16].apply(lambda row: ';\n'.join(row), axis=1)

        df['Возрастная_категория'] = df['Возрастная_категория'].apply(
            lambda x: re.sub(r'error[;]?', '', x))  # очищаем от лишнего

        df['Возрастная_категория'] = df['Возрастная_категория'].apply(lambda x: re.sub(r'[;]?$', '', x))
        df['Возрастная_категория'] = df['Возрастная_категория'].apply(lambda x:x.strip()) # убираем знаки переноса


        # Соединияем значения из колонок с нозологиями
        df['Допустимые_нозологии'] = df.iloc[:, 16:25].apply(lambda row: ';'.join(row), axis=1)

        df['Допустимые_нозологии'] = df['Допустимые_нозологии'].apply(
            lambda x: re.sub(r'error[;]?', '', x))  # очищаем строку от error

        df['Допустимые_нозологии'] = df['Допустимые_нозологии'].apply(lambda x: re.sub(r'[;]?$', '', x))

        df = df.drop(df.columns[13:16], axis=1)  # последовательно удаляем лишние колонки

        df = df.drop(df.columns[13:22], axis=1)  # последовательно удаляем лишние колонки

        lst_name_columns = ['ID', 'Время_создания', 'Профессиональная_проба',
                            'Наименование_профессионального_направления', 'ФИО',
                            'Должность', 'Регион', 'Город', 'Электронная_почта', 'Контактный_телефон',
                            'Уровень_сложности', 'Формат_проведения', 'Время_проведения',
                            'Спец_условия', 'Возможность_проведения', 'Краткое_описание', 'Перспективы',
                            'Навыки_знания', 'Интересные_факты',
                            'Связь_пробы', 'Постановка_цели', 'Демонстрация', 'Инструкция', 'Рекомендации_организация',
                            'Критерии', 'Рекомендации_контроль',
                            'Рефлексия', 'Инфраструктурный_лист', 'Доп_источники', 'Доп_файлы', 'Возрастная_категория',
                            'Допустимые_нозологии']

        df.columns = lst_name_columns

        df = df.applymap(lambda x: x.replace("\u00A0", " "))  # удаляем символ неразрывного пробела
        df = df.applymap(lambda x: x.replace("_x000D_", ""))  # удаляем
        df = df.applymap(lambda x: x.replace("error", ""))  # очищаем от слова error меняя на пустую строку

        # Конвертируем датафрейм в список словарей
        data = df.to_dict('records')
        error_df = pd.DataFrame(columns=['Автор', 'Наименование', 'Ошибка'])

        for idx, row in enumerate(data):
            flag_error = False
            name_author = row['ФИО']
            name_prob = row['Наименование_профессионального_направления']
            inf_lst = row['Инфраструктурный_лист'].split('*')

            inf_lst = list(map(str.strip, inf_lst))  # очищаем от пробельных символов
            inf_lst = [value for value in inf_lst if value]  # очищаем от пустого пробела в конце списка
            for value in inf_lst:
                tmp_lst = value.split(';')
                tmp_lst = [val for val in tmp_lst if val]
                if len(tmp_lst) != 4:
                    error_df.loc[len(error_df.index)] = [name_author, name_prob,
                                                         'Ошибка в инфраструктурном листе. Не хватает значений. Проверьте значение соответствующей колонки в таблице\n'
                                                         'Каждые четыре значения должны разделятся символом звездочка (*)\n'
                                                         'Между собой значения должны разделятся точкой с запятой.']
                    flag_error =True
            split_data = [item.split(';') for item in inf_lst]  # создаем список списков
            # создаем датафрейм для хранения инфраструктурника
            if len(split_data) != 4:
                if not flag_error:
                    error_df.loc[len(error_df.index)] = [name_author, name_prob,
                                                         'Ошибка в инфраструктурном листе. Не хватает значений. Проверьте значение соответствующей колонки в таблице\n'
                                                         'Каждые четыре значения должны разделятся символом звездочка (*)\n'
                                                         'Между собой значения должны разделятся точкой с запятой.']
                inf_df =pd.DataFrame(data=[['Проверьте правильность заполнения поля с данными инфраструктурного листа для этой пробы','Ошибка','Ошибка','Ошибка']],columns=['Наименование', 'Характеристика', 'Количество', 'Распределение'])
            else:
                inf_df = pd.DataFrame(split_data, columns=['Наименование', 'Характеристика', 'Количество', 'Распределение'])
            # Обрабатываем дополнительные ссылки
            url_lst = row['Доп_источники'].split(';')
            url_lst = list(map(str.strip, url_lst))  # очищаем от пробельных символов
            url_lst = list(map(lambda x: x.replace('•\t', ''), url_lst))  # очищаем от пробельных символов

            url_lst = [value for value in url_lst if value]  # очищаем от пустого пробела в конце списка

            doc = DocxTemplate(name_file_template_doc)
            context = row
            context['inf_lst'] = inf_df.to_dict('records')
            context['url_lst'] = url_lst

            doc.render(context)
            name_file = f'{name_prob} {name_author}'
            name_file = re.sub(r'[<> :"?*|\\/]', ' ', name_file)
            # проверяем файл на наличие, если файл с таким названием уже существует то добавляем окончание
            if os.path.exists(f'{path_to_end_folder_doc}/{name_file}.docx'):
                doc.save(f'{path_to_end_folder_doc}/{name_file}_{idx}.docx')
            else:
                doc.save(f'{path_to_end_folder_doc}/{name_file}.docx')
            # создаем pdf
            if mode_pdf == 'Yes':
                if os.path.exists(f'{path_to_end_folder_doc}/{name_file}.pdf'):
                    convert(f'{path_to_end_folder_doc}/{name_file}.docx', f'{path_to_end_folder_doc}/{name_file}_{idx}.pdf',
                            keep_active=True)
                else:
                    convert(f'{path_to_end_folder_doc}/{name_file}.docx', f'{path_to_end_folder_doc}/{name_file}.pdf',
                        keep_active=True)

        error_df.to_excel(f'{path_to_end_folder_doc}/Файлы в которых есть ошибки.xlsx', index=False)
    except NameError as e:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError as e:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'В таблице не найдена указанная колонка {e.args}')
    except PermissionError:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'Закройте все файлы созданные Минервой')
        logging.exception('AN ERROR HAS OCCURRED')
    except FileNotFoundError:
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')

    else:
        if error_df.shape[0] != 0:
            messagebox.showerror('Минерва Отчеты и пробы Билет в будущее',
                                 f'В некоторых файлах обнаружены проблемы. Проверьте данные для указанных в файле Ошибки ФИО и названий проб ')
        else:
            messagebox.showinfo('Минерва Отчеты и пробы Билет в будущее',
                            'Создание документов завершено!')


if __name__ == '__main__':
    path_template_main = 'data/Создание программ профроб/Шаблон профессиональной пробы 19.06.docx'
    path_data_main = 'data/Создание программ профроб/probs.xlsx'
    path_end_main = 'data/Создание программ профроб/result'
    mode_pdf_main = 'No'
    generate_program_prob(path_template_main,path_data_main,path_end_main,mode_pdf_main)

    print('Lindy Booth')
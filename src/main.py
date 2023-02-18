#!/usr/bin/python3
# -*- coding: utf-8 -*-
#
#   Author          : Viacheslav Zamaraev
#   email           : zamaraev@gmail.com
#   Script Name     : main.py.py
#   Created         : 18.02.2020
#   Last Modified	: 18.02.2020
#   Version		    : 1.0
#   PIP             : pip install
#   RESULT          : csv file with columns: FIO;...COMM2
# Modifications	: 1.1 -
#               : 1.2 -
#
# Description   : find resume (*.doc) and make xlsr


# import pythoncom
# import textract
# import fnmatch
# import pythoncom
import os
import win32com.client
import re
import csv
import pandas as pd

from src import cfg
from src.log import set_logger


def get_list_files_by_ext(folder_start: str = '', ext: str = 'txt'):
    info_doc = []
    myDir = folder_start
    log = set_logger('find_files.log')
    for subdir, dirs, files in os.walk(myDir):
        for file in files:
            file_path = subdir + os.path.sep + file
            file_to_seek = str(file).lower()
            if file_to_seek[-4:] != ext:
                continue
            info_doc.append(file_path)
            str_q = 'Found: ' + file_path
            print(str_q)
            log.info(str_q)
    return info_doc


def doc2txt(file_path=''):
    # folder_out = cfg.FOLDER_IN
    log = set_logger('find_files.log')
    print(file_path)
    if len(str(cfg.FOLDER_IN)) < 3:
        return
    app = win32com.client.Dispatch('Word.Application')
    try:
        file_out_txt = file_path + '.txt'
        if os.path.isfile(file_out_txt):  # Если выходной LOG файл существует - удаляем его
            os.remove(file_out_txt)
        doc = app.Documents.Open(file_path, Visible=False)
        doc = app.Documents.Open(file_path)
        doc.SaveAs(file_out_txt, FileFormat=7)
        app.Quit()
    except Exception as e:
        str_err = "Exception occurred " + str(e) + ' File: ' + file_path
        print(str_err)  # , exc_info=True
        log.error(str_err)
    finally:
        app.Quit()


# def valid_email(email):
#     str_regex_email = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)+$'
#     return bool(re.search(str_regex_email, email))


def get_extension(filename=''):
    basename = os.path.basename(filename)  # os independent
    ffile = filename.split('\\').pop().split('/').pop()
    ext = '.'.join(ffile.split('.')[1:])

    if len(ext):
        return '.' + ext if ext else None
    else:
        return ''


def get_file_name_without_extension(path=''):
    ext = get_extension(path)
    if len(ext):
        return path.split('\\').pop().split('/').pop().rsplit(ext, 1)[0]
    else:
        return path.split('\\').pop().split('/').pop()


def csv_file_out_create():
    csv_dict = cfg.CSV_DICT
    file_csv = str(os.path.join(os.getcwd(), cfg.CSV_FILE))  # from cfg.file
    # Если выходной CSV файл существует - удаляем его
    if os.path.isfile(file_csv):
        os.remove(file_csv)
    with open(file_csv, 'w', newline='', encoding='utf-8') as csv_file:  # Just use 'w' mode in 3.x
        csv_file_open = csv.DictWriter(csv_file, csv_dict.keys(), delimiter=cfg.CSV_DELIMITER)
        csv_file_open.writeheader()


def txt2xls(f):
    # folder_out = cfg.FOLDER_IN
    filename = f
    str_regex_email = r'[\w.+-]+@[\w-]+\.[\w.-]+'
    str_regex_tel = r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]'

    str_regex_city = r'^Проживает:'
    str_regex_gender_m = r'^Мужчина,'
    str_regex_gender_f = r'^Женщина,'
    str_regex_gr = r'^Гражданство:'
    str_regex_zan = r'^Занятость:'
    str_regex_obr = r'^Образование'
    str_regex_nav = r'^Навыки'

    str_fio = get_file_name_without_extension(filename)
    str_tel = ''
    str_email = ''
    str_city = ''
    str_gender = ''
    str_age = ''
    str_gr = ''
    str_zan = ''
    str_obr = ''
    str_nav = ''

    with open(filename, 'r', encoding='utf-16-le') as file:
        lines = file.readlines()
        cnt_line = 0
        is_obr = False
        is_nav = False

        for line in lines:
            cnt_line += 1

            # ищем в первых 10 строках
            match = re.search(str_regex_email, line)
            if match and cnt_line < 10:
                str_email = match.group(0)

            match1 = re.search(str_regex_tel, line)
            if match1 and cnt_line < 10:
                str_tel = match1.group(0)

            match2 = re.search(str_regex_city, line)
            if match2 and cnt_line < 10:
                str_city = line.replace("Проживает:", "").strip()

            match3 = re.search(str_regex_gender_m, line)
            if match3 and cnt_line < 10:
                str_gender = 'Мужчина'
                str_tmp = line.replace("Мужчина,", "")
                arr_tmp = str_tmp.split()
                if len(arr_tmp[0]):
                    str_age = arr_tmp[0]

            match4 = re.search(str_regex_gender_f, line)
            if match4 and cnt_line < 10:
                str_gender = 'Женщина'
                str_tmp = line.replace("Женщина,", "")
                arr_tmp = str_tmp.split()
                if len(arr_tmp[0]):
                    str_age = arr_tmp[0]

            match5 = re.search(str_regex_gr, line)
            if match5:
                str_tmp = line.replace("Гражданство:", "").strip()
                arr_tmp = str_tmp.split(",")
                if len(arr_tmp[0]):
                    str_gr = arr_tmp[0]

            match6 = re.search(str_regex_zan, line)
            if match6:
                str_zan = line.replace("Занятость:", "").strip()

            if is_obr:
                str_obr = line.strip()
                is_obr = False

            if is_nav:
                str_nav = line.strip()
                is_nav = False

            match7 = re.search(str_regex_obr, line)
            if match7:
                is_obr = True

            match8 = re.search(str_regex_nav, line)
            if match8:
                is_nav = True

    print(f"Found FIO: {str_fio}")
    # print(f"Found EMAIL: {str_email}")
    # print(f"Found TEL: {str_tel}")
    # print(f"Found City: {str_city}")
    # print(f"Found Gender: {str_gender}")
    # print(f"Found Age: {str_age}")
    # print(f"Found Gr: {str_gr}")
    # print(f"Found Zan: {str_zan}")
    # print(f"Found OBR: {str_obr}")
    # print(f"Found NAV: {str_nav}")

    csv_dict = cfg.CSV_DICT
    file_csv = cfg.CSV_FILE
    for key in csv_dict:
        csv_dict[key] = ''

    csv_dict['FIO'] = str_fio
    csv_dict['EMAIL'] = str_email
    csv_dict['TEL'] = str_tel
    csv_dict['CITY'] = str_city
    csv_dict['GENDER'] = str_gender
    csv_dict['AGE'] = str_age
    csv_dict['GR'] = str_gr
    csv_dict['ZAN'] = str_zan
    csv_dict['OBR'] = str_obr
    csv_dict['NAVIK'] = str_nav

    with open(file_csv, 'a', newline='\n', encoding='utf-8') as csv_file:  # Just use 'w' mode in 3.x
        csv_file_open = csv.DictWriter(csv_file, csv_dict.keys(), delimiter=cfg.CSV_DELIMITER)
        try:
            # print(csv_dict['FULLNAME'])
            csv_file_open.writerow(csv_dict)
        except Exception as e:
            print("Exception occurred " + str(e))  # , exc_info=True
        csv_file.close()


def folder_scan():
    files_list_doc = get_list_files_by_ext(cfg.FOLDER_IN, '.doc')
    for f in files_list_doc:
        doc2txt(f)

    csv_file_out_create()

    files_list_txt = get_list_files_by_ext(cfg.FOLDER_IN, '.txt')
    for f in files_list_txt:
        txt2xls(f)

    read_file = pd.read_csv(cfg.CSV_FILE)
    read_file.to_excel(cfg.CSV_FILE + '.xlsx', index=None, header=True)


if __name__ == "__main__":
    folder_scan()

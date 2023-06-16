#!/usr/bin/env python
# -*- coding: utf8 -*-


""" Анализ файла Microsoft Office 2007.

Данный скрипт позволяет произвести анализ данных файла Microsoft Office 2007, 
для поиска следов, которые появляются при изменении метаданных файла.
"""
__author__ = "Андрей Кравцов"
__version__ = "1.0"

import os
import sys
import copy
import zipfile
import csv
import argparse
import xml.etree.ElementTree as ET
from ctypes import Structure, LittleEndianStructure, sizeof, c_uint16, c_char, c_int16, c_uint
from dataclasses import dataclass
from datetime import datetime


# Парсинг каталогов ZIP

class DOS_date(LittleEndianStructure):
    """Структура DOS-time"""
    _pack_ = 1
    _fields_ = [('day', c_uint16, 5),
                ('month', c_uint16, 4),
                ('year', c_uint16, 7)]


class DOS_time(LittleEndianStructure):
    """Структура DOS-date"""
    _pack_ = 1
    _fields_ = [('second', c_uint16, 5),
                ('min', c_uint16, 6),
                ('hour', c_uint16, 5)]


class FileEntry(Structure):
    """Структура локального каталога ZIP"""
    _pack_ = 1
    _fields_ = [('header_signature', c_char * 4),
                ('version_extract', c_uint16),
                ('bit_flag', c_int16),
                ('compression_method', c_int16),
                ('mod_time', DOS_time),
                ('mod_date', DOS_date),
                ('crc32', c_char * 4),
                ('compressed_size', c_uint),
                ('uncompressed_size', c_uint),
                ('file_name_length', c_int16),
                ('extra_field_length', c_int16)]


class CentralDirectory(Structure):
    """Структура главного каталога ZIP"""
    _pack_ = 1
    _fields_ = [('central_file_header_signature', c_char * 4),
                ('version_made_by', c_uint16),
                ('version_needed_to_extract', c_uint16),
                ('general_purpose_bit_flag', c_uint16),
                ('compression_method', c_uint16),
                ('mod_time', DOS_time),
                ('mod_date', DOS_date),
                ('crc32', c_char * 4),
                ('compressed_size', c_uint),
                ('uncompressed_size', c_uint),
                ('file_name_length', c_int16),
                ('extra_field_length', c_int16),
                ('file_comment_length', c_int16),
                ('disk_number_start', c_int16),
                ('internal_file_attributes', c_int16),
                ('external_file_attributes', c_uint),
                ('offset_local_header', c_uint)]


class EndOFCentralDirectory(Structure):
    """Конец главного каталога ZIP"""
    _pack_ = 1
    _fields_ = [('end_of_central_dir_signature', c_char * 4),
                ('number_of_this_disk', c_uint16),
                ('number_of_this_disk', c_uint16),
                ('start_of_the_central_directory', c_uint16),
                ('total_number_of_the_central_directory', c_uint16),
                ('size_central_directory', c_uint),
                ('offset_central_directory', c_uint),
                ('ZIP_file_comment_length', c_uint16)]


def pars_central_dir(path_to_file: str, obj_list: list):
    """Анализ главного каталога ZIP. Первым аргументом принимает путь к файлу, вторым список, 
    в который будут записаны данные после анализа."""

    central_dir = CentralDirectory()
    end_central_dir = EndOFCentralDirectory()
    try:
        with open(path_to_file, 'rb') as f:
            f_size = f.seek(0, os.SEEK_END)
            offset_end_dir = f_size - sizeof(EndOFCentralDirectory)
            f.seek(f_size - sizeof(EndOFCentralDirectory))
            f.readinto(end_central_dir)
            f.seek(end_central_dir.offset_central_directory)
            if end_central_dir.end_of_central_dir_signature != b'PK\x05\x06':
                raise Exception
                

            while True:
                # парсинг главного каталога с описанием всех файлов и создание списка
                f.readinto(central_dir)
                central_dir.file_name = f.read(central_dir.file_name_length)
                central_dir.extra_filed = f.read(central_dir.extra_field_length)
                central_dir.file_comment = f.read(
                    central_dir.file_comment_length)
                obj_list.append(copy.deepcopy(central_dir))
                if f.tell() == offset_end_dir:
                    break
    except Exception:
        print('Файл поврежден или имеет не верное расширение.')
        sys.exit()


# Парсинг данных XML

def get_xml_data(central_file_list: list, metadata_list_ms: list, file_to_parse: str):
    """Функция принимает в качестве аргументов список файлов из главного каталога ZIP, 
    список для сохранения свойств файла, путь к файлу"""
    
    CONST_MS = 'core.xml'
    CORE_XML = ''
    for i in central_file_list:
        name = str(i.file_name, encoding='utf-8')
        if CONST_MS in name:
            with zipfile.ZipFile(file_to_parse, 'r') as zip:
                CORE_XML = str(zip.read(name), encoding='utf-8')
        del (name)
    root_xml = ET.fromstring(CORE_XML)
    for child in root_xml:
        metadata_list_ms.append((child.tag, child.text))


# классы содержащие данные для создания отчета

@dataclass
class MetadataFileMS:
    """Объект класса содержит метаданные файла MS_OFFICE"""
    title: str = ''
    subject: str = ''
    creator: str = ''
    keywords: str = ''
    description: str = ''
    lastModifiedBy: str = ''
    revision: str = ''
    lastPrinted: str = ''
    created: str = ''
    modified: str = ''

    def set_metadata_MS(self, _list):
        """Заполняет поля объекта класса MetadataFileMS. В качестве параметров 
        принимает список с данными анализа главного каталога"""

        def get_dt(string_dt):
            string_dt = datetime.fromisoformat(i[1])
            string_dt = string_dt.astimezone().strftime('%d.%m.%Y %H:%M:%S')
            return string_dt

        for i in _list:
            if 'title' in i[0]:
                self.title = i[1]
            elif 'subject' in i[0]:
                self.subject = i[1]
            elif 'creator' in i[0]:
                self.creator = i[1]
            elif 'keywords' in i[0]:
                self.keywords = i[1]
            elif 'description' in i[0]:
                self.description = i[1]
            elif 'lastModifiedBy' in i[0]:
                self.lastModifiedBy = i[1]
            elif 'revision' in i[0]:
                self.revision = i[1]
            elif 'lastPrinted' in i[0]:
                self.lastPrinted = get_dt(i[1])
            elif 'created' in i[0]:
                self.created = get_dt(i[1])
            elif 'modified' in i[0]:
                self.modified = get_dt(i[1])


@dataclass
class MetadataFilesZipDir:
    """Объект класса содержит метаданные файлов, описанные в главном каталоге ZIP"""
    file_name: str = None
    offset: str = None
    crc32: str = None
    mod_date_time: str = None

    def set_metadata_zip(self, _object):
        """Заполняет поля объекта класса MetadataFilesZipDir. В качестве параметров 
        принимает объект класса CentralDirectory
        - _object - заполненный объект класса CentralDirectory
        """
        def date_time_DOS(_object):
            dt = datetime(year=_object.mod_date.year+1980, 
                          month=_object.mod_date.month, 
                          day=_object.mod_date.day, 
                          hour=_object.mod_time.hour,
                          minute=_object.mod_time.min, 
                          second=_object.mod_time.second).strftime('%d.%m.%Y %H:%M:%S')
            return dt

        self.file_name = str(_object.file_name, encoding='utf-8')
        self.offset = str(_object.offset_local_header)
        self.crc32 = _object.crc32.hex()
        self.mod_date_time = date_time_DOS(_object)


# парсинг файла

def pars_file(file_name: str, central_file_list: list, metadata_list_ms: list, metadata_file_zip: list, metadata_file_ms: object):
    """
    - file_name - путь к анализируемому файлу
    - central_file_list - список который будет содержать данные о записях в главном каталоге ZIP
    - metadata_list_ms - список с метаданными из файла core.xml
    - metadata_file_zip - список с заполненными объектами класса MetadataFilesZipDir
    - metadata_file_ms - объект класса MetadataFileMS
    """

    pars_central_dir(file_name, central_file_list)
    get_xml_data(central_file_list, metadata_list_ms, file_name)
    for i in central_file_list:
        local_zip_file = MetadataFilesZipDir()
        local_zip_file.set_metadata_zip(i)
        metadata_file_zip.append(copy.deepcopy(local_zip_file))
        del(local_zip_file)
    metadata_file_ms.set_metadata_MS(metadata_list_ms)


# создание отчета 

def create_report (metadata_file_zip: list, metadata_file_ms: object, file_name: str):
    """
    - file_name - путь к файлу
    - central_file_list - список в котором содержаться заполненные объекты класса MetadataFilesZipDir
    - metadata_list_ms - заполненный объект класса MetadataFileMS
    
    """
    name = os.path.basename(file_name)
    if not os.path.exists('report'):
        os.mkdir('report')
    with open (f'report\\[ {name} ]_metadata_file_report.txt', '+w', encoding='utf-8') as f:
        f.write(f'Название: {metadata_file_ms.title}\n')
        f.write(f'Тема: {metadata_file_ms.subject}\n')
        f.write(f'Теги: {metadata_file_ms.keywords}\n')
        f.write(f'Комментарии: {metadata_file_ms.description}\n')
        f.write(f'Авторы: {metadata_file_ms.creator}\n')
        f.write(f'Последнее редактирование: {metadata_file_ms.lastModifiedBy}\n')
        f.write(f'Номер редакции: {metadata_file_ms.revision}\n')
        f.write(f'Дата создания содержимого: {metadata_file_ms.created}\n')
        f.write(f'Дата последнего сохранения: {metadata_file_ms.modified}\n')
        f.write(f'Последний вывод на печать: {metadata_file_ms.lastPrinted}\n')
    
    with open (f'report\\[ {name} ]_metadata_dir_in_file_report.csv', '+w', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';', lineterminator='\n')
        writer.writerow(['Имя файла', 'Контрольная сумма CRC32', 'Дата модификации', 'Смещение'])
        for i in metadata_file_zip:
            writer.writerow([i.file_name, i.crc32, i.mod_date_time, hex(int(i.offset))])



if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', type=str,  nargs='?', help='python scrypt.py --file path to file')
    args = parser.parse_args()

    file_name = ''
    central_file_list = []
    metadata_list_ms = []
    metadata_file_ms = MetadataFileMS()
    metadata_file_zip = []
    extension = ['.docx', '.xlsx', '.pptx']

    if not args.file:
        print('Введите путь к файлу: ', end='')
        file_name = input()
    else:
        file_name = args.file
    
    file_name, file_extension = os.path.splitext(file_name)
    if file_extension not in extension:
        print('Данный тип файлов не поддерживается')
    
    pars_file(file_name, central_file_list, metadata_list_ms, metadata_file_zip, metadata_file_ms)
    create_report(metadata_file_zip, metadata_file_ms, file_name)
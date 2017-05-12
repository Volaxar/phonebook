"""
Скрипт создания и выгрузки телефонного справочника, выгружается в формате таблицы Excel.

Список организаций, отделов и пользователей (ФИО, вн. номера, эл. почта) загружаются из AD.
Соответствующие им входящие городские номера загружаются из БД основного Астериска и
файлов extensions.ael периферийных Астерисков.

Для работы скрипта необходим файл настроек: config.ini. Формат файла:
[main] -- секция основных настроек
xls_path -- путь и имя выгружаемого файла, для разделения каталогов используется слэш "/"
m_list -- список секций через пробел, используется для создания структуры каталогов, с настройками Астерисков
 
[ad] -- секция параметров подключения к Active Directory
host -- ip адрес или имя сервера
user -- имя пользователя в виде Домен\Логин
password -- пароль пользователя

[asterisk]
host -- ip адрес или имя сервера
user -- имя пользователя
password -- пароль пользователя
db -- имя базы данных с настройками Астериска

[имя_указанное_в_m_list] -- секция параметров для подключения к периферийным Астерискам, для каждого своя секция
host -- ip адрес или имя сервера
user -- имя пользователя
password -- пароль пользователя

Дополнительные файлы:
importer.py - функции импорта данных из АД, БД и файлов и функция чтения параметров конфигурации config.ini
exporter.py - функция экспорта сформированного словаря телефонной книги в excel
utils.py - вспомогательные функции: ведение логов, загрузка конфигурации, преобразование списка номеров
style.py - стили оформления ячеек для выгружаемого файла

"""

import logging
import re
from collections import defaultdict
from datetime import datetime

import exporter
import importer
import utils
from logger import DiffFileHandler


def main():
    """
    Основная функция модуля, формирует полученные данные из AD и АТС и отправляет на сохранение

    """
    log = logging.getLogger('phonebook')
    log.setLevel(logging.INFO)

    formatter = logging.Formatter('[%(asctime)s] %(levelname)-8s %(filename)s[LINE:%(lineno)d]# %(message)s')

    handler = DiffFileHandler()
    handler.setLevel(logging.INFO)
    handler.setFormatter(formatter)
    log.addHandler(handler)

    utils.log = log

    sections = utils.get_options('main', 'm_list', True)
    
    if not sections:
        log.critical('Ошибка чтения конфигурационного файла, см. ошибки выше')
        exit()
        
    sections = sections.split(' ')
    
    importer.update_ael(sections)

    ad_list = importer.get_ad_list()  # Импортируем список сотрудников из AD
    if not ad_list:
        log.critical('Не удалось загрузить список сотрудников из AD')
        exit()
    
    at_list = importer.get_at_list()  # Импортируем список гор. номеров из БД Астериска
    at_list.update(importer.get_pars_list(sections))  # Импортируем список гор. номеров из конф. файлов Астерисков
    
    raw = defaultdict(dict)
    re_exp_date = re.compile(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}')

    # Формируем список для выгрузки
    for x in ad_list:
        # Если заданы поля "Отображаемое имя" и "Организация"
        if x.displayName and x.company:
            exp_date = re_exp_date.match(str(x.accountExpires))

            # Если истекло время действия учётной записи, пропускаем
            if exp_date:
                if datetime.strptime(exp_date.group(), '%Y-%m-%d %H:%M:%S') < datetime.now():
                    continue

            org = str(x.company)                            # Организация
            dep = str(x.department).replace('[]', '')       # Отдел
            itn = str(x.telephoneNumber).replace('[]', '')  # Внутренние номера

            if dep not in raw[org]:
                raw[org][dep] = []

            # Составляем список полей для выгрузки в xls, порядок полей важен!
            raw[org][dep].append([
                str(x.displayName),                         # ФИО
                str(x.title).replace('[]', ''),             # Должность
                itn,                                        # Внутренние номера
                utils.get_city(itn, at_list),               # Городские номера
                str(x.mail).lower().replace('[]', '')       # Эл. почта
            ])

    if exporter.export_xls(raw):
        log.info('Телефонная книга выгружена успешно')
    else:
        log.critical('Ошибка выгрузки телефонной книги')


if __name__ == '__main__':
    main()

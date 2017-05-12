"""
Функции импорта данных для телефонной книги

update_ael - загрузка файлов конфигурации extensions.ael с периферийных Астерисков
get_ad_list - загрузка информации о сотрудниках из AD
get_at_list - загрузка списка гор. номеров из БД Астериска
get_pars_list - загрузка списка гор. номеров из файлов extensions.ael
"""
import os
import re
from collections import defaultdict

import paramiko
import pymysql
from ldap3 import Server, Connection, ALL, NTLM
from ldap3.core.exceptions import LDAPSocketOpenError, LDAPBindError
from pymysql.err import OperationalError
from scp import SCPClient, SCPException

from utils import log, get_options


def update_ael(sections):
    """
    Загружаем и обновляем конф. файлы из АТС, в случае ошибки обновления, будут использоваться файлы загруженные ранее

    :param sections: [string], список секций из конф. файла с параметрами подключения к АТС
    """
    for section in sections:
        options_list = get_options(section)
        
        if options_list:
            host, user, password = options_list
            
            try:
                with paramiko.SSHClient() as ssh:
                    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                    ssh.connect(hostname=host, username=user, password=password)
                    
                    remotepath = '/etc/asterisk/extensions.ael'
                    localpath = 'asterisk/%s/extensions.ael' % section
                    
                    if not os.path.exists(os.path.dirname(localpath)):
                        os.makedirs(os.path.dirname(localpath))
                    
                    try:
                        with SCPClient(ssh.get_transport()) as client:
                            client.get(remotepath, localpath)
                    except SCPException as e:
                        log.error(e)
            
            except TimeoutError as e:
                log.error('%s - %s' % (host, e.strerror))
            except paramiko.ssh_exception.AuthenticationException:
                log.error('Ошибка авторизации на %s' % host)
    
    return True


def get_ad_list():
    """"
    Импортируем список сотрудников из AD
    
     Фильтр !(userAccountControl:1.2.840.113556.1.4.803:=2) исключает заблокированные учётные записи

     :return: [{}], cписок словарей с аттрибутами пользователей
    """
    raw = []
    options_list = get_options('ad')
    ad_search = get_options('main', 'ad_search', True)
    
    if options_list:
        host, user, password = options_list
        
        server = Server(host, get_info=ALL)
        
        try:
            with Connection(server, user, password, authentication=NTLM, auto_bind=True) as conn:
                filter_str = '(&(objectclass=person)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'
                
                conn.search(ad_search, filter_str, attributes=[
                    'displayName',          # ФИО
                    'telephoneNumber',      # Внутренние номера
                    'mail',                 # Эл. почта
                    'title',                # Должность
                    'department',           # Отдел
                    'company',              # Организация
                    'accountExpires'        # Дата блокировки учётной записи
                ])
                
                raw = conn.entries
        except LDAPSocketOpenError as e:
            log.error(e)
        except LDAPBindError:
            log.error('Ошибка доменной авторизации')
    
    return raw


def get_at_list():
    """
    Импортируем список городских номеров из БД АТС
    
    :return: {string: {string}}, {вн_номер: {гор_номер, ...}}
    """
    raw = defaultdict(set)
    options_list = get_options('asterisk', 'db')
    
    if options_list:
        host, user, password, db = options_list
        
        try:
            with pymysql.connect(host, user, password, db) as cur:
                
                re_group = re.compile(r'ext-group,(\d{3}),1')
                re_im = re.compile(r'from-did-direct,(\d{4}),1')
                re_ivr = re.compile(r'ivr-(\d+),s,1')

                # Запрос, привязка групп к вн. номерам
                cur.execute("SELECT grpnum, grplist FROM ringgroups ")
                
                # Составляем словарь {группа: set(список_вн_номеров)}
                groups_list = dict((k, v.replace('#', '').split('-')) for k, v in cur)
                
                cur.execute("SELECT ivr_id, selection, dest FROM ivr_dests WHERE ivr_ret = 0")

                ivr_list = defaultdict(dict)

                for ivr_id, sel, dest in cur:
                    ivr_group = re_group.match(dest)

                    if ivr_group:
                        ivr_list[str(ivr_id)][sel] = groups_list[ivr_group.group(1)]
                    else:
                        ivr_im = re_im.match(dest)

                        if ivr_im:
                            ivr_list[str(ivr_id)][sel] = [ivr_im.group(1)]

                # Запрос, привязка гор. номеров к вн. номерам или группам
                cur.execute("SELECT extension, destination FROM incoming "
                            "WHERE LENGTH(extension) = 7 OR LENGTH(extension) = 11")
                
                for ext, des in cur:
                    
                    # Формат гор. номера "###-##-##"
                    cm = '%s-%s-%s' % (ext[-7:-4], ext[-4:-2], ext[-2:])
                    
                    # Выбираем только группы
                    group = re_group.match(des)
                    
                    if group:
                        for x in groups_list[group.group(1)]:
                            # Добавляем все вн. номера группы связанные с гор. номером
                            raw[x].add(cm)
                    
                    else:
                        # Выбираем только прямые вн. номера
                        im = re_im.match(des)
                        
                        if im:
                            # Добавляем вн. номер связанный с гор. номером
                            raw[im.group(1)].add(cm)
                        else:
                            ivr = re_ivr.match(des)

                            if ivr and ivr.group(1) in ivr_list:
                                for ik, iv in ivr_list[ivr.group(1)].items():
                                    for x in iv:
                                        ivr_cm = '%s/%s' % (cm, ik)

                                        raw[x].add(ivr_cm)

        except OperationalError as e:
            log.error(e)
            
    return raw


def _get_num(section):
    """
    Импортируем список городских номеров из конф. файла АТС
    
    :param section: string, имя секции, определяет путь к файлу конфигурации АТС
    :return: {string: [string]}, {гор_номер: [вн_номер]}
    """
    items = {}
    
    try:
        with open('./asterisk/%s/extensions.ael' % section) as f:
    
            text = re.search(r'context temp_incom\s+\{(\s*_[\dX]+\s*=>\s*\{.*?\}\s)+\s+\}', f.read(), re.DOTALL)
    
            if text:
                num_text = re.findall(r'\s*_[\d]{7,11}\s*=>\s*\{.*?\}\s', text.group(), re.DOTALL)
    
                for t in num_text:
                    header = re.search(r'_(\d+) =>', t)
                    dials = re.search(r'Dial\(.*?,,tTr\);', t)
    
                    if header and dials:
                        items[header.group(1)] = re.findall(r'\d{4}', dials.group())
    except FileNotFoundError as e:
        log.error('Конфигурационный файл АТС не найден: %s' % e.filename)
    
    return items


def get_pars_list(sections):
    """
    Импортируем список городских номеров из конф. файлов АТС
    
    :param sections: [string], список секций из конф. файла (перечень конф. файлов АТС)
    :return: {string: {string}}, словарь множеств, {вн_номер: {гор_номер}}
    """
    raw = defaultdict(set)
    
    for section in sections:
        
        # Получаем список гор. и вн. номеров
        num_list = _get_num(section)
        
        for k, v in num_list.items():
    
            # Форматируем Челябинский гор. номер "###-##-##"
            if len(k) == 7:
                fk = '%s-%s-%s' % (k[0:3], k[3:5], k[5:7])
    
            # Форматируем Магнитогорский гор. номер "#(####)###-###"
            else:
                fk = '%s(%s)%s-%s' % (k[0:1], k[1:5], k[5:8], k[8:11])
    
            for x in v:
                raw[x].add(fk)
    
    return raw

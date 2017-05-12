import xlwt

import style as ts
from utils import log, get_options


def export_xls(raw):
    """
    Выгрузка структуры телефонной книги в excel

    :param raw: {string: {string: [[string]]}}, структура тел. книги {организация: {отдел: [[данные_сотрудника]]}}
    :return: bool, True - если тел. книга выгружена, False - если произошла ошибка выгрузки
    """

    wb = xlwt.Workbook()
    ws = wb.add_sheet('Телефонная книга')

    # Заголовок
    ws.write_merge(0, 0, 0, 1, '', ts.sh0)
    ws.write(0, 2, 'ФИО', ts.sh1)
    ws.write(0, 3, 'Должность', ts.sh2)
    ws.write(0, 4, 'Вн. тел.', ts.sh2)
    ws.write(0, 5, 'Гор. тел.', ts.sh2)
    ws.write(0, 6, 'Почта', ts.sh3)

    line = 1

    item_len = [0, 0, 0, 0, 0, 0]  # Для подсчёта ширины колонок
    item_style = [ts.ss1, ts.ss2, ts.ss2, ts.ss2, ts.ss3]  # Стили для полей сотрудников

    for kc in sorted(raw):
        # Наименование организации
        ws.write_merge(line, line, 0, 6, kc, ts.so0)
        ws.row(line).level = 0
        line += 1
        ws.row(line).level = 1

        for kd in sorted(raw[kc]):
            # Наименование отдела
            ws.write(line, 0, '', ts.sd0)
            ws.write_merge(line, line, 1, 6, kd, ts.sd1)
            ws.row(line).level = 1
            line += 1
            ws.row(line).level = 2

            for item in sorted(raw[kc][kd]):
                ws.write_merge(line, line, 0, 1, '', ts.ss0)

                for k, field in enumerate(item):
                    # Поля сотрудника: ФИО, Должность, Вн. тел., Гор. тел., Почта
                    ws.write(line, k + 2, field, item_style[k])

                    # Считаем максимальную ширину для каждой колонки
                    item_len[k] = max(item_len[k], len(field))

                line += 1
                ws.row(line).level = 2

    # Подчёркиваем таблицу
    ws.write_merge(line, line, 0, 6, '', ts.sb0)

    ws.row(line).level = 0

    # Задаем ширину колонок, для организации и отдела ширина фиксирована
    ws.col(0).width = int(36.5 * 17)
    ws.col(1).width = int(36.5 * 17)

    # Для остальных колонок ширина вычисляется в зависимости от длинны полей
    for k, v in enumerate(item_len):
        ws.col(k + 2).width = int(36.5 * 7.3 * v)

    # Закрепляем заголовок
    ws.panes_frozen = True
    ws.horz_split_pos = 1

    path = get_options('main', 'xls_path', True)

    if not path:
        log.critical('Ошибка чтения конфигурационного файла, см. ошибки выше')
        return

    try:
        wb.save(path)
    except PermissionError as e:
        log.error('Недостаточно прав для сохранения файла: %s' % e.filename)
        return
    except FileNotFoundError as e:
        log.error('Неверный путь или имя файла: %s' % e.filename)
        return

    return True

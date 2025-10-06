import openpyxl
import re
from django.shortcuts import render
from django.contrib import messages

from TatOMS import settings


def parse_houses(houses_str):
    """Разбивает строку домов по запятой и точке с запятой, обрабатывает диапазоны через тире и дроби через /"""
    if not houses_str:
        return []

    houses_str = str(houses_str).strip()

    # Проверяем на наличие "Лесной городок" в таблице
    if houses_str.lower() in ['лесной городок']:
        return ['Лесной городок']  # Возвращаем специальный маркер
    # Проверяем на наличие "ВСЕ" в таблице
    if houses_str.upper() == "ВСЕ":
        return ["ВСЕ"]  # Возвращаем специальный маркер

    houses = []
    # Разбиваем строку по разделителям
    for part in re.split('[,;]', houses_str):
        part = part.strip()
        if not part:
            continue

        # Обработка диапазонов
        if '-' in part:
            try:
                start, end = map(str.strip, part.split('-'))
                if start.isdigit() and end.isdigit():
                    start_num = int(start)
                    end_num = int(end)
                    # Генерируем все числа в диапазоне
                    houses.extend(str(i) for i in range(min(start_num, end_num),
                                                        max(start_num, end_num) + 1))
                else:
                    # Если не числа, добавляем как есть
                    houses.append(part)
            except:
                # В случае ошибки добавляем как есть
                houses.append(part)

        # Обработка дробей "/"
        elif '/' in part:
            main_part, fraction = part.split('/', 1)
            main_part = main_part.strip()
            fraction = fraction.strip()

            # Проверяем, что обе части валидны
            if (re.match(r'^[\dА-Яа-я]+$', main_part) and
                    re.match(r'^[\dА-Яа-я]+$', fraction)):
                houses.append(f"{main_part}/{fraction}")
            else:
                return ["INVALID"]
        else:
            # Проверка на допустимые символы в номере дома
            if not re.match(r'^[\dА-Яа-я]+$', part):
                return ["INVALID"]
            houses.append(part)

    return houses


def upload_file(request):
    """Обрабатывает загруженный Excel файл с адресами"""
    addresses = [] # Список валидных адресов
    filename = None # Имя загруженного файла
    all_houses_rows = [] # Строки с "ВСЕ" в номерах домов
    all_streets_rows = [] # Строки с "ВСЕ" в названиях улиц
    empty_field_rows = [] # Строки с пустыми обязательными полями
    invalid_houses_rows = [] # Строки с невалидными номерами домов
    not_found_addresses_rows = [] # Строки с адресами, которые не найдены на карте

    if request.method == "POST" and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        filename = excel_file.name

        # Загрузка Excel файла
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active

        # Пропускаем заголовок (первая строка)
        for row_num, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            # Проверка на пустые поля
            has_empty_fields = False
            if not row[0]:  # Номер участка
                has_empty_fields = True
            if not row[1] or not str(row[1]).strip():  # Город
                has_empty_fields = True
            if not row[2] or not str(row[2]).strip():  # Улица
                has_empty_fields = True
            if not row[3] or not str(row[3]).strip():  # Дома
                has_empty_fields = True

            if has_empty_fields:
                empty_field_rows.append(row_num)
                continue  # Пропускаем строку с пустыми полями

            # Извлечение данных из строки
            group = int(row[0]) if row[0] else None
            city = row[1].strip() if row[1] else ''
            street = row[2].strip() if row[2] else ''
            houses_str = str(row[3]) if row[3] else ''

            # Проверяем поле улицы на "ВСЕ"
            if street.upper() == "ВСЕ":
                all_streets_rows.append(row_num)
                continue  # Пропускаем строку с "ВСЕ" в поле улицы

            # Парсинг номеров домов
            houses = parse_houses(houses_str)

            # Проверяем на невалидные значения
            if "INVALID" in houses:
                invalid_houses_rows.append(row_num)
                continue

            # Проверяем поле списка домов на "ВСЕ"
            if "ВСЕ" in houses:
                all_houses_rows.append(row_num)
                continue

            # Обработка специального случая
            if "Лесной городок" in houses:
                addresses.append({
                    'group': group,
                    'city': city,
                    'street': street,
                    'house': 'Лесной городок',
                    'is_forest_town': True,
                    'row_num': row_num
                })
                continue

            # Добавление всех домов из строки
            if houses:
                for house in houses:
                    addresses.append({
                        'group': group,
                        'city': city,
                        'street': street,
                        'house': house,
                        'row_num': row_num
                    })
            else:
                # Если домов нет, добавляем только улицу
                addresses.append({
                    'group': group,
                    'city': city,
                    'street': street,
                    'house': None
                })

        # Формирование сообщений об ошибках
        # Сообщение о строках с "ВСЕ" в улицах
        if all_streets_rows:
            if len(all_streets_rows) == 1:
                row_word = "строку"
                err_street_word = "названия улицы"
            else:
                row_word = "строки"
                err_street_word = "названия улицы"

            messages.error(
                request,
                f'Обнаружена ошибка! Вместо {err_street_word} написано "ВСЕ". Эти адреса не были обработаны. '
                    f'Исправьте, пожалуйста, {row_word}:'
            )

        # Сообщение о строках с "ВСЕ" в домах
        if all_houses_rows:
            if len(all_houses_rows) == 1:
                row_word = "строку"
                err_houses_word = "номеров домов"
            else:
                row_word = "строки"
                err_houses_word = "номеров домов"

            messages.error(
                request,
                f'Обнаружена ошибка! Вместо {err_houses_word} написано "ВСЕ". Эти адреса не были обработаны. '
                f'Исправьте, пожалуйста, {row_word}:'
            )

        # Сообщение о строках с пустыми полями
        if empty_field_rows:
            if len(empty_field_rows) == 1:
                row_word = "строку"
                field_word = "пустыми полями"
            else:
                row_word = "строки"
                field_word = "пустыми полями"

            messages.error(
                request,
                f'Обнаружена ошибка! Есть строки с {field_word}. Эти адреса не были обработаны. '
                f'Исправьте, пожалуйста, {row_word}:'
            )

        # Сообщение о невалидных номерах домов
        if invalid_houses_rows:
            if len(invalid_houses_rows) == 1:
                row_word = "строку"
                error_word = "некорректные значения"
            else:
                row_word = "строки"
                error_word = "некорректные значения"

            messages.error(
                request,
                f'Обнаружена ошибка! В файле указаны {error_word}. Эти адреса не были обработаны. '
                f'Исправьте, пожалуйста, {row_word}:'
            )

    return render(request, 'upload.html', {
        'addresses': addresses,
        'filename': filename,
        'yandex_api_key': settings.YANDEX_MAPS_API_KEY,
        'all_houses_rows': all_houses_rows,
        'all_streets_rows': all_streets_rows,
        'empty_field_rows': empty_field_rows,
        'invalid_houses_rows': invalid_houses_rows,
        'not_found_addresses_rows': not_found_addresses_rows,
    })
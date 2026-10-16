"""
Модуль для генерации реалистичных данных и создания стилизованного отчета
в Google Sheets, оформленного как презентабельный официальный документ (Word/Excel бланки).
"""

from __future__ import annotations

import datetime
import random
from typing import Any, Callable, Dict, List, Optional, Tuple

from google_sheets_client import GoogleSheetsClient


def hex_to_color(hex_str: str) -> dict[str, float]:
    """Конвертирует HEX цвет (#RRGGBB) в формат RGBA Google Sheets API (0.0 - 1.0)."""
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return {"red": round(r, 3), "green": round(g, 3), "blue": round(b, 3)}


def grid_range(sheet_id: int, start_row: int, end_row: int, start_col: int, end_col: int) -> dict[str, int]:
    """Формирует корректный объект GridRange для Google Sheets API v4."""
    return {
        "sheetId": sheet_id,
        "startRowIndex": start_row,
        "endRowIndex": end_row,
        "startColumnIndex": start_col,
        "endColumnIndex": end_col,
    }


# Список реалистичных контрагентов для симуляции
SAMPLE_CLIENTS = [
    "ООO «СеверТрансЛогистик»",
    "ПАО «ПромИнвестХолдинг»",
    "АО «ТехноСфера Инновации»",
    "ООО «Альфа Дистрибьюшн»",
    "ИП Смирнов А.В.",
    "ООО «ГазЭнергоСтрой»",
    "АО «Русский Ритейл»",
    "ООО «МедиаКом Консалтинг»",
    "ИП Воронова Е.М.",
    "ООО «Вектор Развития»",
    "АО «СпецСнаб Комплект»",
    "ООО «АстраТелеком»",
]

# Категории и номенклатура по департаментам
DEPARTMENT_DATA = {
    "Отдел корпоративных продаж": {
        "manager_pool": ["Иванов Алексей Сергеевич", "Кузнецова Ольга Павловна", "Петров Денис Игоревич"],
        "items": [
            ("Поставка серверного оборудования Enterprise", 180000, 750000),
            ("Лицензия ПО «Корпоративная безопасность» (1 год)", 95000, 340000),
            ("Внедрение CRM-системы и обучение персонала", 250000, 680000),
            ("Пакет облачных вычислительных мощностей vCPU", 45000, 180000),
            ("Сетевые коммутаторы 10Gbps и монтаж СКС", 120000, 410000),
        ],
    },
    "Департамент логистики и ВЭД": {
        "manager_pool": ["Морозов Артем Викторович", "Соколова Марина Юрьевна", "Григорьев Роман Олегович"],
        "items": [
            ("Мультимодальная контейнерная перевозка (40ft)", 210000, 520000),
            ("Услуги таможенного оформления и брокерский сбор", 45000, 130000),
            ("Ответственное складское хранение (Класс А, месяц)", 160000, 390000),
            ("Срочная авиадоставка комплектующих карго", 85000, 270000),
            ("Экспедиторское сопровождение спецгруза", 65000, 195000),
        ],
    },
    "Финансово-аналитический департамент": {
        "manager_pool": ["Васильева Екатерина Николаевна", "Федоров Павел Сергеевич", "Орлова Анна Михайловна"],
        "items": [
            ("Комплексный финансовый аудит за отчетный период", 320000, 850000),
            ("Разработка инвестиционного меморандума", 180000, 420000),
            ("Налоговое консультирование и структурирование сделки", 90000, 260000),
            ("Оценка рыночной стоимости нематериальных активов", 140000, 380000),
            ("Построение финансово-экономической модели проекта", 210000, 490000),
        ],
    },
    "Департамент маркетинга и развития": {
        "manager_pool": ["Дмитриев Максим Андреевич", "Зайцева Ксения Викторовна", "Белов Сергей Александрович"],
        "items": [
            ("Проведение таргетированной рекламной кампании Q3", 150000, 480000),
            ("Разработка брендбука и фирменного стиля холдинга", 220000, 550000),
            ("Организация участия в отраслевой выставке Expo", 350000, 920000),
            ("SEO-оптимизация и контентная стратегия портала", 80000, 210000),
            ("Маркетинговое исследование рынка конкурентов", 130000, 340000),
        ],
    },
}

STATUS_LIST = ["Оплачено", "Выполнено", "Оплачено", "Акт подписан", "В обработке"]


def generate_random_report_data(
    date_from: datetime.date,
    date_to: datetime.date,
    department: str,
    author: str,
    row_count: int = 15,
) -> Dict[str, Any]:
    """
    Генерирует согласованный набор реалистичных данных для отчета.
    """
    if date_from > date_to:
        date_from, date_to = date_to, date_from

    total_days = max(1, (date_to - date_from).days)

    dept_info = DEPARTMENT_DATA.get(
        department,
        DEPARTMENT_DATA["Отдел корпоративных продаж"]
    )
    items_pool = dept_info["items"]
    managers = dept_info["manager_pool"]

    rows = []
    total_sum = 0
    total_qty = 0

    for i in range(1, row_count + 1):
        # Случайная дата в интервале
        offset = random.randint(0, total_days)
        op_date = date_from + datetime.timedelta(days=offset)
        op_date_str = op_date.strftime("%d.%m.%Y")

        doc_num = f"ДОК-{op_date.strftime('%y%m')}-{random.randint(1001, 9999)}"
        client = random.choice(SAMPLE_CLIENTS)
        item_title, min_price, max_price = random.choice(items_pool)
        manager = random.choice(managers)
        qty = random.randint(1, 5)
        unit_price = round(random.randint(min_price, max_price), -2)
        amount = qty * unit_price
        status = random.choice(STATUS_LIST)

        total_sum += amount
        total_qty += qty

        rows.append({
            "idx": i,
            "date": op_date_str,
            "doc_num": doc_num,
            "client": client,
            "item": item_title,
            "manager": manager,
            "qty": qty,
            "amount": amount,
            "status": status,
        })

    # Сортировка по дате
    rows.sort(key=lambda x: datetime.datetime.strptime(x["date"], "%d.%m.%Y"))
    for idx, r in enumerate(rows, start=1):
        r["idx"] = idx

    avg_check = total_sum / max(1, row_count)
    completion_rate = round(random.uniform(98.5, 114.2), 1)

    return {
        "date_from": date_from.strftime("%d.%m.%Y"),
        "date_to": date_to.strftime("%d.%m.%Y"),
        "created_at": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
        "department": department,
        "author": author or "Руководитель аналитического отдела",
        "doc_id": f"REP-{datetime.datetime.now().strftime('%Y%m%d')}-{random.randint(100, 999)}",
        "rows": rows,
        "total_sum": total_sum,
        "total_qty": total_qty,
        "avg_check": avg_check,
        "completion_rate": completion_rate,
        "row_count": row_count,
    }


def build_sheet_report(
    client: GoogleSheetsClient,
    spreadsheet_id: str,
    sheet_title: str,
    data: Dict[str, Any],
    progress_callback: Optional[Callable[[str, int], None]] = None,
) -> Tuple[int, str]:
    """
    Создает в Google Sheets новый лист с оформлением под официальный документ Word/Excel:
    - Заголовок организации и официальный гриф
    - Название и метаданные документа
    - 4 стильные карточки KPI показателей
    - Реестр детальных хозяйственных операций
    - Формулы Google Sheets для автоматического подсчета итогов
    - Заключение и блок официальных подписей с печатями
    - Идеальные ширины колонок, отступы, шрифты, границы и цветовые решения

    :return: (sheet_id, sheet_url)
    """

    def notify(msg: str, step: int):
        if progress_callback:
            progress_callback(msg, step)

    notify("Подготовка листа в Google Таблице...", 10)

    # 1. Проверяем или создаем лист
    existing_sheets = client.get_sheet_names(spreadsheet_id)
    if sheet_title in existing_sheets:
        notify(f"Очистка существующего листа «{sheet_title}»...", 15)
        sheet_id = client.get_sheet_id(sheet_title, spreadsheet_id)
        client.clear_sheet(sheet_title, spreadsheet_id)
    else:
        notify(f"Создание новой вкладки «{sheet_title}»...", 15)
        res = client.create_sheet(sheet_title, spreadsheet_id)
        added_sheet = res["replies"][0]["addSheet"]["properties"]
        sheet_id = int(added_sheet["sheetId"])

    notify("Формирование документарного макета...", 25)

    # Колонки:
    # A (отступ)
    # B (№), C (Дата), D (Документ), E (Контрагент), F (Номенклатура), G (Менеджер), H (Кол-во), I (Сумма, руб), J (Статус)
    # Всего колонок J = 10 (индексы 0..9)

    matrix: List[List[Any]] = []

    # Строка 1: пустой отступ
    matrix.append([""] * 10)

    # Строка 2: Надпись организации (Word Header)
    matrix.append(["", "ХОЛДИНГОВАЯ КОМПАНИЯ «ИНВЕСТ-ТЕХНОЛОДЖИ ГРУПП»", "", "", "", "", "", "", "", ""])

    # Строка 3: Официальный гриф документа
    matrix.append(["", "Форма внутреннего учета № 14-БИС | УТВЕРЖДЕНО РЕШЕНИЕМ СОВЕТА ДИРЕКТОРОВ", "", "", "", "", "", "", "", ""])

    # Строка 4: разделитель
    matrix.append([""] * 10)

    # Строка 5: Главное название отчета
    matrix.append(["", "СВОДНЫЙ ОТЧЕТ ОБ ОПЕРАЦИОННОЙ ДЕЯТЕЛЬНОСТИ И РАСЧЕТАХ", "", "", "", "", "", "", "", ""])

    # Строка 6: Период и департамент
    period_str = f"Отчетный период: с {data['date_from']} по {data['date_to']}   •   Подразделение: {data['department']}"
    matrix.append(["", period_str, "", "", "", "", "", "", "", ""])

    # Строка 7: разделитель
    matrix.append([""] * 10)

    # Строка 8-9: Паспорт документа (Инфо-панель)
    matrix.append(["", "Номер документа:", data["doc_id"], "", "Дата формирования:", data["created_at"], "", "Статус документа:", "УТВЕРЖДЕНО", ""])
    matrix.append(["", "Ответственный:", data["author"], "", "Объем выборки:", f"{data['row_count']} операций", "", "Гриф:", "ДЛЯ ВНУТРЕННЕГО ПОЛЬЗОВАНИЯ", ""])

    # Строка 10: разделитель
    matrix.append([""] * 10)

    # Строки 11-13: Карточки KPI (4 карточки)
    # B11..C11, D11..E11, F11..G11, H11..I11
    matrix.append(["", "ОБЩИЙ ОБОРОТ", "", "КОЛИЧЕСТВО СДЕЛОК", "", "СРЕДНИЙ ЧЕК", "", "ВЫПОЛНЕНИЕ ПЛАНА", "", ""])
    matrix.append(["", data["total_sum"], "", data["row_count"], "", data["avg_check"], "", f"{data['completion_rate']}%", "", ""])
    matrix.append(["", "рублей с учетом НДС", "", "хозяйственных операций", "", "средняя сумма операции", "", "целевой KPI департамента", "", ""])

    # Строка 14: разделитель
    matrix.append([""] * 10)

    # Строка 15: Заголовок раздела таблицы
    matrix.append(["", "I. РЕЕСТР ПРОВЕДЕННЫХ ОПЕРАЦИЙ И ВЗАИМОРАСЧЕТОВ", "", "", "", "", "", "", "", ""])

    # Строка 16: пустая
    matrix.append([""] * 10)

    # Строка 17: Шапка таблицы (Row 17, 1-based)
    table_header = ["", "№ п/п", "Дата", "Документ", "Контрагент / Клиент", "Номенклатура / Услуга", "Ответственный", "Кол-во", "Сумма, ₽", "Статус"]
    matrix.append(table_header)

    # Строки 18..N: Табличная часть
    data_start_row = 18  # 1-indexed
    for r in data["rows"]:
        matrix.append([
            "",
            r["idx"],
            r["date"],
            r["doc_num"],
            r["client"],
            r["item"],
            r["manager"],
            r["qty"],
            r["amount"],
            r["status"],
        ])

    data_end_row = data_start_row + len(data["rows"]) - 1
    total_row = data_end_row + 1

    # Строка Итого (живые формулы Google Sheets)
    formula_sum = f"=SUM(I{data_start_row}:I{data_end_row})"
    formula_qty = f"=SUM(H{data_start_row}:H{data_end_row})"
    matrix.append([
        "",
        "ИТОГО ПО РЕЕСТРУ:",
        "",
        "",
        "",
        "",
        "",
        formula_qty,
        formula_sum,
        "",
    ])

    # Разделители после таблицы
    matrix.append([""] * 10)

    # Раздел 2: Заключительные положения и подписи
    matrix.append(["", "II. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ И ПОДПИСИ СТОРОН", "", "", "", "", "", "", "", ""])
    matrix.append([
        "",
        "Все представленные операции подтверждены первичными учетными документами (товарные накладные, акты выполненных работ, УПД). Расхождения и претензии со стороны контрагентов отсутствуют.",
        "", "", "", "", "", "", "", "",
    ])
    matrix.append([""] * 10)

    # Блок подписей (2 строки)
    sign_row_1 = total_row + 5
    matrix.append([
        "",
        "Составитель отчета:",
        f"_________________ / {data['author']} /",
        "",
        "",
        "Руководитель финансовой службы:",
        f"_________________ / Воропаев М.А. /",
        "",
        "",
        "",
    ])
    matrix.append([
        "",
        f"Дата подписи: {data['created_at'][:10]} г.",
        "М.П.",
        "",
        "",
        "Отметка аудитора: Замечаний нет",
        "Штамп системы: ПРОВЕРЕНО",
        "",
        "",
        "",
    ])

    notify("Запись данных в Google Таблицу...", 45)

    # Записываем матрицу значений в лист
    target_range = f"'{sheet_title}'!A1:J{len(matrix)}"
    client.update_range(cell_range=target_range, values=matrix, spreadsheet_id=spreadsheet_id)

    notify("Применение стилей и форматирования Word-документа...", 70)

    # 2. Формируем batchUpdate для глубокого визуального оформления
    requests: List[Dict[str, Any]] = []

    # 2.1 Настройки сетки, цвет ярлыка вкладки и закрепление строк
    requests.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "title": sheet_title,
                "tabColor": hex_to_color("#1E3A8A"),
                "gridProperties": {
                    "hideGridlines": False,
                    "frozenRowCount": 17,  # Закрепляем строки шапки таблицы
                },
            },
            "fields": "tabColor,gridProperties(hideGridlines,frozenRowCount)",
        }
    })

    # 2.2 Размеры колонок
    col_widths = [
        (0, 1, 28),    # A (отступ слева)
        (1, 2, 55),    # B (№)
        (2, 3, 105),   # C (Дата)
        (3, 4, 135),   # D (Документ)
        (4, 5, 230),   # E (Контрагент)
        (5, 6, 250),   # F (Номенклатура)
        (6, 7, 185),   # G (Менеджер)
        (7, 8, 85),    # H (Кол-во)
        (8, 9, 140),   # I (Сумма)
        (9, 10, 125),  # J (Статус)
    ]
    for start_c, end_c, pixel_size in col_widths:
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_c,
                    "endIndex": end_c,
                },
                "properties": {"pixelSize": pixel_size},
                "fields": "pixelSize",
            }
        })

    # Высоты особых строк
    row_heights = [
        (1, 2, 24),   # Шапка организации
        (4, 5, 38),   # Заголовок отчета
        (10, 11, 24), # KPI верх
        (11, 12, 34), # KPI значение
        (12, 13, 22), # KPI низ
        (16, 17, 34), # Шапка таблицы
        (total_row - 1, total_row, 30), # Итого
    ]
    for start_r, end_r, pixel_size in row_heights:
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start_r,
                    "endIndex": end_r,
                },
                "properties": {"pixelSize": pixel_size},
                "fields": "pixelSize",
            }
        })

    # 2.3 Объединение ячеек (mergeCells) с правильными startRowIndex / endRowIndex
    merges = [
        # Верхний колонтитул
        (1, 2, 1, 10),
        (2, 3, 1, 10),
        # Главный заголовок
        (4, 5, 1, 10),
        (5, 6, 1, 10),
        # Паспорт документа
        (7, 8, 2, 4),
        (7, 8, 5, 7),
        (7, 8, 8, 10),
        (8, 9, 2, 4),
        (8, 9, 5, 7),
        (8, 9, 8, 10),
        # KPI Карточки (B..C, D..E, F..G, H..I)
        (10, 11, 1, 3), (11, 12, 1, 3), (12, 13, 1, 3),
        (10, 11, 3, 5), (11, 12, 3, 5), (12, 13, 3, 5),
        (10, 11, 5, 7), (11, 12, 5, 7), (12, 13, 5, 7),
        (10, 11, 7, 9), (11, 12, 7, 9), (12, 13, 7, 9),
        # Раздел 1
        (14, 15, 1, 10),
        # Итого надпись
        (total_row - 1, total_row, 1, 7),
        # Раздел 2
        (total_row + 1, total_row + 2, 1, 10),
        (total_row + 2, total_row + 3, 1, 10),
        # Подписи
        (sign_row_1 - 1, sign_row_1, 2, 5),
        (sign_row_1 - 1, sign_row_1, 6, 10),
        (sign_row_1, sign_row_1 + 1, 2, 5),
        (sign_row_1, sign_row_1 + 1, 6, 10),
    ]
    for start_r, end_r, start_c, end_c in merges:
        requests.append({
            "mergeCells": {
                "range": grid_range(sheet_id, start_r, end_r, start_c, end_c),
                "mergeType": "MERGE_ALL",
            }
        })

    # 2.4 Стилизация ячеек

    # Базовый шрифт на всем листе
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, 0, len(matrix) + 10, 0, 10),
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "fontFamily": "Arial",
                        "fontSize": 10,
                        "foregroundColor": hex_to_color("#334155"),
                    },
                    "verticalAlignment": "MIDDLE",
                }
            },
            "fields": "userEnteredFormat(textFormat,verticalAlignment)",
        }
    })

    # Шапка компании (Строка 2)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, 1, 2, 1, 10),
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "fontFamily": "Arial",
                        "fontSize": 10,
                        "bold": True,
                        "foregroundColor": hex_to_color("#1E3A8A"),
                    },
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(textFormat,horizontalAlignment)",
        }
    })

    # Подзаголовок компании (Строка 3)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, 2, 3, 1, 10),
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "fontFamily": "Arial",
                        "fontSize": 8,
                        "italic": True,
                        "foregroundColor": hex_to_color("#64748B"),
                    },
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(textFormat,horizontalAlignment)",
        }
    })

    # Главный заголовок документа (Строка 5)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, 4, 5, 1, 10),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": hex_to_color("#0F172A"),
                    "textFormat": {
                        "fontFamily": "Arial",
                        "fontSize": 14,
                        "bold": True,
                        "foregroundColor": hex_to_color("#FFFFFF"),
                    },
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
        }
    })

    # Период (Строка 6)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, 5, 6, 1, 10),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": hex_to_color("#F1F5F9"),
                    "textFormat": {
                        "fontFamily": "Arial",
                        "fontSize": 10,
                        "bold": True,
                        "foregroundColor": hex_to_color("#1E293B"),
                    },
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
        }
    })

    # Паспорт документа (Строки 8-9)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, 7, 9, 1, 10),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": hex_to_color("#F8FAFC"),
                    "textFormat": {
                        "fontSize": 9,
                    },
                    "borders": {
                        "top": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                        "bottom": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                        "left": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                        "right": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                    },
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,borders)",
        }
    })

    # Карточки KPI (Строки 11-13)
    kpi_configs = [
        (1, 3, "#EFF6FF", "#1D4ED8", True),   # Выручка (Синий)
        (3, 5, "#ECFDF5", "#047857", False),  # Сделки (Зеленый)
        (5, 7, "#FFFBEB", "#B45309", True),   # Средний чек (Янтарный)
        (7, 9, "#F5F3FF", "#6D28D9", False),  # План (Фиолетовый)
    ]
    for start_c, end_c, bg_color, text_color, is_currency in kpi_configs:
        # Верхний заголовок карточки
        requests.append({
            "repeatCell": {
                "range": grid_range(sheet_id, 10, 11, start_c, end_c),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": hex_to_color(bg_color),
                        "textFormat": {
                            "fontSize": 9,
                            "bold": True,
                            "foregroundColor": hex_to_color(text_color),
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
            }
        })
        # Центральное число карточки
        number_format = {
            "type": "CURRENCY",
            "pattern": '#,##0.00 "₽"'
        } if is_currency else {"type": "NUMBER", "pattern": "#,##0"}

        requests.append({
            "repeatCell": {
                "range": grid_range(sheet_id, 11, 12, start_c, end_c),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": hex_to_color(bg_color),
                        "textFormat": {
                            "fontSize": 15,
                            "bold": True,
                            "foregroundColor": hex_to_color(text_color),
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "numberFormat": number_format,
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)",
            }
        })
        # Нижняя подпись карточки
        requests.append({
            "repeatCell": {
                "range": grid_range(sheet_id, 12, 13, start_c, end_c),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": hex_to_color(bg_color),
                        "textFormat": {
                            "fontSize": 8,
                            "italic": True,
                            "foregroundColor": hex_to_color(text_color),
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
            }
        })
        # Рамка для всей карточки
        requests.append({
            "repeatCell": {
                "range": grid_range(sheet_id, 10, 13, start_c, end_c),
                "cell": {
                    "userEnteredFormat": {
                        "borders": {
                            "top": {"style": "SOLID_MEDIUM", "color": hex_to_color(text_color)},
                            "bottom": {"style": "SOLID", "color": hex_to_color(text_color)},
                            "left": {"style": "SOLID", "color": hex_to_color(text_color)},
                            "right": {"style": "SOLID", "color": hex_to_color(text_color)},
                        }
                    }
                },
                "fields": "userEnteredFormat(borders)",
            }
        })

    # Разделы I и II (Заголовки секций)
    for section_row in [14, total_row + 1]:
        requests.append({
            "repeatCell": {
                "range": grid_range(sheet_id, section_row, section_row + 1, 1, 10),
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "fontSize": 11,
                            "bold": True,
                            "foregroundColor": hex_to_color("#0F172A"),
                        },
                        "borders": {
                            "bottom": {"style": "SOLID_MEDIUM", "color": hex_to_color("#2563EB")},
                        },
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(textFormat,borders,verticalAlignment)",
            }
        })

    # Шапка таблицы (Строка 17, индексы 16..17)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, 16, 17, 1, 10),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": hex_to_color("#1E3A8A"),
                    "textFormat": {
                        "fontSize": 10,
                        "bold": True,
                        "foregroundColor": hex_to_color("#FFFFFF"),
                    },
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "borders": {
                        "top": {"style": "SOLID", "color": hex_to_color("#172554")},
                        "bottom": {"style": "SOLID_MEDIUM", "color": hex_to_color("#172554")},
                        "left": {"style": "SOLID", "color": hex_to_color("#3B82F6")},
                        "right": {"style": "SOLID", "color": hex_to_color("#3B82F6")},
                    },
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)",
        }
    })

    # Чередующаяся заливка строк таблицы (Зебра) и границы
    for r_idx in range(data_start_row - 1, data_end_row):
        is_even = (r_idx % 2 == 0)
        row_bg = "#FFFFFF" if is_even else "#F8FAFC"
        requests.append({
            "repeatCell": {
                "range": grid_range(sheet_id, r_idx, r_idx + 1, 1, 10),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": hex_to_color(row_bg),
                        "borders": {
                            "top": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                            "bottom": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                            "left": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                            "right": {"style": "SOLID", "color": hex_to_color("#E2E8F0")},
                        },
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,borders)",
            }
        })

    # Выравнивание и формат чисел колонок таблицы
    # B: № (центр)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, data_start_row - 1, data_end_row, 1, 2),
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment",
        }
    })
    # C: Дата (центр)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, data_start_row - 1, data_end_row, 2, 3),
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment",
        }
    })
    # D: Документ (центр)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, data_start_row - 1, data_end_row, 3, 4),
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment",
        }
    })
    # H: Кол-во (центр/право)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, data_start_row - 1, total_row, 7, 8),
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "RIGHT",
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0"},
                }
            },
            "fields": "userEnteredFormat(horizontalAlignment,numberFormat)",
        }
    })
    # I: Сумма (право + денежный формат)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, data_start_row - 1, total_row, 8, 9),
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "RIGHT",
                    "numberFormat": {"type": "CURRENCY", "pattern": '#,##0.00 "₽"'},
                }
            },
            "fields": "userEnteredFormat(horizontalAlignment,numberFormat)",
        }
    })
    # J: Статус (центр, бейдж)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, data_start_row - 1, data_end_row, 9, 10),
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER",
                    "textFormat": {"bold": True, "fontSize": 9},
                }
            },
            "fields": "userEnteredFormat(horizontalAlignment,textFormat)",
        }
    })

    # Строка ИТОГО (total_row)
    requests.append({
        "repeatCell": {
            "range": grid_range(sheet_id, total_row - 1, total_row, 1, 10),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": hex_to_color("#E2E8F0"),
                    "textFormat": {
                        "fontSize": 11,
                        "bold": True,
                        "foregroundColor": hex_to_color("#0F172A"),
                    },
                    "horizontalAlignment": "RIGHT",
                    "verticalAlignment": "MIDDLE",
                    "borders": {
                        "top": {"style": "SOLID_MEDIUM", "color": hex_to_color("#475569")},
                        "bottom": {"style": "DOUBLE", "color": hex_to_color("#0F172A")},
                    },
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)",
        }
    })

    # Выполняем все стили разом через batchUpdate
    notify("Отправка пакета оформления в Google Sheets API...", 85)
    client.service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()

    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={sheet_id}"
    notify("Отчет успешно сформирован!", 100)

    return sheet_id, sheet_url

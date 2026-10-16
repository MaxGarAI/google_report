"""
Модуль для работы с Google Sheets API v4 через сервисный аккаунт.
Предоставляет класс GoogleSheetsClient для выполнения CRUD-операций:
- Create: добавление строк (append), создание новых листов.
- Read: чтение ячеек, диапазонов, всего листа, чтение данных как словарей (dict).
- Update: обновление отдельных ячеек, диапазонов, целых строк.
- Delete: очистка диапазонов (clear), физическое удаление строк, удаление листов.

Модуль готов к импорту в любые другие проекты.
"""

from __future__ import annotations

import glob
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

# Настройка безопасного UTF-8 вывода для Windows консоли
if sys.platform == "win32":
    try:
        if sys.stdout and hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if sys.stderr and hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import Resource, build
from googleapiclient.errors import HttpError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


class GoogleSheetsClient:
    """
    Клиент для управления Google Таблицами через официальный Google Sheets API v4.
    """

    DEFAULT_SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
    ]

    def __init__(
        self,
        credentials_path: Optional[Union[str, Path]] = None,
        spreadsheet_id: Optional[str] = None,
        scopes: Optional[List[str]] = None,
    ) -> None:
        """
        Инициализирует подключение к Google Sheets API через сервисный аккаунт.

        :param credentials_path: Путь к JSON-файлу ключа сервисного аккаунта.
                                 Если не указан, производится автопоиск в текущей папке.
        :param spreadsheet_id: ID Google Таблицы (из URL таблицы: /d/<SPREADSHEET_ID>/edit).
        :param scopes: OAuth 2.0 области видимости (по умолчанию full access к sheets).
        """
        self.scopes = scopes or self.DEFAULT_SCOPES
        self.credentials_path = self._resolve_credentials_path(credentials_path)
        self.spreadsheet_id = spreadsheet_id

        # Авторизация и создание сервиса
        self._credentials = Credentials.from_service_account_file(
            str(self.credentials_path),
            scopes=self.scopes,
        )
        self.service: Resource = build(
            "sheets", "v4", credentials=self._credentials, cache_discovery=False
        )
        self._sheet_resource = self.service.spreadsheets()

        logger.info(
            "GoogleSheetsClient успешно инициализирован. Аккаунт: %s",
            self.service_account_email,
        )

    @property
    def service_account_email(self) -> str:
        """Возвращает email сервисного аккаунта, которому нужно дать доступ к таблице."""
        return self._credentials.service_account_email

    def set_spreadsheet_id(self, spreadsheet_id: str) -> None:
        """Устанавливает или изменяет рабочий ID таблицы."""
        self.spreadsheet_id = spreadsheet_id

    def _require_spreadsheet_id(self, spreadsheet_id: Optional[str] = None) -> str:
        """Проверяет наличие ID таблицы."""
        target_id = spreadsheet_id or self.spreadsheet_id
        if not target_id:
            raise ValueError(
                "ID таблицы не указан. Передайте spreadsheet_id в конструктор или в метод."
            )
        return target_id

    @staticmethod
    def _resolve_credentials_path(credentials_path: Optional[Union[str, Path]]) -> Path:
        """Находит и валидирует путь к JSON-ключу сервисного аккаунта."""
        if credentials_path:
            path = Path(credentials_path)
            if not path.is_file():
                raise FileNotFoundError(f"Файл ключа не найден по пути: {path.resolve()}")
            return path

        # Проверка переменной окружения
        env_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
        if env_path and Path(env_path).is_file():
            return Path(env_path)

        # Автопоиск *.json в текущей рабочей директории и рядом со скриптом
        search_dirs = [Path.cwd(), Path(__file__).resolve().parent]
        for directory in search_dirs:
            for json_file in directory.glob("*.json"):
                try:
                    with open(json_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if data.get("type") == "service_account" and "private_key" in data:
                            logger.info("Автоматически обнаружен ключ сервисного аккаунта: %s", json_file.name)
                            return json_file
                except Exception:
                    continue

        raise FileNotFoundError(
            "Не удалось найти JSON-файл ключа сервисного аккаунта. "
            "Передайте путь явно в credentials_path или задайте переменную GOOGLE_APPLICATION_CREDENTIALS."
        )

    # =========================================================================
    # METADATA & HELPER METHODS
    # =========================================================================

    def get_spreadsheet_metadata(self, spreadsheet_id: Optional[str] = None) -> Dict[str, Any]:
        """
        Возвращает метаданные таблицы (название, список листов, их ID и размеры).
        """
        sid = self._require_spreadsheet_id(spreadsheet_id)
        try:
            return self._sheet_resource.get(spreadsheetId=sid).execute()
        except HttpError as error:
            logger.error("Ошибка при получении метаданных таблицы: %s", error)
            raise

    def get_sheet_names(self, spreadsheet_id: Optional[str] = None) -> List[str]:
        """Возвращает список имен всех листов (вкладок) в таблице."""
        metadata = self.get_spreadsheet_metadata(spreadsheet_id)
        return [sheet["properties"]["title"] for sheet in metadata.get("sheets", [])]

    def get_default_sheet_name(self, spreadsheet_id: Optional[str] = None) -> str:
        """Возвращает имя первого листа в таблице."""
        sheet_names = self.get_sheet_names(spreadsheet_id)
        if not sheet_names:
            raise ValueError("Таблица не содержит ни одного листа.")
        return sheet_names[0]

    def get_sheet_id(self, sheet_name: Optional[str] = None, spreadsheet_id: Optional[str] = None) -> int:
        """
        Возвращает числовой ID листа (sheetId) по его названию.
        Если sheet_name не передан, возвращается ID первого листа.
        """
        metadata = self.get_spreadsheet_metadata(spreadsheet_id)
        sheets = metadata.get("sheets", [])
        if not sheets:
            raise ValueError("Таблица не содержит листов.")

        if sheet_name is None:
            return int(sheets[0]["properties"]["sheetId"])

        for sheet in sheets:
            props = sheet.get("properties", {})
            if props.get("title") == sheet_name:
                return int(props["sheetId"])

        raise ValueError(f"Лист с названием '{sheet_name}' не найден.")

    def _format_range(self, sheet_name: Optional[str], cell_range: Optional[str] = None) -> str:
        """Форматирует диапазон A1 с учетом экранирования имени листа кавычками."""
        name = sheet_name or self.get_default_sheet_name()
        escaped_name = f"'{name}'" if "'" not in name else name
        if cell_range:
            return f"{escaped_name}!{cell_range}"
        return escaped_name

    # =========================================================================
    # CREATE (Добавление строк, создание листов)
    # =========================================================================

    def append_row(
        self,
        values: List[Any],
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Добавляет одну строку в конец заполненной области листа.

        :param values: Список значений для строки, например: ["Иван", 30, "ivan@example.com"].
        :param sheet_name: Название листа (если None, берется первый лист).
        :param spreadsheet_id: ID таблицы.
        :param value_input_option: "USER_ENTERED" (парсит числа/даты) или "RAW".
        """
        return self.append_rows(
            rows=[values],
            sheet_name=sheet_name,
            spreadsheet_id=spreadsheet_id,
            value_input_option=value_input_option,
        )

    def append_rows(
        self,
        rows: List[List[Any]],
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Добавляет несколько строк в конец заполненной области листа.

        :param rows: Список строк (матрица значений), например: [[1, "A"], [2, "B"]].
        """
        sid = self._require_spreadsheet_id(spreadsheet_id)
        target_range = self._format_range(sheet_name)

        body = {"values": rows}
        try:
            result = (
                self._sheet_resource.values()
                .append(
                    spreadsheetId=sid,
                    range=target_range,
                    valueInputOption=value_input_option,
                    insertDataOption="INSERT_ROWS",
                    body=body,
                )
                .execute()
            )
            logger.info("Успешно добавлено строк: %s в '%s'", len(rows), target_range)
            return result
        except HttpError as error:
            logger.error("Ошибка при добавлении строк: %s", error)
            raise

    def create_sheet(self, title: str, spreadsheet_id: Optional[str] = None) -> Dict[str, Any]:
        """
        Создает новую вкладку (лист) в таблице с указанным названием.
        """
        sid = self._require_spreadsheet_id(spreadsheet_id)
        request_body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": title,
                        }
                    }
                }
            ]
        }
        try:
            result = self._sheet_resource.batchUpdate(spreadsheetId=sid, body=request_body).execute()
            logger.info("Создан новый лист: '%s'", title)
            return result
        except HttpError as error:
            logger.error("Ошибка при создании листа '%s': %s", title, error)
            raise

    # =========================================================================
    # READ (Чтение ячеек, диапазонов, всего листа)
    # =========================================================================

    def read_range(
        self,
        cell_range: str,
        spreadsheet_id: Optional[str] = None,
    ) -> List[List[Any]]:
        """
        Считывает диапазон ячеек в формате A1 (например: 'Лист1!A1:D10' или 'A1:C5').
        Возвращает двумерный список значений.
        """
        sid = self._require_spreadsheet_id(spreadsheet_id)
        try:
            result = (
                self._sheet_resource.values()
                .get(spreadsheetId=sid, range=cell_range)
                .execute()
            )
            return result.get("values", [])
        except HttpError as error:
            logger.error("Ошибка при чтении диапазона '%s': %s", cell_range, error)
            raise

    def read_all(
        self,
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
    ) -> List[List[Any]]:
        """
        Считывает все заполненные ячейки указанного листа (или первого листа по умолчанию).
        """
        target_range = self._format_range(sheet_name)
        return self.read_range(cell_range=target_range, spreadsheet_id=spreadsheet_id)

    def read_records(
        self,
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        Считывает лист как список словарей. Первая строка считается заголовками колонок.
        Если в строке меньше колонок, чем в заголовке, недостающие заполняются None.
        """
        rows = self.read_all(sheet_name=sheet_name, spreadsheet_id=spreadsheet_id)
        if not rows:
            return []

        headers = [str(col).strip() for col in rows[0]]
        records: List[Dict[str, Any]] = []

        for row in rows[1:]:
            record = {}
            for i, header in enumerate(headers):
                record[header] = row[i] if i < len(row) else None
            records.append(record)

        return records

    def read_cell(
        self,
        cell: str,
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
    ) -> Any:
        """
        Считывает значение одиночной ячейки (например: cell='B2').
        """
        full_range = self._format_range(sheet_name, cell)
        data = self.read_range(cell_range=full_range, spreadsheet_id=spreadsheet_id)
        if data and data[0]:
            return data[0][0]
        return None

    # =========================================================================
    # UPDATE (Обновление ячеек, диапазонов, строк)
    # =========================================================================

    def update_range(
        self,
        cell_range: str,
        values: List[List[Any]],
        spreadsheet_id: Optional[str] = None,
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Обновляет диапазон ячеек новыми данными.

        :param cell_range: Диапазон ячеек (например 'Лист1!A1:B2' или 'B5:D6').
        :param values: Двумерный список значений.
        :param value_input_option: "USER_ENTERED" или "RAW".
        """
        sid = self._require_spreadsheet_id(spreadsheet_id)
        body = {"values": values}
        try:
            result = (
                self._sheet_resource.values()
                .update(
                    spreadsheetId=sid,
                    range=cell_range,
                    valueInputOption=value_input_option,
                    body=body,
                )
                .execute()
            )
            logger.info("Успешно обновлен диапазон '%s'", cell_range)
            return result
        except HttpError as error:
            logger.error("Ошибка при обновлении диапазона '%s': %s", cell_range, error)
            raise

    def update_cell(
        self,
        cell: str,
        value: Any,
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Обновляет значение одиночной ячейки (например: cell='A1', value='Тест').
        """
        target_range = self._format_range(sheet_name, cell)
        return self.update_range(
            cell_range=target_range,
            values=[[value]],
            spreadsheet_id=spreadsheet_id,
            value_input_option=value_input_option,
        )

    def update_row(
        self,
        row_index: int,
        values: List[Any],
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Перезаписывает значения конкретной строки по ее порядковому номеру (1-indexed).

        :param row_index: Номер строки (например, 2 для второй строки).
        :param values: Список значений для строки.
        """
        if row_index < 1:
            raise ValueError("Номер строки row_index должен быть >= 1.")
        target_range = self._format_range(sheet_name, f"A{row_index}")
        return self.update_range(
            cell_range=target_range,
            values=[values],
            spreadsheet_id=spreadsheet_id,
            value_input_option=value_input_option,
        )

    # =========================================================================
    # DELETE (Очистка диапазонов, удаление строк, удаление листов)
    # =========================================================================

    def clear_range(
        self,
        cell_range: str,
        spreadsheet_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Очищает содержимое ячеек в указанном диапазоне (например 'Лист1!A2:D100').
        """
        sid = self._require_spreadsheet_id(spreadsheet_id)
        try:
            result = (
                self._sheet_resource.values()
                .clear(spreadsheetId=sid, range=cell_range)
                .execute()
            )
            logger.info("Очищен диапазон '%s'", cell_range)
            return result
        except HttpError as error:
            logger.error("Ошибка при очистке диапазона '%s': %s", cell_range, error)
            raise

    def clear_sheet(
        self,
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Очищает все значения на указанном листе.
        """
        target_range = self._format_range(sheet_name)
        return self.clear_range(cell_range=target_range, spreadsheet_id=spreadsheet_id)

    def delete_row(
        self,
        row_index: int,
        sheet_name: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Физически удаляет строку из таблицы со сдвигом остальных строк вверх.

        :param row_index: Номер строки (1-indexed, например 5).
        :param sheet_name: Имя листа (если None, берется первый лист).
        """
        if row_index < 1:
            raise ValueError("Номер строки row_index должен быть >= 1.")

        sid = self._require_spreadsheet_id(spreadsheet_id)
        sheet_id = self.get_sheet_id(sheet_name=sheet_name, spreadsheet_id=sid)

        # В API индексы 0-based: [startIndex, endIndex)
        start_index = row_index - 1
        end_index = row_index

        request_body = {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": start_index,
                            "endIndex": end_index,
                        }
                    }
                }
            ]
        }
        try:
            result = self._sheet_resource.batchUpdate(spreadsheetId=sid, body=request_body).execute()
            logger.info("Удалена строка %s на листе '%s' (sheetId=%s)", row_index, sheet_name or "default", sheet_id)
            return result
        except HttpError as error:
            logger.error("Ошибка при удалении строки %s: %s", row_index, error)
            raise

    def delete_sheet(
        self,
        sheet_name: str,
        spreadsheet_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Удаляет лист (вкладку) из таблицы по его имени.
        """
        sid = self._require_spreadsheet_id(spreadsheet_id)
        sheet_id = self.get_sheet_id(sheet_name=sheet_name, spreadsheet_id=sid)

        request_body = {
            "requests": [
                {
                    "deleteSheet": {
                        "sheetId": sheet_id,
                    }
                }
            ]
        }
        try:
            result = self._sheet_resource.batchUpdate(spreadsheetId=sid, body=request_body).execute()
            logger.info("Лист '%s' (sheetId=%s) успешно удален", sheet_name, sheet_id)
            return result
        except HttpError as error:
            logger.error("Ошибка при удалении листа '%s': %s", sheet_name, error)
            raise


# =============================================================================
# ТОЧКА ВХОДА (Для проверки работы и чтения всех ячеек)
# =============================================================================
if __name__ == "__main__":
    import sys

    print("=" * 60)
    print("Проверка GoogleSheetsClient через сервисный аккаунт")
    print("=" * 60)

    # 1. Поиск ключа сервисного аккаунта (в текущей папке проекта)
    # По умолчанию клиент сам найдет файл excel-factory-511811-6090cec269a3.json в папке проекта
    client = GoogleSheetsClient()

    print(f"Сервисный аккаунт: {client.service_account_email}")
    print("-" * 60)

    # 2. Получение SPREADSHEET_ID
    # Можно передать аргументом командной строки: python google_sheets_client.py <SPREADSHEET_ID>
    # Либо задать в переменной окружения SPREADSHEET_ID
    # Либо указать в коде ниже
    spreadsheet_id = None

    if len(sys.argv) > 1 and sys.argv[1].strip():
        spreadsheet_id = sys.argv[1].strip()
    elif os.getenv("SPREADSHEET_ID"):
        spreadsheet_id = os.getenv("SPREADSHEET_ID").strip()
    else:
        # Вставьте ваш ID таблицы сюда, если не используете параметры командной строки
        spreadsheet_id = ""

    if not spreadsheet_id:
        print("[!] ВНИМАНИЕ: Для выполнения запроса требуется SPREADSHEET_ID.")
        print("  1. Откройте нужную Google Таблицу в браузере.")
        print("  2. Скопируйте ID из адресной строки: https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit")
        print(f"  3. ОБЯЗАТЕЛЬНО предоставьте доступ Редактора сервисному аккаунту:")
        print(f"     --> {client.service_account_email}")
        print("  4. Запустите скрипт с указанием ID:")
        print("     python google_sheets_client.py <SPREADSHEET_ID>")
        print("     или введите ID сейчас.")
        print("-" * 60)
        try:
            user_input = input("Введите SPREADSHEET_ID (или Enter для выхода): ").strip()
            if user_input:
                spreadsheet_id = user_input
            else:
                print("Завершение работы без чтения таблицы.")
                sys.exit(0)
        except (EOFError, KeyboardInterrupt):
            sys.exit(0)

    client.set_spreadsheet_id(spreadsheet_id)

    print(f"\nПодключение к таблице с ID: {spreadsheet_id}...")
    try:
        # Метаданные таблицы и список листов
        metadata = client.get_spreadsheet_metadata()
        title = metadata.get("properties", {}).get("title", "Без названия")
        sheets = client.get_sheet_names()

        print(f"Таблица успешно открыта: «{title}»")
        print(f"Доступные листы ({len(sheets)}): {', '.join(sheets)}")
        print("-" * 60)

        # Чтение всех ячеек первого листа
        first_sheet = sheets[0]
        print(f"Чтение всех ячеек с листа «{first_sheet}»:\n")

        all_cells = client.read_all(sheet_name=first_sheet)

        if not all_cells:
            print(f"[i] Лист «{first_sheet}» пуст (нет данных).")
        else:
            print(f"Найдено строк: {len(all_cells)}")
            for idx, row in enumerate(all_cells, start=1):
                print(f"Строка {idx:3d}: {row}")

        print("\n" + "=" * 60)
        print("Проверка чтения завершена успешно!")
        print("=" * 60)

    except HttpError as err:
        print(f"\n[X] Ошибка Google API (код {err.resp.status}):")
        if err.resp.status == 404:
            print("Таблица с указанным ID не найдена. Проверьте правильность SPREADSHEET_ID.")
        elif err.resp.status == 403:
            print("Отказано в доступе! Убедитесь, что вы расшарили таблицу (Share -> Добавить пользователя)")
            print(f"для сервисного аккаунта: {client.service_account_email} с правами 'Редактор' или 'Читатель'.")
        else:
            print(err)
    except Exception as exc:
        print(f"\n[X] Ошибка: {exc}")


        # Пример добавления строки при ручном запуске
        # client.append_row(["Иван Иванов", "Программист", 150000])
        pass
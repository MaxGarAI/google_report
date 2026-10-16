"""
Графическое приложение на Tkinter для генерации отчетов в Google Sheets.
Симулирует данные и строит красивый документ с богатым форматированием (как в Word/Excel).
"""

from __future__ import annotations

import datetime
import json
import os
import random
import re
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import webbrowser
from pathlib import Path
from typing import Optional

from google_sheets_client import GoogleSheetsClient
from report_generator import (
    DEPARTMENT_DATA,
    build_sheet_report,
    generate_random_report_data,
)

CONFIG_FILE = Path("app_config.json")


def extract_spreadsheet_id(input_value: str) -> str:
    """Извлекает ID таблицы из сырой строки или ссылки Google Sheets."""
    input_value = input_value.strip()
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", input_value)
    if match:
        return match.group(1)
    return input_value


class ModernReportApp(tk.Tk):
    """Главное окно приложения генератора отчетов."""

    def __init__(self) -> None:
        super().__init__()

        self.title("Генератор официальных отчетов в Google Sheets")
        self.geometry("920x820")
        self.minsize(860, 720)

        # Переменные состояния
        self.client: Optional[GoogleSheetsClient] = None
        self.last_sheet_url: Optional[str] = None
        self.is_processing = False

        # Загрузка конфигурации
        self.config_data = self._load_config()

        # Попытка инициализировать клиент Google Sheets
        self._init_sheets_client()

        # Настройка стилей ttk
        self._setup_styles()

        # Создание интерфейса
        self._create_widgets()

        # Заполнение значений по умолчанию
        self._set_default_values()

    def _load_config(self) -> dict:
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_config(self) -> None:
        data = {
            "spreadsheet_id": self.var_spreadsheet_id.get().strip(),
            "department": self.var_department.get(),
            "author": self.var_author.get().strip(),
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _init_sheets_client(self) -> None:
        try:
            self.client = GoogleSheetsClient()
            self.service_email = self.client.service_account_email
        except Exception as e:
            self.client = None
            self.service_email = f"Ошибка поиска ключа сервисного аккаунта: {e}"

    def _setup_styles(self) -> None:
        self.style = ttk.Style(self)
        # Использование современной доступной темы
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.configure(bg="#F8FAFC")

        # Настройка шрифтов и цветов ttk
        self.style.configure("TFrame", background="#F8FAFC")
        self.style.configure("Card.TFrame", background="#FFFFFF", relief="solid", borderwidth=1)
        self.style.configure("TLabel", background="#F8FAFC", foreground="#1E293B", font=("Segoe UI", 9))
        self.style.configure("CardLabel.TLabel", background="#FFFFFF", foreground="#1E293B", font=("Segoe UI", 9))
        self.style.configure("Header.TLabel", background="#F8FAFC", foreground="#0F172A", font=("Segoe UI", 16, "bold"))
        self.style.configure("SubHeader.TLabel", background="#F8FAFC", foreground="#64748B", font=("Segoe UI", 9))
        self.style.configure("Section.TLabel", background="#FFFFFF", foreground="#1E3A8A", font=("Segoe UI", 11, "bold"))

        # Кнопки
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#2563EB",
            foreground="#FFFFFF",
            padding=(12, 8),
            borderwidth=0,
        )
        self.style.map("Primary.TButton", background=[("active", "#1D4ED8"), ("disabled", "#94A3B8")])

        self.style.configure(
            "Success.TButton",
            font=("Segoe UI", 9, "bold"),
            background="#059669",
            foreground="#FFFFFF",
            padding=(10, 6),
            borderwidth=0,
        )
        self.style.map("Success.TButton", background=[("active", "#047857"), ("disabled", "#94A3B8")])

        self.style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            background="#F1F5F9",
            foreground="#334155",
            padding=(8, 4),
            borderwidth=1,
        )
        self.style.map("Secondary.TButton", background=[("active", "#E2E8F0")])

        # Treeview (таблица предпросмотра)
        self.style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background="#E2E8F0",
            foreground="#1E293B",
            relief="flat",
        )
        self.style.configure("Treeview", font=("Segoe UI", 8), rowheight=22)

    def _create_widgets(self) -> None:
        # Главный скроллируемый или растягиваемый контейнер
        main_container = ttk.Frame(self, padding=16)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 1. Шапка приложения
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 12))

        lbl_title = ttk.Label(header_frame, text="📊 Генератор официальных отчетов в Google Sheets", style="Header.TLabel")
        lbl_title.pack(anchor=tk.W)

        lbl_desc = ttk.Label(
            header_frame,
            text="Формирование презентабельных корпоративных отчетов с автоматическим расчетом и дизайном Word-документа.",
            style="SubHeader.TLabel",
        )
        lbl_desc.pack(anchor=tk.W, pady=(2, 0))

        # 2. Карточка 1: Настройки подключения к Google Sheets
        conn_card = ttk.Frame(main_container, style="Card.TFrame", padding=14)
        conn_card.pack(fill=tk.X, pady=(0, 10))

        lbl_conn_title = ttk.Label(conn_card, text="🔑 1. Параметры Google Таблицы", style="Section.TLabel")
        lbl_conn_title.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        # Сервисный аккаунт
        ttk.Label(conn_card, text="Сервисный email бота:", style="CardLabel.TLabel").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.lbl_email_val = ttk.Label(
            conn_card,
            text=self.service_email,
            font=("Consolas", 8, "bold"),
            foreground="#2563EB",
            background="#FFFFFF",
        )
        self.lbl_email_val.grid(row=1, column=1, sticky=tk.W, padx=6, pady=3)

        btn_copy_email = ttk.Button(conn_card, text="📋 Скопировать email", style="Secondary.TButton", command=self._copy_email)
        btn_copy_email.grid(row=1, column=2, sticky=tk.E, pady=3)

        # Подсказка
        lbl_hint = ttk.Label(
            conn_card,
            text="⚠️ Важно: Откройте доступ 'Редактор' для этого email в вашей Google Таблице (кнопка 'Настройки доступа' / 'Share').",
            font=("Segoe UI", 8, "italic"),
            foreground="#D97706",
            background="#FFFFFF",
        )
        lbl_hint.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(0, 6))

        # Поле ввода Spreadsheet ID / URL
        ttk.Label(conn_card, text="ID или ссылка на Таблицу:", style="CardLabel.TLabel").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.var_spreadsheet_id = tk.StringVar()
        self.ent_spreadsheet = ttk.Entry(conn_card, textvariable=self.var_spreadsheet_id, font=("Segoe UI", 9), width=50)
        self.ent_spreadsheet.grid(row=3, column=1, sticky=tk.EW, padx=6, pady=4)

        # Название листа
        ttk.Label(conn_card, text="Имя листа для отчета:", style="CardLabel.TLabel").grid(row=4, column=0, sticky=tk.W, pady=4)
        sheet_frame = ttk.Frame(conn_card, style="Card.TFrame")
        sheet_frame.grid(row=4, column=1, columnspan=2, sticky=tk.W, padx=6, pady=4)

        self.var_sheet_title = tk.StringVar()
        self.ent_sheet_title = ttk.Entry(sheet_frame, textvariable=self.var_sheet_title, font=("Segoe UI", 9), width=28)
        self.ent_sheet_title.pack(side=tk.LEFT)

        btn_regen_title = ttk.Button(sheet_frame, text="🔄 Новое имя", style="Secondary.TButton", command=self._generate_sheet_name)
        btn_regen_title.pack(side=tk.LEFT, padx=6)

        conn_card.columnconfigure(1, weight=1)

        # 3. Карточка 2: Параметры отчета и фильтры
        params_card = ttk.Frame(main_container, style="Card.TFrame", padding=14)
        params_card.pack(fill=tk.X, pady=(0, 10))

        lbl_params_title = ttk.Label(params_card, text="⚙️ 2. Параметры и период формирования", style="Section.TLabel")
        lbl_params_title.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 8))

        # Диапазон дат
        ttk.Label(params_card, text="Дата начала (от):", style="CardLabel.TLabel").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.var_date_from = tk.StringVar()
        self.ent_date_from = ttk.Entry(params_card, textvariable=self.var_date_from, width=14, font=("Segoe UI", 9))
        self.ent_date_from.grid(row=1, column=1, sticky=tk.W, padx=4, pady=3)

        ttk.Label(params_card, text="Дата окончания (до):", style="CardLabel.TLabel").grid(row=1, column=2, sticky=tk.W, padx=(10, 0), pady=3)
        self.var_date_to = tk.StringVar()
        self.ent_date_to = ttk.Entry(params_card, textvariable=self.var_date_to, width=14, font=("Segoe UI", 9))
        self.ent_date_to.grid(row=1, column=3, sticky=tk.W, padx=4, pady=3)

        # Быстрые пресеты дат
        presets_frame = ttk.Frame(params_card, style="Card.TFrame")
        presets_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(4, 8))

        ttk.Label(presets_frame, text="Быстрый выбор периода:", style="CardLabel.TLabel", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(presets_frame, text="Текущий месяц", style="Secondary.TButton", command=self._preset_current_month).pack(side=tk.LEFT, padx=2)
        ttk.Button(presets_frame, text="Прошлый месяц", style="Secondary.TButton", command=self._preset_last_month).pack(side=tk.LEFT, padx=2)
        ttk.Button(presets_frame, text="Текущий квартал", style="Secondary.TButton", command=self._preset_quarter).pack(side=tk.LEFT, padx=2)
        ttk.Button(presets_frame, text="Случайные 30 дней", style="Secondary.TButton", command=self._preset_random_period).pack(side=tk.LEFT, padx=2)

        # Департамент
        ttk.Label(params_card, text="Подразделение / Департамент:", style="CardLabel.TLabel").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.var_department = tk.StringVar()
        self.cmb_department = ttk.Combobox(
            params_card,
            textvariable=self.var_department,
            values=list(DEPARTMENT_DATA.keys()),
            state="readonly",
            width=36,
            font=("Segoe UI", 9),
        )
        self.cmb_department.grid(row=3, column=1, columnspan=3, sticky=tk.W, padx=4, pady=4)

        # Составитель
        ttk.Label(params_card, text="Ответственный составитель:", style="CardLabel.TLabel").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.var_author = tk.StringVar()
        self.ent_author = ttk.Entry(params_card, textvariable=self.var_author, font=("Segoe UI", 9), width=38)
        self.ent_author.grid(row=4, column=1, columnspan=3, sticky=tk.W, padx=4, pady=4)

        # Количество строк (Scale)
        ttk.Label(params_card, text="Количество операций в отчете:", style="CardLabel.TLabel").grid(row=5, column=0, sticky=tk.W, pady=4)
        scale_frame = ttk.Frame(params_card, style="Card.TFrame")
        scale_frame.grid(row=5, column=1, columnspan=3, sticky=tk.W, padx=4, pady=4)

        self.var_rows_count = tk.IntVar(value=18)
        self.scale_rows = ttk.Scale(scale_frame, from_=5, to=40, orient=tk.HORIZONTAL, variable=self.var_rows_count, length=180, command=self._on_scale_change)
        self.scale_rows.pack(side=tk.LEFT)

        self.lbl_rows_display = ttk.Label(scale_frame, text="18 операций", font=("Segoe UI", 9, "bold"), background="#FFFFFF", foreground="#2563EB")
        self.lbl_rows_display.pack(side=tk.LEFT, padx=10)

        # Кнопка рандомизации
        btn_randomize = ttk.Button(params_card, text="🎲 Случайные параметры", style="Secondary.TButton", command=self._randomize_all)
        btn_randomize.grid(row=5, column=3, sticky=tk.E, pady=4)

        # 4. Блок кнопок действий
        actions_frame = ttk.Frame(main_container)
        actions_frame.pack(fill=tk.X, pady=(4, 10))

        self.btn_submit = ttk.Button(
            actions_frame,
            text="🚀 Сформировать и отправить отчет в Google Sheets",
            style="Primary.TButton",
            command=self._start_generation_thread,
        )
        self.btn_submit.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_preview = ttk.Button(
            actions_frame,
            text="👁️ Предпросмотр данных",
            style="Secondary.TButton",
            command=self._generate_and_show_preview,
        )
        self.btn_preview.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_browser = ttk.Button(
            actions_frame,
            text="🌐 Открыть таблицу в браузере",
            style="Success.TButton",
            command=self._open_in_browser,
        )
        self.btn_open_browser.pack(side=tk.RIGHT)

        # 5. Прогресс-бар и статусная строка
        self.var_status = tk.StringVar(value="Готов к созданию отчета.")
        self.lbl_status = ttk.Label(main_container, textvariable=self.var_status, font=("Segoe UI", 9, "bold"), foreground="#475569")
        self.lbl_status.pack(anchor=tk.W, pady=(2, 4))

        self.progress_bar = ttk.Progressbar(main_container, orient=tk.HORIZONTAL, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        # 6. Предпросмотр (Treeview таблица)
        preview_label = ttk.Label(main_container, text="Интерактивный предпросмотр операций (Word-макет в Excel):", font=("Segoe UI", 9, "bold"))
        preview_label.pack(anchor=tk.W, pady=(2, 4))

        tree_frame = ttk.Frame(main_container)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("idx", "date", "doc_num", "client", "item", "manager", "qty", "amount", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=8)

        self.tree.heading("idx", text="№")
        self.tree.heading("date", text="Дата")
        self.tree.heading("doc_num", text="Документ")
        self.tree.heading("client", text="Контрагент")
        self.tree.heading("item", text="Номенклатура / Услуга")
        self.tree.heading("manager", text="Ответственный")
        self.tree.heading("qty", text="Кол-во")
        self.tree.heading("amount", text="Сумма, ₽")
        self.tree.heading("status", text="Статус")

        self.tree.column("idx", width=35, anchor=tk.CENTER)
        self.tree.column("date", width=80, anchor=tk.CENTER)
        self.tree.column("doc_num", width=110, anchor=tk.CENTER)
        self.tree.column("client", width=170, anchor=tk.W)
        self.tree.column("item", width=200, anchor=tk.W)
        self.tree.column("manager", width=140, anchor=tk.W)
        self.tree.column("qty", width=55, anchor=tk.E)
        self.tree.column("amount", width=105, anchor=tk.E)
        self.tree.column("status", width=95, anchor=tk.CENTER)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        tree_scroll_y.grid(row=0, column=1, sticky=tk.NS)
        tree_scroll_x.grid(row=1, column=0, sticky=tk.EW)

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

    def _set_default_values(self) -> None:
        """Устанавливает исходные значения полей."""
        saved_sid = self.config_data.get("spreadsheet_id", "")
        self.var_spreadsheet_id.set(saved_sid)

        saved_dept = self.config_data.get("department", "Отдел корпоративных продаж")
        if saved_dept in DEPARTMENT_DATA:
            self.var_department.set(saved_dept)
        else:
            self.var_department.set(list(DEPARTMENT_DATA.keys())[0])

        saved_author = self.config_data.get("author", "Иванов Алексей Сергеевич")
        self.var_author.set(saved_author)

        self._preset_current_month()
        self._generate_sheet_name()
        self._generate_and_show_preview()

    def _generate_sheet_name(self) -> None:
        now = datetime.datetime.now()
        name = f"Отчет_{now.strftime('%d.%m_%H%M')}"
        self.var_sheet_title.set(name)

    def _copy_email(self) -> None:
        if self.service_email and "@" in self.service_email:
            self.clipboard_clear()
            self.clipboard_append(self.service_email)
            messagebox.showinfo("Скопировано", f"Email скопирован в буфер обмена:\n\n{self.service_email}")
        else:
            messagebox.showwarning("Внимание", "Email сервисного аккаунта недоступен.")

    def _on_scale_change(self, val: str) -> None:
        cnt = int(float(val))
        self.lbl_rows_display.config(text=f"{cnt} операций")

    def _preset_current_month(self) -> None:
        today = datetime.date.today()
        first_day = today.replace(day=1)
        self.var_date_from.set(first_day.strftime("%d.%m.%Y"))
        self.var_date_to.set(today.strftime("%d.%m.%Y"))

    def _preset_last_month(self) -> None:
        today = datetime.date.today()
        first_of_this_month = today.replace(day=1)
        last_day_prev_month = first_of_this_month - datetime.timedelta(days=1)
        first_day_prev_month = last_day_prev_month.replace(day=1)
        self.var_date_from.set(first_day_prev_month.strftime("%d.%m.%Y"))
        self.var_date_to.set(last_day_prev_month.strftime("%d.%m.%Y"))

    def _preset_quarter(self) -> None:
        today = datetime.date.today()
        quarter = (today.month - 1) // 3 + 1
        first_month = (quarter - 1) * 3 + 1
        first_day = datetime.date(today.year, first_month, 1)
        self.var_date_from.set(first_day.strftime("%d.%m.%Y"))
        self.var_date_to.set(today.strftime("%d.%m.%Y"))

    def _preset_random_period(self) -> None:
        today = datetime.date.today()
        start_offset = random.randint(30, 180)
        duration = random.randint(15, 45)
        d_from = today - datetime.timedelta(days=start_offset)
        d_to = d_from + datetime.timedelta(days=duration)
        self.var_date_from.set(d_from.strftime("%d.%m.%Y"))
        self.var_date_to.set(d_to.strftime("%d.%m.%Y"))

    def _randomize_all(self) -> None:
        """Случайно генерирует все параметры."""
        self._preset_random_period()
        dept = random.choice(list(DEPARTMENT_DATA.keys()))
        self.var_department.set(dept)
        mgr = random.choice(DEPARTMENT_DATA[dept]["manager_pool"])
        self.var_author.set(mgr)
        count = random.randint(12, 28)
        self.var_rows_count.set(count)
        self.lbl_rows_display.config(text=f"{count} операций")
        self._generate_sheet_name()
        self._generate_and_show_preview()

    def _parse_dates(self) -> Tuple[datetime.date, datetime.date]:
        from_str = self.var_date_from.get().strip()
        to_str = self.var_date_to.get().strip()
        try:
            d_from = datetime.datetime.strptime(from_str, "%d.%m.%Y").date()
        except ValueError:
            raise ValueError(f"Неверный формат даты начала: '{from_str}'. Используйте ДД.ММ.ГГГГ")

        try:
            d_to = datetime.datetime.strptime(to_str, "%d.%m.%Y").date()
        except ValueError:
            raise ValueError(f"Неверный формат даты окончания: '{to_str}'. Используйте ДД.ММ.ГГГГ")

        return d_from, d_to

    def _generate_and_show_preview(self) -> Optional[dict]:
        """Генерирует данные и заполняет дерево предпросмотра."""
        try:
            d_from, d_to = self._parse_dates()
        except Exception as e:
            messagebox.showwarning("Ошибка ввода даты", str(e))
            return None

        data = generate_random_report_data(
            date_from=d_from,
            date_to=d_to,
            department=self.var_department.get(),
            author=self.var_author.get().strip(),
            row_count=self.var_rows_count.get(),
        )

        # Очистка дерева
        for row in self.tree.get_children():
            self.tree.delete(row)

        for r in data["rows"]:
            formatted_amount = f"{r['amount']:,.2f} ₽".replace(",", " ")
            self.tree.insert(
                "",
                tk.END,
                values=(
                    r["idx"],
                    r["date"],
                    r["doc_num"],
                    r["client"],
                    r["item"],
                    r["manager"],
                    r["qty"],
                    formatted_amount,
                    r["status"],
                ),
            )

        self.var_status.set(f"Предпросмотр готов: {len(data['rows'])} операций. Общая сумма: {data['total_sum']:,.2f} ₽")
        return data

    def _start_generation_thread(self) -> None:
        """Запускает процесс отправки в Google Sheets в фоновом потоке."""
        if self.is_processing:
            return

        sid_raw = self.var_spreadsheet_id.get().strip()
        if not sid_raw:
            messagebox.showerror(
                "ID таблицы не указан",
                "Пожалуйста, введите ID вашей Google Таблицы или вставьте полную ссылку на нее.\n\n"
                "Пример ссылки: https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit",
            )
            self.ent_spreadsheet.focus_set()
            return

        spreadsheet_id = extract_spreadsheet_id(sid_raw)
        sheet_title = self.var_sheet_title.get().strip()
        if not sheet_title:
            self._generate_sheet_name()
            sheet_title = self.var_sheet_title.get().strip()

        try:
            d_from, d_to = self._parse_dates()
        except Exception as e:
            messagebox.showerror("Ошибка даты", str(e))
            return

        # Сохранение конфига
        self._save_config()

        self.is_processing = True
        self.btn_submit.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0

        # Запуск в отдельном потоке
        threading.Thread(
            target=self._run_report_generation,
            args=(spreadsheet_id, sheet_title, d_from, d_to),
            daemon=True,
        ).start()

    def _run_report_generation(
        self,
        spreadsheet_id: str,
        sheet_title: str,
        d_from: datetime.date,
        d_to: datetime.date,
    ) -> None:
        try:
            def update_progress(msg: str, step: int):
                self.after(0, lambda: self._update_ui_progress(msg, step))

            update_progress("Проверка подключения к Google Sheets API...", 5)

            if not self.client:
                self.client = GoogleSheetsClient()

            # Генерируем свежие данные
            data = generate_random_report_data(
                date_from=d_from,
                date_to=d_to,
                department=self.var_department.get(),
                author=self.var_author.get().strip(),
                row_count=self.var_rows_count.get(),
            )

            # Создаем и оформляем Word-отчет в таблице
            sheet_id, sheet_url = build_sheet_report(
                client=self.client,
                spreadsheet_id=spreadsheet_id,
                sheet_title=sheet_title,
                data=data,
                progress_callback=update_progress,
            )

            self.last_sheet_url = sheet_url
            self.after(0, lambda: self._on_generation_success(sheet_title, sheet_url, data))

        except Exception as err:
            err_msg = str(err)
            self.after(0, lambda: self._on_generation_error(err_msg))

    def _update_ui_progress(self, msg: str, step: int) -> None:
        self.var_status.set(msg)
        self.progress_bar["value"] = step

    def _on_generation_success(self, sheet_title: str, sheet_url: str, data: dict) -> None:
        self.is_processing = False
        self.btn_submit.config(state=tk.NORMAL)
        self.progress_bar["value"] = 100
        self.var_status.set(f"✅ Успешно! Отчет «{sheet_title}» готов в Google Sheets.")

        answer = messagebox.askyesno(
            "Отчет успешно создан!",
            f"Отчет «{sheet_title}» успешно загружен и стилизован под Word-документ!\n\n"
            f"• Подразделение: {data['department']}\n"
            f"• Период: с {data['date_from']} по {data['date_to']}\n"
            f"• Всего операций: {data['row_count']}\n"
            f"• Итоговая сумма: {data['total_sum']:,.2f} ₽\n\n"
            f"Открыть Google Таблицу в браузере прямо сейчас?",
        )
        if answer:
            self._open_in_browser()

    def _on_generation_error(self, err_msg: str) -> None:
        self.is_processing = False
        self.btn_submit.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.var_status.set("❌ Произошла ошибка при формировании отчета.")

        advice = ""
        if "403" in err_msg or "The caller does not have permission" in err_msg:
            advice = (
                "\n\nВозможная причина: нет прав доступа к таблице.\n"
                f"Убедитесь, что вы предоставили доступ с ролью 'Редактор' для сервисного аккаунта:\n{self.service_email}"
            )
        elif "404" in err_msg or "not found" in err_msg.lower():
            advice = "\n\nВозможная причина: таблица с таким ID не существует или ссылка неверна."

        messagebox.showerror(
            "Ошибка создания отчета",
            f"Не удалось записать отчет в Google Sheets:\n{err_msg}{advice}",
        )

    def _open_in_browser(self) -> None:
        if self.last_sheet_url:
            webbrowser.open(self.last_sheet_url)
        else:
            sid = self.var_spreadsheet_id.get().strip()
            if sid:
                real_id = extract_spreadsheet_id(sid)
                url = f"https://docs.google.com/spreadsheets/d/{real_id}/edit"
                webbrowser.open(url)
            else:
                messagebox.showinfo("Информация", "Сначала укажите ID таблицы или сгенерируйте отчет.")


def main() -> None:
    app = ModernReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()

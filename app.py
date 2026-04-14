import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import time
import json

from core.config import SettingsManager
from gui.controller import Controller
from gui.table_editor_widget import TableEditorWidget
from core.table_template import TableTemplate


class TkLogHandler(logging.Handler):
    def __init__(self, text_widget: tk.Text):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.after(0, self._append, msg)

    def _append(self, msg: str):
        self.text_widget.insert(tk.END, msg + "\n")
        self.text_widget.see(tk.END)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AhoTender - Обработка документов")

        # Установка иконки для окна и панели задач
        self._set_app_icon()

        # Для защиты от двойных вызовов горячих клавиш
        self._last_command_time = 0
        self._last_command = None

        # Определяем пути
        if getattr(sys, 'frozen', False):
            # Если приложение собрано
            self.base_dir = os.path.dirname(sys.executable)
        else:
            # Если запускается из исходников
            self.base_dir = os.path.dirname(__file__)

        # Создаем пути к директориям
        self.data_dir = os.path.join(self.base_dir, "data")
        self.tenders_dir = os.path.join(self.data_dir, "tenders")
        self.organizations_dir = os.path.join(self.data_dir, "organizations")
        self.templates_dir = os.path.join(self.data_dir, "templates")

        # Создаем папки, если их нет
        self._create_required_folders()

        # Минимальный и стартовый размер окна
        self.geometry("1600x950")
        self.minsize(1600, 950)
        self.resizable(True, True)

        # Определяем пути для работы в собранном приложении
        self._init_paths()

        self.settings_manager = SettingsManager("gui/settings.json")
        self.settings_manager.load()

        self.controller = Controller(data_dir=self.data_dir, settings_manager=self.settings_manager)

        # отображаемые имена компаний (можно дополнять)
        self.company_display_names = {
            "Rosneft": "Роснефть",
            "Gazprom": "Газпром",
            "Lukoil": "Лукойл",
            "Sber": "Сбербанк",
            "Starlit": "Старит",
        }

        self.current_replacements_json = None

        self._build_ui()
        self._init_logging()
        self._load_initial_settings()
        self._bind_hotkeys()

    # ---------------------------------------------------------
    # Установка иконки приложения
    # ---------------------------------------------------------
    def _set_app_icon(self):
        """Устанавливает иконку для окна и панели задач (Windows и Mac)"""

        if sys.platform == "win32":
            # Для Windows используем .ico
            icon_paths = [
                "assets/icon.ico",
                "icon.ico",
            ]

            for path in icon_paths:
                if os.path.exists(path):
                    try:
                        self.iconbitmap(default=path)
                        print(f"Иконка загружена (Windows): {path}")
                        break
                    except Exception as e:
                        print(f"Не удалось загрузить иконку {path}: {e}")

        elif sys.platform == "darwin":  # macOS
            icon_paths = [
                "assets/icon.icns",
                "assets/icon.png",
                "icon.icns",
                "icon.png",
            ]

            for path in icon_paths:
                if os.path.exists(path):
                    try:
                        if path.endswith('.icns'):
                            self.iconbitmap(default=path)
                            print(f"Иконка загружена (Mac): {path}")
                            break
                        else:
                            icon = tk.PhotoImage(file=path)
                            self.iconphoto(True, icon)
                            print(f"Иконка загружена (Mac PNG): {path}")
                            break
                    except Exception as e:
                        print(f"Не удалось загрузить иконку {path}: {e}")

        else:  # Linux и другие
            icon_paths = [
                "assets/icon.png",
                "icon.png",
            ]

            for path in icon_paths:
                if os.path.exists(path):
                    try:
                        icon = tk.PhotoImage(file=path)
                        self.iconphoto(True, icon)
                        print(f"Иконка загружена (Linux): {path}")
                        break
                    except Exception as e:
                        print(f"Не удалось загрузить иконку {path}: {e}")

    # ---------------------------------------------------------
    # Инициализация путей (работает в собранном приложении)
    # ---------------------------------------------------------
    def _create_required_folders(self):
        """Создает все необходимые папки, если они не существуют"""
        folders = [
            self.data_dir,
            self.tenders_dir,
            self.organizations_dir,
            self.templates_dir
        ]

        for folder in folders:
            try:
                os.makedirs(folder, exist_ok=True)
                print(f"Создана папка: {folder}")
            except Exception as e:
                print(f"Ошибка при создании папки {folder}: {e}")

    def _init_paths(self):
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(__file__)

        self.data_dir = os.path.join(self.base_dir, "data")
        self.tenders_dir = os.path.join(self.data_dir, "tenders")
        self.organizations_dir = os.path.join(self.data_dir, "organizations")
        self.templates_dir = os.path.join(self.data_dir, "templates")

        self._ensure_directories()

    def _ensure_directories(self):
        dirs = [
            self.data_dir,
            self.tenders_dir,
            self.organizations_dir,
            self.templates_dir
        ]

        for dir_path in dirs:
            try:
                os.makedirs(dir_path, exist_ok=True)
            except Exception as e:
                logging.error(f"Не удалось создать директорию {dir_path}: {e}")
                if dir_path == self.data_dir:
                    self._create_fallback_directories()
                    break

    def _create_fallback_directories(self):
        home_dir = os.path.expanduser("~")
        self.data_dir = os.path.join(home_dir, ".docx_processor", "data")
        self.tenders_dir = os.path.join(self.data_dir, "tenders")
        self.organizations_dir = os.path.join(self.data_dir, "organizations")
        self.templates_dir = os.path.join(self.data_dir, "templates")

        for dir_path in [self.data_dir, self.tenders_dir, self.organizations_dir, self.templates_dir]:
            os.makedirs(dir_path, exist_ok=True)

        logging.info(f"Используются резервные директории в: {self.data_dir}")

    # ---------------------------------------------------------
    # Горячие клавиши
    # ---------------------------------------------------------
    def _bind_hotkeys(self):
        for seq in ('<Control-c>', '<Control-v>', '<Control-x>',
                    '<Control-a>', '<Control-z>', '<Control-y>'):
            self.bind_class('Text', seq, lambda e: 'break')
            self.bind_class('Entry', seq, lambda e: 'break')
            self.bind_class('Spinbox', seq, lambda e: 'break')

        self.bind_all("<Control-KeyPress>", self._on_ctrl_keycode)
        self.bind_all("<Command-KeyPress>", self._on_ctrl_keycode)

    def _on_ctrl_keycode(self, event):
        current_time = time.time()

        if current_time - self._last_command_time < 0.2:
            if self._last_command == event.keycode:
                return "break"

        widget = self.focus_get()
        if widget is None:
            return "break"

        if not isinstance(widget, (tk.Text, tk.Entry, tk.Spinbox, ttk.Combobox, tk.Listbox)):
            return

        keycode_commands = {
            67: 'copy',
            86: 'paste',
            88: 'cut',
            65: 'select_all',
            90: 'undo',
            89: 'redo',
        }

        if event.keycode in keycode_commands:
            cmd = keycode_commands[event.keycode]

            self._last_command_time = current_time
            self._last_command = event.keycode

            if cmd == 'copy':
                widget.event_generate("<<Copy>>")
            elif cmd == 'paste':
                self.after(10, lambda w=widget: w.event_generate("<<Paste>>"))
            elif cmd == 'cut':
                widget.event_generate("<<Cut>>")
            elif cmd == 'select_all':
                self._select_all(widget)
            elif cmd == 'undo':
                self._undo(widget)
            elif cmd == 'redo':
                self._redo(widget)

            return "break"

        return

    def _select_all(self, widget):
        try:
            if isinstance(widget, tk.Text):
                widget.tag_add(tk.SEL, "1.0", tk.END)
                widget.mark_set(tk.INSERT, "1.0")
                widget.see(tk.INSERT)
            elif isinstance(widget, tk.Entry):
                widget.select_range(0, tk.END)
                widget.icursor(tk.END)
            elif isinstance(widget, ttk.Combobox):
                if widget.winfo_children():
                    entry = widget.winfo_children()[0]
                    if isinstance(entry, tk.Entry):
                        entry.select_range(0, tk.END)
            elif isinstance(widget, tk.Listbox):
                widget.selection_set(0, tk.END)
            elif isinstance(widget, tk.Spinbox):
                widget.select_range(0, tk.END)
        except Exception as e:
            logging.debug(f"Ошибка при выделении: {e}")

    def _undo(self, widget):
        try:
            if hasattr(widget, 'edit_undo'):
                widget.edit_undo()
            else:
                widget.event_generate("<<Undo>>")
        except:
            pass

    def _redo(self, widget):
        try:
            if hasattr(widget, 'edit_redo'):
                widget.edit_redo()
            else:
                widget.event_generate("<<Redo>>")
        except:
            pass

    # ---------------------------------------------------------
    # Вспомогательное: скроллируемый фрейм
    # ---------------------------------------------------------
    def _make_scrollable(self, parent):
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)

        v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        v_scrollbar.pack(side="right", fill="y")

        h_scrollbar = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        h_scrollbar.pack(side="bottom", fill="x")

        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        inner_frame = ttk.Frame(canvas)
        canvas_frame = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        def configure_inner(event):
            bbox = canvas.bbox("all")
            if bbox:
                canvas.configure(scrollregion=bbox)
            canvas.itemconfig(canvas_frame, width=canvas.winfo_width())

        def configure_canvas(event):
            canvas.itemconfig(canvas_frame, width=canvas.winfo_width())
            bbox = canvas.bbox("all")
            if bbox:
                canvas.configure(scrollregion=bbox)

        inner_frame.bind("<Configure>", configure_inner)
        canvas.bind("<Configure>", configure_canvas)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def on_shift_mousewheel(event):
            canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Shift-MouseWheel>", on_shift_mousewheel)
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        return inner_frame

    # ---------------------------------------------------------
    # UI
    # ---------------------------------------------------------
    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_process = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_process, text="Обработка документов")
        self.process_frame = self._make_scrollable(self.tab_process)

        self.tab_fill = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_fill, text="Заполнение шаблонов")
        self.fill_frame = self._make_scrollable(self.tab_fill)

        self.tab_settings = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_settings, text="Настройки")
        self.settings_frame = self._make_scrollable(self.tab_settings)

        self.tab_log = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_log, text="Журнал")

        self._build_process_tab()
        self._build_fill_tab()
        self._build_settings_tab()
        self._build_log_tab()

    # ---------------------------------------------------------
    # Вкладка "Обработка"
    # ---------------------------------------------------------
    def _build_process_tab(self):
        frame = self.process_frame
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        # --- Разделение документа ---
        splitter_frame = ttk.LabelFrame(frame, text="Разделение документа")
        splitter_frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        splitter_frame.columnconfigure(1, weight=1)
        splitter_frame.rowconfigure(5, weight=1)

        ttk.Label(splitter_frame, text="Файл для разделения:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.split_input_entry = ttk.Entry(splitter_frame)
        self.split_input_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(splitter_frame, text="...", width=3,
                   command=lambda: self._choose_file_into(self.split_input_entry)).grid(
            row=0, column=2, padx=4, pady=4
        )

        ttk.Label(splitter_frame, text="Папка для сохранения:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.split_output_entry = ttk.Entry(splitter_frame)
        self.split_output_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(splitter_frame, text="...", width=3,
                   command=lambda: self._choose_folder_into(self.split_output_entry)).grid(
            row=1, column=2, padx=4, pady=4
        )

        ttk.Label(splitter_frame, text="Метод разделения:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.split_method_combo = ttk.Combobox(
            splitter_frame, values=["По умолчанию", "Роснефть"], state="readonly"
        )
        self.split_method_combo.grid(row=2, column=1, sticky="ew", padx=4, pady=4)
        self.split_method_combo.current(0)

        self.split_clean_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(splitter_frame, text="Очищать скрытые символы",
                        variable=self.split_clean_var).grid(
            row=3, column=1, sticky="w", padx=4, pady=4
        )

        self.split_button = ttk.Button(splitter_frame, text="Разделить документ", command=self._on_split_docx)
        self.split_button.grid(row=4, column=1, pady=6)

        self.split_result_box = tk.Text(splitter_frame, height=6)
        self.split_result_box.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=4, pady=4)

        # --- Контейнер для замены текста и редактора таблиц ---
        self.process_pane = ttk.PanedWindow(frame, orient=tk.HORIZONTAL)
        self.process_pane.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        frame.rowconfigure(1, weight=1)

        self.process_left = ttk.Frame(self.process_pane)
        self.process_right = ttk.Frame(self.process_pane)

        self.process_pane.add(self.process_left, weight=1)
        self.process_pane.add(self.process_right, weight=1)

        # --- Замена текста ---
        replace_frame = ttk.LabelFrame(self.process_left, text="Замена текста")
        replace_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        replace_frame.columnconfigure(1, weight=1)

        ttk.Label(replace_frame, text="Организация:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.company_combo = ttk.Combobox(replace_frame, state="readonly")
        self.company_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        self.company_combo.bind("<<ComboboxSelected>>", self._on_company_selected)
        ttk.Button(replace_frame, text="Обновить", command=self._refresh_company_list).grid(
            row=0, column=2, padx=4, pady=4
        )

        ttk.Label(replace_frame, text="Шаблон документа:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.input_entry = ttk.Entry(replace_frame)
        self.input_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(replace_frame, text="...", width=3, command=self._choose_input).grid(
            row=1, column=2, padx=4, pady=4
        )

        self.progress = ttk.Progressbar(replace_frame, mode="indeterminate")
        self.progress.grid(row=2, column=0, columnspan=3, sticky="ew", padx=4, pady=6)

        self.btn_process_docx = ttk.Button(replace_frame, text="Обработать документ", command=self._on_process_docx)
        self.btn_process_docx.grid(row=3, column=1, pady=4)

        self._refresh_company_list()

        # --- Редактор таблиц ---
        editor_frame = ttk.LabelFrame(self.process_right, text="Работа с таблицами")
        editor_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        table_top = ttk.Frame(editor_frame)
        table_top.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(table_top, text="Открыть документ", command=self._open_docx_for_tables).pack(side=tk.LEFT, padx=3)

        tmpl_box_frame = ttk.LabelFrame(editor_frame, text="Шаблоны таблиц (data/templates)")
        tmpl_box_frame.pack(fill=tk.X, padx=4, pady=4)

        self.template_listbox = tk.Listbox(tmpl_box_frame, height=5)
        self.template_listbox.pack(fill=tk.X, padx=4, pady=4)

        btn_tmpl_list = ttk.Frame(tmpl_box_frame)
        btn_tmpl_list.pack(fill=tk.X, padx=4, pady=(0, 4))

        ttk.Button(btn_tmpl_list, text="Обновить список", command=self._refresh_template_list).pack(
            side=tk.LEFT, padx=3
        )
        ttk.Button(btn_tmpl_list, text="Применить шаблон", command=self._apply_selected_template).pack(
            side=tk.LEFT, padx=3
        )

        self.table_editor = TableEditorWidget(editor_frame)
        self.table_editor.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self._refresh_template_list()

    # ---------------------------------------------------------
    # Шаблоны таблиц
    # ---------------------------------------------------------
    def _open_docx_for_tables(self):
        path = filedialog.askopenfilename(filetypes=[("DOCX файлы", "*.docx")])
        if not path:
            return
        try:
            from docx import Document
            doc = Document(path)
            self.table_editor.load_document(doc, file_path=path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить документ:\n{e}")

    def _refresh_template_list(self):
        try:
            os.makedirs(self.templates_dir, exist_ok=True)
            self.template_listbox.delete(0, tk.END)
            for fname in os.listdir(self.templates_dir):
                if fname.lower().endswith(".json"):
                    self.template_listbox.insert(tk.END, fname)
        except Exception as e:
            logging.error(f"Ошибка при обновлении списка шаблонов: {e}")

    def _apply_selected_template(self):
        selection = self.template_listbox.curselection()
        if not selection:
            messagebox.showwarning("Нет выбора", "Выберите шаблон из списка.")
            return

        fname = self.template_listbox.get(selection[0])
        tmpl_path = os.path.join(self.templates_dir, fname)

        try:
            tmpl = TableTemplate.load_template(tmpl_path)
            self.table_editor.template = tmpl
            self.table_editor.fill_table()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось применить шаблон:\n{e}")

    # ---------------------------------------------------------
    # Вкладка "Заполнение" (ОБНОВЛЕННАЯ - без статусов)
    # ---------------------------------------------------------
    def _build_fill_tab(self):
        frame = self.fill_frame
        frame.columnconfigure(0, weight=1)

        fill_frame = ttk.LabelFrame(frame, text="Заполнение документа")
        fill_frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        fill_frame.columnconfigure(1, weight=1)

        # Организация
        ttk.Label(fill_frame, text="Организация:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.fill_company_combo = ttk.Combobox(fill_frame, state="readonly")
        self.fill_company_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        self.fill_company_combo.bind("<<ComboboxSelected>>", self._on_fill_company_selected)
        ttk.Button(fill_frame, text="Обновить", command=self._refresh_fill_company_list).grid(
            row=0, column=2, padx=4, pady=4
        )

        # Файл замен (JSON)
        ttk.Label(fill_frame, text="Файл замен (JSON):").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.fill_json_combo = ttk.Combobox(fill_frame, state="readonly")
        self.fill_json_combo.grid(row=1, column=1, sticky="ew", padx=4, pady=4)

        # Шаблон документа
        ttk.Label(fill_frame, text="Шаблон документа:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.fill_input_entry = ttk.Entry(fill_frame)
        self.fill_input_entry.grid(row=2, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(fill_frame, text="...", width=3,
                   command=lambda: self._choose_file_into(self.fill_input_entry)).grid(
            row=2, column=2, padx=4, pady=4
        )

        # Выходной PDF
        ttk.Label(fill_frame, text="Выходной PDF:").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        self.fill_output_pdf_entry = ttk.Entry(fill_frame)
        self.fill_output_pdf_entry.grid(row=3, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(fill_frame, text="...", width=3, command=self._choose_fill_output_pdf).grid(
            row=3, column=2, padx=4, pady=4
        )

        # Поля для заполнения
        fields_frame = ttk.LabelFrame(fill_frame, text="Поля для заполнения")
        fields_frame.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=4, pady=6)
        fields_frame.columnconfigure(1, weight=1)

        self.fill_fields = {}

        fields_spec = [
            ("Тендер", "{tender}", "entry"),
            ("Ссылка на закупку", "{link}", "entry"),
            ("Категория участника", "{participant_category}", "combo_category"),
            ("Организатор закупки", "{organizer_name}", "entry"),
            ("Док-ты должной осмотрительности", "{verification_docs_ref}", "entry"),
            ("Последняя отчетная дата", "{reporting_period}", "entry"),
            ("Заказчик", "{zakaz}", "entry"),
            ("Адрес заказчика", "{zakaz_adr}", "entry"),
            ("Аккредитация", "{accreditation}", "entry"),
            ("Способ закупки", "{purchase_method}", "entry"),
            ("Предмет закупки", "{purchase_item}", "entry"),
            ("Перечень расходов", "{expenses}", "entry"),
            ("Организация договора", "{buyer_name}", "entry"),
            ("Договор", "{contract_subject}", "entry"),
            ("Срок действия заявки", "{offer_validity}", "entry"),
            ("Организатор", "{organizer}", "entry"),
            ("Итоговая цена", "{final_price}", "entry"),
            ("Цена без НДС", "{price_without_vat}", "entry"),
            ("Срок выполнения", "{deadline}", "entry"),
            ("Форма оплаты", "{payment_form}", "entry"),
            ("Сроки оплаты", "{payment_terms}", "entry"),
            ("Порядок оплаты", "{payment_order}", "entry"),
            ("Валюта", "{currency}", "combo_currency"),
            ("Дата оферты", "{offerta}", "entry")
        ]

        category_values = [
            "Производитель МТР",
            "Посредник",
            "Дилер",
            "Дистрибьютор",
            "Сбытовая организация производителя (Торговый дом)",
            "Исполнитель услуг (собственными силами)",
            "Исполнитель услуг (с привлечением субисполнителей)",
            "Подрядчик (собственными силами)",
            "Генеральный подрядчик",
            "Пэкиджер",
            "Прочие Поставщики",
            "Производитель импортозамещающей продукции",
            "Дистрибьютор импортозамещающей продукции",
            "Сервисная компания, сопровождающая импортозамещающую продукцию",
            "Компания-инвестор, финансирующая разработку импортозамещающей продукции",
        ]
        currency_values = ["RUB", "USD", "EUR", "CNY"]

        row = 0
        for label, tag, kind in fields_spec:
            ttk.Label(fields_frame, text=f"{label}:").grid(row=row, column=0, sticky="w", padx=4, pady=3)
            var = tk.StringVar()
            if kind == "entry":
                ttk.Entry(fields_frame, textvariable=var).grid(row=row, column=1, sticky="ew", padx=4, pady=3)
            elif kind == "combo_category":
                cb = ttk.Combobox(fields_frame, textvariable=var, state="readonly", values=category_values)
                cb.grid(row=row, column=1, sticky="ew", padx=4, pady=3)
            elif kind == "combo_currency":
                cb = ttk.Combobox(fields_frame, textvariable=var, state="readonly", values=currency_values)
                cb.grid(row=row, column=1, sticky="ew", padx=4, pady=3)
            self.fill_fields[label] = (tag, var)
            row += 1

        self.fill_progress = ttk.Progressbar(fill_frame, mode="indeterminate")
        self.fill_progress.grid(row=5, column=0, columnspan=3, sticky="ew", padx=4, pady=6)

        # Кнопки
        btn_frame = ttk.Frame(fill_frame)
        btn_frame.grid(row=6, column=0, columnspan=3, pady=6)

        self.btn_fill_docx = ttk.Button(btn_frame, text="Заполнить документ", command=self._on_fill_docx)
        self.btn_fill_docx.grid(row=0, column=0, padx=4)

        self.btn_fill_convert_pdf = ttk.Button(
            btn_frame, text="Конвертировать в PDF", command=self._on_fill_convert_pdf
        )
        self.btn_fill_convert_pdf.grid(row=0, column=1, padx=4)

        self._refresh_fill_company_list()

    # ---------------------------------------------------------
    # Вкладка "Настройки"
    # ---------------------------------------------------------
    def _build_settings_tab(self):
        frame = self.settings_frame
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Файл подписи:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.signature_entry = ttk.Entry(frame)
        self.signature_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(frame, text="...", width=3, command=self._choose_signature).grid(
            row=0, column=2, padx=4, pady=4
        )

        ttk.Label(frame, text="Файл печати:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.stamp_entry = ttk.Entry(frame)
        self.stamp_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(frame, text="...", width=3, command=self._choose_stamp).grid(
            row=1, column=2, padx=4, pady=4
        )

        ttk.Label(frame, text="ФИО и должность:").grid(row=2, column=0, sticky="nw", padx=6, pady=4)
        self.fullname_text = tk.Text(frame, height=4)
        self.fullname_text.grid(row=2, column=1, sticky="ew", padx=4, pady=4)

        ttk.Button(frame, text="Сохранить настройки", command=self._save_settings).grid(
            row=3, column=0, columnspan=3, pady=8
        )

    # ---------------------------------------------------------
    # Вкладка "Журнал"
    # ---------------------------------------------------------
    def _build_log_tab(self):
        frame = self.tab_log
        self.log_text = tk.Text(frame, state="normal")
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ---------------------------------------------------------
    # Логирование
    # ---------------------------------------------------------
    def _init_logging(self):
        handler = TkLogHandler(self.log_text)
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        handler.setFormatter(formatter)

        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(handler)

    # ---------------------------------------------------------
    # Начальные настройки
    # ---------------------------------------------------------
    def _load_initial_settings(self):
        try:
            sig = self.controller.get_signature_config()
            if sig.signature_path:
                self.signature_entry.insert(0, sig.signature_path)
            if sig.stamp_path:
                self.stamp_entry.insert(0, sig.stamp_path)
            self.fullname_text.insert("1.0", sig.fullname_text)

            last_in, last_out_docx, last_out_pdf = self.controller.get_last_paths()
            if last_in:
                self.input_entry.insert(0, last_in)
            if last_out_pdf:
                self.fill_output_pdf_entry.insert(0, last_out_pdf)
        except Exception as e:
            logging.error(f"Ошибка при загрузке начальных настроек: {e}")

    # ---------------------------------------------------------
    # Потоки
    # ---------------------------------------------------------
    def _run_in_thread(self, job, on_done=None, progressbar=None, buttons_to_disable=None):
        if buttons_to_disable is None:
            buttons_to_disable = []

        def worker():
            error = None
            try:
                job()
            except Exception as e:
                logging.exception("Ошибка в задаче")
                error = e

            def finish():
                if progressbar is not None:
                    progressbar.stop()
                for btn in buttons_to_disable:
                    btn.config(state=tk.NORMAL)
                if error:
                    messagebox.showerror("Ошибка", str(error))
                elif on_done:
                    on_done()

            self.after(0, finish)

        for btn in buttons_to_disable:
            btn.config(state=tk.DISABLED)

        if progressbar is not None:
            progressbar.start(10)

        threading.Thread(target=worker, daemon=True).start()

    # ---------------------------------------------------------
    # Выбор файлов / папок
    # ---------------------------------------------------------
    def _choose_file_into(self, entry: ttk.Entry):
        path = filedialog.askopenfilename(filetypes=[("DOCX файлы", "*.docx")])
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def _choose_folder_into(self, entry: ttk.Entry):
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def _choose_input(self):
        path = filedialog.askopenfilename(filetypes=[("DOCX файлы", "*.docx")])
        if path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, path)
            self.controller.set_last_paths(input_docx=path)

            try:
                from docx import Document
                doc = Document(path)
                self.table_editor.load_document(doc, file_path=path)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить таблицы:\n{e}")

    def _choose_fill_output_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if path:
            self.fill_output_pdf_entry.delete(0, tk.END)
            self.fill_output_pdf_entry.insert(0, path)
            self.controller.set_last_paths(output_pdf=path)

    def _choose_signature(self):
        path = filedialog.askopenfilename(filetypes=[("Изображения", "*.png;*.jpg;*.jpeg;*.bmp")])
        if path:
            self.signature_entry.delete(0, tk.END)
            self.signature_entry.insert(0, path)

    def _choose_stamp(self):
        path = filedialog.askopenfilename(filetypes=[("Изображения", "*.png;*.jpg;*.jpeg;*.bmp")])
        if path:
            self.stamp_entry.delete(0, tk.END)
            self.stamp_entry.insert(0, path)

    # ---------------------------------------------------------
    # Настройки
    # ---------------------------------------------------------
    def _save_settings(self):
        sig = self.signature_entry.get().strip() or None
        stamp = self.stamp_entry.get().strip() or None
        fullname = self.fullname_text.get("1.0", tk.END).strip()

        self.controller.set_signature_config(sig, stamp, fullname)
        messagebox.showinfo("Настройки", "Настройки сохранены")

    # ---------------------------------------------------------
    # Компании и JSON (Обработка, data/tenders)
    # ---------------------------------------------------------
    def _list_tender_companies_from_fs(self):
        try:
            os.makedirs(self.tenders_dir, exist_ok=True)
            companies = []
            for name in os.listdir(self.tenders_dir):
                full = os.path.join(self.tenders_dir, name)
                if os.path.isdir(full):
                    companies.append(name)
            companies.sort()
            return companies
        except Exception as e:
            logging.error(f"Ошибка при чтении списка компаний: {e}")
            return []

    def _refresh_company_list(self):
        try:
            companies = self._list_tender_companies_from_fs()
            display_list = [self.company_display_names.get(c, c) for c in companies]
            self.company_combo["values"] = display_list

            if display_list:
                self.company_combo.current(0)
                self._on_company_selected()
            else:
                self.company_combo.set("")
                self.current_replacements_json = None
        except Exception as e:
            logging.error(f"Ошибка при обновлении списка компаний: {e}")

    def _on_company_selected(self, event=None):
        display_name = self.company_combo.get().strip()
        if not display_name:
            self.current_replacements_json = None
            return

        company = None
        for folder, name in self.company_display_names.items():
            if name == display_name:
                company = folder
                break
        if company is None:
            company = display_name

        json_path = os.path.join(self.tenders_dir, company, "replacements.json")

        if not os.path.exists(json_path):
            messagebox.showwarning(
                "Файл замен не найден",
                f"Для организации '{display_name}' отсутствует файл:\n{json_path}"
            )
            self.current_replacements_json = None
            return

        self.current_replacements_json = json_path

    # ---------------------------------------------------------
    # Компании и JSON (Заполнение, data/organizations)
    # ---------------------------------------------------------
    def _list_organizations_from_fs(self):
        try:
            os.makedirs(self.organizations_dir, exist_ok=True)
            companies = []
            for name in os.listdir(self.organizations_dir):
                full = os.path.join(self.organizations_dir, name)
                if os.path.isdir(full):
                    companies.append(name)
            companies.sort()
            return companies
        except Exception as e:
            logging.error(f"Ошибка при чтении списка организаций: {e}")
            return []

    def _refresh_fill_company_list(self):
        try:
            companies = self._list_organizations_from_fs()
            display_list = [self.company_display_names.get(c, c) for c in companies]
            self.fill_company_combo["values"] = display_list

            if display_list:
                self.fill_company_combo.current(0)
                self._on_fill_company_selected()
            else:
                self.fill_json_combo["values"] = []
        except Exception as e:
            logging.error(f"Ошибка при обновлении списка организаций: {e}")

    def _on_fill_company_selected(self, event=None):
        display_name = self.fill_company_combo.get().strip()
        if not display_name:
            self.fill_json_combo["values"] = []
            return

        company = None
        for folder, name in self.company_display_names.items():
            if name == display_name:
                company = folder
                break
        if company is None:
            company = display_name

        org_dir = os.path.join(self.organizations_dir, company)
        if not os.path.isdir(org_dir):
            self.fill_json_combo["values"] = []
            self.fill_json_combo.set("")
            messagebox.showwarning(
                "Организация не найдена",
                f"Папка организации не найдена:\n{org_dir}"
            )
            return

        try:
            json_files = [f for f in os.listdir(org_dir) if f.lower().endswith(".json")]
            json_files.sort()

            if json_files:
                self.fill_json_combo["values"] = json_files
                self.fill_json_combo.current(0)
            else:
                self.fill_json_combo["values"] = []
                self.fill_json_combo.set("")
                messagebox.showwarning(
                    "Файлы замен не найдены",
                    f"В папке организации нет JSON-файлов:\n{org_dir}"
                )

        except Exception as e:
            logging.error(f"Ошибка при чтении JSON файлов: {e}")
            self.fill_json_combo["values"] = []
            self.fill_json_combo.set("")

    # ---------------------------------------------------------
    # Разделение документа
    # ---------------------------------------------------------
    def _on_split_docx(self):
        src = self.split_input_entry.get().strip()
        out = self.split_output_entry.get().strip()
        method = self.split_method_combo.get()
        if method == "По умолчанию":
            method = None

        if not src or not out:
            messagebox.showwarning("Ошибка", "Укажите исходный файл и папку вывода")
            return

        self.split_result_box.delete("1.0", tk.END)

        def progress_callback(current, total, message):
            value = current / total * 100 if total else 0
            self.progress["value"] = value
            self.progress.update()
            self.split_result_box.insert(tk.END, message + "\n")
            self.split_result_box.see(tk.END)

        def job():
            files = self.controller.split_docx(
                source_docx=src,
                output_dir=out,
                method=method,
                clean_hidden_chars=self.split_clean_var.get(),
                progress_callback=progress_callback,
            )
            self.split_result_box.insert(tk.END, "\nГотово! Созданные файлы:\n")
            for f in files:
                self.split_result_box.insert(tk.END, f + "\n")

        self._run_in_thread(
            job,
            on_done=None,
            progressbar=self.progress,
            buttons_to_disable=[self.split_button],
        )

    # ---------------------------------------------------------
    # Обработка DOCX (Обработка)
    # ---------------------------------------------------------
    def _on_process_docx(self):
        input_path = self.input_entry.get().strip()
        display_name = self.company_combo.get().strip()

        if not input_path or not display_name:
            messagebox.showwarning("Ошибка", "Укажите документ и организацию")
            return

        json_path = self.current_replacements_json
        if not json_path:
            messagebox.showwarning("Ошибка", "Файл замен не найден для выбранной организации.")
            return

        output_docx = self.controller.processor.make_output_path(input_path)

        def job():
            replacements = self.controller.load_replacements(json_path)
            self.controller.process_docx(input_path, output_docx, replacements, tab_number=1)
            self.controller.set_last_paths(input_docx=input_path, output_docx=output_docx)

        def on_done():
            messagebox.showinfo("Готово", f"Документ сохранён:\n{output_docx}")

        self._run_in_thread(
            job,
            on_done=on_done,
            progressbar=self.progress,
            buttons_to_disable=[self.btn_process_docx],
        )

    # ---------------------------------------------------------
    # Валидация для заполнения (data/organizations)
    # ---------------------------------------------------------
    def _validate_company_json_and_input(self, input_entry, company_combo, json_combo):
        input_path = input_entry.get().strip()
        display_name = company_combo.get().strip()
        json_file = json_combo.get().strip()

        if not input_path or not display_name or not json_file:
            messagebox.showwarning("Ошибка", "Укажите документ, организацию и JSON")
            return None, None

        company = None
        for folder, name in self.company_display_names.items():
            if name == display_name:
                company = folder
                break
        if company is None:
            company = display_name

        json_path = os.path.join(self.organizations_dir, company, json_file.replace("/", os.sep))
        return input_path, json_path

    # ---------------------------------------------------------
    # Заполнение DOCX (ОБНОВЛЕННОЕ - без статусов)
    # ---------------------------------------------------------
    def _on_fill_docx(self):
        input_path, json_path = self._validate_company_json_and_input(
            self.fill_input_entry, self.fill_company_combo, self.fill_json_combo
        )
        if not input_path:
            return

        output_docx = self.controller.processor.make_output_path(input_path)
        company_name = self.fill_company_combo.get().strip()

        # Получаем название компании для processor
        company = None
        for folder, name in self.company_display_names.items():
            if name == company_name:
                company = folder
                break
        if company is None:
            company = company_name

        def job():
            replacements = self.controller.load_replacements(json_path)
            for label, (tag, var) in self.fill_fields.items():
                value = var.get().strip()
                if value:
                    replacements[tag] = value

            # Загружаем параметры из image_params.json, если есть
            org_dir = os.path.join(self.organizations_dir, company)

            # Пути к изображениям подписи и печати
            signature_path = os.path.join(org_dir, "sign.png")
            stamp_path = os.path.join(org_dir, "stamp.png")

            # Параметры по умолчанию
            sig_width = 30
            sig_offset_x = 0
            sig_offset_y = -20
            stamp_width = 30
            stamp_offset_x = 0
            stamp_offset_y = -20

            # Загружаем параметры из файла, если он существует
            params_path = os.path.join(org_dir, "image_params.json")
            if os.path.exists(params_path):
                try:
                    with open(params_path, 'r', encoding='utf-8') as f:
                        params = json.load(f)

                    if 'signature' in params:
                        sig = params['signature']
                        sig_width = sig.get('width', 30)
                        sig_offset_x = sig.get('offset_x', 0)
                        sig_offset_y = sig.get('offset_y', -20)

                    if 'stamp' in params:
                        st = params['stamp']
                        stamp_width = st.get('width', 30)
                        stamp_offset_x = st.get('offset_x', 0)
                        stamp_offset_y = st.get('offset_y', -20)

                except Exception as e:
                    logging.error(f"Ошибка при загрузке параметров изображений: {e}")

            # Добавляем пути к изображениям в replacements, если файлы существуют
            if os.path.exists(signature_path):
                replacements["{sign}"] = signature_path
                # Добавляем параметры в специальный словарь processor-а
                self.controller.processor.set_signature_image_params(
                    path=signature_path,
                    width=sig_width,
                    offset_x=sig_offset_x,
                    offset_y=sig_offset_y
                )
                logging.info(f"Добавлена подпись в замены: {signature_path}")

            if os.path.exists(stamp_path):
                replacements["{stamp}"] = stamp_path
                self.controller.processor.set_stamp_image_params(
                    path=stamp_path,
                    width=stamp_width,
                    offset_x=stamp_offset_x,
                    offset_y=stamp_offset_y
                )
                logging.info(f"Добавлена печать в замены: {stamp_path}")

            # Обрабатываем документ (tab_number=2 для вставки изображений)
            self.controller.process_docx(input_path, output_docx, replacements, tab_number=2)
            self.controller.set_last_paths(input_docx=input_path, output_docx=output_docx)

        def on_done():
            messagebox.showinfo("Готово", f"Документ сохранён:\n{output_docx}")

        self._run_in_thread(
            job,
            on_done=on_done,
            progressbar=self.fill_progress,
            buttons_to_disable=[self.btn_fill_docx, self.btn_fill_convert_pdf],
        )

    # ---------------------------------------------------------
    # Конвертация в PDF
    # ---------------------------------------------------------
    def _on_fill_convert_pdf(self):
        input_path = self.fill_input_entry.get().strip()
        pdf_path = self.fill_output_pdf_entry.get().strip()

        if not input_path or not pdf_path:
            messagebox.showwarning("Ошибка", "Укажите DOCX и PDF")
            return

        # Используем выходной DOCX файл, если он существует
        updated_docx = self.controller.processor.make_output_path(input_path)

        def job():
            if os.path.exists(updated_docx):
                self.controller.convert_to_pdf(updated_docx, pdf_path)
                logging.info(f"Конвертация {updated_docx} в {pdf_path}")
            else:
                self.controller.convert_to_pdf(input_path, pdf_path)
                logging.info(f"Конвертация {input_path} в {pdf_path}")
            self.controller.set_last_paths(output_pdf=pdf_path)

        def on_done():
            messagebox.showinfo("Готово", f"PDF создан:\n{pdf_path}")

        self._run_in_thread(
            job,
            on_done=on_done,
            progressbar=self.fill_progress,
            buttons_to_disable=[self.btn_fill_convert_pdf, self.btn_fill_docx],
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
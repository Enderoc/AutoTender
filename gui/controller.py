import os
import logging

from core.processor import DocxProcessor
from core.config import SignatureConfig, SettingsManager
from core.splitter import DocxSplitter


class Controller:
    def __init__(self, data_dir: str, settings_manager: SettingsManager):
        self.data_dir = data_dir
        self.settings_manager = settings_manager

        # Настройки и подпись живут в SettingsManager
        self.signature_config: SignatureConfig = self.settings_manager.settings.signature

        self.processor = DocxProcessor(
            signature_config=self.signature_config,
            data_dir=self.data_dir,
        )
        self.splitter = DocxSplitter()

    # ---------------------------------------------------------
    # Компании и JSON-файлы
    # ---------------------------------------------------------
    def list_companies(self):
        if not os.path.exists(self.data_dir):
            return []
        return [
            name
            for name in os.listdir(self.data_dir)
            if os.path.isdir(os.path.join(self.data_dir, name))
        ]

    def list_json_files_for_company(self, company: str):
        folder = os.path.join(self.data_dir, company)
        if not os.path.exists(folder):
            return []
        return [
            f for f in os.listdir(folder)
            if f.lower().endswith(".json")
        ]

    # ---------------------------------------------------------
    # Разделение DOCX
    # ---------------------------------------------------------
    def split_docx(self, source_docx, output_dir, method=None, clean_hidden_chars=True, progress_callback=None):
        return self.splitter.split_document(
            source_docx=source_docx,
            output_dir=output_dir,
            method=method,
            clean_hidden_chars=clean_hidden_chars,
            progress_callback=progress_callback,
        )

    # ---------------------------------------------------------
    # Загрузка JSON
    # ---------------------------------------------------------
    def load_replacements(self, json_path):
        return self.processor.load_replacements_from_json(json_path)

    # ---------------------------------------------------------
    # Обработка DOCX
    # ---------------------------------------------------------
    def process_docx(self, input_path, output_path, replacements, tab_number=1):
        self.processor.replace_text(input_path, output_path, replacements, tab_number)

    def process_docx_safe(self, input_path, output_path, replacements, tab_number=1):
        """
        Безопасная обработка DOCX с защитой от дублирования
        """
        # Сначала обрабатываем как обычно
        self.processor.replace_text(input_path, output_path, replacements, tab_number)

        # Затем проверяем и исправляем дублирование
        temp_path = output_path + ".temp"
        self.processor.fix_duplicated_tables(output_path, temp_path)

        # Заменяем файл
        os.replace(temp_path, output_path)

    # ---------------------------------------------------------
    # Специализированная обработка для форм Роснефти
    # ---------------------------------------------------------
    def process_rosneft_form(self, input_path, output_path,
                             organization_name="ПАО «НК «Роснефть»",
                             zakupki_url="https://zakupki.rosneft.ru",
                             additional_replacements=None,
                             tab_number=2):
        """
        Упрощенный метод для обработки форм Роснефти

        Args:
            organization_name: название организации для замены в [скобках]
            zakupki_url: URL для замены http://zakupki.rosneft.ru
            additional_replacements: дополнительные замены
            tab_number: номер вкладки (по умолчанию 2 для подписей)
        """
        replacements = {
            "organization_name": organization_name,
            "zakupki_url": zakupki_url,
        }

        if additional_replacements:
            replacements.update(additional_replacements)

        # Обрабатываем документ
        self.processor.replace_text(input_path, output_path, replacements, tab_number)

        return output_path

    # ---------------------------------------------------------
    # Замена текста с цветным фоном
    # ---------------------------------------------------------
    def replace_text_with_background(self, input_path, output_path, replacements):
        """
        Заменяет текст, у которого есть цветной фон
        """
        self.processor.replace_text_with_background(input_path, output_path, replacements)

    def replace_text_with_background_smart(self, input_path, output_path, replacements, partial_match=True):
        """
        Умная замена текста с цветным фоном
        """
        return self.processor.replace_text_with_background_smart(input_path, output_path, replacements, partial_match)

    # ---------------------------------------------------------
    # Исправление дублирования в существующих файлах
    # ---------------------------------------------------------
    def fix_existing_duplication(self, input_path, output_path=None):
        """
        Исправляет дублирование в существующем файле
        """
        if output_path is None:
            output_path = self.processor.make_output_path(input_path, "_deduplicated")

        return self.processor.fix_duplicated_tables(input_path, output_path)

    # ---------------------------------------------------------
    # PDF
    # ---------------------------------------------------------
    def convert_to_pdf(self, input_docx, output_pdf):
        self.processor.convert_to_pdf(input_docx, output_pdf)

    # ---------------------------------------------------------
    # Настройки подписи
    # ---------------------------------------------------------
    def get_signature_config(self) -> SignatureConfig:
        return self.signature_config

    def set_signature_config(self, signature_path, stamp_path, fullname_text):
        self.signature_config.signature_path = signature_path
        self.signature_config.stamp_path = stamp_path
        self.signature_config.fullname_text = fullname_text
        self._save_settings()

    # ---------------------------------------------------------
    # Последние пути
    # ---------------------------------------------------------
    def set_last_paths(self, input_docx=None, output_docx=None, output_pdf=None):
        if input_docx:
            self.settings_manager.settings.last_input_docx = input_docx
        if output_docx:
            self.settings_manager.settings.last_output_docx = output_docx
        if output_pdf:
            self.settings_manager.settings.last_output_pdf = output_pdf
        self._save_settings()

    def get_last_paths(self):
        s = self.settings_manager.settings
        return s.last_input_docx, s.last_output_docx, s.last_output_pdf

    # ---------------------------------------------------------
    # Внутреннее
    # ---------------------------------------------------------
    def _save_settings(self):
        try:
            self.settings_manager.save()
        except Exception:
            logging.exception("Не удалось сохранить настройки")
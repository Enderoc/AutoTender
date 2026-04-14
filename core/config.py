# config.py
import json
import os
import sys
from dataclasses import dataclass, asdict, field
from typing import Optional


@dataclass
class SignatureConfig:
    signature_path: Optional[str] = None
    stamp_path: Optional[str] = None
    fullname_text: str = "Иванов И.И.\nГенеральный директор"


@dataclass
class AppSettings:
    signature: SignatureConfig = field(default_factory=SignatureConfig)
    last_input_docx: str = ""
    last_output_docx: str = ""
    last_output_pdf: str = ""


class SettingsManager:
    def __init__(self, path: str = "settings.json"):
        # Определяем правильный путь к файлу настроек
        if getattr(sys, 'frozen', False):
            # Если приложение собрано
            base_dir = os.path.dirname(sys.executable)
            self.path = os.path.join(base_dir, "gui", os.path.basename(path))
        else:
            # Если запускается из исходников
            self.path = path

        # Создаем директорию, если ее нет
        os.makedirs(os.path.dirname(self.path), exist_ok=True)

        self.settings = AppSettings()

    def load(self):
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)

            sig = data.get("signature", {})
            self.settings.signature = SignatureConfig(
                signature_path=sig.get("signature_path"),
                stamp_path=sig.get("stamp_path"),
                fullname_text=sig.get("fullname_text", "Иванов И.И.\nГенеральный директор"),
            )
            self.settings.last_input_docx = data.get("last_input_docx", "")
            self.settings.last_output_docx = data.get("last_output_docx", "")
            self.settings.last_output_pdf = data.get("last_output_pdf", "")
        except Exception as e:
            print(f"Ошибка при загрузке настроек: {e}")

    def save(self):
        try:
            data = {
                "signature": asdict(self.settings.signature),
                "last_input_docx": self.settings.last_input_docx,
                "last_output_docx": self.settings.last_output_docx,
                "last_output_pdf": self.settings.last_output_pdf,
            }
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка при сохранении настроек: {e}")
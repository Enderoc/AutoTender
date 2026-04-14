import sys
import os
from cx_Freeze import setup, Executable

# 1. Определяем базовую платформу (Win32GUI убирает консоль при запуске)
base = None
if sys.platform == "win32":
    base = "Win32GUI"

main_script = "app.py"
ASSETS_FOLDER = "assets"

# Путь к иконке для EXE файла
# Убедитесь, что файл nn_logo.ico лежит в папке assets
EXE_ICON = os.path.join(ASSETS_FOLDER, "new_logo.ico")


def find_data_files():
    """Собирает все файлы (настройки, картинки, папки) для сборки"""
    data_files = []

    # --- Добавляем настройки ---
    if os.path.exists("gui/settings.json"):
        # Формат: (откуда_взять, куда_положить)
        data_files.append(("gui/settings.json", "gui/settings.json"))

    # --- Добавляем ВСЕ картинки из assets ---
    if os.path.exists(ASSETS_FOLDER):
        for root, dirs, files in os.walk(ASSETS_FOLDER):
            for file in files:
                # Берем только изображения и иконки
                if file.lower().endswith(('.png', '.jpg', '.jpeg', '.ico', '.gif', '.bmp')):
                    source_path = os.path.join(root, file)
                    # rel_path сохранит структуру подпапок внутри assets, если они есть
                    rel_path = os.path.relpath(source_path, start='.')
                    data_files.append((source_path, rel_path))

    # --- Создаем структуру пустых папок data ---
    data_folders = [
        "data",
        "data/tenders",
        "data/organizations",
        "data/templates"
    ]

    for folder in data_folders:
        if not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)
        # Чтобы папка попала в сборку, добавляем её саму
        data_files.append((folder, folder))

    return data_files


# Опции сборки
build_exe_options = {
    "packages": [
        "tkinter", "os", "sys", "json", "re", "threading",
        "logging", "shutil", "tempfile", "datetime", "uuid",
        "zipfile", "lxml", "lxml.etree", "PIL", "docx",
        "docx2pdf", "subprocess", "typing", "dataclasses"
    ],
    "excludes": [
        "unittest", "email", "http", "xmlrpc", "pydoc",
        "test", "tkinter.test", "numpy", "scipy", "matplotlib",
        "pandas", "pygame", "curses", "sqlite3"
    ],
    "include_files": find_data_files(),
    "optimize": 2,
    "include_msvcr": True,  # Добавляет необходимые библиотеки Microsoft C++
}

# Настройка самого запускаемого файла
executables = [
    Executable(
        main_script,
        base=base,
        target_name="AutoTender.exe",
        # ВОТ ЗДЕСЬ УСТАНАВЛИВАЕТСЯ ИКОНКА EXE
        icon=EXE_ICON if os.path.exists(EXE_ICON) else None,
        copyright="Copyright © 2026",
    )
]

setup(
    name="AutoTender",
    version="1.0.0",
    description="Приложение для обработки тендерных документов",
    options={"build_exe": build_exe_options},
    executables=executables,
)
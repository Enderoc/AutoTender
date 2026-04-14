import sys
import os
from cx_Freeze import setup, Executable

# Определяем основную точку входа
main_script = "app.py"

# Путь к иконке для macOS (обязательно .icns!)
ICON_PATH = "assets/icon.icns"

# Функция для сбора всех файлов и папок
def collect_data_files():
    data_files = []
    
    # Папка assets
    if os.path.exists("assets"):
        for root, dirs, files in os.walk("assets"):
            for file in files:
                src = os.path.join(root, file)
                dst = os.path.join("assets", os.path.relpath(src, "assets"))
                data_files.append((src, dst))
    
    # Папка gui (для settings.json и других файлов)
    if os.path.exists("gui"):
        for root, dirs, files in os.walk("gui"):
            for file in files:
                if file.endswith('.json'):
                    src = os.path.join(root, file)
                    dst = os.path.join("gui", os.path.basename(file))
                    data_files.append((src, dst))
    
    # Папка core (если есть json или другие данные)
    if os.path.exists("core"):
        for root, dirs, files in os.walk("core"):
            for file in files:
                if file.endswith('.json'):
                    src = os.path.join(root, file)
                    dst = os.path.join("core", os.path.basename(file))
                    data_files.append((src, dst))
    
    # Папка data (создаём структуру)
    data_folders = ["data", "data/tenders", "data/organizations", "data/templates"]
    for folder in data_folders:
        os.makedirs(folder, exist_ok=True)
        data_files.append((folder, folder))
    
    return data_files

# Настройки сборки
build_exe_options = {
    "packages": [
        "tkinter", "os", "sys", "json", "re", "threading",
        "logging", "shutil", "tempfile", "datetime", "uuid",
        "zipfile", "lxml", "lxml.etree", "PIL", "PIL.Image",
        "docx", "docx2pdf", "subprocess", "typing"
    ],
    "excludes": [
        "unittest", "email", "http", "xmlrpc", "pydoc",
        "test", "tkinter.test", "numpy", "scipy", "matplotlib",
        "pandas", "pygame", "curses", "sqlite3"
    ],
    "include_files": collect_data_files(),
    "optimize": 2,
}

# Настройки для macOS бандла
bdist_mac_options = {
    "bundle_name": "AhoTender",
    "iconfile": ICON_PATH if os.path.exists(ICON_PATH) else None,
    "plist_items": [
        ("CFBundleIdentifier", "com.yourcompany.ahotender"),
        ("CFBundleVersion", "1.0.0"),
        ("CFBundleShortVersionString", "1.0.0"),
        ("NSHighResolutionCapable", True),
    ],
}

executables = [
    Executable(
        main_script,
        base=None,  # на macOS не нужен Win32GUI
        target_name="AhoTender",
        icon=ICON_PATH if os.path.exists(ICON_PATH) else None,
        copyright="Copyright © 2026",
    )
]

setup(
    name="AhoTender",
    version="1.0.0",
    description="Приложение для обработки тендерных документов",
    options={
        "build_exe": build_exe_options,
        "bdist_mac": bdist_mac_options,
    },
    executables=executables,
)

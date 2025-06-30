import sys
from cx_Freeze import setup, Executable

# Зависимости для включения
build_exe_options = {
    "packages": [
        "tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox", "tkinter.scrolledtext",
        "requests", "pandas", "json", "time", "threading", "os", "sys", "pathlib",
        "bs4", "re", "urllib.parse", "random"
    ],
    "include_files": [
        # Добавьте дополнительные файлы если нужно
        # ("icon.ico", "icon.ico"),
        # ("target.xlsx", "target.xlsx")
    ],
    "excludes": [
        # Исключаем ненужные модули для уменьшения размера
        "unittest", "test", "distutils", "setuptools", "pydoc"
    ]
}

# Опционально добавляем PIL и OpenCV если доступны
try:
    import PIL
    build_exe_options["packages"].extend(["PIL", "PIL.Image", "PIL.ImageTk", "PIL.ImageOps"])
except ImportError:
    pass

try:
    import cv2
    import numpy
    build_exe_options["packages"].extend(["cv2", "numpy"])
except ImportError:
    pass

# Настройки для разных ОС
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Убирает консольное окно в Windows

setup(
    name="ImageProcessor",
    version="2.0",
    description="Обработчик изображений с парсингом",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, icon="icon.ico" if os.path.exists("icon.ico") else None)]
)
from cx_Freeze import setup, Executable
import sys

# Основные настройки приложения
base = None
if sys.platform == "win32":
    # Для WINDOWS GUI-приложения используем 'gui'
    base = "gui"

# Список исполняемых файлов
executables = [Executable(
    "doctofb2.py",        # Главный файл вашего приложения
    base=base,            # База (консольная или графическая)
    target_name="DOCtoFB2.exe",  # Имя итогового EXE-файла
    icon="icon.ico"    # Раскомментируйте, если у вас есть файл иконки
)]

# Опции сборки
build_exe_options = {
    "packages": ["os", "sys", "json", "lxml", "docx", "PIL", "PyQt5"],
    "excludes": ["tkinter"],
    "include_files": [],  # Сюда можно добавить дополнительные файлы (иконки, данные)
}

# Вызов setup
setup(
    name="DOCtoFB2",
    version="1.0",
    description="Конвертер DOC/DOCX в FB2 для Литрес",
    options={"build_exe": build_exe_options},
    executables=executables
)
"""Макрос 2. Формирование ZIP-пакета по Excel-таблице."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile

try:
    from openpyxl import load_workbook
except ModuleNotFoundError as error:
    if error.name != "openpyxl":
        raise
    load_workbook = None


OPENPYXL_REQUIREMENT = "openpyxl>=3.1,<4"


def ensure_dependencies() -> str | None:
    """
    Проверяет наличие openpyxl и при необходимости устанавливает зависимости.

    Установка выполняется через тот же интерпретатор Python, которым запущен
    текущий макрос.
    """
    global load_workbook

    if load_workbook is not None:
        return None

    install_command = [
        sys.executable,
        "-m",
        "pip",
        "install",
        OPENPYXL_REQUIREMENT,
    ]

    try:
        subprocess.check_call(install_command)
    except (OSError, subprocess.CalledProcessError) as error:
        command_text = " ".join(f'"{part}"' if " " in part else part for part in install_command)
        return (
            "Не удалось автоматически установить зависимости.\n\n"
            f"Команда: {command_text}\n\n"
            f"Подробности: {error}"
        )

    try:
        from openpyxl import load_workbook as workbook_loader
    except ModuleNotFoundError as error:
        return f"Зависимость установлена, но openpyxl не загрузился: {error}"

    load_workbook = workbook_loader
    return None



# ============================================================================
# НАСТРОЙКИ МАКРОСА
#
# Здесь собраны параметры, которые можно менять без поиска по всему файлу:
#
# PROJECT_ROOT
#     Корень проекта, относительно которого ищутся основные папки.
#
# PACKAGE_FOLDER_NAME
#     Папка, в которую складываются готовые папки пакетов и ZIP-архивы.
#
# DATE_FORMAT
#     Формат даты и времени в имени создаваемого пакета.
#
# ENABLED_STATUS
#     Значение во втором столбце таблицы, означающее "файл включить".
#     Для таблиц нового генератора это число 1.
#
# LOG_FILE_NAME
#     Имя текстового лога внутри собранного пакета.
# ============================================================================
PROJECT_ROOT = Path.cwd().resolve()


PACKAGE_FOLDER_NAME = "1_Пакеты на отправку"


DATE_FORMAT = "%Y-%m-%d_%H-%M-%S"


ENABLED_STATUS = "1"


LOG_FILE_NAME = "состав_пакета.txt"


# ============================================================================
# ДИАЛОГ ВЫБОРА ТАБЛИЦЫ
# ============================================================================
def choose_manifest() -> Path | None:
    """Показывает окно выбора Excel-файла состава пакета."""
    root = tk.Tk()
    root.withdraw()

    selected = filedialog.askopenfilename(
        title="Выберите таблицу состава пакета",
        initialdir=PROJECT_ROOT,
        filetypes=[("Excel workbook", "*.xlsx *.xlsm")],
    )

    root.destroy()

    return Path(selected) if selected else None


# ============================================================================
# ЧТЕНИЕ И ПРОВЕРКА ТАБЛИЦЫ
# ============================================================================
def read_manifest(
    path: Path,
    root: Path = PROJECT_ROOT,
) -> tuple[str, dict[str, list[str]], int, int]:
    """
    Читает таблицу и превращает ее в структуру "папка -> включенные файлы".

    Ожидаемый формат:
    - A1: название пакета;
    - строка с названием реальной папки первого уровня: главная папка;
    - любая другая строка: файл и его статус;
    - файл копируется только при статусе ENABLED_STATUS (True/1).

    Папка определяется по имени каталога в root, а не по значению B.
    Это важно для Excel-чекбоксов: связанная ячейка B у флажка содержит
    TRUE/FALSE и больше не может использоваться как признак строки папки.
    """

    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active

    package_name = str(sheet["A1"].value or "").strip()

    if not package_name:
        raise ValueError("Ячейка A1 должна содержать название пакета.")

    folders: dict[str, list[str]] = defaultdict(list)

    current_folder = None
    total_file_count = 0
    required_file_count = 0

    root_folder_names = {
        item.name.casefold()
        for item in root.iterdir()
        if item.is_dir() and item.name not in {PACKAGE_FOLDER_NAME, "0_Архив"}
    }

    for row_number in range(2, sheet.max_row + 1):
        name = str(sheet.cell(row_number, 1).value or "").strip()
        status = str(sheet.cell(row_number, 2).value or "").strip().upper()

        if not name:
            continue

        if name.casefold() in root_folder_names:
            if name in folders:
                raise ValueError(f"Папка повторяется в строке {row_number}: {name}")

            folders[name] = []
            current_folder = name
            continue

        if current_folder is None:
            raise ValueError(f"Файл указан до первой папки, строка {row_number}.")

        total_file_count += 1

        if status in {"1", "1.0", "ВКЛ", "TRUE"}:
            status = ENABLED_STATUS

        if status == ENABLED_STATUS:
            folders[current_folder].append(name)
            required_file_count += 1

    if not folders:
        raise ValueError("В таблице не найдено ни одной папки.")

    return package_name, dict(folders), total_file_count, required_file_count


# ============================================================================
# ПОИСК ИСХОДНОГО ФАЙЛА
# ============================================================================
def find_file(folder: Path, file_name: str) -> Path | None:
    """Ищет файл по имени на любой глубине внутри главной папки."""
    matches = sorted(
        (
            path for path in folder.rglob("*")
            if path.is_file() and path.name.casefold() == file_name.casefold()
        ),
        key=lambda path: str(path).casefold(),
    )

    return matches[0] if matches else None


def to_extended_windows_path(path: Path) -> Path:
    """
    Преобразует путь Windows в расширенный формат без ограничения MAX_PATH.

    Это особенно важно для сетевых UNC-путей, где длинное имя проекта,
    папки пакета и файла вместе могут превысить стандартный лимит Windows.
    """
    path_text = str(path)

    if path_text.startswith("\\\\?\\"):
        return path

    if path_text.startswith("\\\\"):
        return Path("\\\\?\\UNC\\" + path_text[2:])

    if len(path_text) >= 248 and len(path_text) >= 2 and path_text[1] == ":":
        return Path("\\\\?\\" + path_text)

    return path



# ============================================================================
# СБОРКА ПАПКИ ПАКЕТА, ЛОГА И ZIP-АРХИВА
# ============================================================================
def build_package(
    manifest_path: Path,
    root: Path = PROJECT_ROOT,
) -> tuple[Path, int, int, int]:
    """
    Выполняет основной пайплайн формирования пакета:

    1. Читает и проверяет Excel-таблицу.
    2. Создает папку пакета с датой и временем.
    3. Создает внутри нее основные папки.
    4. Ищет и копирует отмеченные файлы.
    5. Записывает результат каждого файла в лог.
    6. Упаковывает готовую папку в ZIP.
    """
    
    package_name, folders, total_file_count, required_file_count = read_manifest(
        manifest_path,
        root,
    )

    destination_root = root / PACKAGE_FOLDER_NAME
    destination_root.mkdir(parents=True, exist_ok=True)

    # Недопустимые для имени Windows символы заменяем на "_".
    safe_name = "".join("_" if char in '<>:"/\\|?*' else char for char in package_name)
    package_path = destination_root / f"{safe_name}_{datetime.now().strftime(DATE_FORMAT)}"
    to_extended_windows_path(package_path).mkdir(parents=True, exist_ok=True)

    log_lines = [f"Пакет: {package_name}", f"Таблица: {manifest_path.name}", ""]
    missing = 0
    copied = 0

    # ------------------------------------------------------------------------
    # Обходим папки и включенные строки таблицы.
    # ------------------------------------------------------------------------
    for folder_name, files in folders.items():
        source_folder = root / folder_name
        target_folder = package_path / folder_name
        to_extended_windows_path(target_folder).mkdir(parents=True, exist_ok=True)

        for file_name in files:
            source_file = (
                find_file(source_folder, file_name)
                if source_folder.is_dir()
                else None
            )

            if source_file is None:
                missing += 1
                log_lines.append(f"ОТСУТСТВУЕТ | {folder_name} | {file_name}")
                continue

            target_file = target_folder / source_file.name
            extended_target_file = to_extended_windows_path(target_file)
            extended_target_file.parent.mkdir(parents=True, exist_ok=True)

            try:
                shutil.copy2(source_file, extended_target_file)
            except OSError as error:
                raise OSError(
                    f"Не удалось скопировать файл.\n"
                    f"Источник: {source_file}\n"
                    f"Назначение: {target_file}\n"
                    f"Причина: {error}"
                ) from error

            relative_source = source_file.relative_to(root)
            log_lines.append(f"СКОПИРОВАН | {folder_name} | {file_name} | {relative_source}")
            copied += 1

    # ------------------------------------------------------------------------
    # Сохраняем чек-лист: по нему можно проверить, что планировалось забрать
    # и что действительно было найдено и скопировано.
    # ------------------------------------------------------------------------
    log_lines.extend(
        [
            "",
            f"Всего файлов в полном пакете: {total_file_count}",
            f"Включено в пакет: {required_file_count}",
            f"Выключено: {total_file_count - required_file_count}",
            f"Добавлено фактически: {copied}",
            f"Отсутствует: {missing}",
        ]
    )
    to_extended_windows_path(package_path / LOG_FILE_NAME).write_text(
        "\n".join(log_lines) + "\n",
        encoding="utf-8",
    )

    # ------------------------------------------------------------------------
    # Архивируем содержимое папки пакета.
    # ------------------------------------------------------------------------
    archive_path = destination_root / f"{package_path.name}.zip"

    archived_file_count = 0

    with zipfile.ZipFile(
        to_extended_windows_path(archive_path),
        "w",
        zipfile.ZIP_DEFLATED,
    ) as archive:
        extended_package_path = to_extended_windows_path(package_path)
        extended_archive_root = to_extended_windows_path(package_path.parent)

        # Не используем pathlib.rglob: на длинных UNC-путях он может
        # перечислить только часть дерева, хотя файлы уже скопированы.
        for current_folder, _, file_names in os.walk(extended_package_path):
            for file_name in file_names:
                source_path = Path(current_folder) / file_name
                archive_name = os.path.relpath(source_path, extended_archive_root)
                archive.write(source_path, archive_name)
                archived_file_count += 1

    if archived_file_count != copied + 1:
        raise RuntimeError(
            "ZIP создан неполностью: "
            f"в архив добавлено {archived_file_count} файлов, "
            f"ожидалось {copied + 1} (включая лог)."
        )

    # После упаковки временная папка больше не нужна: оставляем только ZIP.
    shutil.rmtree(to_extended_windows_path(package_path))

    if missing:
        raise RuntimeError(
            f"Пакет создан с отсутствующими файлами: {missing}.\n"
            f"Всего файлов: {total_file_count}; "
            f"включено: {required_file_count}; "
            f"выключено: {total_file_count - required_file_count}.\n"
            f"Лог: {package_path / LOG_FILE_NAME}"
        )
    return archive_path, copied, required_file_count, total_file_count


def show_result_window(title: str, text: str, success: bool) -> None:
    """Показывает итог с явным зеленым или красным оформлением."""
    root = tk.Tk()
    root.withdraw()

    window = tk.Toplevel(root)
    window.title(title)
    window.configure(bg="#C6EFCE" if success else "#FFC7CE")
    window.resizable(False, False)

    label = tk.Label(
        window,
        text=text,
        bg="#C6EFCE" if success else "#FFC7CE",
        fg="#006100" if success else "#9C0006",
        justify="left",
        padx=24,
        pady=20,
        font=("Segoe UI", 11),
    )
    label.pack()

    button = tk.Button(window, text="OK", command=root.destroy, width=12)
    button.pack(pady=(0, 16))
    window.protocol("WM_DELETE_WINDOW", root.destroy)
    window.grab_set()
    root.mainloop()


# ============================================================================
# ТОЧКА ВХОДА
#
# Запуск без аргументов: выбор таблицы через окно Windows.
# Запуск с аргументом: путь к таблице передается через командную строку.
# ============================================================================
def main() -> int:
    dependency_error = ensure_dependencies()
    if dependency_error is not None:
        if len(sys.argv) > 1:
            print(dependency_error, file=sys.stderr)
        else:
            messagebox.showerror("Ошибка установки зависимостей", dependency_error)
        return 1

    manifest_path = Path(sys.argv[1]) if len(sys.argv) > 1 else choose_manifest()

    if manifest_path is None:
        return 0

    try:
        archive_path, copied, required, total = build_package(manifest_path)

    except (OSError, ValueError, RuntimeError) as error:
        if len(sys.argv) > 1:
            print(f"Ошибка формирования пакета: {error}", file=sys.stderr)
        else:
            show_result_window("Ошибка формирования пакета", str(error), success=False)
        return 1

    result_message = (
        f"Пакет сформирован.\n\n"
        f"Всего файлов: {total}.\n"
        f"Включено в пакет: {required}.\n"
        f"Выключено: {total - required}.\n"
        f"Добавлено фактически: {copied}.\n\n"
        f"ZIP:\n{archive_path}"
    )

    if len(sys.argv) > 1:
        print(result_message)
    else:
        show_result_window("Пакет сформирован", result_message, success=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
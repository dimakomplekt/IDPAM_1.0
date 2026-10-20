from pathlib import Path
from datetime import datetime
import zipfile
import os
import re


# ============================================================
# НАСТРОЙКИ
# ============================================================

# Папка, из которой запускается скрипт.
PROJECT_ROOT = Path.cwd().resolve()


# Папка, в которой хранятся архивы.
ARCHIVE_FOLDER_NAME = "0_Архив"


# Формат даты и времени в имени архива.
ARCHIVE_DATE_FORMAT = "%Y-%m-%d_%H-%M-%S"


# Расширение файлов, которые не добавляются в архив.
EXCLUDED_FILE_EXTENSION = ".py"


# Шаблон имени архива с версией.
#
# Пример:
#
# V_1.00_2026-09-01_14-58-23.zip
#
ARCHIVE_VERSION_PATTERN = re.compile(
    r"^V_(\d+)\.(\d{2})_.*\.zip$",
    re.IGNORECASE,
)


# ============================================================
# ПРОВЕРКА ИСКЛЮЧЕНИЙ
# ============================================================

def should_exclude_file(
    file_path: Path,
) -> bool:
    """
    Проверяет, нужно ли исключить файл из архива.

    Исключаются все .py-файлы без учёта регистра.
    """

    return (
        file_path.suffix.lower()
        == EXCLUDED_FILE_EXTENSION
    )


def should_exclude_folder(
    folder_path: Path,
    archive_folder: Path,
) -> bool:
    """
    Проверяет, нужно ли исключить папку из архива.

    Полностью исключается папка 0_Архив.
    """

    return (
        folder_path.resolve()
        == archive_folder.resolve()
    )


# ============================================================
# ОПРЕДЕЛЕНИЕ СЛЕДУЮЩЕЙ ВЕРСИИ
# ============================================================

def get_next_archive_version(
    archive_folder: Path,
) -> int:
    """
    Находит максимальную существующую версию архива
    и возвращает следующую.

    Версия хранится в сотых:

        1.00 -> 100
        1.01 -> 101
        1.99 -> 199
        2.00 -> 200

    Если архивов ещё нет, возвращается 100,
    то есть версия 1.00.
    """

    maximum_version = 99

    if not archive_folder.exists():
        return 100

    for file_name in os.listdir(
        archive_folder
    ):

        # ----------------------------------------------------
        # ПРОВЕРЯЕМ, ЧТО ЭТО ФАЙЛ
        # ----------------------------------------------------

        file_path = (
            archive_folder
            / file_name
        )

        if not file_path.is_file():
            continue

        # ----------------------------------------------------
        # ПРОВЕРЯЕМ ИМЯ
        # ----------------------------------------------------

        match = ARCHIVE_VERSION_PATTERN.match(
            file_name
        )

        if match is None:
            continue

        # ----------------------------------------------------
        # ИЗВЛЕКАЕМ ВЕРСИЮ
        # ----------------------------------------------------

        major_version = int(
            match.group(1)
        )

        minor_version = int(
            match.group(2)
        )

        version_number = (
            major_version * 100
            + minor_version
        )

        if version_number > maximum_version:
            maximum_version = version_number

    # Следующая версия.
    return maximum_version + 1


def format_archive_version(
    version_number: int,
) -> str:
    """
    Преобразует внутреннее числовое представление версии
    в формат V_X.XX.

    Например:

        100 -> V_1.00
        101 -> V_1.01
        199 -> V_1.99
        200 -> V_2.00
    """

    major_version = version_number // 100
    minor_version = version_number % 100

    return (
        f"V_{major_version}.{minor_version:02d}"
    )


# ============================================================
# СОЗДАНИЕ АРХИВА
# ============================================================

def create_archive() -> None:
    """
    Создаёт ZIP-архив всего содержимого папки запуска.

    Исключает:
        - папку 0_Архив;
        - все .py-файлы.

    Имя архива:

        V_X.XX_ГГГГ-ММ-ДД_ЧЧ-ММ-СС.zip
    """

    # --------------------------------------------------------
    # ПАПКА АРХИВА
    # --------------------------------------------------------

    archive_folder = (
        PROJECT_ROOT
        / ARCHIVE_FOLDER_NAME
    )

    archive_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # ОПРЕДЕЛЯЕМ ВЕРСИЮ
    # --------------------------------------------------------

    next_version_number = (
        get_next_archive_version(
            archive_folder
        )
    )

    archive_version = (
        format_archive_version(
            next_version_number
        )
    )

    # --------------------------------------------------------
    # ДАТА И ВРЕМЯ
    # --------------------------------------------------------

    current_datetime = datetime.now()

    archive_datetime = (
        current_datetime.strftime(
            ARCHIVE_DATE_FORMAT
        )
    )

    # --------------------------------------------------------
    # ИМЯ АРХИВА
    # --------------------------------------------------------

    archive_name = (
        f"{archive_version}_"
        f"{archive_datetime}.zip"
    )

    archive_path = (
        archive_folder
        / archive_name
    )

    # --------------------------------------------------------
    # ИНФОРМАЦИЯ
    # --------------------------------------------------------

    print("=" * 70)
    print("СОЗДАНИЕ ZIP-АРХИВА")
    print("=" * 70)
    print()

    print(
        f"Папка запуска: {PROJECT_ROOT}"
    )

    print(
        f"Версия      : {archive_version}"
    )

    print(
        f"Архив       : {archive_path}"
    )

    print()

    archived_files_count = 0
    skipped_files_count = 0

    # --------------------------------------------------------
    # СОЗДАНИЕ ZIP
    # --------------------------------------------------------

    with zipfile.ZipFile(
        archive_path,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:

        # ====================================================
        # РЕКУРСИВНЫЙ ОБХОД
        # ====================================================

        for current_folder, folder_names, file_names in os.walk(
            PROJECT_ROOT,
            topdown=True,
        ):

            current_folder = Path(
                current_folder
            )

            # ------------------------------------------------
            # ИСКЛЮЧАЕМ 0_АРХИВ
            # ------------------------------------------------

            allowed_folder_names = []

            for folder_name in folder_names:

                folder_path = (
                    current_folder
                    / folder_name
                )

                if should_exclude_folder(
                    folder_path,
                    archive_folder,
                ):

                    print(
                        f"[ПАПКА пропущена] "
                        f"{folder_path}"
                    )

                    continue

                allowed_folder_names.append(
                    folder_name
                )

            # Не позволяем os.walk() заходить
            # в 0_Архив.
            folder_names[:] = (
                allowed_folder_names
            )

            # ------------------------------------------------
            # ФАЙЛЫ
            # ------------------------------------------------

            for file_name in file_names:

                source_file = (
                    current_folder
                    / file_name
                )

                # --------------------------------------------
                # ИСКЛЮЧАЕМ .PY
                # --------------------------------------------

                if should_exclude_file(
                    source_file
                ):

                    print(
                        f"[ФАЙЛ пропущен] "
                        f"{source_file}"
                    )

                    skipped_files_count += 1

                    continue

                # --------------------------------------------
                # ОТНОСИТЕЛЬНЫЙ ПУТЬ
                # --------------------------------------------

                relative_path = (
                    source_file.relative_to(
                        PROJECT_ROOT
                    )
                )

                # --------------------------------------------
                # ДОБАВЛЯЕМ В ZIP
                # --------------------------------------------

                archive.write(
                    source_file,
                    arcname=relative_path,
                )

                print(
                    f"[ДОБАВЛЕН] "
                    f"{relative_path}"
                )

                archived_files_count += 1

    # ========================================================
    # РЕЗУЛЬТАТ
    # ========================================================

    print()
    print("=" * 70)
    print("ZIP-АРХИВ СОЗДАН")
    print("=" * 70)
    print()

    print(
        f"Файл архива : {archive_path}"
    )

    print(
        f"Версия      : {archive_version}"
    )

    print(
        f"Добавлено   : {archived_files_count}"
    )

    print(
        f"Пропущено .py: {skipped_files_count}"
    )

    print()


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    """
    Главная функция.
    """

    create_archive()


if __name__ == "__main__":
    main()

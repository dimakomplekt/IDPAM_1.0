from pathlib import Path
import shutil
import os


# ============================================================
# НАСТРОЙКИ
# ============================================================

# Скрипт запускается из папки "3_Вывод".
PROJECT_ROOT = Path.cwd().resolve()


# ------------------------------------------------------------
# ИСХОДНЫЕ ПАПКИ
# ------------------------------------------------------------

SOURCE_STEP = Path("..") / "0_Модели" / "STP"

SOURCE_PDF = (
    Path("..")
    / "1_Чертежи"
    / "2_PDF"
)

SOURCE_DXF = (
    Path("..")
    / "1_Чертежи"
    / "3_DXF"
)

SOURCE_SPECIFICATION = (
    Path("..")
    / "1_Чертежи"
    / "4_СП"
)


# ------------------------------------------------------------
# ЦЕЛЕВЫЕ ПАПКИ
# ------------------------------------------------------------

OUTPUT_STEP = Path("STEP")
OUTPUT_PDF = Path("PDF")
OUTPUT_DXF = Path("DXF")
OUTPUT_SPECIFICATION = Path("Спецификация")


# ------------------------------------------------------------
# ФИЧА НАСЛЕДИЯ ДОКУМЕНТОВ
# ------------------------------------------------------------

SOURCE_INHERITANCE = (
    # Path("..")
    # / ".."
    # / "00153_PZ-EQT-INL-WAA-A"
    # / "3_Вывод"
)

INHERITED_FILES = {

    OUTPUT_PDF: (
        Path("PDF"),
        [
            # "PROJECT_Z.0005",
            # "PROJECT_Z.0014",
            # "PROJECT_Z.0015",
            # "PROJECT_Z.0020",
            # "PROJECT_Z.0021",
            # "PROJECT_Z.0024",
            # "PROJECT_Z.0033",
            # "PROJECT_Z.0050",
            # "PROJECT_Z.0070",
            # "PROJECT_Z.0073",
            # "PROJECT_Z.0074",
            # "PROJECT_Z.0085",
            # "PROJECT_Z.0086",
            # "PROJECT_Z.0087",
            # "PROJECT_Z.0089",
            # "PROJECT_Z.0090",
            # "PROJECT_Z.0091",
            # "PROJECT_Z.0092",
            # "PROJECT_Z.0095",
            # "PROJECT_Z.0096",
            # "PROJECT_Z.0105",
        ],
    ),


    OUTPUT_SPECIFICATION: (
        Path("Спецификация"),
        [
            # "PZ.0060.0031"
        ],
    ),
}

# ------------------------------------------------------------
# ФИЧА НАСЛЕДИЯ ДОКУМЕНТОВ
# ------------------------------------------------------------


# ------------------------------------------------------------
# ИСКЛЮЧЕНИЯ ДЛЯ ПАПОК
# ------------------------------------------------------------
#
# Если название папки содержит хотя бы одно из этих значений
# в любом регистре, вся папка и всё её содержимое
# полностью исключаются.
#

EXCLUDED_FOLDER_WORDS = [
    "архив",
    "oldversions",
    "редактируемый формат",
]


# ------------------------------------------------------------
# ИСКЛЮЧЕНИЯ ДЛЯ ФАЙЛОВ
# ------------------------------------------------------------
#
# Если имя файла содержит хотя бы одно из этих значений
# в любом регистре, файл не копируется.
#

EXCLUDED_FILE_WORDS = [
    "MASTER_FILE",
    "drawing_generator",
    "вывод dxf",
    "Под резку",
    "step_export_log",
]


# ============================================================
# СЛУЖЕБНЫЕ ФУНКЦИИ
# ============================================================

def contains_excluded_word(
    name: str,
    excluded_words: list[str],
) -> bool:
    """
    Проверяет, содержит ли имя хотя бы одно исключаемое слово.

    Проверка выполняется без учёта регистра.
    """

    name_lower = name.lower()

    for excluded_word in excluded_words:
        if excluded_word.lower() in name_lower:
            return True

    return False


def should_exclude_folder(folder_path: Path) -> bool:
    """
    Проверяет, нужно ли полностью исключить папку.
    """

    return contains_excluded_word(
        folder_path.name,
        EXCLUDED_FOLDER_WORDS,
    )


def should_exclude_file(file_path: Path) -> bool:
    """
    Проверяет, нужно ли исключить файл.
    """

    return contains_excluded_word(
        file_path.name,
        EXCLUDED_FILE_WORDS,
    )


def file_matches_identifier(
    file_path: Path,
    identifier: str,
) -> bool:
    """
    Проверяет, что имя файла начинается с кода детали целиком,
    а не с более длинного кода с тем же префиксом.
    """

    remainder = file_path.stem[len(identifier):]

    return (
        file_path.stem[:len(identifier)].upper()
        == identifier.upper()
        and (
            not remainder
            or remainder[0] in " _"
        )
    )


def inherit_selected_files(
    source_folder: Path,
    output_folder: Path,
    identifiers: list[str],
) -> None:
    """
    Копирует перечисленные файлы наследия, заменяя в назначении
    файлы с теми же кодами.
    """

    source_folder = source_folder.resolve()
    output_folder = output_folder.resolve()

    if not source_folder.is_dir():
        print("[ОШИБКА] Папка наследия не существует:")
        print(f"         {source_folder}")
        return

    output_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    print()
    print("=" * 70)
    print(f"Наследие  : {source_folder}")
    print(f"Назначение: {output_folder}")
    print("=" * 70)

    for identifier in identifiers:
        source_files = [
            file_path
            for file_path in source_folder.iterdir()
            if file_path.is_file()
            and file_matches_identifier(file_path, identifier)
        ]

        if not source_files:
            print(
                f"[ОШИБКА] Файл наследия для {identifier} "
                f"не найден в {source_folder}"
            )
            continue

        for target_file in output_folder.iterdir():
            if (
                target_file.is_file()
                and file_matches_identifier(target_file, identifier)
            ):
                target_file.unlink()

        for source_file in source_files:
            target_file = output_folder / source_file.name
            shutil.copy2(
                source_file,
                target_file,
            )
            print(f"[УНАСЛЕДОВАНО]   {source_file}")
            print(f"                 -> {target_file}")


def copy_folder_contents(
    source_folder: Path,
    output_folder: Path,
) -> None:
    """
    Рекурсивно копирует всё содержимое исходной папки
    в целевую папку.

    Структура подпапок сохраняется.

    Существующие файлы заменяются.

    Исключённые папки пропускаются целиком.
    Исключённые файлы пропускаются.
    """

    source_folder = source_folder.resolve()
    output_folder = output_folder.resolve()

    if not source_folder.exists():
        print("[ОШИБКА] Исходная папка не существует:")
        print(f"         {source_folder}")
        return

    if not source_folder.is_dir():
        print("[ОШИБКА] Исходный путь не является папкой:")
        print(f"         {source_folder}")
        return

    output_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    print()
    print("=" * 70)
    print(f"Источник  : {source_folder}")
    print(f"Назначение: {output_folder}")
    print("=" * 70)

    copied_files_count = 0
    skipped_files_count = 0
    skipped_folders_count = 0

    # ========================================================
    # РЕКУРСИВНЫЙ ОБХОД
    # ========================================================
    #
    # Используем os.walk(), поскольку он работает и в старых
    # версиях Python, где Path.walk() ещё отсутствует.
    #
    # topdown=True позволяет изменять folder_names.
    # Удалённые из folder_names папки os.walk() обходить не будет.
    #

    for current_folder, folder_names, file_names in os.walk(
        source_folder,
        topdown=True,
    ):

        current_folder = Path(current_folder)

        # ----------------------------------------------------
        # ФИЛЬТРАЦИЯ ПАПОК
        # ----------------------------------------------------

        allowed_folder_names = []

        for folder_name in folder_names:

            folder_path = current_folder / folder_name

            if should_exclude_folder(folder_path):

                print(f"[ПАПКА пропущена] {folder_path}")

                skipped_folders_count += 1

                continue

            allowed_folder_names.append(folder_name)

        # Очень важный момент:
        #
        # os.walk() использует этот список для дальнейшего
        # рекурсивного обхода.
        #
        # Поэтому исключённые папки вообще не будут посещены.

        folder_names[:] = allowed_folder_names

        # ----------------------------------------------------
        # ОПРЕДЕЛЯЕМ ОТНОСИТЕЛЬНЫЙ ПУТЬ
        # ----------------------------------------------------

        relative_folder = current_folder.relative_to(
            source_folder
        )

        target_folder = output_folder / relative_folder

        target_folder.mkdir(
            parents=True,
            exist_ok=True,
        )

        # ----------------------------------------------------
        # ОБРАБОТКА ФАЙЛОВ
        # ----------------------------------------------------

        for file_name in file_names:

            source_file = current_folder / file_name

            # ------------------------------------------------
            # ПРОВЕРКА ИСКЛЮЧЕНИЯ ФАЙЛА
            # ------------------------------------------------

            if should_exclude_file(source_file):

                print(f"[ФАЙЛ пропущен]   {source_file}")

                skipped_files_count += 1

                continue

            # ------------------------------------------------
            # ЦЕЛЕВОЙ ПУТЬ
            # ------------------------------------------------

            target_file = target_folder / file_name

            target_file.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            # ------------------------------------------------
            # КОПИРОВАНИЕ С ЗАМЕНОЙ
            # ------------------------------------------------
            #
            # copy2() автоматически заменит существующий файл.
            #

            shutil.copy2(
                source_file,
                target_file,
            )

            print(f"[СКОПИРОВАН]      {source_file}")
            print(f"                 -> {target_file}")

            copied_files_count += 1

    # ========================================================
    # СТАТИСТИКА
    # ========================================================

    print()
    print(
        f"Результат: "
        f"скопировано файлов: {copied_files_count}, "
        f"пропущено файлов: {skipped_files_count}, "
        f"пропущено папок: {skipped_folders_count}"
    )


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    """
    Главная функция.
    """

    print("=" * 70)
    print("КОПИРОВАНИЕ ФАЙЛОВ ПРОЕКТА")
    print("=" * 70)
    print()

    print(f"Точка запуска: {PROJECT_ROOT}")
    print()

    # --------------------------------------------------------
    # STEP
    # --------------------------------------------------------

    copy_folder_contents(
        source_folder=PROJECT_ROOT / SOURCE_STEP,
        output_folder=PROJECT_ROOT / OUTPUT_STEP,
    )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    copy_folder_contents(
        source_folder=PROJECT_ROOT / SOURCE_PDF,
        output_folder=PROJECT_ROOT / OUTPUT_PDF,
    )

    # --------------------------------------------------------
    # DXF
    # --------------------------------------------------------

    copy_folder_contents(
        source_folder=PROJECT_ROOT / SOURCE_DXF,
        output_folder=PROJECT_ROOT / OUTPUT_DXF,
    )

    # --------------------------------------------------------
    # СПЕЦИФИКАЦИЯ
    # --------------------------------------------------------

    copy_folder_contents(
        source_folder=PROJECT_ROOT / SOURCE_SPECIFICATION,
        output_folder=PROJECT_ROOT / OUTPUT_SPECIFICATION,
    )




    # --------------------------------------------------------
    # НАСЛЕДИЕ
    # --------------------------------------------------------

    # for output_folder, (source_subfolder, identifiers) in (
    #     INHERITED_FILES.items()
    # ):
    #     inherit_selected_files(
    #         source_folder=(
    #             PROJECT_ROOT
    #             / SOURCE_INHERITANCE
    #             / source_subfolder
    #         ),
    #         output_folder=PROJECT_ROOT / output_folder,
    #         identifiers=identifiers,
    #     )



    print()
    print("=" * 70)
    print("КОПИРОВАНИЕ ЗАВЕРШЕНО")
    print("=" * 70)


if __name__ == "__main__":
    main()

import os
import win32com.client


def convert_excel_to_pdf_final():
    # ============================================================
    # PATHS
    # ============================================================

    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)

    # ============================================================
    # EXCEL
    # ============================================================

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    print(f"Поиск файлов в: {current_dir}")

    # ============================================================
    # PROCESS FILES
    # ============================================================

    for file_name in os.listdir(current_dir):

        if not (
            file_name.endswith((".xlsx", ".xls", ".xlsm"))
            and not file_name.startswith("~$")
        ):
            continue

        excel_path = os.path.join(current_dir, file_name)

        name_without_ext = os.path.splitext(file_name)[0]

        pdf_path = os.path.join(
            parent_dir,
            f"{name_without_ext}.pdf"
        )

        workbook = None

        try:
            print(f"\nОбработка: {file_name}...")

            # ----------------------------------------------------
            # OPEN WORKBOOK
            # ----------------------------------------------------

            workbook = excel.Workbooks.Open(excel_path)

            print(
                f"Листов в книге: {workbook.Worksheets.Count}"
            )

            # ----------------------------------------------------
            # PAGE SETUP FOR EVERY WORKSHEET
            # ----------------------------------------------------

            for sheet in workbook.Worksheets:

                print(
                    f"  Настройка листа: {sheet.Name}"
                )

                page_setup = sheet.PageSetup

                # ------------------------------------------------
                # MARGINS
                #
                # Excel PageSetup uses POINTS.
                # 
                # ШОК НО:
                #
                # 29 = 10 points
                # 0  = 0 points
                # ------------------------------------------------

                page_setup.TopMargin = 29
                page_setup.BottomMargin = 0

                page_setup.LeftMargin = 29
                page_setup.RightMargin = 0

                page_setup.HeaderMargin = 29
                page_setup.FooterMargin = 0

            # ----------------------------------------------------
            # SAVE CHANGES TO EXCEL
            # ----------------------------------------------------

            print("  Сохранение настроек страницы...")

            workbook.Save()

            # ----------------------------------------------------
            # EXPORT ENTIRE WORKBOOK TO PDF
            # ----------------------------------------------------

            print("  Экспорт в PDF...")

            workbook.ExportAsFixedFormat(
                Type=0,                   # xlTypePDF
                Filename=pdf_path,
                Quality=0,                # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )

            # ----------------------------------------------------
            # CLOSE WORKBOOK
            # ----------------------------------------------------

            workbook.Close(SaveChanges=False)
            workbook = None

            print(
                f"  Успешно сохранен: {name_without_ext}.pdf"
            )

        except Exception as e:

            print(
                f"  ОШИБКА при обработке {file_name}:"
            )
            print(
                f"  {e}"
            )

            # ----------------------------------------------------
            # TRY TO CLOSE WORKBOOK AFTER ERROR
            # ----------------------------------------------------

            if workbook is not None:

                try:
                    workbook.Close(
                        SaveChanges=False
                    )

                except Exception:
                    pass

    # ============================================================
    # CLOSE EXCEL
    # ============================================================

    excel.Quit()

    print("\n========================================")
    print("Работа завершена!")
    print("========================================")


# ================================================================
# ENTRY POINT
# ================================================================

if __name__ == "__main__":
    convert_excel_to_pdf_final()
    
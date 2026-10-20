Option Explicit


' =========================================================
' PDF EXPORT (ROOT IDW → PDF)
' =========================================================
'
' Поток макроса:
'
' ROOT (папка запуска)
'     ↓
' Берём все .idw из корня
'     ↓
' Исключаем служебные чертежи
'     ↓
' Создаём папку 2_PDF (если её нет)
'     ↓
' Открываем каждый .idw
'     ↓
' Настраиваем PDF Translator
'     ↓
' Чистим имя файла
'     ↓
' Экспортируем PDF через SaveCopyAs
'     ↓
' Закрываем документ
'     ↓
' Считаем успехи
'     ↓
' Открываем итоговую папку


Sub Main()

' =========================================================
' 1. НАСТРОЙКИ ПАПОК
' =========================================================

' Базовая папка под PDF (выходная структура)
Dim base_PDF_folder_name As String
base_PDF_folder_name = "2_PDF"


' =========================================================
' 2. ОПРЕДЕЛЕНИЕ КОРНЯ
' =========================================================

' Папка, где лежит текущий документ/запуск макроса
Dim root_dir As String
root_dir = ThisDoc.Path


' =========================================================
' 3. СПИСОК ИСКЛЮЧЕНИЙ
' =========================================================

Dim excluded_files As New System.Collections.Generic.HashSet(Of String)( _
    System.StringComparer.OrdinalIgnoreCase)

excluded_files.Add("0.0_MASTER_FILE.idw")
excluded_files.Add("0.1_ШАБЛОН ДЕТАЛЬ.idw")
excluded_files.Add("0.2_ШАБЛОН СБОРКА.idw")


' =========================================================
' 4. СОЗДАНИЕ PDF-КОРНЯ
' =========================================================

' Формируем путь:
' ROOT → .. → 2_PDF
'
' (поднимаемся на уровень вверх относительно чертежей)

Dim base_PDF_dir As String
base_PDF_dir = System.IO.Path.GetFullPath( _
    System.IO.Path.Combine(root_dir, "..\" & base_PDF_folder_name) _
)

' Если папки нет → создаём
If Not System.IO.Directory.Exists(base_PDF_dir) Then
    System.IO.Directory.CreateDirectory(base_PDF_dir)
End If


' =========================================================
' 5. PDF TRANSLATOR (ENGINE EXPORTA)
' =========================================================

Dim pdf_add_in As TranslatorAddIn
pdf_add_in = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

Dim open_context As TranslationContext
open_context = ThisApplication.TransientObjects.CreateTranslationContext
open_context.Type = IOMechanismEnum.kFileBrowseIOMechanism


' =========================================================
' 6. СЧЁТЧИК УСПЕШНЫХ ЭКСПОРТОВ
' =========================================================

Dim total_count As Integer
total_count = 0


' =========================================================
' 7. ПОЛУЧЕНИЕ ЧЕРТЕЖЕЙ
' =========================================================

' Все чертежи (*.idw) в корневой папке
' Вложенные папки не обходятся

Dim files() As String
files = System.IO.Directory.GetFiles(root_dir, "*.idw")


' =========================================================
' 8. ОБХОД ЧЕРТЕЖЕЙ
' =========================================================

Dim curr_file_path As String

For Each curr_file_path In files

    ' Имя файла
    Dim file_name As String
    file_name = System.IO.Path.GetFileName(curr_file_path)

    ' Пропускаем служебные чертежи
    If excluded_files.Contains(file_name) Then
        Continue For
    End If

    Dim doc As DrawingDocument
    doc = Nothing


    ' =====================================================
    ' 8.1 ОТКРЫТИЕ ЧЕРТЕЖА
    ' =====================================================

    Try

        doc = ThisApplication.Documents.Open(curr_file_path, False)


        ' =============================================
        ' 8.2 ИМЯ ФАЙЛА
        ' =============================================

        Dim curr_file_name As String
        curr_file_name = System.IO.Path.GetFileNameWithoutExtension(curr_file_path)


        ' =============================================
        ' 8.3 SANITIZER (защита ФС)
        ' =============================================

        Dim invalid_chars As Char()
        invalid_chars = System.IO.Path.GetInvalidFileNameChars()

        Dim c As Char
        For Each c In invalid_chars
            curr_file_name = curr_file_name.Replace(c, "_"c)
        Next


        ' =============================================
        ' 8.4 ПУТЬ PDF
        ' =============================================

        Dim pdf_path As String
        pdf_path = System.IO.Path.Combine(base_PDF_dir, curr_file_name & ".pdf")


        ' =============================================
        ' 8.5 ОПЦИИ ЭКСПОРТА
        ' =============================================

        Dim opened_options As NameValueMap
        opened_options = ThisApplication.TransientObjects.CreateNameValueMap

        If pdf_add_in.HasSaveCopyAsOptions(doc, open_context, opened_options) Then

            opened_options.Value("All_Color_AS_Black") = False
            opened_options.Value("Remove_Line_Weights") = False
            opened_options.Value("Vector_Resolution") = 400
            opened_options.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets

        End If


        ' =============================================
        ' 8.6 EXPORT ENGINE
        ' =============================================

        Dim opened_data_env As DataMedium
        opened_data_env = ThisApplication.TransientObjects.CreateDataMedium

        opened_data_env.FileName = pdf_path

        pdf_add_in.SaveCopyAs(doc, open_context, opened_options, opened_data_env)


        ' =============================================
        ' 8.7 СЧЁТЧИК
        ' =============================================

        total_count = total_count + 1


    Catch ex As Exception

        MessageBox.Show("Ошибка: " & curr_file_path & vbCrLf & ex.Message)

    Finally

        ' Закрываем документ всегда, даже при ошибке
        If Not doc Is Nothing Then doc.Close(True)

    End Try

Next


' =========================================================
' 9. ФИНАЛ
' =========================================================

MessageBox.Show("Готово. Всего PDF: " & total_count)

System.Diagnostics.Process.Start("explorer.exe", base_PDF_dir)

End Sub
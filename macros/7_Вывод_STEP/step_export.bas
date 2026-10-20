Option Explicit

' ###########################################################################
' #
' # ГЕНЕРАТОР STEP-ФАЙЛОВ INVENTOR
' #
' # ОСНОВНОЙ PIPELINE
' #
' #   ACTIVE DOCUMENT
' #        |
' #        v
' #   ПРОВЕРКА, ЧТО ЭТО СБОРКА
' #        |
' #        v
' #   ПОЛУЧЕНИЕ ПУТИ ОСНОВНОЙ СБОРКИ
' #        |
' #        v
' #   DISCOVERY
' #   Поиск всех подходящих моделей в дереве сборки
' #        |
' #        v
' #   UNIQUE FILTER
' #   Удаление повторов по полному пути модели
' #        |
' #        v
' #   CANDIDATES
' #   Формирование списка моделей для экспорта
' #        |
' #        v
' #   EXPORT
' #   Для каждого кандидата:
' #
' #       open model
' #          |
' #          v
' #       export to STEP
' #          |
' #          v
' #       close document
' #
' #        |
' #        v
' #   REPORT
' #
' ###########################################################################

Sub Main()

    ' =====================================================
    ' БАЗОВЫЕ НАСТРОЙКИ МАКРОСА
    ' =====================================================

    Dim controlled_code As String
    controlled_code = "PZ.0060."

    Dim step_relative_path As String
    step_relative_path = "/STP"

    Dim enable_logging As Boolean
    enable_logging = True

    Dim log_file_name As String
    log_file_name = "step_export_log.txt"


    ' =====================================================
    ' 1. ПРОВЕРКА ВХОДНОГО ДОКУМЕНТА
    ' =====================================================

    Dim assembly_error As Boolean
    assembly_error = False

    check_active_document_is_assembly(assembly_error)

    If assembly_error Then
        Exit Sub
    End If


    ' =====================================================
    ' 2. ПОЛУЧЕНИЕ ПУТИ ОСНОВНОЙ СБОРКИ
    ' =====================================================

    Dim current_assembly_path As String
    current_assembly_path = ThisApplication.ActiveDocument.FullFileName

    If current_assembly_path = "" Then
        MsgBox( _
            "Основная сборка ещё не сохранена." & vbCrLf & _
            "Сохраните сборку и запустите макрос повторно.", _
            vbExclamation, _
            "STEP Export")
        Exit Sub
    End If

    ' =====================================================
    ' ИСХОДНЫЙ ДОКУМЕНТ
    ' =====================================================

    Dim source_document As Object
    source_document = ThisApplication.ActiveDocument

    ' =====================================================
    ' 3. ПОЛУЧЕНИЕ ПАПКИ ОСНОВНОЙ СБОРКИ
    ' =====================================================

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim assembly_folder_path As String
    assembly_folder_path = fso.GetParentFolderName(current_assembly_path)

    If assembly_folder_path = "" Then
        MsgBox( _
            "Не удалось определить папку основной сборки.", _
            vbCritical, _
            "STEP Export")
        Exit Sub
    End If


    ' =====================================================
    ' 4. ПОЛУЧЕНИЕ ПАПКИ STEP
    ' =====================================================

    Dim step_folder_path As String
    step_folder_path = resolve_step_folder_path(assembly_folder_path, step_relative_path, fso)

    If step_folder_path = "" Then
        Exit Sub
    End If


    ' =====================================================
    ' 5. СОЗДАНИЕ СТРУКТУР DISCOVERY
    ' =====================================================

    Dim generation_candidates As Object
    generation_candidates = CreateObject("System.Collections.ArrayList")

    Dim processed_model_paths As Object
    processed_model_paths = CreateObject("Scripting.Dictionary")

    Dim exported_files As Object
    exported_files = CreateObject("System.Collections.ArrayList")

    Dim skipped_files As Object
    skipped_files = CreateObject("System.Collections.ArrayList")

    Dim failed_files As Object
    failed_files = CreateObject("System.Collections.ArrayList")


    ' =====================================================
    ' 6. DISCOVERY
    ' =====================================================

    process_root_model_for_step_export( _
        ThisApplication.ActiveDocument, _
        generation_candidates, _
        processed_model_paths, _
        controlled_code)

    collect_step_candidates( _
        ThisApplication.ActiveDocument.ComponentDefinition.Occurrences, _
        generation_candidates, _
        processed_model_paths, _
        controlled_code)


    ' =====================================================
    ' 7. ПРОВЕРКА КАНДИДАТОВ
    ' =====================================================

    If generation_candidates.Count = 0 Then
        MsgBox( _
            "В дереве сборки не найдено моделей с шифром " & controlled_code & _
            vbCrLf & "Экспорт STEP не выполнен.", _
            vbInformation, _
            "STEP Export")
        Exit Sub
    End If


    ' =====================================================
    ' 8. EXPORT
    ' =====================================================

    generate_step_files( _
        generation_candidates, _
        step_folder_path, _
        fso, _
        exported_files, _
        skipped_files, _
        failed_files, _
        source_document)


    ' =====================================================
    ' 9. REPORT
    ' =====================================================

    If enable_logging Then
        write_step_export_report( _
            step_folder_path, _
            generation_candidates, _
            exported_files, _
            skipped_files, _
            failed_files, _
            log_file_name, _
            controlled_code)
    End If

    show_step_export_result( _
        exported_files, _
        skipped_files, _
        failed_files, _
        step_folder_path)

End Sub


' =========================================================
' ПРОВЕРКА АКТИВНОГО ДОКУМЕНТА
' =========================================================

Sub check_active_document_is_assembly(assembly_error As Boolean)

    assembly_error = False

    If ThisApplication.ActiveDocument Is Nothing Then
        MsgBox( _
            "Не удалось получить активный документ Inventor.", _
            vbCritical, _
            "STEP Export")
        assembly_error = True
        Exit Sub
    End If

    If ThisApplication.ActiveDocument.DocumentType <> _
        DocumentTypeEnum.kAssemblyDocumentObject Then

        MsgBox( _
            "Правило запускается только из сборки (.iam)." & vbCrLf & _
            "Откройте основную сборку и запустите правило повторно.", _
            vbExclamation, _
            "STEP Export")

        assembly_error = True
        Exit Sub
    End If

End Sub


' =========================================================
' ПОЛУЧЕНИЕ ПАПКИ STEP
' =========================================================

Function resolve_step_folder_path( _
    assembly_folder_path As String, _
    relative_path As String, _
    fso As Object) As String

    resolve_step_folder_path = ""

    If assembly_folder_path = "" Then
        MsgBox( _
            "Пустой путь к папке основной сборки.", _
            vbCritical, _
            "STEP Export")
        Exit Function
    End If

    If relative_path = "" Then
        MsgBox( _
            "Не задан относительный путь к папке STEP.", _
            vbCritical, _
            "STEP Export")
        Exit Function
    End If

    On Error GoTo handle_error

    Dim normalized_relative_path As String
    normalized_relative_path = relative_path

    If Left(normalized_relative_path, 1) = "/" Then
        normalized_relative_path = Mid(normalized_relative_path, 2)
    End If

    If Left(normalized_relative_path, 1) = "\" Then
        normalized_relative_path = Mid(normalized_relative_path, 2)
    End If

    Dim step_folder_path As String
    step_folder_path = fso.GetAbsolutePathName( _
        assembly_folder_path & "\" & normalized_relative_path)

    If Not fso.FolderExists(step_folder_path) Then
        fso.CreateFolder(step_folder_path)
    End If

    If Not fso.FolderExists(step_folder_path) Then
        MsgBox( _
            "Не удалось создать папку STEP:" & vbCrLf & step_folder_path, _
            vbCritical, _
            "STEP Export")
        Exit Function
    End If

    resolve_step_folder_path = step_folder_path
    Exit Function

handle_error:

    MsgBox( _
        "Ошибка при создании папки STEP." & vbCrLf & _
        "Путь сборки: " & assembly_folder_path & vbCrLf & _
        "Ошибка: " & Err.Description, _
        vbCritical, _
        "STEP Export")

End Function


' =========================================================
' ДОБАВЛЕНИЕ КОРНЕВОЙ СБОРКИ
' =========================================================

Sub process_root_model_for_step_export( _
    root_assembly_document As Object, _
    generation_candidates As Object, _
    processed_model_paths As Object, _
    controlled_code As String)

    On Error GoTo handle_error

    If root_assembly_document Is Nothing Then
        Exit Sub
    End If

    If root_assembly_document.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Exit Sub
    End If

    Dim model_path As String
    model_path = root_assembly_document.FullFileName

    If model_path = "" Then
        Exit Sub
    End If

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim file_name As String
    file_name = fso.GetFileName(model_path)

    If contains_controlled_code(file_name, controlled_code) = False Then
        Exit Sub
    End If

    append_unique_model_path( _
        model_path, _
        generation_candidates, _
        processed_model_paths)

    Exit Sub

handle_error:

    ' Ошибка корневой модели не должна останавливаться
    ' дальнейшей генерацией.

End Sub


' =========================================================
' ОБХОД ДЕРЕВА СБОРКИ
' =========================================================

Sub collect_step_candidates( _
    occurrences As Object, _
    generation_candidates As Object, _
    processed_model_paths As Object, _
    controlled_code As String)

    If occurrences Is Nothing Then
        Exit Sub
    End If

    Dim occurrence As Object

    For Each occurrence In occurrences

        If occurrence Is Nothing Then
            GoTo continue_loop
        End If

        If occurrence.Suppressed Then
            GoTo continue_loop
        End If

        Dim model_document As Object
        model_document = Nothing

        On Error Resume Next
        model_document = occurrence.Definition.Document
        On Error GoTo 0

        If model_document Is Nothing Then
            GoTo continue_loop
        End If

        If model_document.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            GoTo continue_loop
        End If

        process_model_for_step_export( _
            occurrence, _
            generation_candidates, _
            processed_model_paths, _
            controlled_code)

        If Not occurrence.Definition Is Nothing Then
            If Not occurrence.Definition.Occurrences Is Nothing Then
                collect_step_candidates( _
                    occurrence.Definition.Occurrences, _
                    generation_candidates, _
                    processed_model_paths, _
                    controlled_code)
            End If
        End If

continue_loop:
    Next

End Sub


' =========================================================
' ПРОВЕРКА ОДНОЙ МОДЕЛИ
' =========================================================

Sub process_model_for_step_export( _
    occurrence As Object, _
    generation_candidates As Object, _
    processed_model_paths As Object, _
    controlled_code As String)

    On Error GoTo handle_error

    If occurrence Is Nothing Then
        Exit Sub
    End If

    Dim model_document As Object
    model_document = occurrence.Definition.Document

    If model_document Is Nothing Then
        Exit Sub
    End If

    If model_document.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Exit Sub
    End If

    Dim model_path As String
    model_path = model_document.FullFileName

    If model_path = "" Then
        Exit Sub
    End If

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim model_file_name As String
    model_file_name = fso.GetFileName(model_path)

    If contains_controlled_code(model_file_name, controlled_code) = False Then
        Exit Sub
    End If

    append_unique_model_path( _
        model_path, _
        generation_candidates, _
        processed_model_paths)

    Exit Sub

handle_error:

    ' Ошибка одного компонента не должна останавливаться
    ' для остальных моделей.

End Sub


' =========================================================
' УНИКАЛЬНЫЙ ПОИСК И ДОБАВЛЕНИЕ КАНДИДАТА
' =========================================================

Sub append_unique_model_path( _
    model_path As String, _
    generation_candidates As Object, _
    processed_model_paths As Object)

    If model_path = "" Then
        Exit Sub
    End If

    Dim model_key As String
    model_key = LCase(model_path)

    If processed_model_paths.Exists(model_key) Then
        Exit Sub
    End If

    processed_model_paths.Add(model_key, True)
    generation_candidates.Add(model_path)

End Sub


' =========================================================
' КОНТРОЛЬ НАЛИЧИЯ ШИФРА
' =========================================================

Function contains_controlled_code( _
    file_name As String, _
    controlled_code As String) As Boolean

    contains_controlled_code = False

    If file_name = "" Then
        Exit Function
    End If

    If controlled_code = "" Then
        Exit Function
    End If

    If InStr(1, file_name, controlled_code, vbTextCompare) > 0 Then
        contains_controlled_code = True
    End If

End Function


' =========================================================
' ГЕНЕРАЦИЯ STEP-ФАЙЛОВ
' =========================================================

Sub generate_step_files( _
    generation_candidates As Object, _
    step_folder_path As String, _
    fso As Object, _
    exported_files As Object, _
    skipped_files As Object, _
    failed_files As Object, _
    source_document As Object)

    If generation_candidates Is Nothing Then Exit Sub
    If generation_candidates.Count = 0 Then Exit Sub

    Dim progress_bar As Object
    progress_bar = ThisApplication.CreateProgressBar( _
        False, _
        generation_candidates.Count, _
        "Экспорт STEP")

    Dim candidate_path As Object
    Dim candidate_index As Integer
    candidate_index = 0

    On Error GoTo handle_common_error

    ThisApplication.ScreenUpdating = False

    For Each candidate_path In generation_candidates

        candidate_index = candidate_index + 1

        If Not progress_bar Is Nothing Then
            progress_bar.Message = _
                "Экспорт STEP: " & candidate_index & " / " & _
                generation_candidates.Count

            progress_bar.UpdateProgress
        End If

        If candidate_path <> "" Then
            export_single_step_file( _
                CStr(candidate_path), _
                step_folder_path, _
                fso, _
                exported_files, _
                skipped_files, _
                failed_files, _
                source_document)
        End If

    Next

    ThisApplication.ScreenUpdating = True

    If Not progress_bar Is Nothing Then
        progress_bar.Close
    End If

    Exit Sub

handle_common_error:

    ThisApplication.ScreenUpdating = True

    If Not progress_bar Is Nothing Then
        progress_bar.Close
    End If

    MsgBox( _
        "Критическая ошибка генерации STEP." & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "STEP Export")

End Sub


' =========================================================
' ЭКСПОРТ ОДНОГО DOCUMENT В STEP
' =========================================================

Sub export_single_step_file( _
    model_path As String, _
    step_folder_path As String, _
    fso As Object, _
    exported_files As Object, _
    skipped_files As Object, _
    failed_files As Object, _
    source_document As Object)

    If model_path = "" Then
        Exit Sub
    End If


    ' =====================================================
    ' 1. ФОРМИРОВАНИЕ ИМЕНИ STEP-ФАЙЛА
    ' =====================================================

    Dim model_file_name As String
    model_file_name = fso.GetFileName(model_path)

    Dim step_file_name As String
    step_file_name = fso.GetBaseName(model_file_name) & ".stp"

    Dim step_file_path As String
    step_file_path = fso.BuildPath( _
        step_folder_path, _
        step_file_name)


    ' =====================================================
    ' 2. ПРОВЕРКА СУЩЕСТВУЮЩЕГО STEP
    ' =====================================================

    ' Блокировка перезаписи - удалить комменты
    ' If fso.FileExists(step_file_path) Then
    '     skipped_files.Add(step_file_path)
    '     Exit Sub
    ' End If


    On Error GoTo handle_error


    ' =====================================================
    ' 3. ПОЛУЧЕНИЕ STEP TRANSLATOR
    ' =====================================================

    Dim step_translator As Object

    step_translator = ThisApplication.ApplicationAddIns.ItemById( _
        "{90AF7F40-0C01-11D5-8E83-0010B541CD80}")


    If step_translator Is Nothing Then

        failed_files.Add( _
            step_file_path & _
            " | STEP translator not available")

        Exit Sub

    End If


    ' =====================================================
    ' 4. ОТКРЫТИЕ МОДЕЛИ
    ' =====================================================

    Dim model_document As Object

    model_document = ThisApplication.Documents.Open(model_path)


    If model_document Is Nothing Then

        failed_files.Add( _
            step_file_path & _
            " | document could not be opened")

        Exit Sub

    End If


    ' =====================================================
    ' 5. СОЗДАНИЕ TRANSLATION CONTEXT
    ' =====================================================

    Dim translation_context As Object

    translation_context = _
        ThisApplication.TransientObjects.CreateTranslationContext

    translation_context.Type = kFileBrowseIOMechanism


    ' =====================================================
    ' 6. СОЗДАНИЕ EXPORT OPTIONS
    ' =====================================================

    Dim export_options As Object

    export_options = _
        ThisApplication.TransientObjects.CreateNameValueMap


    ' STEP translator может иметь собственные настройки.
    '
    ' В частности, Autodesk использует:
    '
    '   ApplicationProtocolType = 3
    '
    ' где:
    '
    '   2 = AP 203
    '   3 = AP 214
    '
    ' Пока используем AP 214.

    If step_translator.HasSaveCopyAsOptions( _
        model_document, _
        translation_context, _
        export_options) Then

        export_options.Value("ApplicationProtocolType") = 3

    End If


    ' =====================================================
    ' 7. СОЗДАНИЕ TARGET DATA
    ' =====================================================

    Dim target_data As Object

    target_data = _
        ThisApplication.TransientObjects.CreateDataMedium

    target_data.FileName = step_file_path


    ' =====================================================
    ' 8. ЭКСПОРТ STEP
    ' =====================================================

    step_translator.SaveCopyAs( _
        model_document, _
        translation_context, _
        export_options, _
        target_data)


    ' =====================================================
    ' 9. ПРОВЕРКА РЕЗУЛЬТАТА
    ' =====================================================

    If fso.FileExists(step_file_path) Then

        exported_files.Add(step_file_path)

    Else

        failed_files.Add( _
            step_file_path & _
            " | STEP file was not created")

    End If


    ' =====================================================
    ' 10. ЗАКРЫТИЕ ДОКУМЕНТА
    ' =====================================================

    If Not model_document Is source_document Then
        model_document.Close(False)
    End If

    Exit Sub


' =========================================================
' ОБРАБОТКА ОШИБКИ
' =========================================================

handle_error:

    Dim error_description As String

    error_description = Err.Description

    On Error Resume Next

    If Not model_document Is Nothing Then
        If Not model_document Is source_document Then
            model_document.Close(False)
        End If
    End If

    On Error GoTo 0


    failed_files.Add( _
        step_file_path & _
        " | " & _
        error_description)

End Sub


' =========================================================
' ЗАПИСЬ ФИНАЛЬНОГО ОТЧЁТА
' =========================================================

Sub write_step_export_report( _
    step_folder_path As String, _
    generation_candidates As Object, _
    exported_files As Object, _
    skipped_files As Object, _
    failed_files As Object, _
    log_file_name As String, _
    controlled_code As String)

    Dim log_path As String
    log_path = step_folder_path & "\" & log_file_name

    Dim log_content As String
    log_content = "STEP export report" & vbCrLf & _
        "Generated at: " & Now & vbCrLf & _
        "Controlled code: " & controlled_code & vbCrLf & _
        "Target folder: " & step_folder_path & vbCrLf & _
        "Candidates total: " & generation_candidates.Count & vbCrLf & _
        "Exported: " & exported_files.Count & vbCrLf & _
        "Skipped: " & skipped_files.Count & vbCrLf & _
        "Failed: " & failed_files.Count & vbCrLf & vbCrLf

    Dim item As Object

    If exported_files.Count > 0 Then
        log_content = log_content & "Exported files:" & vbCrLf
        For Each item In exported_files
            log_content = log_content & " - " & CStr(item) & vbCrLf
        Next
        log_content = log_content & vbCrLf
    End If

    If skipped_files.Count > 0 Then
        log_content = log_content & "Skipped files:" & vbCrLf
        For Each item In skipped_files
            log_content = log_content & " - " & CStr(item) & vbCrLf
        Next
        log_content = log_content & vbCrLf
    End If

    If failed_files.Count > 0 Then
        log_content = log_content & "Failed files:" & vbCrLf
        For Each item In failed_files
            log_content = log_content & " - " & CStr(item) & vbCrLf
        Next
        log_content = log_content & vbCrLf
    End If

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim output_stream As Object
    output_stream = fso.CreateTextFile(log_path, True, False)
    output_stream.Write(log_content)
    output_stream.Close

End Sub


' =========================================================
' ПОКАЗ РЕЗУЛЬТАТА
' =========================================================

Sub show_step_export_result( _
    exported_files As Object, _
    skipped_files As Object, _
    failed_files As Object, _
    step_folder_path As String)

    Dim message As String
    message = "Экспорт STEP завершён." & vbCrLf & vbCrLf
    message = message & "Создано: " & exported_files.Count & vbCrLf
    message = message & "Пропущено: " & skipped_files.Count & vbCrLf
    message = message & "Ошибок: " & failed_files.Count & vbCrLf & vbCrLf
    message = message & "Папка: " & step_folder_path

    MsgBox(message, vbInformation, "STEP Export")

End Sub

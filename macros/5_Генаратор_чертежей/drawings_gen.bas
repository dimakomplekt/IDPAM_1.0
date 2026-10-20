Option Explicit


' ###########################################################################
' #
' # ГЕНЕРАТОР ЗАГОТОВОК ЧЕРТЕЖЕЙ INVENTOR
' #
' # ОБЩИЙ PIPELINE
' #
' #   ACTIVE DOCUMENT
' #        |
' #        v
' #   ПРОВЕРКА, ЧТО ЭТО СБОРКА
' #        |
' #        v
' #   ПОЛУЧЕНИЕ ПУТЕЙ
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
' #   Формирование списка того, что нужно создать
' #        |
' #        v
' #   GENERATION
' #   Для каждого кандидата:
' #
' #       template
' #          |
' #          v
' #       copy template -> target .idw
' #          |
' #          v
' #       open model
' #          |
' #          v
' #       open drawing
' #          |
' #          v
' #       create main view
' #          |
' #          v
' #       copy iProperties
' #          |
' #          v
' #       save
' #          |
' #          v
' #       close
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
    controlled_code = "PZ.0066."

    Dim drawings_relative_path As String
    drawings_relative_path = "..\1_Чертежи\1_DWG"

    Dim detail_template_path As String
    detail_template_path = _
        "Y:\000_Библиотека\400_RnD\099_iLogicRepository\Templates\ru-RU\kriskovets\detail.idw"

    Dim assembly_template_path As String
    assembly_template_path = _
        "Y:\000_Библиотека\400_RnD\099_iLogicRepository\Templates\ru-RU\kriskovets\assembly.idw"


    ' =====================================================
    ' 1. ПРОВЕРКА ВХОДНОГО ДОКУМЕНТА
    ' =====================================================
    '
    ' Pipeline:
    '
    '   Inventor
    '      |
    '      v
    '   ActiveDocument
    '      |
    '      v
    '   Проверка: это Assembly?
    '
    ' Если нет, весь pipeline останавливается здесь.
    '

    Dim assembly_error As Boolean
    assembly_error = False

    check_active_document_is_assembly(assembly_error)

    If assembly_error Then
        Exit Sub
    End If


    ' =====================================================
    ' 2. ПОЛУЧЕНИЕ ПУТИ ОСНОВНОЙ СБОРКИ
    ' =====================================================
    '
    ' На этом этапе нам нужен физический путь основной сборки.
    '
    ' Он используется как отправная точка для:
    '
    '   - определения папки сборки;
    '   - построения относительного пути к DWG;
    '   - дальнейшего формирования абсолютных путей.
    '

    Dim current_assembly_path As String
    current_assembly_path = ThisApplication.ActiveDocument.FullFileName

    If current_assembly_path = "" Then

        MsgBox( _
            "Основная сборка ещё не сохранена." & vbCrLf & _
            "Перед запуском макроса необходимо сохранить сборку.", _
            vbExclamation, _
            "Генератор чертежей")

        Exit Sub

    End If


    ' =====================================================
    ' 3. ПОЛУЧЕНИЕ ПАПКИ ОСНОВНОЙ СБОРКИ
    ' =====================================================

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim assembly_folder_path As String
    assembly_folder_path = _
        fso.GetParentFolderName(current_assembly_path)

    If assembly_folder_path = "" Then

        MsgBox( _
            "Не удалось определить папку основной сборки.", _
            vbCritical, _
            "Генератор чертежей")

        Exit Sub

    End If


    ' =====================================================
    ' 4. ПОЛУЧЕНИЕ ПАПКИ ЧЕРТЕЖЕЙ
    ' =====================================================
    '
    ' Здесь относительный путь превращается в абсолютный.
    '
    ' Например:
    '
    '   assembly_folder
    '          |
    '          +-- ..\1_Чертежи\1_DWG
    '
    '                |
    '                v
    '
    '          абсолютный путь к DWG
    '

    Dim drawings_folder_path As String

    drawings_folder_path = get_drawings_folder_path( _
        assembly_folder_path, _
        drawings_relative_path)

    If drawings_folder_path = "" Then
        Exit Sub
    End If


    ' =====================================================
    ' 5. СОЗДАНИЕ СТРУКТУР DISCOVERY
    ' =====================================================
    '
    ' generation_candidates
    '
    '   Список моделей, для которых потенциально нужно
    '   создать чертёж.
    '
    '
    ' processed_model_paths
    '
    '   Dictionary для защиты от повторной обработки
    '   одной и той же модели.
    '
    ' Почему путь, а не имя?
    '
    ' Потому что два разных файла теоретически могут иметь
    ' одинаковое имя, находясь в разных каталогах.
    '
    ' Поэтому уникальность определяем по полному пути.
    '

    Dim generation_candidates As Object
    generation_candidates = CreateObject( _
        "System.Collections.ArrayList")

    Dim processed_model_paths As Object
    processed_model_paths = CreateObject( _
        "Scripting.Dictionary")


    ' =====================================================
    ' 6. ПОЛУЧЕНИЕ КОРНЕВОЙ СБОРКИ
    ' =====================================================

    Dim root_assembly_document As Object
    root_assembly_document = ThisApplication.ActiveDocument

    Dim root_assembly_definition As Object
    root_assembly_definition = _
        root_assembly_document.ComponentDefinition


    ' =====================================================
    ' 7. DISCOVERY
    ' =====================================================
    '
    ' ВАЖНЫЙ МОМЕНТ АРХИТЕКТУРЫ:
    '
    ' Здесь мы ещё НИЧЕГО НЕ СОЗДАЁМ.
    '
    ' Discovery отвечает только на вопрос:
    '
    '   "Какие модели существуют в дереве и подходят
    '    под наши условия?"
    '
    ' Результатом Discovery является:
    '
    '   generation_candidates
    '
    ' После завершения этого этапа список кандидатов уже
    ' полностью сформирован.
    '

    process_root_model_for_generation( _
        root_assembly_document, _
        generation_candidates, _
        processed_model_paths, _
        controlled_code)

    collect_generation_candidates( _
        root_assembly_definition.Occurrences, _
        generation_candidates, _
        processed_model_paths, _
        controlled_code)


    ' =====================================================
    ' 8. ОТЧЁТ DISCOVERY
    ' =====================================================
    '
    ' Показываем, что именно Discovery нашёл.
    '
    ' Это полезно отделять от Generation:
    '
    '   Discovery = ЧТО нужно сделать
    '   Generation = СДЕЛАТЬ это
    '

    log_generation_candidates( _
        generation_candidates, _
        drawings_folder_path, _
        controlled_code)


    ' =====================================================
    ' 9. GENERATION
    ' =====================================================
    '
    ' Только здесь начинается фактическое создание файлов.
    '
    ' Для каждого candidate:
    '
    '   candidate
    '       |
    '       v
    '   определить имя .idw
    '       |
    '       v
    '   проверить существование
    '       |
    '       v
    '   выбрать template
    '       |
    '       v
    '   скопировать template
    '       |
    '       v
    '   открыть model
    '       |
    '       v
    '   открыть drawing
    '       |
    '       v
    '   создать основной вид
    '       |
    '       v
    '   перенести iProperties
    '       |
    '       v
    '   сохранить
    '
    generate_drawings_from_candidates( _
        generation_candidates, _
        drawings_folder_path, _
        detail_template_path, _
        assembly_template_path)

End Sub


' =========================================================
' ПРОВЕРКА АКТИВНОГО ДОКУМЕНТА
' =========================================================

Sub check_active_document_is_assembly( _
    ByRef assembly_error As Boolean)

    assembly_error = False

    If ThisApplication.ActiveDocument Is Nothing Then

        MsgBox( _
            "Не удалось получить активный документ Inventor.", _
            vbCritical, _
            "Генератор чертежей")

        assembly_error = True
        Exit Sub

    End If


    If ThisApplication.ActiveDocument.DocumentType <> _
        DocumentTypeEnum.kAssemblyDocumentObject Then

        MsgBox( _
            "Правило предназначено для запуска строго из сборки (.iam)." & _
            vbCrLf & _
            "Откройте основную сборку и запустите правило повторно.", _
            vbExclamation, _
            "Генератор чертежей")

        assembly_error = True
        Exit Sub

    End If

End Sub


' =========================================================
' ДОБАВЛЕНИЕ КОРНЕВОЙ СБОРКИ
' =========================================================
'
' Корневая сборка не является occurrence внутри самой себя,
' поэтому её нужно обработать отдельно.
'
' Таким образом Discovery охватывает:
'
'   1. root assembly
'   2. все occurrences
'   3. все вложенные occurrences подсборок
'

Sub process_root_model_for_generation( _
    root_assembly_document As Object, _
    generation_candidates As Object, _
    processed_model_paths As Object, _
    controlled_code As String)

    On Error GoTo handle_error

    Dim model_path As String
    model_path = root_assembly_document.FullFileName

    If model_path = "" Then Exit Sub

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim model_file_name As String
    model_file_name = fso.GetFileName(model_path)

    If InStr(1, model_file_name, controlled_code, vbTextCompare) = 0 Then
        Exit Sub
    End If

    Dim model_key As String
    model_key = LCase(model_path)

    If processed_model_paths.Exists(model_key) Then Exit Sub

    processed_model_paths.Add(model_key, True)

    Dim candidate(2) As Object

    candidate(0) = model_path
    candidate(1) = model_file_name
    candidate(2) = "Assembly"

    generation_candidates.Add(candidate)

    Exit Sub


handle_error:

    ' Ошибка корневой модели не должна останавливать
    ' дальнейшую работу макроса.

End Sub


' =========================================================
' ПОЛУЧЕНИЕ ПАПКИ ЧЕРТЕЖЕЙ
' =========================================================

Function get_drawings_folder_path( _
    assembly_folder_path As String, _
    relative_path As String) As String

    get_drawings_folder_path = ""

    If assembly_folder_path = "" Then

        MsgBox( _
            "Пустой путь к папке основной сборки.", _
            vbCritical, _
            "Генератор чертежей")

        Exit Function

    End If


    If relative_path = "" Then

        MsgBox( _
            "Не задан относительный путь к папке чертежей.", _
            vbCritical, _
            "Генератор чертежей")

        Exit Function

    End If


    On Error GoTo handle_error


    ' -----------------------------------------------------
    ' Формируем абсолютный путь.
    ' -----------------------------------------------------

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim drawings_folder_path As String

    drawings_folder_path = fso.GetAbsolutePathName( _
        assembly_folder_path & "\" & relative_path)


    ' -----------------------------------------------------
    ' Создаём папку, если она отсутствует.
    ' -----------------------------------------------------

    If Not fso.FolderExists(drawings_folder_path) Then

        fso.CreateFolder(drawings_folder_path)

    End If


    ' -----------------------------------------------------
    ' Проверяем результат.
    ' -----------------------------------------------------

    If Not fso.FolderExists(drawings_folder_path) Then

        MsgBox( _
            "Не удалось создать папку чертежей:" & vbCrLf & _
            drawings_folder_path, _
            vbCritical, _
            "Генератор чертежей")

        Exit Function

    End If


    get_drawings_folder_path = drawings_folder_path

    Exit Function


handle_error:

    MsgBox( _
        "Ошибка при определении папки чертежей." & vbCrLf & _
        "Путь сборки: " & assembly_folder_path & vbCrLf & _
        "Относительный путь: " & relative_path & vbCrLf & _
        "Ошибка: " & Err.Description, _
        vbCritical, _
        "Генератор чертежей")

End Function


' =========================================================
' ОБХОД ДЕРЕВА СБОРКИ
' =========================================================
'
' Рекурсивный Discovery.
'
' Для каждого occurrence:
'
'   SUPPRESSED?
'       |
'       +-- Да --> пропускаем
'       |
'       +-- Нет
'            |
'            +-- Assembly --> проверяем + спускаемся глубже
'            |
'            +-- Part    --> проверяем
'
' Здесь также НЕ происходит генерации файлов.
'

Sub collect_generation_candidates( _
    occurrences As Object, _
    generation_candidates As Object, _
    processed_model_paths As Object, _
    controlled_code As String)

    If occurrences Is Nothing Then
        Exit Sub
    End If


    Dim occurrence As Object

    For Each occurrence In occurrences

        ' -------------------------------------------------
        ' Пропускаем подавленные компоненты.
        ' -------------------------------------------------

        If occurrence.Suppressed Then
            Continue For
        End If


        ' -------------------------------------------------
        ' Сборка.
        ' -------------------------------------------------

        If occurrence.DefinitionDocumentType = _
            DocumentTypeEnum.kAssemblyDocumentObject Then

            process_model_for_generation( _
                occurrence, _
                "Assembly", _
                generation_candidates, _
                processed_model_paths, _
                controlled_code)


            ' -------------------------------------------------
            ' Рекурсивно спускаемся в подсборку.
            ' -------------------------------------------------

            If occurrence.SubOccurrences.Count > 0 Then

                collect_generation_candidates( _
                    occurrence.SubOccurrences, _
                    generation_candidates, _
                    processed_model_paths, _
                    controlled_code)

            End If


        ' -------------------------------------------------
        ' Деталь.
        ' -------------------------------------------------

        ElseIf occurrence.DefinitionDocumentType = _
            DocumentTypeEnum.kPartDocumentObject Then

            process_model_for_generation( _
                occurrence, _
                "Part", _
                generation_candidates, _
                processed_model_paths, _
                controlled_code)

        End If

    Next

End Sub


' =========================================================
' ПРОВЕРКА ОДНОЙ МОДЕЛИ
' =========================================================
'
' Здесь модель превращается в candidate.
'
' Candidate:
'
'   candidate(0) = полный путь модели
'   candidate(1) = имя файла модели
'   candidate(2) = тип модели
'
' Например:
'
'   0 = Y:\Project\PZ.0066.0005 - Корпус в сборе.iam
'   1 = PZ.0066.0005 - Корпус в сборе.iam
'   2 = Assembly
'

Sub process_model_for_generation( _
    occurrence As Object, _
    model_type As String, _
    generation_candidates As Object, _
    processed_model_paths As Object, _
    controlled_code As String)

    On Error GoTo handle_error


    ' -----------------------------------------------------
    ' Получаем документ модели.
    ' -----------------------------------------------------

    Dim model_document As Object
    model_document = occurrence.Definition.Document

    If model_document Is Nothing Then
        Exit Sub
    End If


    ' -----------------------------------------------------
    ' Получаем путь модели.
    ' -----------------------------------------------------

    Dim model_path As String
    model_path = model_document.FullFileName

    If model_path = "" Then
        Exit Sub
    End If


    ' -----------------------------------------------------
    ' Получаем имя файла.
    ' -----------------------------------------------------

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    Dim model_file_name As String
    model_file_name = fso.GetFileName(model_path)


    ' -----------------------------------------------------
    ' Проверяем наличие контролируемого шифра.
    ' -----------------------------------------------------

    If controlled_code = "" Then

        MsgBox( _
            "Не задан контролируемый шифр.", _
            vbCritical, _
            "Генератор чертежей")

        Exit Sub

    End If


    If InStr( _
        1, _
        model_file_name, _
        controlled_code, _
        vbTextCompare) = 0 Then

        Exit Sub

    End If


    ' -----------------------------------------------------
    ' Формируем уникальный ключ модели.
    ' -----------------------------------------------------

    Dim model_key As String
    model_key = LCase(model_path)


    ' -----------------------------------------------------
    ' Защита от повторов.
    ' -----------------------------------------------------

    If processed_model_paths.Exists(model_key) Then
        Exit Sub
    End If

    processed_model_paths.Add(model_key, True)


    ' -----------------------------------------------------
    ' Формируем запись кандидата.
    ' -----------------------------------------------------

    Dim candidate(2) As Object

    candidate(0) = model_path
    candidate(1) = model_file_name
    candidate(2) = model_type

    generation_candidates.Add(candidate)

    Exit Sub


handle_error:

    ' Ошибка одного компонента не должна останавливать
    ' весь обход дерева.

End Sub


' =========================================================
' ЛОГ DISCOVERY
' =========================================================

Sub log_generation_candidates( _
    generation_candidates As Object, _
    drawings_folder_path As String, _
    controlled_code As String)

    Dim log_message As String

    log_message = _
        "=== DISCOVERY ГЕНЕРАТОРА ЧЕРТЕЖЕЙ ===" & _
        vbCrLf & vbCrLf


    log_message = log_message & _
        "Контролируемый шифр: " & _
        controlled_code & _
        vbCrLf

    log_message = log_message & _
        "Папка чертежей: " & _
        drawings_folder_path & _
        vbCrLf & vbCrLf


    log_message = log_message & _
        "Найдено уникальных кандидатов: " & _
        generation_candidates.Count & _
        vbCrLf & vbCrLf


    If generation_candidates.Count = 0 Then

        log_message = log_message & _
            "Подходящих моделей не найдено." & _
            vbCrLf

    Else

        Dim candidate As Object
        Dim candidate_index As Long

        candidate_index = 1


        For Each candidate In generation_candidates

            log_message = log_message & _
                candidate_index & ". " & _
                candidate(2) & " | " & _
                candidate(1) & _
                vbCrLf

            log_message = log_message & _
                "   " & candidate(0) & _
                vbCrLf

            candidate_index = candidate_index + 1

        Next

    End If


    MsgBox( _
        log_message, _
        vbInformation, _
        "Генератор чертежей")

End Sub


' =========================================================
' ГЕНЕРАЦИЯ ЧЕРТЕЖЕЙ ПО СПИСКУ КАНДИДАТОВ
' =========================================================
'
' Это начало второй большой фазы pipeline:
'
'                 DISCOVERY
'                     |
'                     v
'             generation_candidates
'                     |
'                     v
'                 GENERATION
'
' Важно, что Generation уже не занимается поиском дерева.
' Он работает только с готовым списком кандидатов.
'
' Это разделение сильно упрощает дальнейшее развитие правила.
'

Sub generate_drawings_from_candidates( _
    generation_candidates As Object, _
    drawings_folder_path As String, _
    detail_template_path As String, _
    assembly_template_path As String)

    If generation_candidates Is Nothing Then Exit Sub

    If generation_candidates.Count = 0 Then

        MsgBox( _
            "Список кандидатов пуст. Новые чертежи не создавались.", _
            vbInformation, _
            "Генератор чертежей")

        Exit Sub

    End If


    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")


    ' -----------------------------------------------------
    ' Progress bar.
    ' -----------------------------------------------------

    Dim progress_bar As Object

    progress_bar = ThisApplication.CreateProgressBar( _
        False, _
        generation_candidates.Count, _
        "Генерация заготовок чертежей")


    ' -----------------------------------------------------
    ' Результаты Generation.
    '
    ' generated_files
    '     реально созданные файлы
    '
    ' skipped_files
    '     файл уже существовал
    '
    ' failed_files
    '     возникла ошибка
    ' -----------------------------------------------------

    Dim generated_files As Object
    generated_files = CreateObject( _
        "System.Collections.ArrayList")

    Dim skipped_files As Object
    skipped_files = CreateObject( _
        "System.Collections.ArrayList")

    Dim failed_files As Object
    failed_files = CreateObject( _
        "System.Collections.ArrayList")


    Dim candidate As Object
    Dim candidate_index As Integer

    candidate_index = 0


    On Error GoTo handle_common_error

    ThisApplication.ScreenUpdating = False


    ' -----------------------------------------------------
    ' Последовательно обрабатываем candidates.
    ' -----------------------------------------------------

    For Each candidate In generation_candidates

        candidate_index = candidate_index + 1


        ' -------------------------------------------------
        ' Обновляем progress bar.
        ' -------------------------------------------------

        If Not progress_bar Is Nothing Then

            progress_bar.Message = _
                "Обработка " & candidate_index & " из " & _
                generation_candidates.Count & ": " & _
                candidate(1)

            progress_bar.UpdateProgress

        End If


        ' -------------------------------------------------
        ' Генерируем один drawing.
        ' -------------------------------------------------

        generate_single_drawing( _
            candidate, _
            drawings_folder_path, _
            detail_template_path, _
            assembly_template_path, _
            fso, _
            generated_files, _
            skipped_files, _
            failed_files)

    Next


    ThisApplication.ScreenUpdating = True


    If Not progress_bar Is Nothing Then
        progress_bar.Close
    End If


    ' =====================================================
    ' REPORT
    ' =====================================================

    write_generation_report( _
        drawings_folder_path, _
        generated_files, _
        skipped_files, _
        failed_files)

    show_generation_result( _
        generated_files, _
        skipped_files, _
        failed_files, _
        drawings_folder_path)

    Exit Sub


handle_common_error:

    ThisApplication.ScreenUpdating = True

    If Not progress_bar Is Nothing Then
        progress_bar.Close
    End If

    MsgBox( _
        "Критическая ошибка генерации." & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Генератор чертежей")

End Sub


' =========================================================
' ГЕНЕРАЦИЯ ОДНОГО ФАЙЛА
' =========================================================
'
' Здесь происходит непосредственная materialization candidate.
'
' Pipeline одного файла:
'
'   candidate
'      |
'      v
'   drawing filename
'      |
'      v
'   target path
'      |
'      +-- exists --> SKIP
'      |
'      v
'   template
'      |
'      v
'   CopyFile
'      |
'      v
'   Open model
'      |
'      v
'   Open drawing
'      |
'      v
'   Create main view
'      |
'      v
'   Copy iProperties
'      |
'      v
'   Save
'      |
'      v
'   Close
'
' После успешного завершения candidate становится
' generated_file.
'

Sub generate_single_drawing( _
    candidate As Object, _
    drawings_folder_path As String, _
    detail_template_path As String, _
    assembly_template_path As String, _
    fso As Object, _
    generated_files As Object, _
    skipped_files As Object, _
    failed_files As Object)

    Dim model_path As String
    model_path = candidate(0)

    Dim model_file_name As String
    model_file_name = candidate(1)


    ' =====================================================
    ' Формируем имя будущего drawing.
    ' =====================================================

    Dim drawing_file_name As String

    drawing_file_name = get_drawing_file_name( _
        fso.GetBaseName(model_file_name), _
        candidate(2))


    ' =====================================================
    ' Формируем полный путь.
    ' =====================================================

    Dim drawing_path As String

    drawing_path = fso.BuildPath( _
        drawings_folder_path, _
        drawing_file_name)


    On Error GoTo handle_error


    ' =====================================================
    ' ЗАЩИТА ОТ ПОВТОРНОГО СОЗДАНИЯ
    ' =====================================================
    '
    ' Эта проверка повторяется здесь намеренно.
    '
    ' Discovery формирует список кандидатов.
    '
    ' Но между Discovery и Generation состояние файловой
    ' системы теоретически могло измениться.
    '
    ' Поэтому Generation самостоятельно защищается
    ' от существующего файла.
    '

    If fso.FileExists(drawing_path) Then

        skipped_files.Add(drawing_path)

        Exit Sub

    End If


    ' =====================================================
    ' ВЫБОР TEMPLATE
    ' =====================================================

    Dim template_path As String

    If candidate(2) = "Assembly" Then

        template_path = assembly_template_path

    Else

        template_path = detail_template_path

    End If


    ' -----------------------------------------------------
    ' Проверяем наличие template.
    ' -----------------------------------------------------

    If Not fso.FileExists(template_path) Then

        failed_files.Add( _
            drawing_path & _
            " | Не найден шаблон: " & _
            template_path)

        Exit Sub

    End If


    ' =====================================================
    ' MATERIALIZATION
    ' =====================================================
    '
    ' Здесь физически создаётся новый файл.
    '
    ' До этой строки Generation только вычислял,
    ' проверял и подготавливал.
    '
    ' CopyFile превращает template в конкретный
    ' целевой drawing.
    '

    fso.CopyFile( _
        template_path, _
        drawing_path, _
        False)


    ' =====================================================
    ' ОТКРЫТИЕ MODEL
    ' =====================================================

    Dim model_document As Object

    model_document = ThisApplication.Documents.Open( _
        model_path, _
        False)


    ' =====================================================
    ' ОТКРЫТИЕ DRAWING
    ' =====================================================

    Dim drawing_document As Object

    drawing_document = ThisApplication.Documents.Open( _
        drawing_path, _
        False)


    ' =====================================================
    ' СОЗДАНИЕ ОСНОВНОГО ВИДА
    ' =====================================================
    '
    ' На этом этапе template превращается из пустой
    ' заготовки в drawing, связанный с конкретной моделью.
    '

    create_drawing_main_view( _
        drawing_document, _
        model_document)


    ' =====================================================
    ' ПОЛУЧЕНИЕ МОДЕЛИ, НА КОТОРУЮ ССЫЛАЕТСЯ ВИД
    ' =====================================================
    '
    ' Берём ReferencedDocument именно из DrawingView.
    '
    ' Это даёт нам модель, которая реально используется
    ' созданным видом.
    '

    Dim view_model_document As Object

    view_model_document = _
        drawing_document.ActiveSheet.DrawingViews(1). _
        ReferencedDocumentDescriptor.ReferencedDocument


    ' =====================================================
    ' ПЕРЕНОС iPROPERTIES
    ' =====================================================

    copy_model_properties_to_drawing( _
        drawing_document, _
        view_model_document, _
        candidate(2))


    ' =====================================================
    ' СОХРАНЕНИЕ
    ' =====================================================

    drawing_document.Save


    ' =====================================================
    ' ЗАКРЫТИЕ
    ' =====================================================

    drawing_document.Close(True)


    ' =====================================================
    ' ФИКСИРУЕМ УСПЕШНОЕ СОЗДАНИЕ
    ' =====================================================

    generated_files.Add(drawing_path)

    Exit Sub


handle_error:

    failed_files.Add( _
        drawing_path & _
        " | " & _
        Err.Description)


    ' -----------------------------------------------------
    ' Если drawing успел открыться, закрываем его без
    ' сохранения повреждённого/незавершённого результата.
    ' -----------------------------------------------------

    On Error Resume Next

    If Not drawing_document Is Nothing Then
        drawing_document.Close(False)
    End If

    On Error GoTo 0

End Sub


' =========================================================
' ИМЯ ФАЙЛА ЧЕРТЕЖА
' =========================================================
'
' ВАЖНАЯ ЛОГИКА.
'
' Исходное имя модели:
'
'   PZ.0066.0005 - Корпус в сборе
'
' Для детали:
'
'   PZ.0066.0005 - Корпус в сборе.idw
'
'
' Для сборки:
'
'   PZ.0066.0005СБ - Корпус в сборе.idw
'
'
' Раньше логика делала:
'
'   drawing_base_name = model_base_name
'   drawing_base_name = drawing_base_name & "СБ"
'
' Поэтому получалось:
'
'   PZ.0066.0005 - Корпус в сбореСБ
'
' Что неверно.
'
' Теперь мы ищем разделитель " - " и вставляем "СБ"
' непосредственно перед ним.
'
' Если разделителя нет, используем безопасный fallback:
'
'   PZ.0066.0005СБ
'

Function get_drawing_file_name( _
    model_base_name As String, _
    model_type As String) As String

    Dim drawing_base_name As String

    drawing_base_name = model_base_name


    ' =====================================================
    ' СБОРКА
    ' =====================================================

    If model_type = "Assembly" Then

        Dim separator_position As Long

        separator_position = InStr( _
            1, _
            drawing_base_name, _
            " - ", _
            vbTextCompare)


        If separator_position > 0 Then

            ' -------------------------------------------------
            ' Есть стандартный разделитель:
            '
            ' PZ.0066.0005 - Корпус в сборе
            '
            ' separator_position указывает на начало " - ".
            '
            ' Поэтому:
            '
            ' LEFT:
            '   PZ.0066.0005
            '
            ' RIGHT:
            '   " - Корпус в сборе"
            '
            ' Получаем:
            '
            '   PZ.0066.0005СБ - Корпус в сборе
            ' -------------------------------------------------

            Dim number_part As String
            number_part = Left( _
                drawing_base_name, _
                separator_position - 1)


            Dim description_part As String
            description_part = Mid( _
                drawing_base_name, _
                separator_position)


            ' -------------------------------------------------
            ' Защита от повторного добавления СБ.
            ' -------------------------------------------------

            If UCase(Right(number_part, 2)) <> "СБ" Then

                number_part = number_part & "СБ"

            End If


            drawing_base_name = _
                number_part & _
                description_part

        Else

            ' -------------------------------------------------
            ' Если " - " отсутствует, просто добавляем СБ
            ' к концу основной части имени.
            ' -------------------------------------------------

            If UCase(Right(drawing_base_name, 2)) <> "СБ" Then

                drawing_base_name = _
                    drawing_base_name & "СБ"

            End If

        End If

    End If


    ' =====================================================
    ' ФИНАЛЬНОЕ РАСШИРЕНИЕ
    ' =====================================================

    get_drawing_file_name = _
        drawing_base_name & ".idw"

End Function


' =========================================================
' СОЗДАНИЕ ОСНОВНОГО ВИДА
' =========================================================
'
' НАСТРОЙКИ ВИДА:
'
'   Scale:
'       1 : 5
'
'       Inventor принимает масштаб как коэффициент:
'
'           1 : 1 -> 1.0
'           1 : 2 -> 0.5
'           1 : 5 -> 0.2
'
'
'   Position:
'
'       X = 25% ширины листа
'       Y = 50% высоты листа
'
'   То есть центр вида находится примерно здесь:
'
'                  ширина листа
'       +--------------------------------+
'       |                                |
'       |                                |
'       |       X                        |
'       |       |                        |
'       |       |                        |
'       |       v                        |
'       |      [ VIEW ]                  |
'       |                                |
'       |                                |
'       +--------------------------------+
'
'       X = Width * 0.25
'       Y = Height * 0.50
'
' AddBaseView использует view_point как точку размещения
' центра создаваемого вида.
'
' Поэтому нам не требуется отдельно вычислять ширину
' или высоту самой модели.
'

Sub create_drawing_main_view( _
    drawing_document As Object, _
    model_document As Object)

    Dim active_sheet As Object

    active_sheet = drawing_document.ActiveSheet


    ' -----------------------------------------------------
    ' Если в template уже существует вид, ничего не делаем.
    ' -----------------------------------------------------

    If active_sheet.DrawingViews.Count > 0 Then

        Exit Sub

    End If


    ' =====================================================
    ' TRANSIENT GEOMETRY
    ' =====================================================
    '
    ' Используем Inventor TransientGeometry для создания
    ' двумерной точки.
    '

    Dim transient_geometry As Object

    transient_geometry = _
        ThisApplication.TransientGeometry


    ' =====================================================
    ' РАССЧИТЫВАЕМ ПОЗИЦИЮ ВИДА
    ' =====================================================
    '
    ' X:
    '
    '   25% ширины листа
    '
    ' Y:
    '
    '   50% высоты листа
    '
    ' Таким образом центр вида располагается:
    '
    '   слева от геометрического центра листа,
    '   но вертикально по центру.
    '

    Dim view_x As Double
    view_x = active_sheet.Width * 0.25


    Dim view_y As Double
    view_y = active_sheet.Height * 0.5


    Dim view_point As Object

    view_point = transient_geometry.CreatePoint2d( _
        view_x, _
        view_y)


    ' =====================================================
    ' СОЗДАНИЕ ОСНОВНОГО ВИДА
    ' =====================================================
    '
    ' Масштаб:
    '
    '   0.2 = 1 : 5
    '
    ' Ориентация:
    '
    '   Default
    '
    ' Стиль:
    '
    '   Hidden Line Removed
    '

    Dim main_view As Object

    main_view = active_sheet.DrawingViews.AddBaseView( _
        model_document, _
        view_point, _
        0.2, _
        ViewOrientationTypeEnum.kDefaultViewOrientation, _
        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)


    ' =====================================================
    ' НАСТРОЙКИ LABEL
    ' =====================================================

    main_view.ShowLabel = False

End Sub


' =========================================================
' ПЕРЕНОС iPROPERTIES МОДЕЛИ В ЧЕРТЁЖ
' =========================================================

Sub copy_model_properties_to_drawing( _
    drawing_document As Object, _
    model_document As Object, _
    model_type As String)

    ' -----------------------------------------------------
    ' Читаем параметры модели.
    '
    ' Используем ту же систему параметров, что и в
    ' проверенном правиле пакетной синхронизации.
    ' -----------------------------------------------------

    Dim part_number As String
    part_number = clean_string( _
        get_parameters(model_document, "part_number"))

    Dim part_name As String
    part_name = clean_string( _
        get_parameters(model_document, "part_name"))

    Dim part_type As String
    part_type = clean_string( _
        get_parameters(model_document, "part_type"))

    Dim part_developer As String
    part_developer = clean_string( _
        get_parameters(model_document, "part_developer"))

    Dim developer_date As String
    developer_date = clean_string( _
        get_parameters(model_document, "developer_date"))

    Dim part_test As String
    part_test = clean_string( _
        get_parameters(model_document, "part_test"))

    Dim test_date As String
    test_date = clean_string( _
        get_parameters(model_document, "test_date"))

    Dim part_tech_control As String
    part_tech_control = clean_string( _
        get_parameters(model_document, "part_tech_control"))

    Dim tech_control_date As String
    tech_control_date = clean_string( _
        get_parameters(model_document, "tech_control_date"))

    Dim part_department_head As String
    part_department_head = clean_string( _
        get_parameters(model_document, "part_department_head"))

    Dim department_head_date As String
    department_head_date = clean_string( _
        get_parameters(model_document, "department_head_date"))

    Dim part_norms_control As String
    part_norms_control = clean_string( _
        get_parameters(model_document, "part_norms_control"))

    Dim norms_control_date As String
    norms_control_date = clean_string( _
        get_parameters(model_document, "norms_control_date"))

    Dim part_approved_by As String
    part_approved_by = clean_string( _
        get_parameters(model_document, "part_approved_by"))

    Dim part_approved_date As String
    part_approved_date = clean_string( _
        get_parameters(model_document, "part_approved_date"))


    ' -----------------------------------------------------
    ' В исходном рабочем правиле Company не проходит
    ' через clean_string().
    ' -----------------------------------------------------

    Dim part_company As String

    part_company = get_parameters( _
        model_document, _
        "part_company")


    ' =====================================================
    ' ЛОГИКА ТИПА ДОКУМЕНТА
    ' =====================================================

    Dim document_type As String

    document_type = ""


    If part_type = "Сборка" Then

        ' -------------------------------------------------
        ' Защита от повторного добавления "СБ".
        ' -------------------------------------------------

        If UCase(Right(part_number, 2)) <> "СБ" Then

            part_number = _
                part_number & "СБ"

        End If

        document_type = "Сборочный чертеж"

    End If


    ' =====================================================
    ' ЗАПИСЬ В iPROPERTIES ЧЕРТЕЖА
    ' =====================================================

    safe_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Part Number", _
        part_number)


    safe_properties_setter( _
        drawing_document, _
        "Inventor Summary Information", _
        "Title", _
        part_name)


    safe_properties_setter( _
        drawing_document, _
        "Inventor Summary Information", _
        "Author", _
        part_developer)


    safe_date_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Creation Time", _
        developer_date)


    safe_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Checked By", _
        part_test)


    safe_date_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Date Checked", _
        test_date)


    safe_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Tech Checked By", _
        part_tech_control)


    safe_date_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Tech Date Checked", _
        tech_control_date)


    safe_properties_setter( _
        drawing_document, _
        "Inventor Document Summary Information", _
        "Manager", _
        part_department_head)


    safe_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Authority", _
        part_department_head)


    safe_date_properties_setter( _
        drawing_document, _
        "Свойства ГОСТ", _
        "Руководитель Дата", _
        department_head_date)


    safe_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Engr Approved By", _
        part_norms_control)


    safe_date_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Engr Date Approved", _
        norms_control_date)


    safe_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Mfg Approved By", _
        part_approved_by)


    safe_date_properties_setter( _
        drawing_document, _
        "Design Tracking Properties", _
        "Mfg Date Approved", _
        part_approved_date)


    ' =====================================================
    ' ТИП ДОКУМЕНТА
    ' =====================================================

    If document_type <> "" Then

        safe_properties_setter( _
            drawing_document, _
            "Свойства ГОСТ", _
            "Тип документа", _
            document_type)

    End If


    ' =====================================================
    ' КОМПАНИЯ
    ' =====================================================

    safe_properties_setter( _
        drawing_document, _
        "Inventor Document Summary Information", _
        "Company", _
        part_company)

End Sub


' =========================================================
' БЕЗОПАСНОЕ ЧТЕНИЕ ПАРАМЕТРА МОДЕЛИ
' =========================================================

Function get_parameters( _
    model_document As Object, _
    parameter_name As String) As String

    On Error GoTo handle_error

    Dim parameters As Object
    parameters = model_document.ComponentDefinition.Parameters

    Dim parameter As Object
    parameter = parameters.Item(parameter_name)

    get_parameters = CStr(parameter.Value)

    Exit Function


handle_error:

    get_parameters = ""

End Function


' =========================================================
' БЕЗОПАСНАЯ ЗАПИСЬ iPROPERTY
' =========================================================

Sub safe_properties_setter( _
    drawing_document As Object, _
    property_set_name As String, _
    property_name As String, _
    property_value As String)

    On Error Resume Next

    Dim property_set As Object

    property_set = _
        drawing_document.PropertySets.Item( _
            property_set_name)


    Dim normalized_value As String

    normalized_value = clean_string(property_value)


    If normalized_value = "-" Then

        normalized_value = ""

    End If


    property_set.Item(property_name).Value = _
        normalized_value


    On Error GoTo 0

End Sub



' =========================================================
' БЕЗОПАСНАЯ ЗАПИСЬ DATE iPROPERTY
' =========================================================

Sub safe_date_properties_setter( _
    drawing_document As Object, _
    property_set_name As String, _
    property_name As String, _
    property_value As String)

    On Error Resume Next


    ' -----------------------------------------------------
    ' Получаем PropertySet.
    ' -----------------------------------------------------

    Dim property_set As Object

    property_set = _
        drawing_document.PropertySets.Item( _
            property_set_name)


    If property_set Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If


    ' -----------------------------------------------------
    ' Получаем конкретное свойство.
    ' -----------------------------------------------------

    Dim target_property As Object

    target_property = _
        property_set.Item(property_name)


    If target_property Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If


    ' =====================================================
    ' ПУСТАЯ ДАТА
    ' =====================================================
    '
    ' Inventor не использует "" как пустое значение
    ' для Date iProperty.
    '
    ' Специальное значение:
    '
    '     01.01.1601
    '
    ' означает "дата не задана".
    '
    ' В интерфейсе Inventor такая дата отображается
    ' как пустое значение.
    '

    If Trim(property_value) = "" Or _
       Trim(property_value) = "-" Then

        Dim empty_date As Date

        empty_date = DateSerial(1601, 1, 1)

        target_property.Value = empty_date

        On Error GoTo 0
        Exit Sub

    End If


    ' =====================================================
    ' ОБЫЧНАЯ ДАТА
    ' =====================================================
    '
    ' Здесь преобразуем строку модели в Date.
    '

    Dim converted_date As Date

    converted_date = CDate(property_value)

    target_property.Value = converted_date


    On Error GoTo 0

End Sub



' =========================================================
' ОЧИСТКА СТРОКИ
' =========================================================

Function clean_string(value As String) As String

    On Error GoTo handle_error


    If value Is Nothing Then

        clean_string = ""

        Exit Function

    End If


    Dim result As String

    result = value


    ' -----------------------------------------------------
    ' Удаляем управляющие символы.
    ' -----------------------------------------------------

    result = result.Replace(vbCr, "")
    result = result.Replace(vbLf, "")
    result = result.Replace(vbTab, "")


    ' -----------------------------------------------------
    ' Заменяем двойные кавычки.
    ' -----------------------------------------------------

    result = result.Replace(Chr(34), "'")


    ' -----------------------------------------------------
    ' Заменяем символы, которые считаются запрещёнными
    ' для файловых путей.
    ' -----------------------------------------------------

    Dim bad_characters As String

    bad_characters = "/\:*?<>|"


    Dim character_index As Integer

    For character_index = 1 To Len(bad_characters)

        result = result.Replace( _
            Mid(bad_characters, character_index, 1), _
            "_")

    Next


    ' -----------------------------------------------------
    ' Схлопываем двойные пробелы.
    ' -----------------------------------------------------

    Do While result.Contains("  ")

        result = result.Replace("  ", " ")

    Loop


    ' -----------------------------------------------------
    ' Убираем пробелы по краям.
    ' -----------------------------------------------------

    result = result.Trim()


    ' -----------------------------------------------------
    ' Ограничиваем длину.
    ' -----------------------------------------------------

    If result.Length > 250 Then

        result = Left(result, 250)

    End If


    clean_string = result

    Exit Function


handle_error:

    clean_string = ""

End Function


' =========================================================
' ОТЧЁТ РЕЗУЛЬТАТА
' =========================================================

Sub write_generation_report( _
    drawings_folder_path As String, _
    generated_files As Object, _
    skipped_files As Object, _
    failed_files As Object)

    On Error Resume Next

    Dim fso As Object

    fso = CreateObject( _
        "Scripting.FileSystemObject")


    Dim report_path As String

    report_path = fso.BuildPath( _
        drawings_folder_path, _
        "drawing_generator.log")


    Dim writer As Object

    writer = fso.CreateTextFile( _
        report_path, _
        True)


    writer.WriteLine( _
        "=== ГЕНЕРАЦИЯ ЗАГОТОВОК ЧЕРТЕЖЕЙ ===")

    writer.WriteLine( _
        "Создано: " & _
        generated_files.Count)

    writer.WriteLine( _
        "Пропущено: " & _
        skipped_files.Count)

    writer.WriteLine( _
        "Ошибок: " & _
        failed_files.Count)

    writer.WriteLine("")


    writer.WriteLine( _
        "=== СОЗДАННЫЕ ФАЙЛЫ ===")

    write_report_items( _
        writer, _
        generated_files)


    writer.WriteLine( _
        "=== СУЩЕСТВУЮЩИЕ ФАЙЛЫ ===")

    write_report_items( _
        writer, _
        skipped_files)


    writer.WriteLine( _
        "=== ОШИБКИ ===")

    write_report_items( _
        writer, _
        failed_files)


    writer.Close

    On Error GoTo 0

End Sub


' =========================================================
' ЗАПИСЬ ЭЛЕМЕНТОВ В ЛОГ
' =========================================================

Sub write_report_items( _
    writer As Object, _
    items As Object)

    Dim item As Object

    For Each item In items

        writer.WriteLine( _
            CStr(item))

    Next

End Sub


' =========================================================
' ПОКАЗ РЕЗУЛЬТАТА
' =========================================================

Sub show_generation_result( _
    generated_files As Object, _
    skipped_files As Object, _
    failed_files As Object, _
    drawings_folder_path As String)

    Dim result_message As String

    result_message = _
        "Генерация завершена." & _
        vbCrLf & _
        "Создано новых чертежей: " & _
        generated_files.Count & _
        vbCrLf & _
        "Пропущено существующих: " & _
        skipped_files.Count & _
        vbCrLf & _
        "Ошибок: " & _
        failed_files.Count & _
        vbCrLf & _
        "" & _
        vbCrLf & _
        "Список созданных файлов:" & _
        vbCrLf & _
        build_items_message(generated_files) & _
        vbCrLf & _
        "Лог: " & _
        CreateObject( _
            "Scripting.FileSystemObject"). _
            BuildPath( _
                drawings_folder_path, _
                "drawing_generator.log")


    MsgBox( _
        result_message, _
        vbInformation, _
        "Генератор чертежей")

End Sub


' =========================================================
' ПОСТРОЕНИЕ СПИСКА ДЛЯ MESSAGE BOX
' =========================================================

Function build_items_message( _
    items As Object) As String

    Dim result As String

    result = ""


    Dim item As Object

    For Each item In items

        result = _
            result & _
            CStr(item) & _
            vbCrLf

    Next


    If result = "" Then

        result = "нет"

    End If


    build_items_message = result

End Function

Option Explicit

'====================================================
' MAIN
'====================================================
Sub Main()


    '================================================
    ' 0. НАСТРОЙКА ВЫВОДА ФАЙЛА
    '================================================

    ' Показ данных внутри файла
    Dim show_metadata_in_file As Boolean
    show_metadata_in_file = False
    
    ' Показ количества в названии файла     
    Dim show_quantity_in_filename As Boolean
    show_quantity_in_filename = False

    ' Показ толщины металла в названии файла
    Dim show_leaf_thickness As Boolean
    show_leaf_thickness = True


    '================================================
    ' 1. НАСТРОЙКА ОБРАБОТКИ ОШИБОК
    '================================================

    Dim assembly_error As Boolean
    assembly_error = False

    Dim has_error As Boolean
    has_error = False

    Dim export_error As Boolean
    export_error = False

    On Error GoTo handle_common_error


    '================================================
    ' 2. ПРОВЕРКА СБОРКИ
    '================================================

    Call check_assembly(assembly_error)

    If assembly_error Then
        GoTo handle_assembly_error
    End If


    '================================================
    ' 3. ПУТИ
    '================================================

    ' Имя INI-файла для настройки DXF-вывода.
    Dim dxf_ini_name As String
    dxf_ini_name = "Под резку.ini"


    ' Полный путь до текущей сборки.
    Dim current_assembly_path As String
    current_assembly_path = _
        ThisApplication.ActiveDocument.FullFileName


    ' Сохраняем путь, чтобы после выполнения
    ' снова открыть исходную сборку.
    Dim current_assembly_path_tmp As String
    current_assembly_path_tmp = _
        current_assembly_path


    '================================================
    ' 4. ПОИСК КОРНЯ ПРОЕКТА
    '================================================

    Dim base_folder_path As String

    base_folder_path = _
        find_base_folder_path( _
            current_assembly_path)


    If base_folder_path = "" Then

        has_error = True
        GoTo handle_common_error

    End If


    '================================================
    ' 5. ПОИСК ПАПКИ DXF
    '================================================

    Dim dxf_folder_path As String

    dxf_folder_path = _
        find_dxf_folder_path( _
            base_folder_path)


    If dxf_folder_path = "" Then

        has_error = True
        GoTo handle_common_error

    End If


    '================================================
    ' 6. ПОИСК INI
    '================================================

    Dim ini_file_path As String

    ini_file_path = _
        find_ini_file_path( _
            dxf_folder_path, _
            dxf_ini_name)


    If ini_file_path = "" Then

        has_error = True
        GoTo handle_common_error

    End If


    '================================================
    ' 7. СОЗДАНИЕ СПИСКА СБОРОК
    '
    ' Формат:
    '
    ' [
    '   [
    '       имя сборки,
    '       [
    '           путь к детали,
    '           путь к детали
    '       ]
    '   ],
    '   ...
    ' ]
    '
    '================================================

    Dim assemblies_with_sheet_metal As Object

    assemblies_with_sheet_metal = _
        CreateObject( _
            "System.Collections.ArrayList")


    Dim isRoot As Boolean
    isRoot = True


    '================================================
    ' СЛОВАРЬ УЖЕ ОБРАБОТАННЫХ СБОРОК
    '================================================

    Dim addedAssemblies As Object

    addedAssemblies = _
        CreateObject( _
            "Scripting.Dictionary")


    '================================================
    ' 8. DISCOVERY
    '================================================

    Call collect_assemblies_and_sheet_metal( _
        ThisApplication.ActiveDocument.ComponentDefinition.Occurrences, _
        assemblies_with_sheet_metal, _
        isRoot, _
        ThisApplication.ActiveDocument, _
        addedAssemblies)


    '================================================
    ' ЛОГ СПИСКА
    '================================================

    Call LogAssembliesWithSheetMetal( _
        assemblies_with_sheet_metal)


    '================================================
    ' 9. EXPORT
    '================================================

    Call export_dxf_for_curerrent_assembly( _
        assemblies_with_sheet_metal, _
        dxf_folder_path, _
        ini_file_path, _
        export_error, _
        show_metadata_in_file, _
        show_quantity_in_filename, _
        show_leaf_thickness)


    If export_error Then
        GoTo handle_export_error
    End If


    '================================================
    ' УСПЕШНОЕ ЗАВЕРШЕНИЕ
    '================================================

    MsgBox( _
        "Программа выполнена успешно", _
        vbInformation, _
        "DXF Export")


    '================================================
    ' ВОЗВРАЩАЕМ ИСХОДНУЮ СБОРКУ
    '================================================

    Dim doc As Document

    doc = _
        ThisApplication.Documents.Open( _
            current_assembly_path_tmp, _
            True)

    doc.Activate


    Exit Sub


'====================================================
' ОБЩАЯ ОШИБКА
'====================================================

handle_common_error:

    MsgBox( _
        "Произошла системная ошибка:" & _
        vbCrLf & vbCrLf & _
        Err.Number & _
        " | " & _
        Err.Description, _
        vbCritical, _
        "Error")

    Exit Sub


'====================================================
' ОШИБКА СБОРКИ
'====================================================

handle_assembly_error:

    MsgBox( _
        "Open an assembly before running this rule.", _
        vbCritical, _
        "Error")

    Exit Sub


'====================================================
' ОШИБКА ЭКСПОРТА
'====================================================

handle_export_error:

    MsgBox( _
        "Во время экспорта произошли ошибки." & _
        vbCrLf & _
        "Некоторые детали могли быть экспортированы успешно.", _
        vbCritical, _
        "DXF Export Error")

    Exit Sub

End Sub



'====================================================
' ПРОВЕРКА АКТИВНОГО ДОКУМЕНТА
'====================================================
Sub check_assembly( _
    ByRef assembly_error As Boolean)

    assembly_error = False


    ' Нельзя использовать:
    '
    ' If ActiveDocument Is Nothing Or ...
    '
    ' потому что обычный Or не гарантирует
    ' короткое замыкание проверки.

    If ThisApplication.ActiveDocument Is Nothing Then

        assembly_error = True
        Exit Sub

    End If


    If ThisApplication.ActiveDocument.DocumentType <> _
        kAssemblyDocumentObject Then

        assembly_error = True
        Exit Sub

    End If

End Sub



'====================================================
' ПОИСК КОРНЯ ПРОЕКТА
'
' Ищем вверх от папки текущей сборки.
'
' Признаки корня:
'
'   0_Модели
'   1_Чертежи
'
'====================================================
Function find_base_folder_path( _
    current_assembly_path As String) As String

    find_base_folder_path = ""


    Dim fso As Object

    fso = _
        CreateObject( _
            "Scripting.FileSystemObject")


    ' current_assembly_path содержит путь
    ' до файла .iam.
    '
    ' Поэтому первым cur должен быть
    ' родитель этой папки.

    Dim cur As String

    cur = _
        fso.GetParentFolderName( _
            current_assembly_path)


    Dim project_root As String

    project_root = ""


    Do While cur <> ""


        '--------------------------------------------
        ' Проверяем признаки корня проекта
        '--------------------------------------------

        If _
            Dir( _
                cur & "\0_Модели", _
                vbDirectory) <> "" _
            And _
            Dir( _
                cur & "\1_Чертежи", _
                vbDirectory) <> "" Then

            project_root = cur
            Exit Do

        End If


        '--------------------------------------------
        ' Получаем родительскую папку
        '--------------------------------------------

        Dim parent_folder As String

        parent_folder = _
            fso.GetParentFolderName(cur)


        If parent_folder = "" Then
            Exit Do
        End If


        cur = parent_folder

    Loop


    '================================================
    ' РЕЗУЛЬТАТ
    '================================================

    If project_root = "" Then

        MsgBox( _
            "Project root not found." & _
            vbCrLf & vbCrLf & _
            "Expected folders:" & _
            vbCrLf & _
            "0_Модели" & _
            vbCrLf & _
            "1_Чертежи", _
            vbCritical, _
            "Error")

        Exit Function

    End If


    find_base_folder_path = project_root

End Function



'====================================================
' ПОИСК ПАПКИ DXF
'
'   <root>\1_Чертежи\3_DXF
'
'====================================================
Function find_dxf_folder_path( _
    base_folder_path As String) As String

    find_dxf_folder_path = ""


    Dim dxf_folder As String

    dxf_folder = _
        base_folder_path & _
        "\1_Чертежи\3_DXF"


    If Dir( _
        dxf_folder, _
        vbDirectory) = "" Then

        MsgBox( _
            "DXF folder not found." & _
            vbCrLf & vbCrLf & _
            "Expected:" & _
            vbCrLf & _
            dxf_folder, _
            vbCritical, _
            "Error")

        Exit Function

    End If


    find_dxf_folder_path = dxf_folder

End Function



'====================================================
' ПОИСК INI-ФАЙЛА
'====================================================
Function find_ini_file_path( _
    dxf_folder_path As String, _
    dxf_ini_name As String) As String

    find_ini_file_path = ""


    Dim ini_path As String

    ini_path = _
        dxf_folder_path & _
        "\" & _
        dxf_ini_name


    If Dir(ini_path) = "" Then

        MsgBox( _
            "Project file not found." & _
            vbCrLf & vbCrLf & _
            "Expected file:" & _
            vbCrLf & _
            ini_path, _
            vbCritical, _
            "Error")

        Exit Function

    End If


    find_ini_file_path = ini_path

End Function



'====================================================
' DISCOVERY
'
' Собирает:
'
'   сборка
'       ↓
'   уникальные детали листового металла
'
' и рекурсивно проходит подсборки.
'
'====================================================
Sub collect_assemblies_and_sheet_metal( _
    occurrences As Object, _
    assemblies As Object, _
    isRoot As Boolean, _
    rootDoc As Object, _
    addedAssemblies As Object)


    '================================================
    ' 1. ПРОВЕРКА OCCURRENCES
    '================================================

    If occurrences Is Nothing Then
        Exit Sub
    End If


    '================================================
    ' 2. СЛОВАРЬ
    '================================================

    If addedAssemblies Is Nothing Then

        addedAssemblies = _
            CreateObject( _
                "Scripting.Dictionary")

    End If


    '================================================
    ' 3. ОПРЕДЕЛЯЕМ ТЕКУЩУЮ СБОРКУ
    '================================================

    Dim asmPath As String
    Dim asmName As String

    asmPath = ""
    asmName = ""


    If isRoot And Not rootDoc Is Nothing Then

        asmPath = rootDoc.FullFileName
        asmName = rootDoc.DisplayName

    Else

        If occurrences.Count = 0 Then
            Exit Sub
        End If


        On Error GoTo skip_current_assembly


        asmPath = _
            occurrences(1).Parent.Document.FullFileName

        asmName = _
            occurrences(1).Parent.Document.DisplayName

        On Error GoTo 0

    End If


    '================================================
    ' 4. ПРОВЕРКА ПУТИ
    '================================================

    If asmPath = "" Then
        Exit Sub
    End If


    '================================================
    ' ЗАЩИТА ОТ ПОВТОРНОЙ ОБРАБОТКИ
    '================================================

    If addedAssemblies.Exists(asmPath) Then
        Exit Sub
    End If


    Call addedAssemblies.Add( _
        asmPath, _
        True)



    '================================================
    ' СОЗДАЁМ ЗАПИСЬ СБОРКИ
    '
    ' currentEntry(0) = имя
    ' currentEntry(1) = список уникальных деталей
    ' currentEntry(2) = количество вхождений каждой детали
    '
    '================================================

    Dim currentEntry(2) As Object

    currentEntry(0) = asmName

    currentEntry(1) = _
        CreateObject( _
            "System.Collections.ArrayList")

    currentEntry(2) = _
        CreateObject( _
            "Scripting.Dictionary")


    '================================================
    ' СКАНИРУЕМ ДЕТАЛИ ТЕКУЩЕГО УРОВНЯ
    '================================================

    Dim occ As Object

    For Each occ In occurrences

        If occ Is Nothing Then
            GoTo continue_sheet_scan
        End If


        Dim occurrence_definition As Object

        On Error Resume Next

        Err.Clear

        occurrence_definition = _
            occ.Definition

        If Err.Number <> 0 Then

            Err.Clear
            On Error GoTo 0

            GoTo continue_sheet_scan

        End If

        On Error GoTo 0


        If occurrence_definition Is Nothing Then
            GoTo continue_sheet_scan
        End If


        If TypeName( _
            occurrence_definition) = _
            "SheetMetalComponentDefinition" Then


            Dim fullPath As String

            On Error Resume Next

            Err.Clear

            fullPath = _
                occurrence_definition.Document.FullFileName

            If Err.Number <> 0 Then

                Err.Clear
                On Error GoTo 0

                GoTo continue_sheet_scan

            End If

            On Error GoTo 0


            If fullPath <> "" Then

                '--------------------------------------------
                ' Считаем фактическое количество вхождений
                ' данной детали в текущей сборке.
                '--------------------------------------------

                If currentEntry(2).Exists(fullPath) Then

                    currentEntry(2)(fullPath) = _
                        currentEntry(2)(fullPath) + 1

                Else

                    currentEntry(2).Add( _
                        fullPath, _
                        1)

                End If


                '--------------------------------------------
                ' Добавляем только уникальный путь.
                '--------------------------------------------

                If Not currentEntry(1).Contains(fullPath) Then

                    Call currentEntry(1).Add(fullPath)

                End If

            End If

        End If


continue_sheet_scan:

    Next


    '================================================
    ' ДОБАВЛЯЕМ СБОРКУ В СПИСОК
    '================================================

    If currentEntry(1).Count > 0 Then

        Call assemblies.Add( _
            currentEntry)

    End If


    '================================================
    ' РЕКУРСИЯ ПО ПОДСБОРКАМ
    '================================================

    For Each occ In occurrences

        If occ Is Nothing Then
            GoTo continue_recursion
        End If


        Dim recursion_definition As Object

        On Error Resume Next

        Err.Clear

        recursion_definition = _
            occ.Definition

        If Err.Number <> 0 Then

            Err.Clear
            On Error GoTo 0

            GoTo continue_recursion

        End If

        On Error GoTo 0


        If recursion_definition Is Nothing Then
            GoTo continue_recursion
        End If


        If TypeName( _
            recursion_definition) = _
            "AssemblyComponentDefinition" Then


            Dim nested_occurrences As Object

            On Error Resume Next

            Err.Clear

            nested_occurrences = _
                recursion_definition.Occurrences

            If Err.Number <> 0 Then

                Err.Clear
                On Error GoTo 0

                GoTo continue_recursion

            End If

            On Error GoTo 0


            If nested_occurrences Is Nothing Then
                GoTo continue_recursion
            End If


            Call collect_assemblies_and_sheet_metal( _
                nested_occurrences, _
                assemblies, _
                False, _
                Nothing, _
                addedAssemblies)

        End If


continue_recursion:

    Next


    Exit Sub


'====================================================
' Если не удалось определить текущую сборку
'====================================================

skip_current_assembly:

    On Error GoTo 0
    Exit Sub

End Sub



'====================================================
' ЭКСПОРТ ВСЕХ ДЕТАЛЕЙ В DXF
'
' Pipeline:
'
'   assemblies_with_sheet_metal
'       ↓
'   сборка
'       ↓
'   детали листового металла
'       ↓
'   export_single_sheet_metal_part()
'
'====================================================
Sub export_dxf_for_curerrent_assembly( _
    assemblies_with_sheet_metal As Object, _
    dxf_folder_path As String, _
    ini_file_path As String, _
    ByRef export_error As Boolean, _
    show_metadata_in_file As Boolean, _
    show_quantity_in_filename As Boolean, _
    show_leaf_thickness As Boolean)


    export_error = False


    Dim assembly_entry As Object
    Dim sheet_parts As Object
    Dim part_path As Object


    '================================================
    ' ПРОГРЕССБАР
    '================================================

    Dim total_parts As Long
    Dim processed_parts As Long

    total_parts = 0
    processed_parts = 0


    '================================================
    ' ПЕРВЫЙ ПРОХОД:
    ' СЧИТАЕМ ОБЩЕЕ КОЛИЧЕСТВО УНИКАЛЬНЫХ ДЕТАЛЕЙ
    '================================================

    For Each assembly_entry In _
        assemblies_with_sheet_metal

        sheet_parts = _
            assembly_entry(1)

        total_parts = _
            total_parts + sheet_parts.Count

    Next


    Dim progress_bar As Object

    If total_parts > 0 Then

        progress_bar = _
            ThisApplication.CreateProgressBar( _
                False, _
                total_parts, _
                "DXF Export", _
                False)

        progress_bar.Message = _
            "Подготовка экспорта..."

    End If


    '================================================
    ' 1. ПРОХОД ПО СБОРКАМ
    '================================================

    For Each assembly_entry In _
        assemblies_with_sheet_metal


        '================================================
        ' 2. ПОДГОТОВКА ИМЕНИ ПАПКИ СБОРКИ
        '================================================

        Dim clean_sub_name As String

        clean_sub_name = _
            CStr(assembly_entry(0))


        '--------------------------------------------
        ' Убираем расширение
        '--------------------------------------------

        If InStrRev( _
            clean_sub_name, _
            ".") > 0 Then

            clean_sub_name = _
                Left( _
                    clean_sub_name, _
                    InStrRev( _
                        clean_sub_name, _
                        ".") - 1)

        End If


        '--------------------------------------------
        ' Недопустимые символы
        '--------------------------------------------

        Const INVALID_FILE_NAME_CHARACTERS As String = _
            "/\[]*?<>|:"


        Dim character_index As Long

        For character_index = 1 To _
            Len(INVALID_FILE_NAME_CHARACTERS)

            clean_sub_name = _
                Replace( _
                    clean_sub_name, _
                    Mid( _
                        INVALID_FILE_NAME_CHARACTERS, _
                        character_index, _
                        1), _
                    "_")

        Next


        '--------------------------------------------
        ' Убираем последний числовой суффикс
        '--------------------------------------------

        Dim name_parts() As String

        name_parts = _
            Split( _
                clean_sub_name, _
                "_")


        If UBound(name_parts) > 0 Then

            If IsNumeric( _
                name_parts(UBound(name_parts))) Then

                ReDim Preserve _
                    name_parts( _
                        UBound(name_parts) - 1)

                clean_sub_name = _
                    Join( _
                        name_parts, _
                        "_")

            End If

        End If


        '================================================
        ' 3. СОЗДАЁМ ПАПКУ СБОРКИ
        '================================================

        Dim subfolder_path As String

        subfolder_path = _
            dxf_folder_path & _
            "\" & _
            clean_sub_name


        If Dir( _
            subfolder_path, _
            vbDirectory) = "" Then

            MkDir(subfolder_path)

        End If


        '================================================
        ' 4. СПИСОК ДЕТАЛЕЙ
        '================================================

        sheet_parts = _
            assembly_entry(1)


        '================================================
        ' 5. ЭКСПОРТ ДЕТАЛЕЙ
        '================================================

        Dim sheet_parts_count As Long

        For Each part_path In sheet_parts

            ' Количество именно этой детали
            ' в текущей сборке.
            sheet_parts_count = _
                assembly_entry(2)(CStr(part_path))


            If total_parts > 0 Then

                progress_bar.Message = _
                    "Экспорт детали " & _
                    CStr(processed_parts + 1) & _
                    " из " & _
                    CStr(total_parts)

            End If


            Call export_single_sheet_metal_part( _
                CStr(part_path), _
                subfolder_path, _
                ini_file_path, _
                sheet_parts_count, _
                export_error, _
                show_metadata_in_file, _
                show_quantity_in_filename, _
                show_leaf_thickness)


            '================================================
            ' ОБНОВЛЯЕМ ПРОГРЕСС
            '================================================

            processed_parts = _
                processed_parts + 1


            If total_parts > 0 Then

                progress_bar.UpdateProgress

            End If

        Next


    Next


    '================================================
    ' ЕСЛИ БЫЛИ ОШИБКИ
    '================================================

    If export_error Then

        If total_parts > 0 Then
            progress_bar.Close
        End If

        Exit Sub

    End If


    '================================================
    ' ЗАКРЫВАЕМ ПРОГРЕССБАР
    '================================================

    If total_parts > 0 Then
        progress_bar.Close
    End If


    MsgBox( _
        "DXF Export Completed", _
        vbInformation, _
        "DXF Export")

End Sub




'====================================================
' УСТАНОВКА ФЛАГА DXF EXPORT
'
' Создаёт параметр dxf_export_started,
' если он отсутствует.
'
' Значение:
'
'   True  - деталь находится в процессе DXF export
'   False - обычное состояние
'
'====================================================
Sub set_dxf_export_flag_inside_model( _
    part_document As Object, _
    value As Boolean)


    Dim parameters As Object

    parameters = _
        part_document.ComponentDefinitions(1).Parameters


    Dim user_parameters As Object

    user_parameters = _
        parameters.UserParameters


    Dim dxf_parameter As Object

    dxf_parameter = Nothing


    '================================================
    ' ИЩЕМ СУЩЕСТВУЮЩИЙ ПАРАМЕТР
    '================================================

    On Error Resume Next

    Err.Clear

    dxf_parameter = _
        user_parameters.Item( _
            "dxf_export_started")

    On Error GoTo 0


    '================================================
    ' СОЗДАЁМ, ЕСЛИ НЕ СУЩЕСТВУЕТ
    '================================================

    If dxf_parameter Is Nothing Then

        dxf_parameter = _
            user_parameters.AddByValue( _
                "dxf_export_started", _
                value, _
                UnitsTypeEnum.kBooleanUnits)

    Else

        dxf_parameter.Value = value

    End If


    '================================================
    ' ПРИНУДИТЕЛЬНО ПЕРЕСТРАИВАЕМ МОДЕЛЬ при изменении флага
    '================================================

    part_document.Update2(True)

End Sub



'====================================================
' ЭКСПОРТ ОДНОЙ ДЕТАЛИ ЛИСТОВОГО МЕТАЛЛА
'
' Формат DXF:
'
'   имя детали_n шт_t мм.dxf
'
' Например:
'
'   Кронштейн_3 шт_2 мм.dxf
'
'====================================================
Sub export_single_sheet_metal_part( _
    part_path As String, _
    output_folder_path As String, _
    ini_file_path As String, _
    sheet_parts_count As Long, _
    ByRef export_error As Boolean, _
    show_metadata_in_file As Boolean, _
    show_quantity_in_filename As Boolean, _
    show_leaf_thickness As Boolean)



    Dim part_document As Object
    Dim sheet_metal_definition As Object
    Dim flat_pattern As Object
    Dim drawing_document As Object


    ' Используется для понятного сообщения,
    ' если ошибка произошла не из-за исключения,
    ' а из-за отсутствующего объекта.
    Dim error_stage As String

    error_stage = "START"


    On Error GoTo handle_part_error


    '================================================
    ' 1. ОТКРЫВАЕМ IPT
    '================================================

    error_stage = "OPEN PART"


    part_document = _
        ThisApplication.Documents.Open( _
            part_path, _
            False)


    If part_document Is Nothing Then
        GoTo handle_part_error
    End If


    ' УСТАНАВЛИВАЕМ ФЛАГ DXF EXPORT


    error_stage = "SET DXF EXPORT FLAG"

    Call set_dxf_export_flag_inside_model( _
        part_document, _
        True)


    '================================================
    ' 2. ПОЛУЧАЕМ DEFINITION
    '================================================

    error_stage = _
        "GET COMPONENT DEFINITION"


    sheet_metal_definition = _
        part_document.ComponentDefinitions(1)


    If sheet_metal_definition Is Nothing Then
        GoTo handle_part_error
    End If


    '================================================
    ' 3. ПРОВЕРЯЕМ SHEET METAL
    '================================================

    error_stage = _
        "CHECK SHEET METAL DEFINITION"


    If TypeName( _
        sheet_metal_definition) <> _
        "SheetMetalComponentDefinition" Then

        GoTo handle_part_error

    End If


    '================================================
    ' 4. FLAT PATTERN
    '================================================

    error_stage = _
        "GET FLAT PATTERN"


    If sheet_metal_definition.HasFlatPattern Then

        ' Развёртка уже существует.
        flat_pattern = _
            sheet_metal_definition.FlatPattern

    Else

        '--------------------------------------------
        ' Создаём развёртку.
        '--------------------------------------------

        error_stage = _
            "UNFOLD"


        sheet_metal_definition.Unfold()


        '--------------------------------------------
        ' Выходим из режима редактирования.
        '--------------------------------------------

        error_stage = _
            "EXIT FLAT PATTERN EDIT"


        sheet_metal_definition.FlatPattern.ExitEdit()


        flat_pattern = _
            sheet_metal_definition.FlatPattern

    End If


    If flat_pattern Is Nothing Then
        GoTo handle_part_error
    End If


    '================================================
    ' 5. СОЗДАЁМ DRAWING
    '================================================

    error_stage = _
        "CREATE DRAWING"


    drawing_document = _
        ThisApplication.Documents.Add( _
            DocumentTypeEnum.kDrawingDocumentObject, _
            , _
            False)


    If drawing_document Is Nothing Then
        GoTo handle_part_error
    End If


    '================================================
    ' 6. НАСТРОЙКА ЕДИНИЦ DRAWING
    '================================================

    error_stage = _
        "SET DRAWING UNITS"


    drawing_document.UnitsOfMeasure.LengthUnits = _
        UnitsTypeEnum.kMillimeterLengthUnits


    '================================================
    ' 7. TRANSIENT GEOMETRY
    '================================================

    Dim transient_geometry As Object

    error_stage = _
        "GET TRANSIENT GEOMETRY"


    transient_geometry = _
        ThisApplication.TransientGeometry


    If transient_geometry Is Nothing Then
        GoTo handle_part_error
    End If


    '================================================
    ' 8. BASE VIEW OPTIONS
    '================================================

    Dim base_view_options As Object

    error_stage = _
        "CREATE BASE VIEW OPTIONS"


    base_view_options = _
        ThisApplication.TransientObjects _
            .CreateNameValueMap()


    If base_view_options Is Nothing Then
        GoTo handle_part_error
    End If


    Call base_view_options.Add( _
        "SheetMetalFoldedModel", _
        False)


    '================================================
    ' 9. ACTIVE SHEET
    '================================================

    Dim sheet As Object

    error_stage = _
        "GET ACTIVE SHEET"


    sheet = _
        drawing_document.ActiveSheet


    If sheet Is Nothing Then
        GoTo handle_part_error
    End If


    '================================================
    ' 10. СОЗДАЁМ FLAT PATTERN VIEW
    '================================================

    Dim drawing_view As Object

    error_stage = _
        "CREATE BASE VIEW"


    drawing_view = _
        sheet.DrawingViews.AddBaseView( _
            part_document, _
            transient_geometry.CreatePoint2d(0, 0), _
            1, _
            ViewOrientationTypeEnum.kDefaultViewOrientation, _
            DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
            "FlatPattern", _
            , _
            base_view_options)


    If drawing_view Is Nothing Then
        GoTo handle_part_error
    End If


    '================================================
    ' 11. ПОЛУЧАЕМ ТОЛЩИНУ
    '================================================

    error_stage = _
        "GET MATERIAL THICKNESS"


    Dim thickness_mm As Double

    thickness_mm = _
        sheet_metal_definition.Thickness.Value * 10


    '================================================
    ' 12. ФОРМАТИРУЕМ ТОЛЩИНУ
    '================================================

    Dim thickness_string As String
    Dim dot_position As Long


    thickness_string = _
        CStr(thickness_mm)


    dot_position = _
        InStr( _
            thickness_string, _
            ".")


    If dot_position = 0 Then

        dot_position = _
            InStr( _
                thickness_string, _
                ",")

    End If


    If dot_position > 0 Then

        thickness_string = _
            Left( _
                thickness_string, _
                dot_position + 2)

    End If


    '================================================
    ' 13. ИМЯ ВИДА
    '================================================

    error_stage = _
        "SET VIEW NAME"


    drawing_view.Name = _
        part_document.DisplayName & _
        " - " & _
        thickness_string & _
        " мм"


    ' Показывать подпись вида или нет.
    drawing_view.ShowLabel = show_metadata_in_file


    '================================================
    ' 14. УДАЛЯЕМ ЛИНИИ СГИБА
    '================================================

    error_stage = _
        "REMOVE BEND LINES"


    Call RemoveBendLines( _
        drawing_view, _
        flat_pattern)


    '================================================
    ' 15. ПОЛУЧАЕМ ИМЯ IPT
    '================================================

    Dim part_file_name As String


    part_file_name = _
        Mid( _
            part_path, _
            InStrRev( _
                part_path, _
                "\") + 1)


    '--------------------------------------------
    ' Убираем расширение
    '--------------------------------------------

    If InStrRev( _
        part_file_name, _
        ".") > 0 Then

        part_file_name = _
            Left( _
                part_file_name, _
                InStrRev( _
                    part_file_name, _
                    ".") - 1)

    End If


    '================================================
    ' 16. ОЧИЩАЕМ ИМЯ ФАЙЛА
    '================================================

    Const INVALID_FILE_NAME_CHARACTERS As String = _
        "/\[]*?<>|:"


    Dim character_index As Long


    For character_index = 1 To _
        Len(INVALID_FILE_NAME_CHARACTERS)

        part_file_name = _
            Replace( _
                part_file_name, _
                Mid( _
                    INVALID_FILE_NAME_CHARACTERS, _
                    character_index, _
                    1), _
                "_")

    Next


    '================================================
    ' 17. ФОРМИРУЕМ ИМЯ DXF
    '
    ' Форматы на выбор по флагам show_quantity_in_filename и show_leaf_thickness:
    '   Имя.dxf
    '   Имя_2 мм.dxf
    '   Имя_3 шт_2 мм.dxf
    '
    '================================================

    Dim dxf_file_name As String


    ' Шаг 1: Базовое состояние — имя детали
    dxf_file_name = part_file_name

    ' Шаг 2: Проверяем флаг количества
    If show_quantity_in_filename Then
        dxf_file_name = dxf_file_name & "_" & CStr(sheet_parts_count) & " шт"
    End If

    ' Шаг 3: Проверяем флаг толщины
    If show_leaf_thickness Then
        dxf_file_name = dxf_file_name & "_" & thickness_string & " мм"
    End If

    ' Шаг 4: Финальное состояние — добавляем расширение
    dxf_file_name = dxf_file_name & ".dxf"

    '================================================
    ' 18. ФОРМИРУЕМ ПОЛНЫЙ ПУТЬ
    '================================================

    Dim dxf_file_path As String


    dxf_file_path = _
        output_folder_path & _
        "\" & _
        dxf_file_name


    '================================================
    ' 19. СОХРАНЯЕМ DXF
    '================================================

    error_stage = _
        "SAVE DXF"


    Call SaveDXF( _
        drawing_document, _
        dxf_file_path, _
        ini_file_path)


    '================================================
    ' ЗАКРЫВАЕМ DRAWING
    '================================================

    error_stage = _
        "CLOSE DRAWING"


    drawing_document.Close()


    drawing_document = Nothing


    '================================================
    ' ЗАКРЫВАЕМ IPT
    '================================================


    ' Устанавливаем флаг экспорта DXF в false

    error_stage = "RESET DXF EXPORT FLAG"

    Call set_dxf_export_flag_inside_model( _
        part_document, _
        False)


    error_stage = _
        "CLOSE PART"


    part_document.Close(True)


    part_document = Nothing


    '================================================
    ' SUCCESS
    '================================================

    Exit Sub


'====================================================
' ОБРАБОТКА ОШИБКИ
'====================================================

handle_part_error:


    export_error = True


    Dim current_error_number As Long
    Dim current_error_description As String


    current_error_number = _
        Err.Number


    current_error_description = _
        Err.Description


    ' Если переход произошёл через GoTo
    ' без реальной ошибки VBA, Err может быть 0.
    If current_error_number = 0 Then

        current_error_description = _
            "Operation failed without VBA error description."

    End If


    MsgBox( _
        "Ошибка экспорта детали:" & _
        vbCrLf & vbCrLf & _
        part_path & _
        vbCrLf & vbCrLf & _
        "Этап: " & _
        error_stage & _
        vbCrLf & _
        "Error " & _
        CStr(current_error_number) & _
        ": " & _
        current_error_description, _
        vbCritical, _
        "DXF Export Error")


    '================================================
    ' CLEANUP
    '================================================

    On Error Resume Next


    ' Принудительный сброс флага экспорта DXF в модели детали

    If Not part_document Is Nothing Then

        Call set_dxf_export_flag_inside_model( _
            part_document, _
            False)

    End If


    ' Закрываем док

    If Not drawing_document Is Nothing Then

        drawing_document.Close(False)

    End If


    If Not part_document Is Nothing Then

        part_document.Close(False)

    End If


    On Error GoTo 0


End Sub



'====================================================
' УБИРАЕМ ЛИНИИ СГИБА
'====================================================
Sub RemoveBendLines( _
    oView As DrawingView, _
    oFlatPattern As FlatPattern)


    Dim oBendEdgesUp As Edges

    oBendEdgesUp = _
        oFlatPattern.GetEdgesOfType( _
            FlatPatternEdgeTypeEnum.kBendUpFlatPatternEdge)


    Dim oBendEdgesDown As Edges

    oBendEdgesDown = _
        oFlatPattern.GetEdgesOfType( _
            FlatPatternEdgeTypeEnum.kBendDownFlatPatternEdge)


    Dim oEdge As Edge
    Dim oCurve As DrawingCurve
    Dim oSegment As DrawingCurveSegment


    '================================================
    ' BEND UP
    '================================================

    For Each oEdge In oBendEdgesUp

        For Each oCurve In _
            oView.DrawingCurves(oEdge)

            For Each oSegment In _
                oCurve.Segments

                oSegment.Visible = False

            Next

        Next

    Next


    '================================================
    ' BEND DOWN
    '================================================

    For Each oEdge In oBendEdgesDown

        For Each oCurve In _
            oView.DrawingCurves(oEdge)

            For Each oSegment In _
                oCurve.Segments

                oSegment.Visible = False

            Next

        Next

    Next

End Sub



'====================================================
' СОХРАНЕНИЕ DRAWING В DXF
'====================================================
Sub SaveDXF( _
    oDrawing As DrawingDocument, _
    oFileName As String, _
    oIniFile As String)


    '================================================
    ' DXF TRANSLATOR
    '================================================

    Dim DXFAddIn As TranslatorAddIn

    DXFAddIn = _
        ThisApplication.ApplicationAddIns.ItemById( _
            "{C24E3AC4-122E-11D5-8E91-0010B541CD80}")


    '================================================
    ' TRANSLATION CONTEXT
    '================================================

    Dim oContext As TranslationContext

    oContext = _
        ThisApplication.TransientObjects _
            .CreateTranslationContext


    oContext.Type = _
        IOMechanismEnum.kFileBrowseIOMechanism


    '================================================
    ' OPTIONS
    '================================================

    Dim oOptions As NameValueMap

    oOptions = _
        ThisApplication.TransientObjects _
            .CreateNameValueMap


    '================================================
    ' DATA MEDIUM
    '================================================

    Dim oDataMedium As DataMedium

    oDataMedium = _
        ThisApplication.TransientObjects _
            .CreateDataMedium


    '================================================
    ' INI
    '================================================

    If DXFAddIn.HasSaveCopyAsOptions( _
        oDrawing, _
        oContext, _
        oOptions) Then

        oOptions.Value( _
            "Export_Acad_IniFile") = _
            oIniFile

    End If


    '================================================
    ' OUTPUT FILE
    '================================================

    oDataMedium.FileName = _
        oFileName


    '================================================
    ' SAVE
    '================================================

    DXFAddIn.SaveCopyAs( _
        oDrawing, _
        oContext, _
        oOptions, _
        oDataMedium)

End Sub



'====================================================
' СОЗДАНИЕ TXT
'====================================================
Sub CreateTXT( _
    oText As String, _
    oFileName As String)

    Dim oTxtWriter As Object

    oTxtWriter = _
        CreateObject( _
            "Scripting.FileSystemObject") _
        .CreateTextFile( _
            oFileName, _
            True)


    oTxtWriter.WriteLine(oText)
    oTxtWriter.Close

End Sub



'====================================================
' ЛОГИРОВАНИЕ assemblies_root
'====================================================
Sub LogAssembliesRoot( _
    assemblies_root As Object)


    Dim msg As String

    msg = _
        "=== assemblies_root ===" & _
        vbCrLf


    Dim assembly As Object


    For Each assembly In assemblies_root

        msg = _
            msg & _
            "Name: " & _
            assembly(0) & _
            " | Path: " & _
            assembly(1) & _
            vbCrLf

    Next


    MsgBox( _
        msg, _
        vbInformation, _
        "assemblies_root log")

End Sub



'====================================================
' ЛОГИРОВАНИЕ assemblies_with_sheet_metal
'====================================================
Sub LogAssembliesWithSheetMetal( _
    assemblies_with_sheet_metal As Object)


    Dim msg As String

    msg = _
        "=== assemblies_with_sheet_metal ===" & _
        vbCrLf


    Dim entry As Object
    Dim sheetPart As Object


    For Each entry In _
        assemblies_with_sheet_metal


        msg = _
            msg & _
            "Subassembly: " & _
            entry(0) & _
            vbCrLf


        For Each sheetPart In entry(1)

            msg = _
                msg & _
                "  sheet part: " & _
                sheetPart & _
                vbCrLf

        Next

    Next


    MsgBox( _
        msg, _
        vbInformation, _
        "assemblies_with_sheet_metal log")

End Sub

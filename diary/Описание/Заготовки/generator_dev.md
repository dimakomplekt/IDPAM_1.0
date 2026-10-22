# Описание разрабатываемого макроса


___

0. Правила разработки


## Главные правила

      1) snake_case в нейминге
      2) суперпонятные имена (пусть и длинные)
      3) Разделение на отдельные функции каких-то важных частей пайплайна
      4) Защита от "плохого инпута" с выводом сообщений и логированием
      5) Обязательное комментирование пайплайна + обязательное комментирование сложных кусков кода
      6) Все базовые данные макроса (допустим, абсолютные ссылки которые я дам, относительные ссылки, контролируемый шифр - будет ясно далее, какие именно) вынести на самый верх макроса в виде констант, чтобы их было удобно править


## Основная информация по кодированию:


```txt

    1) Никогда не использовать Set для объявления - ТОЛЬКО dim (Set — присвоение объектной переменной – не работает в инвенторе!!!). Объекты делае без Set:

    ``` bas

    fso = CreateObject("Scripting.FileSystemObject") ' Object, без Set

    ```

    2) Не использовать команду либы Format() - менять на свои итераторы
    
    3) НЕ ИСПОЛЬЗОВАТЬ КРУГЛЫЕ СКОБКИ НА КОНЦЕ ФУНКЦИЙ И САБОВ ПРИ ВЫЗОВЕ

    ``` bas

    ' Так нельзя!

    oContext = app.TransientObjects.CreateTranslationContext()
    oOptions = app.TransientObjects.CreateNameValueMap()
    oData = app.TransientObjects.CreateDataMedium()

    ' Надо так!

    oContext = app.TransientObjects.CreateTranslationContext
    oOptions = app.TransientObjects.CreateNameValueMap
    oData = app.TransientObjects.CreateDataMedium

    ```

    4) Использовать круглые скобки при любой передаче параметров / аргументов

    ``` bas

    ' Так нельзя!
    MsgBox log, vbInformation, "DXF Export Completed"

    ' Надо так!
    MsgBox(log, vbInformation, "DXF Export Completed")

    ```

    5) 'Тип "Variant" больше не поддерживается; используйте тип "Object". !!!


```
___

___

1. Описание основного функционала

Разрабатываемый макрос служит для облегчения повседневных дел инженера-конструктора. Его главная функция - генерация заготовок для всех чертежей, которые будут необходимы в конструкторском проекте.

После разработки 3-д моделей изделия, присвоения им шифров, номеров и наименований, а также - формирования набора их метаданных инженер приступает к разработке чертежей на уникальные изделия (чертежи деталей и сборочных единиц).

Формирование заготовок чертежей можно разделить на 3 основных этапа:

0) Анализ дерева корневой сборки
1) Переход в директорию для хранения редактируемого формата (.idw) чертежей
2) Копирования шаблонов чертежей для деталей и сборок под каждое уникальное изделие в дереве (без дублирования) с неймингом файлов чертежей, соответствующим неймингу файлов 3д моделей (грубо говоря - создание соответсвующих моделям проекта файлов чертежей).
3) Вывода одного основного вида связанной с чертежом модели на лист 
4) Заполнения основной надписи рамки по метадате чертежа (принимаемой по основному виду)


Результатом формирования заготовок является наличие в требуемой папке для чертежей всех уникальных изделий (по дереву основной сборки) с выведенным на лист основным видом модели и заполненной по метадате модели основной надписью.

Соответственно - инженеру останется лишь сделать главную работу (образмерить, нанести технические требования).
___


___

2. Самые важные фичи

0) При генерации заготовок чертежей необходимо ориентироваться по дереву основной подсборки (не привязываться к адресам моделей, а смотреть их по дереву)
1) При генерации заготовок чертежей необходимо ориентироваться на шифр - проверять, есть ли в срезке названия 3д модели шифр и генерировать чертеж, только если шифр есть
2) При генерации заготовок чертежей необходимо проверять нет ли в конечной папке чертежа, который должен быть сгенерирован - если есть - не генерировать, чтобы не заменить пустым чертежом тот чертеж, с которым уже работали (защита данных)!!! Соответственно при новой генерации в конце необходимо выводить список сгенерированных файлов (чтобы понимать, какие новые модели были выведены в заготовки чертежей, а какие чертежи остались с прошлых разов)

___


___

3. Основной программный пайплайн


Макрос будет запускаться из 3д-модели основной сборки, на её адрес можно ориентироваться


Папка под .idw-файлы находится по относительному пути: "../1_Чертежи/1_DWG". Если папки по нужному адресу нет, то её нужно создать. Её относительный путь должен быть одной из констант в верхней зоне кода макроса.

Шаблон чертежа под деталь будет находиться по абсолютному пути: "Y:\000_Библиотека\400_RnD\099_iLogicRepository\Templates\ru-RU\kriskovets\detail.idw". Его абсолютный путь должен быть одной из констант в верхней зоне кода макроса.

Шаблон чертежа под сборку будет находиться по абсолютному пути: "Y:\000_Библиотека\400_RnD\099_iLogicRepository\Templates\ru-RU\kriskovets\assembly.idw". Его абсолютный путь должен быть одной из констант в верхней зоне кода макроса.

В них активны первые листы, с ними и надо работать


1) Необходимо обходом дерева (пример дальше) понять какие вообще детали и сборки надо выводить на чертеж - берем шифр из константы и проверям названия деталей и сборок в дереве на предмет содержания данного шифра. Заполняем какой-то список всего, что возможно нужно будет сгенерировать (допустим, абсолютные адреса моделей сборок и деталей, включая основную сборку), исключая дублирования в списке.
2) Чекаем целевую папку, смотрим нет ли в ней уже чего-то из списка на генерацию, если есть - убираем это из списка на генерацию, чтобы не затереть существующий чертеж
3) По финальному списку на генерацию делаем (опираясь на то - деталь модель или сборка):
   3.1) Копию шаблона
   3.2) Переименование скопированного шаблона (станет называться как деталь или сборка)
   3.3) Вывод на первый лист полученного файла основного вида модели (одна проекция, не больше) в универсальном масштабе (1:10)
   3.4) Заполнение метадаты через основной вид (пример дальше)

На этом вроде бы всё, в итоге получим готовые заготовки, с которыми можно будет работать.

Круто было бы в процессе видеть перед собой помимо логирования и ассертов прогресс бар (пример ниже)


Референсов на генерацию файлов и вывода основного вида с нужным масштабом на чертеж у меня нет( хнык. Есть только список свойств, который я тебе копирну ниже

___


___

4. Референсы

Обход дерева 1:

```bas

Option Explicit

Sub Main()

    ' Настройка обработки ошибок
    Dim assembly_error As Boolean: assembly_error = False
    Dim has_error As Boolean: has_error = False
    Dim export_error As Boolean: export_error = False

    On Error GoTo handle_common_error

    ' Проверка сборки
    check_assembly(assembly_error)
    If assembly_error Then GoTo handle_assembly_error


    ' Глобальные переменные

    ' Имя ини-файла для настройки DXF-вывода
    Dim dxf_ini_name As String
    dxf_ini_name = "Под резку.ini"

    Dim current_assembly_path As String ' Путь до текущей сборки
    current_assembly_path = ThisApplication.ActiveDocument.FullFileName

    Dim current_assembly_path_tmp As String ' Путь до текущей сборки - сохранение
    current_assembly_path_tmp = current_assembly_path

    ' Базовая папка из адреса текущей подсборки
    Dim base_folder_path As String ' Основная папка проекта конструкции
    base_folder_path = find_base_folder_path(current_assembly_path)

    ' Проверка структуры проекта
    If (base_folder_path  = "") Then

        has_error = True
        GoTo handle_common_error

    End If


    ' Папка для DXF из адреса основной папки
    Dim dxf_folder_path As String ' Папка для хранения DXF
    dxf_folder_path = find_dxf_folder_path(base_folder_path)

    ' Проверка структуры проекта
    If (dxf_folder_path = "") Then

        has_error = True
        GoTo handle_common_error

    End If


    ' Путь до ини-файла настройки DXF-вывода через папку для хранения DXF
    Dim ini_file_path As String 
    ini_file_path = find_ini_file_path(dxf_folder_path, dxf_ini_name)

    ' Проверка структуры проекта
    If (ini_file_path = "") Then

        has_error = True
        GoTo handle_common_error

    End If


    ' Инициируем список формата [[Имя подсборки, [Путь к детали, Путь к детали], ...] для подсборок внутри списка
    ' assemblies_root, если там внутри есть деталь листовой металл 
    Dim assemblies_with_sheet_metal As Object
    assemblies_with_sheet_metal = CreateObject("System.Collections.ArrayList")

    Dim isRoot As Boolean
    isRoot =  True

    ' Расширяем инициированный список подсборками и деталями
    Dim addedAssemblies As Object
    addedAssemblies = CreateObject("Scripting.Dictionary")

    Call collect_assemblies_and_sheet_metal(ThisApplication.ActiveDocument.ComponentDefinition.Occurrences, assemblies_with_sheet_metal, isRoot, ThisApplication.ActiveDocument, addedAssemblies)

    Call LogAssembliesWithSheetMetal(assemblies_with_sheet_metal)


    ' Производим экспорт всех деталей для всех подсборок в списке, создавая под них папки в папке для DXF
    export_dxf_for_curerrent_assembly(assemblies_with_sheet_metal, dxf_folder_path, ini_file_path, export_error)


    If export_error Then GoTo handle_export_error

    ' --- Успешное завершение ---
    MsgBox("Программа выполнена успешно")


    Dim doc As Document
    doc = ThisApplication.Documents.Open(current_assembly_path_tmp, True) ' True → открыть и активировать
    doc.Activate

    Exit Sub

handle_common_error:
    MsgBox("Произошла системная ошибка: " & Err.Description, vbCritical, "Error")
    Exit Sub

handle_assembly_error:
    MsgBox("Open an assembly before running this rule.", vbCritical, "Error")
    Exit Sub

handle_export_error:
    MsgBox("Export error!", vbCritical, "Error")
    Exit Sub
End Sub


' Вспомогательный саб для чека сборки
Sub check_assembly(ByRef assembly_error As Boolean)

    If ThisApplication.ActiveDocument Is Nothing Or _
       ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then

        assembly_error = True
        Exit Sub

    End If

End Sub


' Вспомогательная функция для поиска корня проекта
Function find_base_folder_path(current_assembly_path As String) As String

    Dim project_root As String
    project_root = ""

    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject") ' Object, без Set

    Dim cur As String
    cur = current_assembly_path

    ' Итеративный спуск до корневой папки
    Do While (cur <> "")

        ' Проверка признаков корня проекта
        If (Dir(cur & "\0_Модели", vbDirectory) <> "") And (Dir(cur & "\1_Чертежи", vbDirectory) <> "") Then

            project_root = cur
            Exit Do

        End If

        ' Если достигли корня диска
        If (fso.GetParentFolderName(cur) = "") Then Exit Do
        
        ' Итерация
        cur = fso.GetParentFolderName(cur)
    Loop

    ' Если не найден корень проекта
    If (project_root = "") Then

        MsgBox("Project root not found (expected folders: 0_Модели and 5_Чертежи).", vbCritical, "Error")
        find_base_folder_path = ""  ' Возврат пустой строки при ошибке
        Exit Function

    End If

    find_base_folder_path = project_root ' Возврат найденного пути

End Function


' Вспомогательная функция для поиска папки для DXF
Function find_dxf_folder_path(base_folder_path As String) As String

    Dim dxf_folder As String
    dxf_folder = base_folder_path & "\1_Чертежи\3_DXF"

    ' Проверка существования папки
    If Dir(dxf_folder, vbDirectory) = "" Then
        ' Папка не найдена — возвращаем пустую строку
         MsgBox("Project folders not found (expected folders: 1_Чертежи and 3_DXF).", vbCritical, "Error")
        find_dxf_folder_path = ""

    Else
        ' Папка существует — возвращаем путь
        find_dxf_folder_path = dxf_folder

    End If

End Function


' Вспомогательная функция для поиска пути к ини-файлу настройки вывода в DXF
Function find_ini_file_path(dxf_folder_path As String, dxf_ini_name As String) As String

    Dim ini_path As String
    ini_path = dxf_folder_path & "\" & dxf_ini_name

    ' Проверка существования файла
    If Dir(ini_path) = "" Then
        ' Файл не найден — выводим название файла в сообщении
        MsgBox("Project file not found (expected file: " & dxf_ini_name & ").", vbCritical, "Error")
        find_ini_file_path = ""
    Else
        ' Файл найден — возвращаем полный путь
        find_ini_file_path = ini_path
    End If

End Function


' Саб для получения списка сборка - детали, содержащих детали из листового металла
' и списка деталей из листового металла в этих сборках
Sub collect_assemblies_and_sheet_metal(occurrences As Object, assemblies As Object, isRoot As Boolean, rootDoc As Object, addedAssemblies As Object)
    
    ' --- Инициализация словаря для защиты от повторов ---
    If addedAssemblies Is Nothing Then addedAssemblies = CreateObject("Scripting.Dictionary")
    
    Dim asmPath As String
    Dim asmName As String
    
    ' --- Определяем имя и путь текущей сборки ---
    If isRoot And Not rootDoc Is Nothing Then
        asmPath = rootDoc.FullFileName
        asmName = rootDoc.DisplayName
    Else
        If occurrences.Count > 0 Then
            asmPath = occurrences(1).Parent.Document.FullFileName
            asmName = occurrences(1).Parent.Document.DisplayName
        Else
            asmPath = ""
            asmName = "Unnamed"
        End If
    End If
    
    ' --- Проверка словаря на повторное добавление ---
    If addedAssemblies.Exists(asmPath) Then Exit Sub
    Call addedAssemblies.Add(asmPath, True)
    
    ' --- Создаём запись сборки ---
    Dim currentEntry(1) As Object
    currentEntry(0) = asmName
    currentEntry(1) = CreateObject("System.Collections.ArrayList")
    
    ' --- 1. Добавляем листовые детали текущего уровня ---
    Dim occ As Object
    For Each occ In occurrences
        If TypeName(occ.Definition) = "SheetMetalComponentDefinition" Then
            Dim fullPath As String
            fullPath = occ.Definition.Document.FullFileName
            If Not currentEntry(1).Contains(fullPath) Then
                Call currentEntry(1).Add(fullPath)
            End If
        End If
    Next
    
    ' --- 2. Добавляем запись в assemblies только если есть листовые детали ---
    If currentEntry(1).Count > 0 Then
        Call assemblies.Add(currentEntry)
    End If
    
    ' --- 3. Рекурсивно идём по подсборкам ---
    For Each occ In occurrences
        If TypeName(occ.Definition) = "AssemblyComponentDefinition" Then
            Call collect_assemblies_and_sheet_metal(occ.Definition.Occurrences, assemblies, False, Nothing, addedAssemblies)
        End If
    Next

End Sub


Sub export_dxf_for_curerrent_assembly(assemblies_with_sheet_metal As Object, dxf_folder_path As String, ini_file_path As String, ByRef export_error As Boolean)

    Dim asm_entry As Object
    Dim i As Long

    For Each asm_entry In assemblies_with_sheet_metal

        ' --- подготовка имени папки сборки ---
        Dim clean_sub_name As String
        clean_sub_name = asm_entry(0)

        If InStrRev(clean_sub_name, ".") > 0 Then
            clean_sub_name = Left(clean_sub_name, InStrRev(clean_sub_name, ".") - 1)
        End If

        Dim invalidChars As String
        invalidChars = "/\[]*?<>|:"
        For k = 1 To Len(invalidChars)
            clean_sub_name = Replace(clean_sub_name, Mid(invalidChars, k, 1), "_")
        Next

        Dim arr() As String
        arr = Split(clean_sub_name, "_")
        If UBound(arr) > 0 Then
            If IsNumeric(arr(UBound(arr))) Then
                ReDim Preserve arr(UBound(arr) - 1)
                clean_sub_name = Join(arr, "_")
            End If
        End If

        Dim subfolder_path As String
        subfolder_path = dxf_folder_path & "\" & clean_sub_name
        If Dir(subfolder_path, vbDirectory) = "" Then MkDir(subfolder_path)

        ' --- подготовка деталей ---
        Dim sheet_parts As Object
        sheet_parts = asm_entry(1)

        Dim part_as_str As Object
        Dim partIndex As Integer
        partIndex = 1

        For Each part_as_str In sheet_parts
            On Error Resume Next

            Dim part_as_component As Object
            part_as_component = ThisApplication.Documents.Open(part_as_str, False)

            If Not part_as_component Is Nothing Then
                Dim oDef As Object
                oDef = part_as_component.ComponentDefinitions(1)

                If TypeName(oDef) = "SheetMetalComponentDefinition" Then
                    Dim smPartDoc As Object
                    smPartDoc = oDef.Document

                    If Not oDef.HasFlatPattern Then
                        oDef.Unfold()
                        oDef.FlatPattern.ExitEdit()
                    End If

                    Dim oDrawing As Object
                    oDrawing = ThisApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, , False)
                    oDrawing.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits

                    Dim oTG As Object
                    oTG = ThisApplication.TransientGeometry
                    Dim oX As Double
                    oX = 0
                    Dim oY As Double
                    oY = 0
                    Dim oBaseViewOptions As Object
                    oBaseViewOptions = ThisApplication.TransientObjects.CreateNameValueMap()
                    Call oBaseViewOptions.Add("SheetMetalFoldedModel", False)

                    Dim oSheet As Object
                    oSheet = oDrawing.ActiveSheet

                    Dim oView As Object
                    oView = oSheet.DrawingViews.AddBaseView(smPartDoc, oTG.CreatePoint2d(oX, oY), 1, _
                        ViewOrientationTypeEnum.kDefaultViewOrientation, _
                        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                        "FlatPattern", , oBaseViewOptions)


                    Dim thickness_mm As Double
                    Dim thickness_str As String
                    Dim dotPos As Long
                    Dim s As String

                    ' Получаем толщину
                    thickness_mm = oDef.Thickness.Value * 10

                    ' Превращаем в строку
                    s = CStr(thickness_mm)

                    ' Ищем точку или запятую
                    dotPos = InStr(s, ".")
                    If dotPos = 0 Then dotPos = InStr(s, ",")  ' на случай локали с запятой

                    If dotPos > 0 Then
                        ' Берём до двух знаков после точки
                        thickness_str = Left(s, dotPos + 2)
                    Else
                        ' Если точки нет, оставляем как есть
                        thickness_str = s
                    End If

                    ' Присваиваем имя виду
                    oView.Name = smPartDoc.DisplayName & " - " & thickness_str & " мм"
                    oView.ShowLabel = True


                    Call RemoveBendLines(oView, oDef.FlatPattern)

                    Dim part_file_name As String
                    part_file_name = Mid(part_as_str, InStrRev(part_as_str, "\") + 1)
                    If InStrRev(part_file_name, ".") > 0 Then
                        part_file_name = Left(part_file_name, InStrRev(part_file_name, ".") - 1)
                    End If

                    ' очистка имени детали
                    For n = 1 To Len(invalidChars)
                        part_file_name = Replace(part_file_name, Mid(invalidChars, n, 1), "_")
                    Next

                    Dim oDXFName As String
                    oDXFName = subfolder_path & "\" & part_file_name & "_развертка.dxf"
                    Call SaveDXF(oDrawing, oDXFName, ini_file_path)

                    Call oDrawing.Close()
                    partIndex = partIndex + 1
                End If

                Call part_as_component.Close(True)
            End If
            On Error GoTo 0
        Next
    Next

    MsgBox("DXF Export Completed", vbInformation)
End Sub



'====================================================
' Убираем линии сгиба из вида
'====================================================
Sub RemoveBendLines(oView As DrawingView, oFlatPattern As FlatPattern)
    Dim oBendEdgesUp As Edges
    oBendEdgesUp = oFlatPattern.GetEdgesOfType(FlatPatternEdgeTypeEnum.kBendUpFlatPatternEdge)

    Dim oBendEdgesDown As Edges
    oBendEdgesDown = oFlatPattern.GetEdgesOfType(FlatPatternEdgeTypeEnum.kBendDownFlatPatternEdge)

    Dim oEdge As Edge
    Dim oCurve As DrawingCurve
    Dim oSegment As DrawingCurveSegment

    For Each oEdge In oBendEdgesUp
        For Each oCurve In oView.DrawingCurves(oEdge)
            For Each oSegment In oCurve.Segments
                oSegment.Visible = False
            Next
        Next
    Next

    For Each oEdge In oBendEdgesDown
        For Each oCurve In oView.DrawingCurves(oEdge)
            For Each oSegment In oCurve.Segments
                oSegment.Visible = False
            Next
        Next
    Next
End Sub

'====================================================
' Сохраняем DrawingDocument в DXF
'====================================================
Sub SaveDXF(oDrawing As DrawingDocument, oFileName As String, oIniFile As String)
	'a reference to the DFX translator
	Dim DXFAddIn As TranslatorAddIn
	DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")

	'Create translation context
	Dim oContext As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext
	oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
	'Create options for the translation
	Dim oOptions As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap
	'Create a DataMedium object
	Dim oDataMedium As DataMedium = ThisApplication.TransientObjects.CreateDataMedium
	'the options (which .ini-file to use)
	If DXFAddIn.HasSaveCopyAsOptions(oDrawing, oContext, oOptions) Then
		oOptions.Value("Export_Acad_IniFile") = oIniFile
	End If
	'the filename property of the DataMedium object
	oDataMedium.FileName = oFileName

    DXFAddIn.SaveCopyAs(oDrawing, oContext, oOptions, oDataMedium)
End Sub

'====================================================
' Создаем текстовый файл
'====================================================
Sub CreateTXT(oText As String, oFileName As String)
    Dim oTxtWriter As Object
    oTxtWriter = CreateObject("Scripting.FileSystemObject").CreateTextFile(oFileName, True)
    oTxtWriter.WriteLine(oText)
    oTxtWriter.Close
End Sub

' Логирование ArrayList assemblies_root
Sub LogAssembliesRoot(assemblies_root As Object)
    Dim msg As String
    msg = "=== assemblies_root ===" & vbCrLf

    Dim assembly As Object
    For Each assembly In assemblies_root
        msg = msg & "Name: " & assembly(0) & " | Path: " & assembly(1) & vbCrLf
    Next
 
    MsgBox(msg, vbInformation, "assemblies_root log")
End Sub

' Логирование ArrayList assemblies_with_sheet_metal
Sub LogAssembliesWithSheetMetal(assemblies_with_sheet_metal As Object)
    Dim msg As String
    msg = "=== assemblies_with_sheet_metal ===" & vbCrLf

    Dim entry As Object
    Dim sheetPart As Object

    For Each entry In assemblies_with_sheet_metal
        msg = msg & "Subassembly: " & entry(0) & vbCrLf
        For Each sheetPart In entry(1)
            msg = msg & "  sheet part: " & sheetPart & vbCrLf
        Next
    Next

    MsgBox(msg, vbInformation, "assemblies_with_sheet_metal log")
End Sub

```

Обход дерева 2:

```bas

Option Explicit On

Sub Main()
    ' Проверяем, что открыта сборка
    Dim oDoc As Document = ThisDoc.Document
    If oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("Пожалуйста, запустите скрипт из файла сборки (.iam)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return
    End If
    
    Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)
    
    ' Формируем путь к текстовому файлу (рядом со сборкой)
    Dim asmPath As String = oAsmDoc.FullFileName
    If String.IsNullOrEmpty(asmPath) Then
        MessageBox.Show("Пожалуйста, сначала сохраните сборку!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return
    End If
    
    Dim txtPath As String = System.IO.Path.ChangeExtension(asmPath, ".txt")
    
    ' Создаем поток для записи файла
    Dim writer As New System.IO.StreamWriter(txtPath, False, System.Text.Encoding.UTF8)
    
    ' Записываем имя главной сборки
    writer.WriteLine(oAsmDoc.DisplayName)
    
    ' Запускаем рекурсивный обход компонентов
    Dim oAsmDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition
    ProcessOccurrences(oAsmDef.Occurrences, 1, writer)
    
    ' Закрываем файл
    writer.Close()
    writer.Dispose()
    
    MessageBox.Show("Дерево успешно выгружено в файл:" & vbCrLf & txtPath, "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information)
End Sub

' Рекурсивная процедура для обхода дерева
Sub ProcessOccurrences(ByVal oOccurrences As ComponentOccurrences, ByVal level As Integer, ByRef writer As System.IO.StreamWriter)
    Dim oOcc As ComponentOccurrence
    
    ' Создаем отступы (4 пробела на уровень)
    Dim indent As String = New String(" ", level * 4)
    
    For Each oOcc In oOccurrences
        ' Проверяем, что компонент не подавлен
        If Not oOcc.Suppressed Then
            ' Записываем имя файла (например, "деталь.ipt") вместо внутреннего имени "деталь:1"
            writer.WriteLine(indent & oOcc.ReferencedDocumentDescriptor.DisplayName)
            
            ' Если это подсборка, спускаемся глубже
            If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                If oOcc.SubOccurrences.Count > 0 Then
                    ProcessOccurrences(oOcc.SubOccurrences, level + 1, writer)
                End If
            End If
        End If
    Next
End Sub

```


Заполнение рамки по метадате из основного вида (много тонкостей, лучше просто дублировать в код с подстройкой):

```bas

Option Explicit


' =========================================================
' ПАКЕТНАЯ СИНХРОНИЗАЦИЯ ЧЕРТЕЖЕЙ
' =========================================================
'
' Общий поток:
'
' ROOT (папка проекта)
'    ↓
' Получаем *.idw из корневой папки
'    ↓
' Исключаем служебные чертежи
'    ↓
' Открываем каждый чертёж
'    ↓
' Синхронизируем с моделью (process_drawing)
'    ↓
' Сохраняем
'    ↓
' Закрываем
'    ↓
' Повторяем для всех файлов
'    ↓
' Показываем итог

Sub Main()

    ' =========================
    ' 1. НАСТРОЙКИ
    ' =========================

    ' Корневая директория (где лежит текущий документ)
    Dim root_dir As String
    root_dir = ThisDoc.Path


    ' Чертежи, которые не нужно обрабатывать
    Dim excluded_files As New System.Collections.Generic.HashSet(Of String)( _
        System.StringComparer.OrdinalIgnoreCase)

    excluded_files.Add("0.0_MASTER_FILE.idw")
    excluded_files.Add("0.1_ШАБЛОН ДЕТАЛЬ.idw")
    excluded_files.Add("0.2_ШАБЛОН СБОРКА.idw")


    ' =========================
    ' 2. СЧЁТЧИК РЕЗУЛЬТАТА
    ' =========================

    ' Сколько чертежей успешно обработано
    Dim total_count As Integer
    total_count = 0


    ' =========================
    ' 3. ПОЛУЧАЕМ ЧЕРТЕЖИ
    ' =========================

    ' Все чертежи (*.idw) в корневой папке
    ' Вложенные папки не обходятся
    Dim files() As String
    files = System.IO.Directory.GetFiles(root_dir, "*.idw")


    ' =========================
    ' 4. ГЛАВНЫЙ TRY (UI-КОНТРОЛЬ)
    ' =========================

    Try

        ' Отключаем обновление интерфейса:
        ' ускоряет массовую обработку и убирает мерцание
        ThisApplication.ScreenUpdating = False


        ' =========================
        ' 5. ОБХОД ВСЕХ ЧЕРТЕЖЕЙ
        ' =========================

        Dim curr_file_path As String

        For Each curr_file_path In files

            ' Получаем имя файла
            Dim file_name As String
            file_name = System.IO.Path.GetFileName(curr_file_path)

            ' Пропускаем файлы-исключения
            If excluded_files.Contains(file_name) Then
                Continue For
            End If


            ' =========================
            ' 6. ОБРАБОТКА КАЖДОГО ФАЙЛА
            ' =========================

            ' Документ чертежа
            Dim doc As DrawingDocument
            doc = Nothing

            Try

                ' =====================
                ' 6.1 ОТКРЫТИЕ
                ' =====================

                ' False → открываем в фоне (быстрее)
                doc = ThisApplication.Documents.Open(curr_file_path, False)


                ' =====================
                ' 6.2 СИНХРОНИЗАЦИЯ
                ' =====================

                ' Вся логика внутри отдельной функции
                process_drawing(doc)


                ' =====================
                ' 6.3 СОХРАНЕНИЕ
                ' =====================

                doc.Save

                total_count = total_count + 1


            Catch ex As Exception

                ' Ошибка по конкретному файлу
                MessageBox.Show("Ошибка: " & curr_file_path & vbCrLf & ex.Message)

            Finally

                ' =====================
                ' 6.4 ЗАКРЫТИЕ
                ' =====================

                ' False → не сохраняем повторно
                If Not doc Is Nothing Then doc.Close(False)

            End Try

        Next

    Finally

        ' ВАЖНО: возвращаем обновление UI
        ThisApplication.ScreenUpdating = True

    End Try


    ' =========================
    ' 7. ФИНАЛ
    ' =========================

    MessageBox.Show("Готово. Обновлено чертежей: " & total_count)

End Sub


' =========================================================
' СИНХРОНИЗАЦИЯ: ЧЕРТЕЖ ← МОДЕЛЬ
' =========================================================
'
' Поток:
'
' Чертёж
'     ↓
' Берём первый вид
'     ↓
' Получаем модель
'     ↓
' Читаем параметры
'     ↓
' Чистим строки
'     ↓
' Формируем логику (тип документа и т.д.)
'     ↓
' Записываем в iProperties


Sub process_drawing(opened_drawing_document As DrawingDocument)

    ' =========================
    ' 1. ПОЛУЧЕНИЕ МОДЕЛИ
    ' =========================

    Dim opened_sheet As Sheet
    opened_sheet = opened_drawing_document.ActiveSheet

    ' Если нет видов — дальше идти бессмысленно
    If opened_sheet.DrawingViews.Count = 0 Then Exit Sub

    ' Берём первый (базовый) вид
    Dim basic_view As DrawingView
    basic_view = opened_sheet.DrawingViews(1)


    Dim main_model_document As Document

    Try

        ' Получаем модель, связанную с видом
        main_model_document = basic_view.ReferencedDocumentDescriptor.ReferencedDocument

    Catch

        ' Ссылка может быть битой или модель не загружена
        MessageBox.Show("Не удалось получить модель из вида")
        Exit Sub

    End Try

    ' Дополнительная защита
    If main_model_document Is Nothing Then

        MessageBox.Show("Модель не найдена или не загружена")
        Exit Sub

    End If



    ' =========================
    ' 2. ЧТЕНИЕ ПАРАМЕТРОВ
    ' =========================
    
    ' Все значения проходят через:
    ' get_parameters → безопасное чтение
    ' clean_string   → очистка строки

    Dim part_number As String = clean_string(get_parameters(main_model_document, "part_number"))
    Dim part_name As String = clean_string(get_parameters(main_model_document, "part_name"))

    Dim part_type As String = clean_string(get_parameters(main_model_document, "part_type"))

    Dim part_developer As String = clean_string(get_parameters(main_model_document, "part_developer"))

    Dim developer_date As String = clean_string(get_parameters(main_model_document, "developer_date"))
    Dim part_test As String = clean_string(get_parameters(main_model_document, "part_test"))
    Dim test_date As String = clean_string(get_parameters(main_model_document, "test_date"))

    Dim part_tech_control As String = clean_string(get_parameters(main_model_document, "part_tech_control"))
    Dim tech_control_date As String = clean_string(get_parameters(main_model_document, "tech_control_date"))

    Dim part_department_head As String = clean_string(get_parameters(main_model_document, "part_department_head"))
    Dim department_head_date As String = clean_string(get_parameters(main_model_document, "department_head_date"))

    Dim part_norms_control As String = clean_string(get_parameters(main_model_document, "part_norms_control"))
    Dim norms_control_date As String = clean_string(get_parameters(main_model_document, "norms_control_date"))

    Dim part_approved_by As String = clean_string(get_parameters(main_model_document, "part_approved_by"))
    Dim part_approved_date As String = clean_string(get_parameters(main_model_document, "part_approved_date"))

    Dim part_company As String = get_parameters(main_model_document, "part_company")


    ' =========================
    ' 3. ЛОГИКА ДОКУМЕНТА
    ' =========================
    
    Dim document_type As String = ""

    ' Если это сборка — модифицируем номер и тип
    If part_type = "Сборка" Then 
        
        ' Защита от повторного добавления "СБ"
        If Not part_number.EndsWith("СБ") Then
            part_number = part_number & "СБ"
        End If

        document_type = "Сборочный чертеж"

    End If


    ' =========================
    ' 4. ЗАПИСЬ В iPROPERTIES
    ' =========================
    
    ' Всё пишется через safe_properties_setter:
    ' - защита от ошибок
    ' - нормализация значений
    ' - Inventor не падает

    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Part Number", part_number)
    safe_properties_setter(opened_drawing_document, "Inventor Summary Information", "Title", part_name)

    safe_properties_setter(opened_drawing_document, "Inventor Summary Information", "Author", part_developer)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Creation Time", developer_date)
    
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Checked By", part_test)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Date Checked", test_date)

    safe_properties_setter(opened_drawing_document, "Inventor Document Summary Information", "Manager", part_department_head)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Authority", part_department_head)
    safe_properties_setter(opened_drawing_document, "Свойства ГОСТ", "Руководитель Дата", department_head_date)

    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Engr Approved By", part_norms_control)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Engr Date Approved", norms_control_date)

    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Mfg Approved By", part_approved_by)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Mfg Date Approved", part_approved_date)

    ' Тип документа (если определён)
    If document_type <> "" Then
        safe_properties_setter(opened_drawing_document, "Свойства ГОСТ", "Тип документа", document_type)
    End If

    ' Компания
    safe_properties_setter(opened_drawing_document, "Inventor Document Summary Information", "Company", part_company)

End Sub


' =========================================================
' SAFE PARAMETER ACCESS (100% SAFE)
' =========================================================
Function get_parameters(doc As Document, paramName As String) As String

' Универсальный безопасный геттер параметров:
'
' 1. Пытается получить параметр из модели
' 2. Если параметра нет → ""
' 3. Если ошибка API → ""
'
' Никогда не падает

    Try

        Dim params As Parameters = doc.ComponentDefinition.Parameters

        Dim p As Parameter = params.Item(paramName)

        Return CStr(p.Value)

    Catch

        Return ""

    End Try

End Function


' =========================================================
' STRING CLEANER
' =========================================================
Function clean_string(value As String) As String

' Нормализует строку перед записью:
'
' - убирает переводы строк и табы
' - заменяет кавычки
' - удаляет запрещённые символы файловой системы
' - схлопывает двойные пробелы
' - обрезает длину
'
' Делает строку "безопасной для Inventor"

    Try
        If value Is Nothing Then Return ""

        Dim s As String = value

        s = s.Replace(vbCr, "")
        s = s.Replace(vbLf, "")
        s = s.Replace(vbTab, "")

        s = s.Replace(Chr(34), "'")

        Dim bad As String = "/\:*?<>|"

        For i As Integer = 1 To Len(bad)
            s = s.Replace(Mid(bad, i, 1), "_")
        Next

        Do While s.Contains("  ")
            s = s.Replace("  ", " ")
        Loop

        s = s.Trim()

        If s.Length > 250 Then
            s = Left(s, 250)
        End If

        Return s

    Catch
        Return ""
    End Try

End Function


' =========================================================
' SAFE iPROPERTY WRITER (NO LOOP, NO SEARCH)
' =========================================================
Sub safe_properties_setter(doc As Document, setName As String, propName As String, value As String)

' Абсолютно безопасная запись iProperty:
'
' 1. Нормализация значения:
'    - Nothing → ""
'    - "-" → ""
' 2. Прямой доступ к PropertySet
' 3. Try/Catch на весь блок
'
' Любая ошибка → игнор (Inventor не падает)

    Try

        ' =========================
        ' NORMALIZATION LAYER
        ' =========================
        If value Is Nothing Then value = ""

        value = value.Trim()

        ' Пустая строка в случае, если в ведомости проставлен прочерк
        If value = "-" Then value = ""

        Dim ps As PropertySet = doc.PropertySets.Item(setName)
        Dim p As [Property] = ps.Item(propName)

        p.Value = value

    Catch
        
    End Try

End Sub


```


Progress bar:

```bas

Imports Inventor
Sub Main()
    Dim oDoc As Document = ThisDoc.Document
    Dim oProgressBar As Inventor.ProgressBar = Nothing
    If oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then 
        MessageBox.Show("Правило предназначено для запуска СТРОГО из сборки!", "Отмена") : Exit Sub
    End If
    Dim docFolder As String = System.IO.Path.GetDirectoryName(oDoc.FullFileName)
    Dim exportFolder As String = System.IO.Path.Combine(docFolder, "DXF")
    If Not System.IO.Directory.Exists(exportFolder) Then System.IO.Directory.CreateDirectory(exportFolder)
    Dim processedGlobalFiles As New HashSet(Of String)()
    Dim missingFlatPatterns As New System.Text.StringBuilder()
    Dim processedFactories As New HashSet(Of String)()
    Dim totalCreated As Integer = 0
    Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)
    Dim totalItems As Integer = oAsmDoc.ComponentDefinition.Occurrences.AllLeafOccurrences.Count
    oProgressBar = ThisApplication.CreateProgressBar(False, totalItems, "Экспорт разверток DXF")
    ThisApplication.ScreenUpdating = False
    Try
        Dim currentProgress As Integer = 0
        ProcessAssemblyTree(oAsmDoc.ComponentDefinition.Occurrences, exportFolder, missingFlatPatterns, processedGlobalFiles, processedFactories, totalCreated, oProgressBar, currentProgress)
        ThisApplication.ScreenUpdating = True
        If oProgressBar IsNot Nothing Then oProgressBar.Close()
        Dim reportMsg As String = "Экспорт разверток завершен!" & vbCrLf & "Успешно создано уникальных файлов DXF: " & totalCreated
        If missingFlatPatterns.Length > 0 Then
            Dim warnMsg As String = "ВНИМАНИЕ! В следующих деталях не создана развертка (не нажата кнопка в модели):" & vbCrLf & missingFlatPatterns.ToString() & vbCrLf & "Эти файлы пропущены."
            MessageBox.Show(warnMsg, "Пропущенные элементы", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        MessageBox.Show(reportMsg, "Результат операции", MessageBoxButtons.OK, MessageBoxIcon.Information)
    Catch ex As System.Exception
        ThisApplication.ScreenUpdating = True
        If oProgressBar IsNot Nothing Then oProgressBar.Close()
        MessageBox.Show("Критический сбой экспорта DXF: " & ex.Message, "Сбой")
    End Try
End Sub
Private Sub ProcessAssemblyTree(ByVal oOccs As ComponentOccurrences, ByVal exportFolder As String, ByRef missingList As System.Text.StringBuilder, ByRef processedGlobalFiles As HashSet(Of String), ByRef processedFactories As HashSet(Of String), ByRef totalCreated As Integer, ByRef oProgressBar As Inventor.ProgressBar, ByRef currentProgress As Integer)
    For Each oOcc As ComponentOccurrence In oOccs
        If Not oOcc.Suppressed Then
            currentProgress += 1
            If oProgressBar IsNot Nothing Then
                oProgressBar.Message = "Сканирование структуры: компонент " & currentProgress
                oProgressBar.UpdateProgress()
            End If
            If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ProcessAssemblyTree(oOcc.SubOccurrences, exportFolder, missingList, processedGlobalFiles, processedFactories, totalCreated, oProgressBar, currentProgress)
            ElseIf oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Try
                    If oOcc.Definition.Type = ObjectTypeEnum.kSheetMetalComponentDefinitionObject Then
                        Dim oSMDef As SheetMetalComponentDefinition = CType(oOcc.Definition, SheetMetalComponentDefinition)
                        Dim oSubDoc As Document = oOcc.Definition.Document
                        Dim isiPartMember As Boolean = oSMDef.IsiPartMember
                        Dim factoryPath As String = ""
                        If isiPartMember Then
                            factoryPath = oSMDef.iPartMember.ParentFactory.Parent.FullFileName
                        ElseIf oSMDef.IsiPartFactory Then
                            factoryPath = oSubDoc.FullFileName
                        End If
                        If factoryPath <> "" Then
                            Dim factoryKey As String = factoryPath.ToLower()
                            If Not processedFactories.Contains(factoryKey) Then
                                processedFactories.Add(factoryKey)
                                Dim oFactoryDoc As PartDocument = CType(ThisApplication.Documents.Open(factoryPath, False), PartDocument)
                                Dim oFactDef As SheetMetalComponentDefinition = CType(oFactoryDoc.ComponentDefinition, SheetMetalComponentDefinition)
                                Dim oFactory As iPartFactory = oFactDef.iPartFactory
                                Dim totalRows As Integer = oFactory.TableRows.Count
                                Dim oOriginalRow As iPartTableRow = oFactory.DefaultRow
                                For i As Integer = 1 To totalRows
                                    Dim oRow As iPartTableRow = oFactory.TableRows.Item(i)
                                    Dim memberName As String = oRow.MemberName
                                    Dim dxfPath As String = System.IO.Path.Combine(exportFolder, memberName & ".dxf")
                                    If Not processedGlobalFiles.Contains(dxfPath.ToLower()) Then
                                        processedGlobalFiles.Add(dxfPath.ToLower())
                                        oFactory.DefaultRow = oRow
                                        oFactoryDoc.Update2(True)
                                        If oProgressBar IsNot Nothing Then
                                            oProgressBar.Message = "Запекание DXF: " & memberName
                                            oProgressBar.UpdateProgress()
                                        End If
                                        Dim isSaved As Boolean = False
                                        ExportFlatPatternEngine(oFactDef, dxfPath, memberName, missingList, isSaved)
                                        If isSaved Then totalCreated += 1
                                    End If
                                Next
                                Try : oFactory.DefaultRow = oOriginalRow : oFactoryDoc.Update2(True) : Catch : End Try
                                oFactoryDoc.Close(True)
                            End If
                        Else
                            Dim dxfName As String = oOcc.Name
                            If dxfName.Contains(":") Then dxfName = dxfName.Split(":"c)(0)
                            For Each invalidChar As Char In System.IO.Path.GetInvalidFileNameChars()
                                dxfName = dxfName.Replace(invalidChar, "_"c)
                            Next
                            Dim dxfPath As String = System.IO.Path.Combine(exportFolder, dxfName & ".dxf")
                            If Not processedGlobalFiles.Contains(dxfPath.ToLower()) Then
                                processedGlobalFiles.Add(dxfPath.ToLower())
                                If oProgressBar IsNot Nothing Then
                                    oProgressBar.Message = "Экспорт развертки: " & dxfName
                                    oProgressBar.UpdateProgress()
                                End If
                                Dim isSaved As Boolean = False
                                ExportFlatPatternEngine(oSMDef, dxfPath, dxfName, missingList, isSaved)
                                If isSaved Then totalCreated += 1
                            End If
                        End If
                    End If
                Catch
                End Try
            End If
        End If
    Next
End Sub
Private Sub ExportFlatPatternEngine(ByVal oSMDef As SheetMetalComponentDefinition, ByVal dxfPath As String, ByVal componentName As String, ByRef missingList As System.Text.StringBuilder, ByRef isSaved As Boolean)
    isSaved = False
    Try
        If oSMDef.HasFlatPattern Then
            Dim sOut As String = "FLAT PATTERN DXF?AcadVersion=R12&OuterProfileLayer=IV_OUTER_PROFILE&InteriorProfilesLayer=IV_INTERIOR_PROFILES&InvisibleLayers=IV_TANGENT;IV_BEND;IV_BEND_DOWN&SimplifySplines=True"
            oSMDef.DataIO.WriteDataToFile(sOut, dxfPath)
            isSaved = True
        Else
            missingList.AppendLine("- " & componentName)
        End If
    Catch
    End Try
End Sub


```

___


5. Список свойств (может поможет)

              Inventor Summary Information - Общая информация Inventor	
 Наименование                   	 iProperties.Value("Inventor Summary Information", "Title")
 Тема                           	 iProperties.Value("Inventor Summary Information", "Subject")
 Разраб.                        	 iProperties.Value("Inventor Summary Information", "Author")
 Ключевые слова                 	 iProperties.Value("Inventor Summary Information", "Keywords")
 Примечание                     	 iProperties.Value("Inventor Summary Information", "Comments")
 Сохранил                       	 iProperties.Value("Inventor Summary Information", "Last Saved By")
 Редакция                       	 iProperties.Value("Inventor Summary Information", "Revision Number")
 Миниатюра                      	 iProperties.Value("Inventor Summary Information", "Thumbnail")
              Inventor Document Summary Information - Информация о документе Inventor	
 Категория                      	 iProperties.Value("Inventor Document Summary Information", "Category")
 Нач. отд.                      	 iProperties.Value("Inventor Document Summary Information", "Manager")
 Компания                       	 iProperties.Value("Inventor Document Summary Information", "Company")
              Design Tracking Properties - Свойства Каталога проекта	
 Дата создания                  	 iProperties.Value("Design Tracking Properties", "Creation Time")
 Обозначение                    	 iProperties.Value("Design Tracking Properties", "Part Number")
 Проект                         	 iProperties.Value("Design Tracking Properties", "Project")
 Сметчик                        	 iProperties.Value("Design Tracking Properties", "Cost Center")
 Проверил                       	 iProperties.Value("Design Tracking Properties", "Checked By")
 Дата проверки                  	 iProperties.Value("Design Tracking Properties", "Date Checked")
 Нормоконтроль                  	 iProperties.Value("Design Tracking Properties", "Engr Approved By")
 Дата нормоконтроля             	 iProperties.Value("Design Tracking Properties", "Engr Date Approved")
 Статус                         	 iProperties.Value("Design Tracking Properties", "User Status")
 Материал                       	 iProperties.Value("Design Tracking Properties", "Material")
 Редакция детали                	 iProperties.Value("Design Tracking Properties", "Part Property Revision Id")
 Веб-сайт                       	 iProperties.Value("Design Tracking Properties", "Catalog Web Link")
 Значок детали                  	 iProperties.Value("Design Tracking Properties", "Part Icon")
 Описание                       	 iProperties.Value("Design Tracking Properties", "Description")
 Поставщик                      	 iProperties.Value("Design Tracking Properties", "Vendor")
 Тип детали                     	 iProperties.Value("Design Tracking Properties", "Document SubType")
 Название типа детали           	 iProperties.Value("Design Tracking Properties", "Document SubType Name")
 Дата обновления прокси         	 iProperties.Value("Design Tracking Properties", "Proxy Refresh Date")
 Утвердил                       	 iProperties.Value("Design Tracking Properties", "Mfg Approved By")
 Дата утверждения               	 iProperties.Value("Design Tracking Properties", "Mfg Date Approved")
 Стоимость                      	 iProperties.Value("Design Tracking Properties", "Cost")
 Стандарт                       	 iProperties.Value("Design Tracking Properties", "Standard")
 Стадия проекта                 	 iProperties.Value("Design Tracking Properties", "Design Status")
 Разработал                     	 iProperties.Value("Design Tracking Properties", "Designer")
 Т. контр.                      	 iProperties.Value("Design Tracking Properties", "Engineer")
 Нач. отдела                    	 iProperties.Value("Design Tracking Properties", "Authority")
 Параметризованный шаблон       	 iProperties.Value("Design Tracking Properties", "Parameterized Template")
 Шаблонная строка               	 iProperties.Value("Design Tracking Properties", "Template Row")
 Внешняя редакция детали        	 iProperties.Value("Design Tracking Properties", "External Property Revision Id")
 Редакция стандарта             	 iProperties.Value("Design Tracking Properties", "Standard Revision")
 Изготовитель                   	 iProperties.Value("Design Tracking Properties", "Manufacturer")
 Организация по стандартизации  	 iProperties.Value("Design Tracking Properties", "Standards Organization")
 Язык                           	 iProperties.Value("Design Tracking Properties", "Language")
 Откладывание изменений         	 iProperties.Value("Design Tracking Properties", "Defer Updates")
 Обозначение размера            	 iProperties.Value("Design Tracking Properties", "Size Designation")
 Категории                      	 iProperties.Value("Design Tracking Properties", "Categories")
 Инвентарный номер              	 iProperties.Value("Design Tracking Properties", "Stock Number")
 Материал сварки                	 iProperties.Value("Design Tracking Properties", "Weld Material")
 Масса                          	 iProperties.Value("Design Tracking Properties", "Mass")
 Площадь поверхности            	 iProperties.Value("Design Tracking Properties", "SurfaceArea")
 Объем                          	 iProperties.Value("Design Tracking Properties", "Volume")
 Плотность                      	 iProperties.Value("Design Tracking Properties", "Density")
 Действительные свойства массы  	 iProperties.Value("Design Tracking Properties", "Valid MassProps")
 Ширина границ развертки        	 iProperties.Value("Design Tracking Properties", "Flat Pattern Width")
 Длина границ развертки         	 iProperties.Value("Design Tracking Properties", "Flat Pattern Length")
 Площадь границ развертки       	 iProperties.Value("Design Tracking Properties", "Flat Pattern Area")
 Правило для листового металла  	 iProperties.Value("Design Tracking Properties", "Sheet Metal Rule")
 Последнее обновление с помощью 	 iProperties.Value("Design Tracking Properties", "Last Updated With")
 Ширина листового металла       	 iProperties.Value("Design Tracking Properties", "Sheet Metal Width")
 Длина листового металла        	 iProperties.Value("Design Tracking Properties", "Sheet Metal Length")
 Площадь листового металла      	 iProperties.Value("Design Tracking Properties", "Sheet Metal Area")
 Идентификатор материала        	 iProperties.Value("Design Tracking Properties", "Material Identifier")
 Представление                  	 iProperties.Value("Design Tracking Properties", "Appearance")
 Обновлять развертку вручную    	 iProperties.Value("Design Tracking Properties", "Flat Pattern Defer Update")
              Inventor User Defined Properties - Пользовательские свойства Inventor	
              Piping Style - Piping Style	
 Name                           	 iProperties.Value("Piping Style", "Name")
 ActiveState                    	 iProperties.Value("Piping Style", "ActiveState")
 PipeType                       	 iProperties.Value("Piping Style", "PipeType")
 RoutePreview                   	 iProperties.Value("Piping Style", "RoutePreview")
 Material                       	 iProperties.Value("Piping Style", "Material")
 Standard                       	 iProperties.Value("Piping Style", "Standard")
 FittingMaterial                	 iProperties.Value("Piping Style", "FittingMaterial")
 FittingStandard                	 iProperties.Value("Piping Style", "FittingStandard")
 RouteDirection90               	 iProperties.Value("Piping Style", "RouteDirection90")
 RouteDirection45               	 iProperties.Value("Piping Style", "RouteDirection45")
 DiameterType                   	 iProperties.Value("Piping Style", "DiameterType")
 Diameter                       	 iProperties.Value("Piping Style", "Diameter")
 Schedule                       	 iProperties.Value("Piping Style", "Schedule")
 OD                             	 iProperties.Value("Piping Style", "OD")
 ID                             	 iProperties.Value("Piping Style", "ID")
 Ending                         	 iProperties.Value("Piping Style", "Ending")
 IntermediateSegmentsMin        	 iProperties.Value("Piping Style", "IntermediateSegmentsMin")
 IntermediateSegmentsMax        	 iProperties.Value("Piping Style", "IntermediateSegmentsMax")
 IntermediateSegmentsInc        	 iProperties.Value("Piping Style", "IntermediateSegmentsInc")
 TerminationSegmentsMin         	 iProperties.Value("Piping Style", "TerminationSegmentsMin")
 TerminationSegmentsMax         	 iProperties.Value("Piping Style", "TerminationSegmentsMax")
 TerminationSegmentsInc         	 iProperties.Value("Piping Style", "TerminationSegmentsInc")
 RoundUpValue                   	 iProperties.Value("Piping Style", "RoundUpValue")
 DefaultBendRadius              	 iProperties.Value("Piping Style", "DefaultBendRadius")
 BendRadius                     	 iProperties.Value("Piping Style", "BendRadius")
 SolidRoutePreviewColor         	 iProperties.Value("Piping Style", "SolidRoutePreviewColor")
 PipingComponentColor           	 iProperties.Value("Piping Style", "PipingComponentColor")
 PipeKey                        	 iProperties.Value("Piping Style", "PipeKey")
 CouplingKey                    	 iProperties.Value("Piping Style", "CouplingKey")
 StartFittingKey                	 iProperties.Value("Piping Style", "StartFittingKey")
 EndFittingKey                  	 iProperties.Value("Piping Style", "EndFittingKey")
 Elbow90Key                     	 iProperties.Value("Piping Style", "Elbow90Key")
 Elbow45Key                     	 iProperties.Value("Piping Style", "Elbow45Key")
 SuppressStartFitting           	 iProperties.Value("Piping Style", "SuppressStartFitting")
 SuppressEndFitting             	 iProperties.Value("Piping Style", "SuppressEndFitting")
 UseSubASM                      	 iProperties.Value("Piping Style", "UseSubASM")



___
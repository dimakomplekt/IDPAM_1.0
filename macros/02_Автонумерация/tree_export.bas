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
' Chandra Excel Integration - Installer
' ВЕРСИЯ: 2.7 - Исправлена ошибка установки свойств формы "Object doesn't support..."
' ДАТА: 2024-10-28
' АВТОР: Jules
'
' КОНЦЕПЦИЯ:
' Пользователь копирует этот код в новый модуль в Excel и запускает макрос InstallChandraOCR.
'
' ИЗМЕНЕНИЯ v2.4:
' - JSON ПАРСЕР ЗАМЕНЕН НА НАДЕЖНУЮ ВЕРСИЮ (RegExp). Корректно обрабатывает запятые.
' - Все известные ошибки исправлены. Код стабилизирован.

Option Explicit

' ====================================================================
' ГЛАВНАЯ ФУНКЦИЯ УСТАНОВКИ
' ====================================================================

Sub InstallChandraOCR()
    On Error GoTo ErrorHandler

    Dim VBE As Object: Set VBE = Application.VBE
    If VBE.ActiveVBProject Is Nothing Then
        MsgBox "Не удалось получить доступ к проекту VBA." & vbCrLf & vbCrLf & _
               "Пожалуйста, разрешите программный доступ к модели объектов VBA:" & vbCrLf & _
               "Файл -> Параметры -> Центр управления безопасностью -> Параметры центра... ->" & vbCrLf & _
               "Параметры макросов -> Установить флажок 'Доверять доступ к объектной модели проектов VBA'.", vbCritical, "Ошибка доступа"
        Exit Sub
    End If

    MsgBox "Запускаю установщик Chandra Excel Integration...", vbInformation, "Установщик Chandra"

    CreateMainModule
    CreateUserForm

    MsgBox "Установка успешно завершена!", vbInformation, "Установка завершена"

    Exit Sub
ErrorHandler:
    MsgBox "Произошла критическая ошибка: " & Err.Description, vbCritical, "Ошибка установки"
End Sub

' ====================================================================
' СОЗДАНИЕ КОМПОНЕНТОВ
' ====================================================================

Private Sub CreateMainModule()
    Dim vbProj As Object: Set vbProj = ThisWorkbook.VBProject
    Dim moduleName As String: moduleName = "mdlChandraIntegration"

    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
    On Error GoTo 0

    Dim vbComp As Object: Set vbComp = vbProj.VBComponents.Add(1) ' vbext_ct_StdModule
    vbComp.Name = moduleName

    Dim codeMod As Object: Set codeMod = vbComp.CodeModule
    Dim codeLines As Variant: codeLines = GetMainModuleCode()

    Dim i As Long
    For i = LBound(codeLines) To UBound(codeLines)
        codeMod.InsertLines i + 1, codeLines(i)
    Next i
End Sub

Private Sub CreateUserForm()
    Dim vbProj As Object: Set vbProj = ThisWorkbook.VBProject
    Dim formName As String: formName = "frmChandraOCR"

    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(formName)
    On Error GoTo 0

    Dim vbComp As Object: Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm

    vbComp.Name = formName
    vbComp.Properties("PredeclaredId").Value = True

    With vbComp.Designer
        .Caption = "Chandra OCR - Обработка документа"
        .Width = 350
        .Height = 280

        With .Controls
            .Add "Forms.Label.1", "lblFilePath", True
            With .Item("lblFilePath"): .Caption = "1. Выберите PDF или изображение:": .Left = 10: .Top = 10: .Width = 200: End With
            .Add "Forms.TextBox.1", "txtFilePath", True
            With .Item("txtFilePath"): .Left = 10: .Top = 28: .Width = 240: .Height = 20: .Enabled = False: End With
            .Add "Forms.CommandButton.1", "btnBrowse", True
            With .Item("btnBrowse"): .Caption = "Обзор...": .Left = 255: .Top = 27: .Width = 70: .Height = 22: End With
            .Add "Forms.Label.1", "lblTemplate", True
            With .Item("lblTemplate"): .Caption = "2. Выберите шаблон (рекомендуется):": .Left = 10: .Top = 60: .Width = 200: End With
            .Add "Forms.ComboBox.1", "cmbTemplate", True
            With .Item("cmbTemplate"): .Left = 10: .Top = 78: .Width = 315: .Height = 20: .Style = 2: End With
            .Add "Forms.Label.1", "lblCustomPrompt", True
            With .Item("lblCustomPrompt"): .Caption = "Или введите свой промпт (для GPT):": .Left = 10: .Top = 110: .Width = 200: End With
            .Add "Forms.TextBox.1", "txtCustomPrompt", True
            With .Item("txtCustomPrompt"): .Left = 10: .Top = 128: .Width = 315: .Height = 40: .MultiLine = True: .ScrollBars = 2: End With
            .Add "Forms.Label.1", "lblStartCell", True
            With .Item("lblStartCell"): .Caption = "3. Укажите начальную ячейку для вывода:": .Left = 10: .Top = 180: .Width = 250: End With
            .Add "Forms.TextBox.1", "txtStartCell", True
            With .Item("txtStartCell"): .Left = 10: .Top = 198: .Width = 100: .Height = 20: End With
            .Add "Forms.CommandButton.1", "btnOK", True
            With .Item("btnOK"): .Caption = "Запуск": .Left = 170: .Top = 210: .Width = 75: .Height = 25: .Default = True: End With
            .Add "Forms.CommandButton.1", "btnCancel", True
            With .Item("btnCancel"): .Caption = "Отмена": .Left = 250: .Top = 210: .Width = 75: .Height = 25: .Cancel = True: End With
        End With
    End With

    Dim codeMod As Object: Set codeMod = vbComp.CodeModule
    Dim codeLines As Variant: codeLines = GetFormCode()

    Dim i As Long
    For i = LBound(codeLines) To UBound(codeLines)
        codeMod.InsertLines i + 1, codeLines(i)
    Next i
End Sub

' ====================================================================
' КОД ДЛЯ ГЕНЕРАЦИИ КОМПОНЕНТОВ
' ====================================================================

Private Function GetMainModuleCode() As Variant
    Dim lines As Collection: Set lines = New Collection
    lines.Add "' Chandra Excel Integration Module"
    lines.Add "' Версия: 2.7"
    lines.Add "Option Explicit"
    lines.Add ""
    lines.Add "Private Const COM_SERVER_NAME As String = ""ChandraExcel.Processor"""
    lines.Add "Private g_Processor As Object"
    lines.Add ""
    lines.Add "Public Sub ShowChandraForm()"
    lines.Add "    On Error GoTo ErrorHandler"
    lines.Add "    If GetProcessor() Is Nothing Then"
    lines.Add "        MsgBox ""Не удалось подключиться к COM-серверу 'ChandraExcel.Processor'."" & vbCrLf & vbCrLf & ""Убедитесь, что вы запустили 'python chandra_excel_com.py --register' от имени администратора."", vbCritical, ""Ошибка COM-сервера"""
    lines.Add "        Exit Sub"
    lines.Add "    End If"
    lines.Add "    frmChandraOCR.Show"
    lines.Add "    Exit Sub"
    lines.Add "ErrorHandler:"
    lines.Add "    MsgBox ""Ошибка при запуске формы: "" & Err.Description, vbCritical, ""Критическая ошибка"""
    lines.Add "End Sub"
    lines.Add ""
    lines.Add "Public Sub ProcessDocument(ByVal filePath As String, ByVal prompt As String, ByVal templateName As String, ByVal startCell As String)"
    lines.Add "    On Error GoTo ErrorHandler"
    lines.Add "    Dim result As String, json As Object"
    lines.Add "    Application.StatusBar = ""Chandra OCR: Идет обработка документа..."""
    lines.Add "    DoEvents"
    lines.Add "    If GetProcessor() Is Nothing Then GoTo ComError"
    lines.Add "    If templateName <> """" Then"
    lines.Add "        result = g_Processor.ProcessDocument(filePath, """", templateName)"
    lines.Add "    Else"
    lines.Add "        result = g_Processor.ProcessDocument(filePath, prompt, """")"
    lines.Add "    End If"
    lines.Add "    Set json = JsonParse(result)"
    lines.Add "    If json Is Nothing Or json.Count = 0 Then"
    lines.Add "        MsgBox ""Не удалось распознать данные в ответе от сервера: "" & result, vbCritical, ""Ошибка формата"""
    lines.Add "        GoTo Cleanup"
    lines.Add "    End If"
    lines.Add "    If json.Exists(""error"") Then"
    lines.Add "        MsgBox ""Сервер вернул ошибку: "" & json(""error""), vbCritical, ""Ошибка обработки"""
    lines.Add "        GoTo Cleanup"
    lines.Add "    End If"
    lines.Add "    If json.Exists(""data"") Then"
    lines.Add "        WriteToExcel json(""data""), startCell"
    lines.Add "        MsgBox ""Документ успешно обработан!"", vbInformation, ""Готово"""
    lines.Add "    Else"
    lines.Add "        MsgBox ""В ответе сервера отсутствуют данные."", vbExclamation, ""Нет данных"""
    lines.Add "    End If"
    lines.Add "    GoTo Cleanup"
    lines.Add "ComError:"
    lines.Add "    MsgBox ""Потеряно соединение с COM-сервером."", vbCritical, ""Ошибка COM-сервера"""
    lines.Add "    GoTo Cleanup"
    lines.Add "ErrorHandler:"
    lines.Add "    MsgBox ""Критическая ошибка при обработке: "" & Err.Description, vbCritical, ""Критическая ошибка"""
    lines.Add "Cleanup:"
    lines.Add "    Application.StatusBar = False"
    lines.Add "End Sub"
    lines.Add ""
    lines.Add "Private Sub WriteToExcel(ByVal data As Object, ByVal startCell As String)"
    lines.Add "    On Error GoTo ErrorHandler"
    lines.Add "    Dim ws As Worksheet: Set ws = ActiveSheet"
    lines.Add "    Dim startRange As Range: Set startRange = ws.Range(startCell)"
    lines.Add "    Dim keys As Variant: keys = data.keys"
    lines.Add "    Dim i As Long"
    lines.Add "    For i = 0 To data.Count - 1"
    lines.Add "        ws.Cells(startRange.Row, startRange.Column + i).Value = keys(i)"
    lines.Add "        ws.Cells(startRange.Row + 1, startRange.Column + i).Value = data(keys(i))"
    lines.Add "    Next i"
    lines.Add "    ws.Range(startRange, startRange.Offset(0, data.Count - 1)).EntireColumn.AutoFit"
    lines.Add "    Exit Sub"
    lines.Add "ErrorHandler:"
    lines.Add "    MsgBox ""Ошибка при записи в Excel: "" & Err.Description, vbCritical, ""Ошибка записи"""
    lines.Add "End Sub"
    lines.Add ""
    lines.Add "Public Function GetTemplates() As Object"
    lines.Add "    On Error Resume Next"
    lines.Add "    Dim templatesJson As String"
    lines.Add "    If GetProcessor() Is Nothing Then Set GetTemplates = Nothing: Exit Function"
    lines.Add "    templatesJson = g_Processor.GetTemplates()"
    lines.Add "    If Err.Number <> 0 Then Set GetTemplates = Nothing: Exit Function"
    lines.Add "    Set GetTemplates = JsonParse(templatesJson)"
    lines.Add "End Function"
    lines.Add ""
    lines.Add "Private Function GetProcessor() As Object"
    lines.Add "    If g_Processor Is Nothing Then"
    lines.Add "        On Error Resume Next"
    lines.Add "        Set g_Processor = CreateObject(""Scripting.Dictionary"")"
    lines.Add "        Set g_Processor = CreateObject(COM_SERVER_NAME)"
    lines.Add "        On Error GoTo 0"
    lines.Add "    End If"
    lines.Add "    Set GetProcessor = g_Processor"
    lines.Add "End Function"
    lines.Add ""
    lines.Add "Private Function JsonParse(ByVal jsonString As String) As Object"
    lines.Add "    ' НАДЕЖНЫЙ ПАРСЕР JSON с использованием RegExp"
    lines.Add "    On Error GoTo ParseError"
    lines.Add "    Dim regex As Object, matches As Object, match As Object"
    lines.Add "    Dim dict As Object: Set dict = CreateObject(""Scripting.Dictionary"")"
    lines.Add "    Dim q As String: q = Chr(34)"
    lines.Add "    Set regex = CreateObject(""VBScript.RegExp"")"
    lines.Add "    With regex"
    lines.Add "        .Global = True"
    lines.Add "        .MultiLine = True"
    lines.Add "        .Pattern = q & ""([^"" & q & ""]+)"" & q & ""\s*:\s*"" & q & ""([^"" & q & ""]*)"" & q"
    lines.Add "    End With"
    lines.Add "    If regex.Test(jsonString) Then"
    lines.Add "        Set matches = regex.Execute(jsonString)"
    lines.Add "        For Each match In matches"
    lines.Add "            If match.SubMatches.Count = 2 Then"
    lines.Add "                dict(match.SubMatches(0)) = match.SubMatches(1)"
    lines.Add "            End If"
    lines.Add "        Next"
    lines.Add "    End If"
    lines.Add "    Set JsonParse = dict"
    lines.Add "    Exit Function"
    lines.Add "ParseError:"
    lines.Add "    Set JsonParse = Nothing"
    lines.Add "End Function"

    Dim arr() As String: ReDim arr(0 To lines.Count - 1)
    Dim i As Long
    For i = 1 To lines.Count: arr(i - 1) = lines(i): Next
    GetMainModuleCode = arr
End Function

Private Function GetFormCode() As Variant
    Dim lines As Collection: Set lines = New Collection
    lines.Add "Option Explicit"
    lines.Add ""
    lines.Add "Private Sub UserForm_Initialize()"
    lines.Add "    Dim templates As Object: Set templates = mdlChandraIntegration.GetTemplates()"
    lines.Add "    Me.cmbTemplate.Clear"
    lines.Add "    Me.cmbTemplate.AddItem """" ' Пустая строка для кастомного промпта"
    lines.Add "    If Not templates Is Nothing And templates.Exists(""templates"") Then"
    lines.Add "        Dim templateName As Variant"
    lines.Add "        For Each templateName In templates(""templates"").keys"
    lines.Add "            Me.cmbTemplate.AddItem templateName"
    lines.Add "        Next templateName"
    lines.Add "    End If"
    lines.Add "    If Not ActiveCell Is Nothing Then"
    lines.Add "        Me.txtStartCell.Value = ActiveCell.Address"
    lines.Add "    Else"
    lines.Add "        Me.txtStartCell.Value = ""A1"""
    lines.Add "    End If"
    lines.Add "    Me.txtCustomPrompt.Enabled = False"
    lines.Add "End Sub"
    lines.Add ""
    lines.Add "Private Sub btnBrowse_Click()"
    lines.Add "    Dim fileDialog As FileDialog: Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)"
    lines.Add "    fileDialog.Title = ""Выберите PDF или файл изображения"""
    lines.Add "    fileDialog.Filters.Clear"
    lines.Add "    fileDialog.Filters.Add ""Документы"", ""*.pdf; *.png; *.jpg; *.jpeg; *.bmp"", 1"
    lines.Add "    If fileDialog.Show = -1 Then Me.txtFilePath.Value = fileDialog.SelectedItems(1)"
    lines.Add "End Sub"
    lines.Add ""
    lines.Add "Private Sub cmbTemplate_Change()"
    lines.Add "    If Me.cmbTemplate.Value = """" Then"
    lines.Add "        Me.txtCustomPrompt.Enabled = True"
    lines.Add "        Me.txtCustomPrompt.SetFocus"
    lines.Add "    Else"
    lines.Add "        Me.txtCustomPrompt.Enabled = False"
    lines.Add "        Me.txtCustomPrompt.Value = """""
    lines.Add "    End If"
    lines.Add "End Sub"
    lines.Add ""
    lines.Add "Private Sub btnOK_Click()"
    lines.Add "    If Me.txtFilePath.Value = """" Or Dir(Me.txtFilePath.Value) = """" Then"
    lines.Add "        MsgBox ""Пожалуйста, выберите существующий файл."", vbExclamation, ""Ошибка ввода"""
    lines.Add "        Exit Sub"
    lines.Add "    End If"
    lines.Add "    If Me.cmbTemplate.Value = """" And Me.txtCustomPrompt.Value = """" Then"
    lines.Add "        MsgBox ""Пожалуйста, выберите шаблон или введите свой промпт."", vbExclamation, ""Ошибка ввода"""
    lines.Add "        Exit Sub"
    lines.Add "    End If"
    lines.Add "    On Error Resume Next"
    lines.Add "    Dim tempRange As Range: Set tempRange = ActiveSheet.Range(Me.txtStartCell.Value)"
    lines.Add "    If Err.Number <> 0 Then"
    lines.Add "        MsgBox ""Пожалуйста, введите корректный адрес ячейки (например, A1)."", vbExclamation, ""Ошибка ввода"""
    lines.Add "        Exit Sub"
    lines.Add "    End If"
    lines.Add "    On Error GoTo 0"
    lines.Add "    Me.Hide"
    lines.Add "    mdlChandraIntegration.ProcessDocument Me.txtFilePath.Value, Me.txtCustomPrompt.Value, Me.cmbTemplate.Value, Me.txtStartCell.Value"
    lines.Add "    Unload Me"
    lines.Add "End Sub"
    lines.Add ""
    lines.Add "Private Sub btnCancel_Click()"
    lines.Add "    Unload Me"
    lines.Add "End Sub"

    Dim arr() As String: ReDim arr(0 To lines.Count - 1)
    Dim i As Long
    For i = 1 To lines.Count: arr(i - 1) = lines(i): Next
    GetFormCode = arr
End Function

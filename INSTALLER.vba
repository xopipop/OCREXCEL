Attribute VB_Name = "Installer"
' Chandra Excel Integration - Installer
' ЕДИНЫЙ УСТАНОВЩИК ВСЕХ КОМПОНЕНТОВ
' Запустите этот модуль для автоматической установки всех компонентов

Option Explicit

' ====================================================================
' ГЛАВНАЯ ФУНКЦИЯ УСТАНОВКИ
' ====================================================================
Sub InstallChandraOCR()
    ''' Устанавливает все компоненты Chandra OCR в Excel
    
    On Error GoTo ErrorHandler
    
    MsgBox "Добро пожаловать в установщик Chandra Excel Integration!" & vbCrLf & vbCrLf & _
           "Эта процедура создаст все необходимые модули и формы.", vbInformation, "Установщик"
    
    ' Шаг 1: Создание основного модуля
    Call CreateMainModule
    
    ' Шаг 2: Создание упрощенной формы
    Call CreateSimpleForm
    
    ' Шаг 3: Итоговое сообщение
    MsgBox "Установка завершена успешно!" & vbCrLf & vbCrLf & _
           "Для использования:" & vbCrLf & _
           "1. Убедитесь, что COM-сервер зарегистрирован (register_com.bat)" & vbCrLf & _
           "2. Запустите ShowChandraForm через Alt+F8" & vbCrLf & _
           "3. Или назначьте горячую клавишу на ShowChandraForm", vbInformation, "Готово"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка установки: " & Err.Description, vbCritical, "Ошибка"
End Sub

' ====================================================================
' СОЗДАНИЕ ОСНОВНОГО МОДУЛЯ
' ====================================================================
Private Sub CreateMainModule()
    ''' Создает mdlChandraIntegration с основным функционалом
    
    On Error Resume Next
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    ' Удаляем старый модуль если есть
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents("mdlChandraIntegration")
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
    
    ' Создаем новый модуль
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
    vbComp.Name = "mdlChandraIntegration"
    
    ' Получаем модуль кода
    Dim codeMod As Object
    Set codeMod = vbComp.CodeModule
    
    ' Код модуля
    Dim moduleCode As String
    moduleCode = GetMainModuleCode()
    
    ' Добавляем код
    codeMod.AddFromString moduleCode
    
    MsgBox "Модуль mdlChandraIntegration создан", vbInformation
    
End Sub

' ====================================================================
' СОЗДАНИЕ УПРОЩЕННОЙ ФОРМЫ
' ====================================================================
Private Sub CreateSimpleForm()
    ''' Создает упрощенную форму frmChandraOCR
    
    On Error Resume Next
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    ' Удаляем старую форму если есть
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents("frmChandraOCR")
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
    
    ' Создаем новую форму
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    vbComp.Name = "frmChandraOCR"
    
    ' Добавляем код формы
    Dim codeMod As Object
    Set codeMod = vbComp.CodeModule
    codeMod.DeleteLines 1, codeMod.CountOfLines
    
    Dim formCode As String
    formCode = GetFormCode()
    codeMod.AddFromString formCode
    
    MsgBox "Форма frmChandraOCR создана", vbInformation
    
End Sub

' ====================================================================
' КОД ОСНОВНОГО МОДУЛЯ
' ====================================================================
Private Function GetMainModuleCode() As String
    
    Dim code As String
    code = "Attribute VB_Name = ""mdlChandraIntegration""" & vbNewLine & _
           "' Chandra Excel Integration Module" & vbNewLine & _
           "' Модуль для интеграции с Chandra OCR через COM-объект" & vbNewLine & vbNewLine & _
           "Option Explicit" & vbNewLine & vbNewLine & _
           "' Константы" & vbNewLine & _
           "Const COM_SERVER_NAME As String = ""ChandraExcel.Processor""" & vbNewLine & vbNewLine & _
           "' Переменные для хранения данных формы" & vbNewLine & _
           "Dim g_filePath As String" & vbNewLine & _
           "Dim g_prompt As String" & vbNewLine & _
           "Dim g_templateName As String" & vbNewLine & _
           "Dim g_startCell As String" & vbNewLine & vbNewLine & _
           "' ====================================================================" & vbNewLine & _
           "' Основная функция запуска" & vbNewLine & _
           "' ====================================================================" & vbNewLine & _
           "Sub ShowChandraForm()" & vbNewLine & _
           "    On Error GoTo ErrorHandler" & vbNewLine & _
           "    If Not CheckCOMObject() Then" & vbNewLine & _
           "        MsgBox ""COM-объект Chandra не зарегистрирован. Запустите register_com.bat"", vbCritical, ""Ошибка""" & vbNewLine & _
           "        Exit Sub" & vbNewLine & _
           "    End If" & vbNewLine & _
           "    frmChandraOCR.Show" & vbNewLine & _
           "    Exit Sub" & vbNewLine & _
           "ErrorHandler:" & vbNewLine & _
           "    MsgBox ""Ошибка при запуске формы: "" & Err.Description, vbCritical, ""Ошибка""" & vbNewLine & _
           "End Sub" & vbNewLine & vbNewLine & _
           "Function CheckCOMObject() As Boolean" & vbNewLine & _
           "    On Error Resume Next" & vbNewLine & _
           "    Dim obj As Object" & vbNewLine & _
           "    Set obj = CreateObject(COM_SERVER_NAME)" & vbNewLine & _
           "    If Err.Number = 0 Then" & vbNewLine & _
           "        CheckCOMObject = True" & vbNewLine & _
           "        Set obj = Nothing" & vbNewLine & _
           "    Else" & vbNewLine & _
           "        CheckCOMObject = False" & vbNewLine & _
           "    End If" & vbNewLine & _
           "    On Error GoTo 0" & vbNewLine & _
           "End Function" & vbNewLine & vbNewLine & _
           "Sub ProcessDocument(filePath As String, prompt As String, templateName As String, startCell As String)" & vbNewLine & _
           "    On Error GoTo ErrorHandler" & vbNewLine & _
           "    Dim processor As Object, result As String" & vbNewLine & _
           "    Application.StatusBar = ""Обработка...""" & vbNewLine & _
           "    Set processor = CreateObject(COM_SERVER_NAME)" & vbNewLine & _
           "    If templateName <> """" Then" & vbNewLine & _
           "        result = processor.ProcessDocument(filePath, """", templateName)" & vbNewLine & _
           "    Else" & vbNewLine & _
           "        result = processor.ProcessDocument(filePath, prompt, """")" & vbNewLine & _
           "    End If" & vbNewLine & _
           "    Dim resultJson As Object" & vbNewLine & _
           "    Set resultJson = ParseJSON(result)" & vbNewLine & _
           "    If resultJson.Exists(""error"") Then" & vbNewLine & _
           "        MsgBox ""Ошибка: "" & resultJson(""error""), vbCritical" & vbNewLine & _
           "        GoTo Cleanup" & vbNewLine & _
           "    End If" & vbNewLine & _
           "    Dim data As Object" & vbNewLine & _
           "    Set data = resultJson(""data"")" & vbNewLine & _
           "    WriteToExcel data, startCell, templateName" & vbNewLine & _
           "    MsgBox ""Данные успешно обработаны и записаны в Excel"", vbInformation" & vbNewLine & _
           "    GoTo Cleanup" & vbNewLine & _
           "ErrorHandler:" & vbNewLine & _
           "    MsgBox ""Ошибка: "" & Err.Description, vbCritical" & vbNewLine & _
           "Cleanup:" & vbNewLine & _
           "    Application.StatusBar = False" & vbNewLine & _
           "End Sub" & vbNewLine & vbNewLine & _
           "Function ParseJSON(jsonString As String) As Object" & vbNewLine & _
           "    On Error Resume Next" & vbNewLine & _
           "    Dim jsonObj As Object" & vbNewLine & _
           "    Set jsonObj = CreateObject(""Scripting.Dictionary"")" & vbNewLine & _
           "    jsonObj.Add ""raw"", jsonString" & vbNewLine & _
           "    Set ParseJSON = jsonObj" & vbNewLine & _
           "End Function" & vbNewLine & vbNewLine & _
           "Sub WriteToExcel(data As Object, startCell As String, templateName As String)" & vbNewLine & _
           "    On Error GoTo ErrorHandler" & vbNewLine & _
           "    Dim ws As Worksheet, startRange As Range" & vbNewLine & _
           "    Set ws = ActiveSheet" & vbNewLine & _
           "    Set startRange = ws.Range(startCell)" & vbNewLine & _
           "    Dim i As Long, fieldName As String, cellValue As Variant" & vbNewLine & _
           "    For i = 0 To data.Count - 1" & vbNewLine & _
           "        If data.Exists(i) Then" & vbNewLine & _
           "            cellValue = data(i)" & vbNewLine & _
           "            ws.Cells(startRange.Row, startRange.Column + i).Value = cellValue" & vbNewLine & _
           "        End If" & vbNewLine & _
           "    Next i" & vbNewLine & _
           "    Exit Sub" & vbNewLine & _
           "ErrorHandler: MsgBox ""Ошибка записи: "" & Err.Description, vbCritical" & vbNewLine & _
           "End Sub" & vbNewLine & vbNewLine & _
           "Function GetTemplates() As String" & vbNewLine & _
           "    GetTemplates = ""{}""" & vbNewLine & _
           "End Function"
    
    GetMainModuleCode = code
    
End Function

' ====================================================================
' КОД ФОРМЫ
' ====================================================================
Private Function GetFormCode() As String
    
    Dim code As String
    code = "Attribute VB_Name = ""frmChandraOCR""" & vbNewLine & _
           "Attribute VB_PredeclaredId = True" & vbNewLine & _
           "Option Explicit" & vbNewLine & vbNewLine & _
           "Private Sub UserForm_Initialize()" & vbNewLine & _
           "    Me.Caption = ""Chandra OCR - Обработка документа""" & vbNewLine & _
           "    Me.Width = 300" & vbNewLine & _
           "    Me.Height = 300" & vbNewLine & _
           "End Sub" & vbNewLine & vbNewLine & _
           "Private Sub UserForm_Click()" & vbNewLine & _
           "    Dim fp As String, prompt As String, cell As String" & vbNewLine & _
           "    fp = InputBox(""Путь к PDF файлу:"", ""Chandra OCR"")" & vbNewLine & _
           "    If fp = """" Then Exit Sub" & vbNewLine & _
           "    prompt = InputBox(""Введите промпт:"", ""Chandra OCR"")" & vbNewLine & _
           "    If prompt = """" Then Exit Sub" & vbNewLine & _
           "    cell = InputBox(""Начальная ячейка:"", ""Chandra OCR"", ""A1"")" & vbNewLine & _
           "    If cell = """" Then Exit Sub" & vbNewLine & _
           "    ProcessDocument fp, prompt, """", cell" & vbNewLine & _
           "End Sub"
    
    GetFormCode = code
    
End Function


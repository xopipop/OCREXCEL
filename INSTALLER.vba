Attribute VB_Name = "Installer"
' Chandra Excel Integration - Installer
' ВЕРСИЯ: 2.0
' ДАТА: 2024-10-27
' АВТОР: Jules
'
' КОНЦЕПЦИЯ:
' Пользователь импортирует этот модуль в Excel и запускает макрос InstallChandraOCR.
' Макрос автоматически создает все необходимые компоненты для работы:
' 1. Модуль 'mdlChandraIntegration' - содержит всю логику для связи с Python COM-сервером.
' 2. Форму 'frmChandraOCR' - красивый и удобный пользовательский интерфейс.
'
' ИЗМЕНЕНИЯ v2.0:
' - Исправлена ошибка "Too many line continuations" путем разбивки генерации кода на строки.
' - Создается полноценная UserForm вместо серии InputBox.
' - Улучшена структура кода и добавлены комментарии.

Option Explicit

' ====================================================================
' ГЛАВНАЯ ФУНКЦИЯ УСТАНОВКИ
' ====================================================================

Sub InstallChandraOCR()
    '''
    ''' Главный макрос для установки всех компонентов Chandra OCR в Excel.
    '''
    On Error GoTo ErrorHandler

    Dim VBE As Object ' VBE
    Set VBE = Application.VBE

    ' Проверка, что доступ к VB проекту разрешен
    If VBE.ActiveVBProject Is Nothing Then
        MsgBox "Не удалось получить доступ к проекту VBA." & vbCrLf & vbCrLf & _
               "Пожалуйста, разрешите программный доступ к модели объектов VBA:" & vbCrLf & _
               "Файл -> Параметры -> Центр управления безопасностью -> Параметры центра управления безопасностью ->" & vbCrLf & _
               "Параметры макросов -> Установить флажок 'Доверять доступ к объектной модели проектов VBA'.", vbCritical, "Ошибка доступа"
        Exit Sub
    End If

    MsgBox "Добро пожаловать в установщик Chandra Excel Integration!" & vbCrLf & vbCrLf & _
           "Сейчас будут созданы все необходимые модули и формы для работы.", vbInformation, "Установщик Chandra"

    ' Шаг 1: Создание основного модуля
    CreateMainModule

    ' Шаг 2: Создание пользовательской формы
    CreateUserForm

    ' Шаг 3: Итоговое сообщение
    MsgBox "Установка успешно завершена!" & vbCrLf & vbCrLf & _
           "Как начать работу:" & vbCrLf & _
           "1. Убедитесь, что вы установили Python-зависимости и зарегистрировали COM-сервер (см. README.md)." & vbCrLf & _
           "2. Откройте редактор макросов (Alt+F11)." & vbCrLf & _
           "3. Запустите макрос 'ShowChandraForm' (через Alt+F8)." & vbCrLf & _
           "4. Рекомендуем назначить на него горячую клавишу для удобства.", vbInformation, "Установка завершена"

    Exit Sub

ErrorHandler:
    MsgBox "Произошла критическая ошибка во время установки: " & vbCrLf & vbCrLf & _
           "Описание: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number, vbCritical, "Ошибка установки"
End Sub

' ====================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ====================================================================

Private Sub CreateMainModule()
    '''
    ''' Создает или обновляет модуль 'mdlChandraIntegration' с основным функционалом.
    '''
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim moduleName As String
    moduleName = "mdlChandraIntegration"

    ' Удаляем старый модуль, если он существует
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
    On Error GoTo 0

    ' Создаем новый модуль
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
    vbComp.Name = moduleName

    ' Добавляем код в модуль
    Dim codeMod As Object
    Set codeMod = vbComp.CodeModule

    Dim codeLines As Variant
    codeLines = GetMainModuleCode()

    Dim i As Long
    For i = LBound(codeLines) To UBound(codeLines)
        codeMod.AddFromString codeLines(i)
    Next i

    Debug.Print "Модуль '" & moduleName & "' успешно создан."
End Sub

Private Sub CreateUserForm()
    '''
    ''' Создает или обновляет пользовательскую форму 'frmChandraOCR'.
    '''
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim formName As String
    formName = "frmChandraOCR"

    ' Удаляем старую форму, если она существует
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(formName)
    On Error GoTo 0

    ' Создаем новую форму
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_MSForm)

    With vbComp
        .Name = formName
        .Properties("Caption") = "Chandra OCR - Обработка документа"
        .Properties("Width") = 350
        .Properties("Height") = 280
    End With

    ' Добавляем элементы управления на форму
    ' --- Label для пути к файлу
    .Controls.Add "Forms.Label.1", "lblFilePath", True
    With .Controls("lblFilePath")
        .Caption = "1. Выберите PDF или изображение:"
        .Left = 10
        .Top = 10
        .Width = 200
    End With

    ' --- TextBox для пути к файлу
    .Controls.Add "Forms.TextBox.1", "txtFilePath", True
    With .Controls("txtFilePath")
        .Left = 10
        .Top = 28
        .Width = 240
        .Height = 20
        .Enabled = False
    End With

    ' --- Кнопка "Обзор"
    .Controls.Add "Forms.CommandButton.1", "btnBrowse", True
    With .Controls("btnBrowse")
        .Caption = "Обзор..."
        .Left = 255
        .Top = 27
        .Width = 70
        .Height = 22
    End With

    ' --- Label для шаблона
    .Controls.Add "Forms.Label.1", "lblTemplate", True
    With .Controls("lblTemplate")
        .Caption = "2. Выберите шаблон (рекомендуется):"
        .Left = 10
        .Top = 60
        .Width = 200
    End With

    ' --- ComboBox для шаблонов
    .Controls.Add "Forms.ComboBox.1", "cmbTemplate", True
    With .Controls("cmbTemplate")
        .Left = 10
        .Top = 78
        .Width = 315
        .Height = 20
        .Style = 2 ' fmStyleDropDownList
    End With

    ' --- Label для своего промпта
    .Controls.Add "Forms.Label.1", "lblCustomPrompt", True
    With .Controls("lblCustomPrompt")
        .Caption = "Или введите свой промпт (для GPT):"
        .Left = 10
        .Top = 110
        .Width = 200
    End With

    ' --- TextBox для своего промпта
    .Controls.Add "Forms.TextBox.1", "txtCustomPrompt", True
    With .Controls("txtCustomPrompt")
        .Left = 10
        .Top = 128
        .Width = 315
        .Height = 40
        .MultiLine = True
        .ScrollBars = 2 ' fmScrollBarsVertical
    End With

    ' --- Label для начальной ячейки
    .Controls.Add "Forms.Label.1", "lblStartCell", True
    With .Controls("lblStartCell")
        .Caption = "3. Укажите начальную ячейку для вывода:"
        .Left = 10
        .Top = 180
        .Width = 250
    End With

    ' --- TextBox для начальной ячейки
    .Controls.Add "Forms.TextBox.1", "txtStartCell", True
    With .Controls("txtStartCell")
        .Left = 10
        .Top = 198
        .Width = 100
        .Height = 20
    End With

    ' --- Кнопка "ОК"
    .Controls.Add "Forms.CommandButton.1", "btnOK", True
    With .Controls("btnOK")
        .Caption = "Запуск"
        .Left = 170
        .Top = 210
        .Width = 75
        .Height = 25
        .Default = True
    End With

    ' --- Кнопка "Отмена"
    .Controls.Add "Forms.CommandButton.1", "btnCancel", True
    With .Controls("btnCancel")
        .Caption = "Отмена"
        .Left = 250
        .Top = 210
        .Width = 75
        .Height = 25
        .Cancel = True
    End With

    ' Добавляем код в форму
    Dim codeMod As Object
    Set codeMod = vbComp.CodeModule

    Dim codeLines As Variant
    codeLines = GetFormCode()

    Dim i As Long
    For i = LBound(codeLines) To UBound(codeLines)
        codeMod.AddFromString codeLines(i)
    Next i

    Debug.Print "Форма '" & formName & "' успешно создана."
End Sub


Private Function GetMainModuleCode() As Variant
    '''
    ''' Возвращает код для основного модуля 'mdlChandraIntegration' в виде массива строк.
    '''
    Dim code As String
    code = "Attribute VB_Name = ""mdlChandraIntegration""" & vbCrLf & _
           "' Chandra Excel Integration Module" & vbCrLf & _
           "' Версия: 2.0" & vbCrLf & _
           "' Этот модуль содержит всю логику для взаимодействия с Chandra OCR COM-сервером." & vbCrLf & _
           "" & vbCrLf & _
           "Option Explicit" & vbCrLf & _
           "" & vbCrLf & _
           "' --- Глобальные переменные ---" & vbCrLf & _
           "Private Const COM_SERVER_NAME As String = ""ChandraExcel.Processor"" ' Имя нашего COM-сервера" & vbCrLf & _
           "Private g_Processor As Object ' Глобальный объект для COM-сервера, чтобы не создавать его каждый раз" & vbCrLf & _
           "" & vbCrLf & _
           "' --- Основные публичные процедуры ---" & vbCrLf & _
           "" & vbCrLf & _
           "Public Sub ShowChandraForm()" & vbCrLf & _
           "    ' Главная процедура, которая проверяет наличие COM-сервера и показывает пользователю форму." & vbCrLf & _
           "    On Error GoTo ErrorHandler" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Проверяем, что COM-сервер доступен" & vbCrLf & _
           "    If GetProcessor() Is Nothing Then" & vbCrLf & _
           "        MsgBox ""Не удалось подключиться к COM-серверу 'ChandraExcel.Processor'."" & vbCrLf & vbCrLf & _
           "               ""Пожалуйста, убедитесь, что вы запустили 'register_com.bat' от имени администратора."", vbCritical, ""Ошибка COM-сервера""" & vbCrLf & _
           "        Exit Sub" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Показываем форму" & vbCrLf & _
           "    frmChandraOCR.Show" & vbCrLf & _
           "    " & vbCrLf & _
           "    Exit Sub" & vbCrLf & _
           "ErrorHandler:" & vbCrLf & _
           "    MsgBox ""Произошла непредвиденная ошибка при запуске формы: "" & Err.Description, vbCritical, ""Критическая ошибка""" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "" & vbCrLf & _
           "Public Sub ProcessDocument(ByVal filePath As String, ByVal prompt As String, ByVal templateName As String, ByVal startCell As String)" & vbCrLf & _
           "    ' Основная логика: вызывает метод COM-сервера и обрабатывает результат." & vbCrLf & _
           "    On Error GoTo ErrorHandler" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim result As String" & vbCrLf & _
           "    Dim json As Object" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' --- Выполнение запроса ---" & vbCrLf & _
           "    Application.StatusBar = ""Chandra OCR: Идет обработка документа... Это может занять некоторое время...""" & vbCrLf & _
           "    DoEvents" & vbCrLf & _
           "    " & vbCrLf & _
           "    If GetProcessor() Is Nothing Then GoTo ComError" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Вызываем метод Python COM-сервера" & vbCrLf & _
           "    If templateName <> """" Then" & vbCrLf & _
           "        result = g_Processor.ProcessDocument(filePath, """", templateName)" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        result = g_Processor.ProcessDocument(filePath, prompt, """")" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' --- Обработка результата ---" & vbCrLf & _
           "    ' Для парсинга JSON используется 'Scripting.Dictionary'" & vbCrLf & _
           "    ' Это требует подключения 'Microsoft Scripting Runtime' в Tools -> References, но мы используем позднее связывание" & vbCrLf & _
           "    Set json = JsonParse(result)" & vbCrLf & _
           "    " & vbCrLf & _
           "    If json Is Nothing Then" & vbCrLf & _
           "        MsgBox ""Не удалось распознать ответ от сервера. Ответ был: "" & vbCrLf & result, vbCritical, ""Ошибка формата данных""" & vbCrLf & _
           "        GoTo Cleanup" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Проверяем, есть ли ошибка в ответе" & vbCrLf & _
           "    If json.Exists(""error"") Then" & vbCrLf & _
           "        MsgBox ""Сервер вернул ошибку: "" & vbCrLf & vbCrLf & json(""error""), vbCritical, ""Ошибка обработки""" & vbCrLf & _
           "        GoTo Cleanup" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Записываем данные в Excel" & vbCrLf & _
           "    If json.Exists(""data"") Then" & vbCrLf & _
           "        WriteToExcel json(""data""), startCell, templateName" & vbCrLf & _
           "        MsgBox ""Документ успешно обработан и данные записаны в Excel!"", vbInformation, ""Готово""" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        MsgBox ""В ответе сервера отсутствуют данные для записи."", vbExclamation, ""Нет данных""" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    GoTo Cleanup" & vbCrLf & _
           "" & vbCrLf & _
           "ComError:" & vbCrLf & _
           "    MsgBox ""Потеряно соединение с COM-сервером. Пожалуйста, перезапустите Excel."", vbCritical, ""Ошибка COM-сервера""" & vbCrLf & _
           "    GoTo Cleanup" & vbCrLf & _
           "" & vbCrLf & _
           "ErrorHandler:" & vbCrLf & _
           "    MsgBox ""Произошла критическая ошибка при обработке: "" & vbCrLf & vbCrLf & Err.Description, vbCritical, ""Критическая ошибка""" & vbCrLf & _
           "" & vbCrLf & _
           "Cleanup:" & vbCrLf & _
           "    ' Очистка" & vbCrLf & _
           "    Application.StatusBar = False" & vbCrLf & _
           "    Set json = Nothing" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "" & vbCrLf & _
           "' --- Вспомогательные функции ---" & vbCrLf & _
           "" & vbCrLf & _
           "Private Sub WriteToExcel(ByVal data As Object, ByVal startCell As String, ByVal templateName As String)" & vbCrLf & _
           "    ' Записывает данные из Scripting.Dictionary в ячейки Excel." & vbCrLf & _
           "    On Error GoTo ErrorHandler" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim ws As Worksheet" & vbCrLf & _
           "    Set ws = ActiveSheet" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim startRange As Range" & vbCrLf & _
           "    Set startRange = ws.Range(startCell)" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim keys As Variant" & vbCrLf & _
           "    keys = data.keys" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim i As Long" & vbCrLf & _
           "    For i = 0 To data.Count - 1" & vbCrLf & _
           "        ' Записываем заголовок (ключ)" & vbCrLf & _
           "        ws.Cells(startRange.Row, startRange.Column + i).Value = keys(i)" & vbCrLf & _
           "        ' Записываем значение" & vbCrLf & _
           "        ws.Cells(startRange.Row + 1, startRange.Column + i).Value = data(keys(i))" & vbCrLf & _
           "    Next i" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Автоподбор ширины столбцов" & vbCrLf & _
           "    ws.Range(startRange, startRange.Offset(0, data.Count - 1)).EntireColumn.AutoFit" & vbCrLf & _
           "    " & vbCrLf & _
           "    Exit Sub" & vbCrLf & _
           "ErrorHandler:" & vbCrLf & _
           "    MsgBox ""Ошибка при записи данных в Excel: "" & Err.Description, vbCritical, ""Ошибка записи""" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "" & vbCrLf & _
           "Public Function GetTemplates() As Object" & vbCrLf & _
           "    ' Получает список шаблонов из COM-сервера." & vbCrLf & _
           "    On Error Resume Next" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim templatesJson As String" & vbCrLf & _
           "    If GetProcessor() Is Nothing Then Set GetTemplates = Nothing: Exit Function" & vbCrLf & _
           "    " & vbCrLf & _
           "    templatesJson = g_Processor.GetTemplates()" & vbCrLf & _
           "    " & vbCrLf & _
           "    If Err.Number <> 0 Then Set GetTemplates = Nothing: Exit Function" & vbCrLf & _
           "    " & vbCrLf & _
           "    Set GetTemplates = JsonParse(templatesJson)" & vbCrLf & _
           "End Function" & vbCrLf & _
           "" & vbCrLf & _
           "Private Function GetProcessor() As Object" & vbCrLf & _
           "    ' Возвращает экземпляр COM-объекта (создает, если его нет)." & vbCrLf & _
           "    If g_Processor Is Nothing Then" & vbCrLf & _
           "        On Error Resume Next" & vbCrLf & _
           "        Set g_Processor = CreateObject(COM_SERVER_NAME)" & vbCrLf & _
           "        On Error GoTo 0" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    Set GetProcessor = g_Processor" & vbCrLf & _
           "End Function" & vbCrLf & _
           "" & vbCrLf & _
           "Private Function JsonParse(ByVal jsonString As String) As Object" & vbCrLf & _
           "    ' Простой парсер JSON на VBA. Ожидает плоский JSON." & vbCrLf & _
           "    ' В реальном проекте лучше использовать библиотеку (напр., VBA-JSON)" & vbCrLf & _
           "    Dim dict As Object" & vbCrLf & _
           "    Set dict = CreateObject(""Scripting.Dictionary"")" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Удаляем крайние скобки" & vbCrLf & _
           "    jsonString = Mid(jsonString, 2, Len(jsonString) - 2)" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim pairs As Variant" & vbCrLf & _
           "    pairs = Split(jsonString, "","")" & vbCrLf & _
           "    " & vbCrLf & _
           "    Dim pair As Variant" & vbCrLf & _
           "    For Each pair In pairs" & vbCrLf & _
           "        Dim kv As Variant" & vbCrLf & _
           "        kv = Split(pair, "":"")" & vbCrLf & _
           "        " & vbCrLf & _
           "        Dim key As String, value As String" & vbCrLf & _
           "        key = Trim(Replace(kv(0), ""\"""", """"))" & vbCrLf & _
           "        value = Trim(Replace(kv(1), ""\"""", """"))" & vbCrLf & _
           "        " & vbCrLf & _
           "        dict(key) = value" & vbCrLf & _
           "    Next pair" & vbCrLf & _
           "    " & vbCrLf & _
           "    Set JsonParse = dict" & vbCrLf & _
           "End Function"

    GetMainModuleCode = Split(code, vbCrLf)
End Function

Private Function GetFormCode() As Variant
    '''
    ''' Возвращает код для формы 'frmChandraOCR' в виде массива строк.
    '''
    Dim code As String
    code = "Attribute VB_Name = ""frmChandraOCR""" & vbCrLf & _
           "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
           "Attribute VB_Creatable = False" & vbCrLf & _
           "Attribute VB_PredeclaredId = True" & vbCrLf & _
           "Attribute VB_Exposed = False" & vbCrLf & _
           "Option Explicit" & vbCrLf & _
           "" & vbCrLf & _
           "Private Sub UserForm_Initialize()" & vbCrLf & _
           "    ' Срабатывает при загрузке формы" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' 1. Заполняем ComboBox шаблонами" & vbCrLf & _
           "    Dim templates As Object" & vbCrLf & _
           "    Set templates = mdlChandraIntegration.GetTemplates()" & vbCrLf & _
           "    " & vbCrLf & _
           "    Me.cmbTemplate.Clear" & vbCrLf & _
           "    Me.cmbTemplate.AddItem """" ' Пустая строка для выбора кастомного промпта" & vbCrLf & _
           "    " & vbCrLf & _
           "    If Not templates Is Nothing And templates.Exists(""templates"") Then" & vbCrLf & _
           "        Dim templateName As Variant" & vbCrLf & _
           "        For Each templateName In templates(""templates"").keys" & vbCrLf & _
           "            Me.cmbTemplate.AddItem templateName" & vbCrLf & _
           "        Next templateName" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' 2. Устанавливаем начальную ячейку в активную" & vbCrLf & _
           "    If Not ActiveCell Is Nothing Then" & vbCrLf & _
           "        Me.txtStartCell.Value = ActiveCell.Address" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        Me.txtStartCell.Value = ""A1""" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' 3. Блокируем поле с кастомным промптом по умолчанию" & vbCrLf & _
           "    Me.txtCustomPrompt.Enabled = False" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "" & vbCrLf & _
           "Private Sub btnBrowse_Click()" & vbCrLf & _
           "    ' Открывает диалог выбора файла" & vbCrLf & _
           "    Dim fileDialog As FileDialog" & vbCrLf & _
           "    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)" & vbCrLf & _
           "    " & vbCrLf & _
           "    fileDialog.Title = ""Выберите PDF или файл изображения""" & vbCrLf & _
           "    fileDialog.Filters.Clear" & vbCrLf & _
           "    fileDialog.Filters.Add ""Документы"", ""*.pdf; *.png; *.jpg; *.jpeg; *.bmp"", 1" & vbCrLf & _
           "    " & vbCrLf & _
           "    If fileDialog.Show = -1 Then" & vbCrLf & _
           "        Me.txtFilePath.Value = fileDialog.SelectedItems(1)" & vbCrLf & _
           "    End If" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "" & vbCrLf & _
           "Private Sub cmbTemplate_Change()" & vbCrLf & _
           "    ' Включает/выключает поле для своего промпта" & vbCrLf & _
           "    If Me.cmbTemplate.Value = """" Then" & vbCrLf & _
           "        Me.txtCustomPrompt.Enabled = True" & vbCrLf & _
           "        Me.txtCustomPrompt.SetFocus" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        Me.txtCustomPrompt.Enabled = False" & vbCrLf & _
           "        Me.txtCustomPrompt.Value = """"" & vbCrLf & _
           "    End If" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "" & vbCrLf & _
           "Private Sub btnOK_Click()" & vbCrLf & _
           "    ' Валидация и запуск обработки" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' 1. Проверка пути к файлу" & vbCrLf & _
           "    If Me.txtFilePath.Value = """" Or Dir(Me.txtFilePath.Value) = """" Then" & vbCrLf & _
           "        MsgBox ""Пожалуйста, выберите существующий файл."", vbExclamation, ""Ошибка ввода""" & vbCrLf & _
           "        Me.btnBrowse.SetFocus" & vbCrLf & _
           "        Exit Sub" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' 2. Проверка промпта" & vbCrLf & _
           "    If Me.cmbTemplate.Value = """" And Me.txtCustomPrompt.Value = """" Then" & vbCrLf & _
           "        MsgBox ""Пожалуйста, выберите шаблон или введите свой промпт."", vbExclamation, ""Ошибка ввода""" & vbCrLf & _
           "        Me.txtCustomPrompt.SetFocus" & vbCrLf & _
           "        Exit Sub" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' 3. Проверка ячейки" & vbCrLf & _
           "    On Error Resume Next" & vbCrLf & _
           "    Dim tempRange As Range" & vbCrLf & _
           "    Set tempRange = ActiveSheet.Range(Me.txtStartCell.Value)" & vbCrLf & _
           "    If Err.Number <> 0 Then" & vbCrLf & _
           "        MsgBox ""Пожалуйста, введите корректный адрес ячейки (например, A1)."", vbExclamation, ""Ошибка ввода""" & vbCrLf & _
           "        Me.txtStartCell.SetFocus" & vbCrLf & _
           "        Exit Sub" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    On Error GoTo 0" & vbCrLf & _
           "    " & vbCrLf & _
           "    ' Прячем форму и запускаем обработку" & vbCrLf & _
           "    Me.Hide" & vbCrLf & _
           "    mdlChandraIntegration.ProcessDocument Me.txtFilePath.Value, Me.txtCustomPrompt.Value, Me.cmbTemplate.Value, Me.txtStartCell.Value" & vbCrLf & _
           "    Unload Me" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "" & vbCrLf & _
           "Private Sub btnCancel_Click()" & vbCrLf & _
           "    ' Закрытие формы" & vbCrLf & _
           "    Unload Me" & vbCrLf & _
           "End Sub"

    GetFormCode = Split(code, vbCrLf)
End Function

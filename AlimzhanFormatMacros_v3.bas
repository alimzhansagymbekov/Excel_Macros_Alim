Attribute VB_Name = "Module1"
Option Explicit
' ============================================================
' ALIMZHAN FORMAT MACROS  (v3)
'
' Ctrl+Shift+Q  >  Number
' Ctrl+Shift+T  >  Percentage
' Ctrl+Shift+R  >  Multiples
' Ctrl+Shift+A  >  Color code selected range
' Ctrl+Shift+Z  >  Restore original colors
' Ctrl+Shift+J  >  Format Chart 7x4 inches (big)
' Ctrl+Shift+M  >  Format Chart 5x4 inches (small)
'
' ЦВЕТА ШРИФТА (color coding):
'   Blue   RGB(0,0,255)    входы: числа, даты, TRUE/FALSE, число-как-текст
'   Green  RGB(58,131,44)  прямая ссылка на другой лист:
'                          =Sheet!A1, =+Sheet!A1, =-Sheet!A1, =Sheet!A1%, ='My Sheet'!C10%
'   Purple RGB(112,48,160) функции FactSet / CapIQ / Bloomberg
'   Black                  все остальные формулы (расчёты)
'   текст в обычных ячейках не трогается
'
' ИЗМЕНЕНИЯ В v3 (логика покраски/форматов НЕ менялась):
'   1. Привязка клавиш переписана через SetupShortcuts (надёжное переназначение).
'   2. Esc больше не прерывает макросы (Application.EnableCancelKey = xlDisabled)
'      — это убирает ошибку "Code execution has been interrupted".
'   3. ScreenUpdating ВСЕГДА восстанавливается, даже если что-то пошло не так
'      (раньше при прерывании экран мог "зависнуть" замороженным).
'   4. Workbook_Open сбрасывает старые назначения перед новыми, чтобы при
'      переоткрытии не оставалось висящего Ctrl+Shift+R (см. модуль ЭтаКнига).
' ============================================================

Public dictOriginalColors As Object

' Цвета как Long-константы (Const не умеет хранить вызов RGB())
Private Const CLR_INPUT As Long = vbBlue        ' RGB(0, 0, 255)
Private Const CLR_FORMULA As Long = vbBlack     ' RGB(0, 0, 0)
Private Const CLR_LINK As Long = 2917178        ' RGB(58, 131, 44)
Private Const CLR_PLUGIN As Long = 10498160     ' RGB(112, 48, 160)

' ============================================================
' ПРИВЯЗКА КЛАВИШ
' Вызывается из Workbook_Open. Вынесено сюда, чтобы переназначение
' можно было запустить вручную (Alt+F8 -> SetupShortcuts) без переоткрытия.
' Сначала очищаем старые назначения, затем ставим новые — иначе при
' повторном запуске на клавишах могут остаться дубли/старые версии.
' ============================================================
Public Sub SetupShortcuts()
    ' --- очистка (на случай повторного запуска или смены клавиш) ---
    Application.OnKey "^+q"
    Application.OnKey "^+t"
    Application.OnKey "^+r"
    Application.OnKey "^+x"     ' на случай, если раньше пробовал вешать мульты на X
    Application.OnKey "^+a"
    Application.OnKey "^+z"
    Application.OnKey "^+j"
    Application.OnKey "^+m"
    ' --- назначение ---
    Application.OnKey "^+q", "Format_Number"
    Application.OnKey "^+t", "Format_Percentage"
    Application.OnKey "^+r", "Format_Multiples"
    Application.OnKey "^+a", "IB_ColorCode_Selection"
    Application.OnKey "^+z", "IB_RestoreColors_Selection"
    Application.OnKey "^+j", "FormatChart_7x4"
    Application.OnKey "^+m", "FormatChart_5x4"
End Sub

' ============================================================
' ФОРМАТЫ ЧИСЕЛ
' ============================================================
Sub Format_Number()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
End Sub

Sub Format_Percentage()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.NumberFormat = "_(* 0.0%_);_(* (0.0%);_(* ""-""??_);_(@_)"
End Sub

Sub Format_Multiples()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.NumberFormat = "0.0""x"";(0.0""x"");_(* ""-""??_);_(@_)"
End Sub

' ============================================================
' COLOR CODING
' ============================================================
Sub IB_ColorCode_Selection()
    Application.EnableCancelKey = xlDisabled          ' (2) Esc не прерывает
    If TypeName(Selection) <> "Range" Then Exit Sub

    Dim targetRange As Range
    Set targetRange = Intersect(Selection, Selection.Worksheet.UsedRange)
    If targetRange Is Nothing Then Exit Sub

    On Error GoTo CleanFail                            ' (3) гарантированно вернуть ScreenUpdating
    Application.ScreenUpdating = False

    Dim numCells As Range, txtCells As Range, fmlCells As Range
    If targetRange.Cells.CountLarge = 1 Then
        Dim c As Range
        Set c = targetRange.Cells(1, 1)
        On Error Resume Next
        If c.HasFormula Then
            Set fmlCells = c
        ElseIf Not IsEmpty(c.Value) Then
            If IsNumeric(c.Value) Or IsDate(c.Value) Then Set numCells = c
        End If
        On Error GoTo CleanFail
    Else
        On Error Resume Next   ' SpecialCells кидает 1004, если категория пуста — это норм
        Set numCells = targetRange.SpecialCells(xlCellTypeConstants, xlNumbers + xlLogical)
        Set txtCells = targetRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        Set fmlCells = targetRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo CleanFail
    End If

    Dim workCells As Range
    Set workCells = AppendRange(AppendRange(numCells, txtCells), fmlCells)
    If workCells Is Nothing Then GoTo CleanDone

    ' Снимок для Ctrl+Shift+Z (откатывает ПОСЛЕДНЮЮ покраску)
    Set dictOriginalColors = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    On Error Resume Next
    For Each cell In workCells
        dictOriginalColors(cell.Address(External:=True)) = cell.Font.Color
    Next cell
    On Error GoTo CleanFail

    ' Покраска
    On Error Resume Next
    If Not numCells Is Nothing Then numCells.Font.Color = CLR_INPUT
    If Not txtCells Is Nothing Then
        For Each cell In txtCells
            If IsNumeric(cell.Value) Then cell.Font.Color = CLR_INPUT
        Next cell
    End If
    If Not fmlCells Is Nothing Then
        For Each cell In fmlCells
            If IsDirectSheetRef(cell.Formula) Then
                cell.Font.Color = CLR_LINK
            ElseIf IsPluginFormula(cell.Formula) Then
                cell.Font.Color = CLR_PLUGIN
            Else
                cell.Font.Color = CLR_FORMULA
            End If
        Next cell
    End If
    On Error GoTo CleanFail

CleanDone:
    Application.ScreenUpdating = True
    Exit Sub
CleanFail:
    Application.ScreenUpdating = True                 ' (3) экран не остаётся замороженным
End Sub

Sub IB_RestoreColors_Selection()
    Application.EnableCancelKey = xlDisabled          ' (2)
    If dictOriginalColors Is Nothing Then
        MsgBox "No saved colors found. Run Ctrl+Shift+A first.", vbInformation
        Exit Sub
    End If
    If dictOriginalColors.Count = 0 Then
        MsgBox "No saved colors found. Run Ctrl+Shift+A first.", vbInformation
        Exit Sub
    End If

    On Error GoTo CleanFail                            ' (3)
    Application.ScreenUpdating = False
    Dim key As Variant
    Dim cell As Range
    On Error Resume Next
    For Each key In dictOriginalColors.Keys
        Set cell = Range(key)
        If Not cell Is Nothing Then
            cell.Font.Color = dictOriginalColors(key)
        End If
    Next key
    On Error GoTo CleanFail
    dictOriginalColors.RemoveAll
    Application.ScreenUpdating = True
    Exit Sub
CleanFail:
    Application.ScreenUpdating = True
End Sub

' ============================================================
' ВСПОМОГАТЕЛЬНЫЕ
' ============================================================
' % на конце допускается, regex компилируется один раз (Static)
Function IsDirectSheetRef(strFormula As String) As Boolean
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "^=[+\-]?('[^!\[\]]+'|[^'!\[\]]+)!\$?[A-Za-z]{1,3}\$?[0-9]+(:\$?[A-Za-z]{1,3}\$?[0-9]+)?%?$"
        rx.IgnoreCase = True
    End If
    IsDirectSheetRef = rx.Test(strFormula)
End Function

' FactSet / CapIQ / Bloomberg — открывающая скобка отсекает ложные совпадения
Function IsPluginFormula(strFormula As String) As Boolean
    Static plugins As Variant
    If IsEmpty(plugins) Then
        plugins = Array("FDS(", "CIQ(", "CIQRANGE(", "CIQINDEX(", "BDP(", "BDH(", "BDS(", "BQL(")
    End If
    Dim f As String
    f = UCase$(CleanFormula(strFormula))   ' убираем строки в кавычках, чтобы не ловить имена плагинов внутри текста
    Dim i As Long
    For i = LBound(plugins) To UBound(plugins)
        If InStr(f, plugins(i)) > 0 Then
            IsPluginFormula = True
            Exit Function
        End If
    Next i
End Function

Function CleanFormula(strIn As String) As String
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = """[^""]*"""
        rx.Global = True
    End If
    CleanFormula = rx.Replace(strIn, "")
End Function

' Объединение диапазонов с учётом Nothing (безопасно)
Private Function AppendRange(baseRng As Range, addRng As Range) As Range
    If baseRng Is Nothing Then
        Set AppendRange = addRng
    ElseIf addRng Is Nothing Then
        Set AppendRange = baseRng
    Else
        Set AppendRange = Union(baseRng, addRng)
    End If
End Function

' ============================================================
' CHART FORMAT
' ============================================================
Sub FormatChart_7x4()
    FormatChartCore 7, 4
End Sub

Sub FormatChart_5x4()
    FormatChartCore 5, 4
End Sub

Private Sub FormatChartCore(widthInches As Double, heightInches As Double)
    Application.EnableCancelKey = xlDisabled          ' (2)
    Dim cht As Chart
    If ActiveChart Is Nothing Then
        If TypeName(Selection) = "ChartObject" Then
            Set cht = Selection.Chart
        Else
            MsgBox "Please click on the chart first, then run the macro.", vbExclamation
            Exit Sub
        End If
    Else
        Set cht = ActiveChart
    End If

    cht.Parent.Width = widthInches * 72
    cht.Parent.Height = heightInches * 72
    If cht.HasTitle Then cht.ChartTitle.Delete
    On Error Resume Next
    cht.Axes(xlValue).MajorGridlines.Delete
    cht.Axes(xlValue).MinorGridlines.Delete
    cht.Axes(xlCategory).MajorGridlines.Delete
    cht.Axes(xlCategory).MinorGridlines.Delete
    On Error GoTo 0
    cht.ChartArea.Border.LineStyle = xlNone
    cht.PlotArea.Border.LineStyle = xlNone
    cht.PlotArea.Format.Fill.Visible = msoFalse
    With cht.Axes(xlCategory)
        .Border.LineStyle = xlNone
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 10
    End With
    With cht.Axes(xlValue)
        .Border.LineStyle = xlNone
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 10
    End With
    cht.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Arial"
    If cht.HasLegend Then
        With cht.Legend.Font
            .Name = "Arial"
            .Size = 9
        End With
    End If
    ' прямые линии — убрать "сглаживание" у каждого ряда
    Dim srs As Series
    On Error Resume Next
    For Each srs In cht.SeriesCollection
        srs.Smooth = False
    Next srs
    On Error GoTo 0

    ' статус-бар вместо MsgBox, очистка через 3 секунды
    Application.StatusBar = "Chart formatted: " & widthInches & " x " & heightInches & " in"
    On Error Resume Next   ' (мелочь) если OnTime не зарегистрируется — не критично
    Application.OnTime Now + TimeSerial(0, 0, 3), "IB_ClearStatusBar"
    On Error GoTo 0
End Sub

Sub IB_ClearStatusBar()
    On Error Resume Next   ' (мелочь) книга могла закрыться — молча игнорируем
    Application.StatusBar = False
    On Error GoTo 0
End Sub

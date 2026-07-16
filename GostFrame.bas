Attribute VB_Name = "GostFrame"
Option Explicit

'=====================================================================
'  Печать таблицы с чертёжной рамкой - Excel 2016
'
'  * Поля листа по ГОСТ: слева 20 мм, сверху/справа/снизу 5 мм.
'    Рамка чертится от края рабочей области листа (по границе полей),
'    поэтому таблица, начинающаяся с ячейки A1, сразу попадает внутрь
'    рамки.
'  * Основная надпись - форма 5 по ГОСТ 21.101 (185х40 мм) на первом
'    листе (по желанию).
'  * Длинная таблица автоматически разбивается на страницы, каждая
'    страница получает свою рамку. Строки, которые попали бы под
'    штамп, переносятся на следующий лист.
'
'  Печать выполняется через временную копию листа, поэтому исходный
'  лист НЕ изменяется.
'
'  Макросы для запуска (Alt+F8):
'    PrintWithFrame - выбор формата (A4/A3), ориентации, штампа,
'                     печать или предварительный просмотр
'    RemoveFrame    - удаление фигур рамки/штампа с активного листа
'                     (для очистки после старых версий макроса)
'=====================================================================

Private Const PFX As String = "GOST_"           ' префикс имён фигур
Private Const TMP_SHEET As String = "GOST_PRINT_TMP"

' поля листа по ГОСТ, мм
Private Const FLD_LEFT_MM As Double = 20
Private Const FLD_OTHER_MM As Double = 5

' основная надпись: форма 5 по ГОСТ 21.101, мм
Private Const STAMP_W As Double = 185
Private Const STAMP_H As Double = 40
Private Const STAMP_GAP As Double = 3           ' зазор таблица-штамп, мм

' толщина линий, пт (основная / тонкая)
Private Const W_MAIN As Double = 2#
Private Const W_THIN As Double = 0.5

Private mSeq As Long                            ' счётчик имён фигур

'---------------------------------------------------------------------
' мм -> пункты
Private Function MM(ByVal mmVal As Double) As Double
    MM = Application.CentimetersToPoints(mmVal / 10)
End Function

Private Function NextName() As String
    mSeq = mSeq + 1
    NextName = PFX & "s" & mSeq
End Function

'---------------------------------------------------------------------
Public Sub PrintWithFrame()
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Активный лист не является рабочим листом.", vbExclamation
        Exit Sub
    End If
    Dim src As Worksheet
    Set src = ActiveSheet
    If src.Name = TMP_SHEET Then
        MsgBox "Активен временный лист печати. Откройте лист с таблицей.", vbExclamation
        Exit Sub
    End If

    ' --- выбор формата ----------------------------------------------
    Dim ans As String
    ans = UCase$(Trim$(InputBox( _
        "Укажите формат листа для печати: A4 или A3", _
        "Чертёжная рамка", "A4")))
    If Len(ans) = 0 Then Exit Sub               ' нажата "Отмена"

    Dim wMM As Double, hMM As Double, paper As XlPaperSize
    Select Case ans
        Case "A4", "А4": wMM = 210: hMM = 297: paper = xlPaperA4
        Case "A3", "А3": wMM = 297: hMM = 420: paper = xlPaperA3
        Case Else
            MsgBox "Неизвестный формат: " & ans & vbCrLf & _
                   "Допустимые значения: A4, A3.", vbExclamation
            Exit Sub
    End Select

    ' --- ориентация -------------------------------------------------
    Dim landscape As Boolean
    landscape = (MsgBox("Альбомная ориентация листа?" & vbCrLf & _
        "Да - альбомная, Нет - книжная.", _
        vbYesNo + vbQuestion, "Ориентация") = vbYes)
    If landscape Then
        Dim t As Double
        t = wMM: wMM = hMM: hMM = t
    End If

    ' --- основная надпись -------------------------------------------
    Dim withStamp As Boolean
    withStamp = (MsgBox("Добавить основную надпись " & _
        "(форма 5 по ГОСТ 21.101, 185х40 мм) на первый лист?", _
        vbYesNo + vbQuestion, "Основная надпись") = vbYes)

    ' размеры рабочей области страницы (внутри полей), пт
    Dim cW As Double, cH As Double
    cW = MM(wMM - FLD_LEFT_MM - FLD_OTHER_MM)
    cH = MM(hMM - 2 * FLD_OTHER_MM)

    On Error GoTo errH
    Application.ScreenUpdating = False
    mSeq = 0

    ' --- временная копия листа --------------------------------------
    KillTmpSheet
    src.Copy After:=src
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Name = TMP_SHEET
    DeleteFrameShapes ws                        ' старые рамки, если были
    ws.ResetAllPageBreaks

    ' --- границы содержимого ----------------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

    ' столбцы области печати - на всю ширину рабочей области
    Dim lastCol As Long, accW As Double
    accW = 0: lastCol = 0
    Do While lastCol < 16000
        If accW + ws.Columns(lastCol + 1).Width > cW + 0.5 Then Exit Do
        lastCol = lastCol + 1
        accW = accW + ws.Columns(lastCol).Width
    Loop
    If lastCol = 0 Then lastCol = 1

    ' --- разбивка на страницы (проход 1: где нужны вставки) ---------
    ' Каждая страница, кроме последней, добивается пустой строкой до
    ' полной высоты рабочей области, чтобы рамки страниц совпадали с
    ' границами бумаги. На первой странице при штампе резервируется
    ' нижняя зона 40+3 мм.
    Dim stampReserve As Double
    If withStamp Then stampReserve = MM(STAMP_H + STAMP_GAP)

    Dim fills As Collection
    Set fills = New Collection                  ' Array(строка, высота вставки)
    Dim r As Long, acc As Double, limit As Double, rowH As Double
    limit = cH - stampReserve                   ' лимит первой страницы
    acc = 0: r = 1
    Do While r <= lastRow
        rowH = ws.Rows(r).Height
        If acc + rowH > limit + 0.3 And acc > 0 Then
            fills.Add Array(r - 1, cH - acc)    ' добить страницу до cH
            limit = cH                          ' следующие страницы - без штампа
            acc = 0
        Else
            acc = acc + rowH
            r = r + 1
        End If
    Loop

    ' вставка пустых строк-заполнителей (снизу вверх)
    Dim i As Long, itm As Variant, inserted As Long
    For i = fills.Count To 1 Step -1
        itm = fills(i)
        inserted = inserted + InsertFiller(ws, CLng(itm(0)), CDbl(itm(1)))
    Next i
    lastRow = lastRow + inserted

    ' --- проход 2: разрывы страниц и рамки ---------------------------
    Dim pages As Collection
    Set pages = New Collection                  ' Array(верх страницы, высота)
    Dim pTop As Double
    acc = 0: pTop = 0
    For r = 1 To lastRow
        rowH = ws.Rows(r).Height
        If acc + rowH > cH + 0.3 And acc > 0 Then
            pages.Add Array(pTop, acc)
            ws.Rows(r).PageBreak = xlPageBreakManual
            pTop = ws.Rows(r).Top
            acc = 0
        End If
        acc = acc + rowH
    Next r
    pages.Add Array(pTop, acc)                  ' последняя страница

    ' рамка на каждой странице
    Dim k As Long, pH As Double
    For k = 1 To pages.Count
        itm = pages(k)
        pTop = itm(0)
        If k = pages.Count Then
            pH = cH                             ' последняя - на всю высоту
        Else
            pH = itm(1) - 0.75                  ' чтобы нижняя линия не ушла на следующий лист
        End If
        DrawFrameRect ws, pTop, cW, pH
    Next k

    ' штамп на первой странице, вплотную к нижней линии рамки
    If withStamp Then
        itm = pages(1)
        Dim sTop As Double
        If pages.Count = 1 Then
            sTop = cH - MM(STAMP_H)
        Else
            sTop = itm(0) + itm(1) - 0.75 - MM(STAMP_H)
        End If
        DrawStamp ws, cW - MM(STAMP_W), sTop
    End If

    ' область печати - до полной высоты последней страницы, чтобы
    ' рамка и штамп гарантированно попали в печать
    Dim lastTop As Double, printLastRow As Long
    itm = pages(pages.Count)
    lastTop = itm(0)
    printLastRow = lastRow
    Do While printLastRow < 100000
        With ws.Rows(printLastRow + 1)
            If .Top + .Height > lastTop + cH - 1 Then Exit Do
        End With
        printLastRow = printLastRow + 1
    Loop

    ' --- параметры страницы ------------------------------------------
    With ws.PageSetup
        .PaperSize = paper
        .Orientation = IIf(landscape, xlLandscape, xlPortrait)
        .LeftMargin = MM(FLD_LEFT_MM)           ' поля по ГОСТ
        .RightMargin = MM(FLD_OTHER_MM)
        .TopMargin = MM(FLD_OTHER_MM)
        .BottomMargin = MM(FLD_OTHER_MM)
        .HeaderMargin = 0: .FooterMargin = 0
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "": .CenterFooter = "": .RightFooter = ""
        .CenterHorizontally = False
        .CenterVertically = False
        .Zoom = 100                             ' печать строго 1:1
        .PrintArea = ws.Range(ws.Cells(1, 1), _
            ws.Cells(printLastRow, lastCol)).Address
    End With

    Application.ScreenUpdating = True

    ' --- печать / предварительный просмотр ---------------------------
    If MsgBox("Листов к печати: " & pages.Count & vbCrLf & vbCrLf & _
        "Отправить на печать сразу?" & vbCrLf & _
        "Нет - открыть предварительный просмотр.", _
        vbYesNo + vbQuestion, "Печать") = vbYes Then
        ws.PrintOut
    Else
        ws.PrintPreview
    End If

    KillTmpSheet
    src.Activate
    Exit Sub

errH:
    Application.ScreenUpdating = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Чертёжная рамка"
    KillTmpSheet
    src.Activate
End Sub

'---------------------------------------------------------------------
' Удаление всех фигур рамки/штампа с активного листа
Public Sub RemoveFrame()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    DeleteFrameShapes ActiveSheet
End Sub

Private Sub DeleteFrameShapes(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PFX)) = PFX Then ws.Shapes(i).Delete
    Next i
End Sub

'---------------------------------------------------------------------
' Удаление временного листа печати
Private Sub KillTmpSheet()
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(TMP_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------
' Вставка пустых строк-заполнителей общей высотой gap (пт) после
' строки afterRow. Возвращает число вставленных строк.
Private Function InsertFiller(ws As Worksheet, ByVal afterRow As Long, _
    ByVal gap As Double) As Long
    Dim remain As Double, h As Double, n As Long
    remain = gap - 1                            ' запас от переполнения страницы
    Do While remain > 0.5
        h = remain
        If h > 400 Then h = 400                 ' предел высоты строки Excel
        ws.Rows(afterRow + 1).Insert Shift:=xlDown
        ws.Rows(afterRow + 1).Clear
        ws.Rows(afterRow + 1).RowHeight = h
        remain = remain - h
        n = n + 1
    Loop
    InsertFiller = n
End Function

'---------------------------------------------------------------------
' Рамка одной страницы: от края рабочей области листа
Private Sub DrawFrameRect(ws As Worksheet, ByVal topPt As Double, _
    ByVal wPt As Double, ByVal hPt As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, topPt, wPt, hPt)
    With shp
        .Name = NextName()
        .Fill.Visible = msoFalse
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With
End Sub

'---------------------------------------------------------------------
' Основная надпись: форма 5 по ГОСТ 21.101 (185х40 мм).
' x0, y0 - левый верхний угол штампа в пунктах.
Private Sub DrawStamp(ws As Worksheet, ByVal x0 As Double, ByVal y0 As Double)
    Dim startIdx As Long
    startIdx = ws.Shapes.Count

    ' наружный контур
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x0, y0, MM(STAMP_W), MM(STAMP_H))
    With shp
        .Name = NextName()
        .Fill.ForeColor.RGB = vbWhite           ' маскирует линии сетки под штампом
        .Fill.Visible = msoTrue
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With

    ' --- левый блок 65 мм: таблица изменений и подписи ---------------
    ' колонки: Изм.(10) Кол.уч.(10) Лист(10) № док.(10) Подп.(15) Дата(10)
    StampLine ws, x0, y0, 10, 0, 10, 15, W_MAIN     ' только зона изменений
    StampLine ws, x0, y0, 30, 0, 30, 15, W_MAIN
    StampLine ws, x0, y0, 20, 0, 20, 40, W_MAIN
    StampLine ws, x0, y0, 40, 0, 40, 40, W_MAIN
    StampLine ws, x0, y0, 55, 0, 55, 40, W_MAIN
    StampLine ws, x0, y0, 65, 0, 65, 40, W_MAIN
    ' строки зоны изменений и заголовок граф
    StampLine ws, x0, y0, 0, 5, 65, 5, W_THIN
    StampLine ws, x0, y0, 0, 10, 65, 10, W_THIN
    StampLine ws, x0, y0, 0, 15, 185, 15, W_MAIN    ' общая линия с графой обозначения
    ' строки подписей
    StampLine ws, x0, y0, 0, 20, 65, 20, W_THIN
    StampLine ws, x0, y0, 0, 25, 65, 25, W_THIN
    StampLine ws, x0, y0, 0, 30, 65, 30, W_THIN
    StampLine ws, x0, y0, 0, 35, 65, 35, W_THIN
    ' заголовки граф таблицы изменений
    StampText ws, x0, y0, 0, 10, 10, 5, "Изм.", 6
    StampText ws, x0, y0, 10, 10, 10, 5, "Кол.уч.", 6
    StampText ws, x0, y0, 20, 10, 10, 5, "Лист", 6
    StampText ws, x0, y0, 30, 10, 10, 5, "№ док.", 6
    StampText ws, x0, y0, 40, 10, 15, 5, "Подп.", 6
    StampText ws, x0, y0, 55, 10, 10, 5, "Дата", 6
    ' подписи (графы 10-13 - должности по усмотрению организации)
    StampText ws, x0, y0, 0, 15, 20, 5, "Разраб.", 7
    StampText ws, x0, y0, 0, 20, 20, 5, "Пров.", 7
    StampText ws, x0, y0, 0, 30, 20, 5, "Н.контр.", 7
    StampText ws, x0, y0, 0, 35, 20, 5, "Утв.", 7

    ' --- средняя часть: обозначение (65..185 x 0..15) и наименование --
    ' (обозначение и наименование заполняются вручную)

    ' --- правый блок 50 мм: Стадия / Лист / Листов, организация -------
    StampLine ws, x0, y0, 135, 15, 135, 40, W_MAIN
    StampLine ws, x0, y0, 150, 15, 150, 30, W_MAIN
    StampLine ws, x0, y0, 165, 15, 165, 30, W_MAIN
    StampLine ws, x0, y0, 135, 20, 185, 20, W_THIN
    StampLine ws, x0, y0, 135, 30, 185, 30, W_MAIN
    StampText ws, x0, y0, 135, 15, 15, 5, "Стадия", 6
    StampText ws, x0, y0, 150, 15, 15, 5, "Лист", 6
    StampText ws, x0, y0, 165, 15, 20, 5, "Листов", 6

    ' --- группировка фигур штампа в одну ------------------------------
    Dim cnt As Long
    cnt = ws.Shapes.Count - startIdx
    If cnt > 1 Then
        Dim arr() As Variant, i As Long
        ReDim arr(0 To cnt - 1)
        For i = 1 To cnt
            arr(i - 1) = ws.Shapes(startIdx + i).Name
        Next i
        With ws.Shapes.Range(arr).Group
            .Name = PFX & "Stamp"
            .Placement = xlFreeFloating
        End With
    End If
End Sub

'---------------------------------------------------------------------
' Линия штампа: координаты в мм относительно левого верхнего угла штампа
Private Sub StampLine(ws As Worksheet, ByVal x0 As Double, ByVal y0 As Double, _
    ByVal x1 As Double, ByVal y1 As Double, _
    ByVal x2 As Double, ByVal y2 As Double, ByVal wt As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddLine(x0 + MM(x1), y0 + MM(y1), x0 + MM(x2), y0 + MM(y2))
    With shp
        .Name = NextName()
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = wt
        .Placement = xlFreeFloating
    End With
End Sub

'---------------------------------------------------------------------
' Надпись штампа: координаты и размеры в мм относительно угла штампа
Private Sub StampText(ws As Worksheet, ByVal x0 As Double, ByVal y0 As Double, _
    ByVal x1 As Double, ByVal y1 As Double, _
    ByVal wd As Double, ByVal ht As Double, _
    ByVal txt As String, ByVal sz As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        x0 + MM(x1), y0 + MM(y1), MM(wd), MM(ht))
    With shp
        .Name = NextName()
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        With .TextFrame
            .MarginLeft = 0: .MarginRight = 0
            .MarginTop = 0: .MarginBottom = 0
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .Characters.Text = txt
            With .Characters.Font
                .Name = "Arial"
                .Size = sz
                .Color = vbBlack
                .Bold = False
            End With
        End With
    End With
End Sub

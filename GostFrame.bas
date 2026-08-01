Attribute VB_Name = "GostFrame"
Option Explicit

'=====================================================================
'  Печать таблицы с чертёжной рамкой - Excel 2016
'
'  * Формат всегда A3, альбомная ориентация.
'  * Поля листа по ГОСТ: слева 20 мм, сверху/справа/снизу 5 мм.
'  * Рамка чертится от края рабочей области листа.
'  * Основная надпись - форма 6 по ГОСТ 21.101-2026 (185x15 мм)
'    на КАЖДОМ листе, у нижней линии рамки.
'  * Область печати - столбцы A..Q, число строк считается по
'    содержимому листа.
'  * Длинная таблица разбивается на страницы, каждая получает
'    свою рамку и свой штамп.
'
'  Печать идёт через временную копию листа - исходный лист НЕ
'  изменяется.
'
'  Макросы для запуска (Alt+F8):
'    PrintWithFrame - построить рамки, штампы и напечатать / открыть
'                     предварительный просмотр
'    RemoveFrame    - удалить фигуры рамки/штампа с активного листа
'=====================================================================

Private Const PFX As String = "GOST_"
Private Const TMP_SHEET As String = "КЖ"

' Данные для основной надписи
Private Const cipher As String = "1.2.1.NT307R.00.001.AK04.GK01"
Private Const Date0 As String = "30.07.26"
Private CountPageAll As Double ' общее количество листов
Private Const PAGE_OFFSET As Long = 3 ' число предшествующих листов

' поля листа по ГОСТ, мм
Private Const FLD_LEFT_MM As Double = 20
Private Const FLD_OTHER_MM As Double = 5

' формат печати - всегда A3 альбомная, мм
Private Const SHEET_W_MM As Double = 420
Private Const SHEET_H_MM As Double = 297

' поправка ширины рамки, мм. Рамка сужается на это значение, а штамп
' привязан к её правому краю и смещается на столько же. Если на вашем
' принтере рамка всё ещё вылезает или, наоборот, не достаёт до края -
' поменяйте только это число.
Private Const FRAME_ADJUST_MM As Double = 28

' нижний отступ рамки, мм (0 = рамка до самого нижнего поля)
Private Const FRAME_BOTTOM_MM As Double = 0

' основная надпись: форма 6 по ГОСТ 21.101-2026, мм
Private Const STAMP_W As Double = 185
Private Const STAMP_H As Double = 15
Private Const STAMP_GAP As Double = 1           ' зазор таблица-штамп

' столбцы области печати: A (1) .. Q (17)
Private Const FIRST_COL As Long = 1
Private Const LAST_COL As Long = 17

' заголовок таблицы, повторяемый на каждом листе (строки 1..N)
Private Const HEADER_ROWS As Long = 4

' --- формирование спецификации ---------------------------------
Private Const SPEC_SRC_SHEET As String = "Таблица кабелей"
Private Const SPEC_DST_SHEET As String = "Спецификация"
Private Const SPEC_TYPE_SRC_COL As Long = 5   ' E - тип кабеля
Private Const SPEC_LEN_SRC_COL As Long = 17   ' Q - длина, м
Private Const SPEC_START_ROW As Long = 2      ' первая строка вывода
Private Const SPEC_TYPE_DST_COL As Long = 1   ' куда писать тип (A)
Private Const SPEC_LEN_DST_COL As Long = 2    ' куда писать сумму (B)

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
    On Error GoTo errH
    Dim src As Worksheet
    Set src = Worksheets("Таблица кабелей")

    ' рабочая область страницы (внутри полей), пт
    Dim cW As Double, cH As Double
    cW = MM(SHEET_W_MM - FLD_LEFT_MM - FLD_OTHER_MM)
    cH = MM(SHEET_H_MM) 'cH = MM(SHEET_H_MM - 1 * FLD_OTHER_MM)

    ' ширина рамки с поправкой; штамп привязан к её правому краю
    Dim frameW As Double
    frameW = cW - MM(FRAME_ADJUST_MM)
    Dim frameH As Double
    frameH = cH - MM(FRAME_BOTTOM_MM)

    Application.ScreenUpdating = False
    mSeq = 0

    ' --- временная копия листа --------------------------------------
    KillTmpSheet
    src.Copy After:=src
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Name = TMP_SHEET
    DeleteFrameShapes ws
    ws.ResetAllPageBreaks
    ws.Tab.Color = RGB(146, 208, 80)
    ws.Columns(1).Insert Shift:=xlToRight
    ws.Columns(1).ColumnWidth = 10

    ' --- последняя строка содержимого -------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

    ' --- разбивка на страницы (проход 1: планирование) --------------
    ' На каждой странице сверху повторяется заголовок (строки 1..
    ' HEADER_ROWS), снизу резервируется зона под штамп, а хвост
    ' страницы добивается пустой строкой до полной высоты листа.
    Dim stampReserve As Double
    stampReserve = MM(STAMP_H + STAMP_GAP)

    ' высота заголовка (строки 1..HEADER_ROWS)
    Dim hHead As Double, hr As Long
    hHead = 0
    For hr = 1 To HEADER_ROWS
        hHead = hHead + ws.Rows(hr).Height
    Next hr

    ' границы страниц: r - первая строка данных следующей страницы,
    ' fillH - чем добить закрываемую страницу до полной высоты листа
    Dim brks As Collection
    Set brks = New Collection       ' Array(r, fillH)
    Dim r As Long, acc As Double, limit As Double, rowH As Double
    limit = cH - hHead - stampReserve
    acc = 0: r = HEADER_ROWS + 1
    Do While r <= lastRow
        rowH = ws.Rows(r).Height
        If acc + rowH > limit + 0.3 And acc > 0 Then
            brks.Add Array(r, cH - hHead - acc)
            acc = 0
        Else
            acc = acc + rowH
            r = r + 1
        End If
    Loop

    ' вставка снизу вверх: на каждой границе - копия заголовка сверху
    ' новой страницы и заполнитель снизу предыдущей
    Dim i As Long, itm As Variant
    For i = brks.Count To 1 Step -1
        itm = brks(i)
        InsertHeaderCopy ws, CLng(itm(0)), HEADER_ROWS
        InsertFiller ws, CLng(itm(0)) - 1, CDbl(itm(1))
    Next i
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    ' --- проход 2: разрывы страниц, рамки, штампы --------------------
    Dim pages As Collection
    Set pages = New Collection      ' верх страницы, пт
    Dim pTop As Double
    acc = 0: pTop = 0
    For r = 1 To lastRow
        rowH = ws.Rows(r).Height
        If acc + rowH > cH + 0.3 And acc > 0 Then
            pages.Add pTop
            ws.Rows(r).PageBreak = xlPageBreakManual
            pTop = ws.Rows(r).Top
            acc = 0
        End If
        acc = acc + rowH
    Next r
    pages.Add pTop

    Dim k As Long, thisH As Double, sTopa As Double, sTopb As Double, thisHDop As Double, CntPgs As Double
    For k = 1 To pages.Count
        pTop = pages(k)
        ' нижняя линия чуть выше края, чтобы не ушла на след. лист
        If k = pages.Count Then
            thisH = frameH - 2
            thisHDop = thisH
        Else
            thisH = frameH ' - 0.75
            thisHDop = frameH - 2.5
        End If
        DrawFrameRect ws, pTop, frameW, thisH
        ' штамп формы 6 у нижней линии рамки
        sTopa = pTop + thisH - MM(STAMP_H)
        sTopb = pTop + thisHDop - MM(85)
        CntPgs = k + PAGE_OFFSET ' номер текущего листа
        DrawStampForm6 ws, frameW - MM(STAMP_W - 20), sTopa, CntPgs
        DrawStampDopS ws, MM(8), sTopb
    Next k
    CountPageAll = pages.Count + PAGE_OFFSET ' общее количество листов
    ' --- область печати ---------------------------------------------
    ' Область печати должна полностью накрывать рамку и штамп, иначе
    ' Excel обрежет их нижнюю и правую границы. Поэтому строки тянем
    ' чуть ниже нижней линии рамки, а столбцы (начиная с Q) - вправо
    ' до правого края рамки.
    Dim lastTop As Double, printLastRow As Long
    lastTop = pages(pages.Count)
    printLastRow = lastRow
    Do While printLastRow < 1000000
        If ws.Rows(printLastRow + 1).Top _
            >= lastTop + frameH + MM(3) Then Exit Do
        printLastRow = printLastRow + 1
    Loop


    ' --- параметры страницы -----------------------------------------
    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .LeftMargin = 0 'MM(FLD_LEFT_MM)
        .RightMargin = 0 'MM(FLD_OTHER_MM)
        .TopMargin = MM(FLD_OTHER_MM)
        .BottomMargin = 0 'MM(FLD_OTHER_MM)
        .HeaderMargin = 0: .FooterMargin = 0
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "":
        .CenterFooter = "&""Times New Roman""&10&F"
        .RightFooter = "&""Times New Roman""&10Формат А3" & space(20)
        .CenterHorizontally = False
        .CenterVertically = False
        .Zoom = 100
        .PrintArea = ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(printLastRow - 1, 18)).Address
    End With
    Application.PrintCommunication = True
    Application.ScreenUpdating = True
    
    src.Activate
    Exit Sub


errH:
    Application.ScreenUpdating = True
    Application.PrintCommunication = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation, _
        "Чертёжная рамка"
    On Error Resume Next
    KillTmpSheet
    src.Activate
End Sub

'---------------------------------------------------------------------
Public Sub PrintWithFrameAnn()
    On Error GoTo errH
    If CountPageAll = 0 Then
        MsgBox "Сначала запустите PrintWithFrame - " & _
            "не определено общее число листов.", vbExclamation
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = Worksheets("Аннотация")

    ' рабочая область страницы (внутри полей), пт
    Dim cW As Double, cH As Double
    cW = MM(210 - FLD_LEFT_MM - FLD_OTHER_MM)
    cH = MM(297 - 1.5)

    ' ширина рамки с поправкой; штамп привязан к её правому краю
    Dim frameW As Double
    frameW = cW - MM(0)
    Dim frameH As Double
    frameH = cH - MM(FRAME_BOTTOM_MM)

    Application.ScreenUpdating = False
    mSeq = 0

    ' --- временная копия листа --------------------------------------
    DeleteFrameShapes ws
    ws.ResetAllPageBreaks
    ws.Tab.Color = RGB(146, 208, 80)

    ' --- последняя строка содержимого -------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
        
    ' --- рисование рамки, штампы --------------------
    Dim sTopa As Double
    DrawFrameRect ws, 0, frameW, frameH
    ' штамп формы 6 у нижней линии рамки
    sTopa = 0 + frameH - MM(40)
    DrawStampForm5 ws, MM(20), sTopa, CountPageAll
    DrawStampDopS ws, MM(8), frameH - MM(85)
    DrawStampDopB ws, MM(5), frameH - MM(85 + 65)

    Dim printLastRow As Long
    printLastRow = lastRow
    Do While printLastRow < 1000000
        If ws.Rows(printLastRow + 1).Top _
            >= frameH + MM(3) Then Exit Do
        printLastRow = printLastRow + 1
    Loop
    
    ' --- параметры страницы -----------------------------------------
    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = MM(FLD_OTHER_MM)
        .BottomMargin = MM(FLD_OTHER_MM)
        .HeaderMargin = 0: .FooterMargin = 0
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "":
        .CenterFooter = "&""Times New Roman""&10&F"
        .RightFooter = "&""Times New Roman""&10Формат А4" & space(20)
        .CenterHorizontally = False
        .CenterVertically = False
        .Zoom = 100
        .PrintArea = ws.Range( _
            ws.Cells(1, FIRST_COL), _
            ws.Cells(printLastRow - 1, 22)).Address
    End With
    Application.PrintCommunication = True
    Application.ScreenUpdating = True
      
    ws.Activate
    Exit Sub

errH:
    Application.ScreenUpdating = True
    Application.PrintCommunication = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation, _
        "Чертёжная рамка"
    On Error Resume Next
    KillTmpSheet
    ws.Activate
End Sub

'---------------------------------------------------------------------
Public Sub RemoveFrame()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    DeleteFrameShapes ActiveSheet
End Sub

'---------------------------------------------------------------------
' Формирование спецификации: суммарная длина кабелей по типам.
' Источник - лист SPEC_SRC_SHEET (тип в столбце E, длина в столбце Q).
' Результат - на листе SPEC_DST_SHEET, столбцы SPEC_TYPE_DST_COL /
' SPEC_LEN_DST_COL начиная со строки SPEC_START_ROW.
Public Sub BuildSpecification()
    On Error GoTo errH
    Dim src As Worksheet, dst As Worksheet
    Set src = Worksheets(SPEC_SRC_SHEET)
    Set dst = Worksheets(SPEC_DST_SHEET)

    ' --- суммирование длин по типам ---------------------------------
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare   ' типы без учёта регистра

    Dim lastRow As Long, r As Long, t As String, v As Variant
    lastRow = src.Cells(src.Rows.Count, SPEC_TYPE_SRC_COL).End(xlUp).Row
    For r = 1 To lastRow
        t = Trim$(CStr(src.Cells(r, SPEC_TYPE_SRC_COL).Value))
        v = src.Cells(r, SPEC_LEN_SRC_COL).Value
        If Len(t) > 0 And IsNumeric(v) Then
            dict(t) = dict(t) + CDbl(v)
        End If
    Next r

    If dict.Count = 0 Then
        MsgBox "На листе '" & SPEC_SRC_SHEET & "' не найдено " & _
            "строк с типом (E) и числовой длиной (Q).", vbExclamation
        Exit Sub
    End If

    ' --- сортировка типов по алфавиту -------------------------------
    Dim keyArr() As String, i As Long, j As Long, tmp As String
    ReDim keyArr(0 To dict.Count - 1)
    Dim key As Variant, idx As Long
    idx = 0
    For Each key In dict.Keys
        keyArr(idx) = key
        idx = idx + 1
    Next key
    For i = 0 To UBound(keyArr) - 1
        For j = i + 1 To UBound(keyArr)
            If keyArr(j) < keyArr(i) Then
                tmp = keyArr(i): keyArr(i) = keyArr(j): keyArr(j) = tmp
            End If
        Next j
    Next i

    Application.ScreenUpdating = False

    ' --- очистка прежнего вывода (только мои столбцы) ---------------
    Dim lastDst As Long
    lastDst = dst.Cells(dst.Rows.Count, SPEC_TYPE_DST_COL).End(xlUp).Row
    If lastDst < SPEC_START_ROW Then lastDst = SPEC_START_ROW
    dst.Range(dst.Cells(SPEC_START_ROW, SPEC_TYPE_DST_COL), _
        dst.Cells(lastDst, SPEC_LEN_DST_COL)).ClearContents

    ' --- заголовок --------------------------------------------------
    dst.Cells(SPEC_START_ROW, SPEC_TYPE_DST_COL).Value = "Тип кабеля"
    dst.Cells(SPEC_START_ROW, SPEC_LEN_DST_COL).Value = _
        "Суммарная длина, м"

    ' --- строки и итог ----------------------------------------------
    Dim total As Double, rowOut As Long
    rowOut = SPEC_START_ROW + 1
    For i = 0 To UBound(keyArr)
        dst.Cells(rowOut, SPEC_TYPE_DST_COL).Value = keyArr(i)
        dst.Cells(rowOut, SPEC_LEN_DST_COL).Value = dict(keyArr(i))
        total = total + dict(keyArr(i))
        rowOut = rowOut + 1
    Next i
    dst.Cells(rowOut, SPEC_TYPE_DST_COL).Value = "Итого"
    dst.Cells(rowOut, SPEC_LEN_DST_COL).Value = total

    Application.ScreenUpdating = True
    dst.Activate
    MsgBox "Спецификация обновлена." & vbCrLf & _
        "Типов кабеля: " & dict.Count & vbCrLf & _
        "Суммарная длина: " & Format(total, "0.###") & " м.", _
        vbInformation, "Спецификация"
    Exit Sub

errH:
    Application.ScreenUpdating = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Спецификация"
End Sub

Private Sub DeleteFrameShapes(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PFX)) = PFX Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

'---------------------------------------------------------------------
' Группировка фигур, добавленных после startIdx, в одну фигуру.
' sendBack=True отправляет группу за содержимое ячеек.
Private Sub GroupShapes(ws As Worksheet, ByVal startIdx As Long, _
    ByVal sendBack As Boolean)
    Dim cnt As Long
    cnt = ws.Shapes.Count - startIdx
    If cnt < 2 Then Exit Sub
    Dim arr() As Variant, i As Long
    ReDim arr(0 To cnt - 1)
    For i = 1 To cnt
        arr(i - 1) = ws.Shapes(startIdx + i).Name
    Next i
    With ws.Shapes.Range(arr).Group
        .Name = NextName()
        .Placement = xlFreeFloating
        If sendBack Then .ZOrder msoSendToBack
    End With
End Sub

'---------------------------------------------------------------------
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
Private Function InsertFiller(ws As Worksheet, _
    ByVal afterRow As Long, ByVal gap As Double) As Long
    Dim remain As Double, h As Double, N As Long
    ' минус 1 пт - запас от переполнения страницы
    remain = gap - 1
    Do While remain > 0.5
        h = remain
        ' предел высоты строки Excel
        If h > 400 Then h = 400
        ws.Rows(afterRow + 1).Insert Shift:=xlDown
        ws.Rows(afterRow + 1).Clear
        ws.Rows(afterRow + 1).RowHeight = h
        remain = remain - h
        N = N + 1
    Loop
    InsertFiller = N
End Function

'---------------------------------------------------------------------
' Вставка копии строк заголовка (1..nRows) перед строкой beforeRow.
Private Sub InsertHeaderCopy(ws As Worksheet, ByVal beforeRow As Long, _
    ByVal nRows As Long)
    ws.Rows("1:" & nRows).Copy
    ws.Rows(beforeRow & ":" & (beforeRow + nRows - 1)).Insert _
        Shift:=xlDown
    Application.CutCopyMode = False
End Sub

'---------------------------------------------------------------------
' Рамка одной страницы: от края рабочей области листа
Private Sub DrawFrameRect(ws As Worksheet, ByVal topPt As Double, _
    ByVal wPt As Double, ByVal hPt As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        MM(20), topPt, wPt, hPt)
    With shp
        .Name = NextName()
        .Fill.Visible = msoFalse
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With
End Sub

'---------------------------------------------------------------------
' Основная надпись: форма 6 по ГОСТ 21.101-2026 (185x15 мм).
' Левый блок - таблица изменений (6 граф), правый блок - графа "Лист",
' центр - обозначение документа (заполняется вручную).
' x0, y0 - левый верхний угол штампа, пт.
Private Sub DrawStampForm6(ws As Worksheet, _
    ByVal x0 As Double, ByVal y0 As Double, _
    ByVal CntPgs As Double)
    Dim startIdx As Long
    startIdx = ws.Shapes.Count

    ' наружный контур (белая заливка маскирует сетку под штампом)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        x0, y0, MM(STAMP_W), MM(STAMP_H))
    With shp
        .Name = NextName()
        .Fill.ForeColor.RGB = vbWhite
        .Fill.Visible = msoTrue
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With

    ' --- левый блок 65 мм: таблица изменений ------------------------
    ' колонки: Изм.(10) Кол.уч.(10) Лист(10) № док.(10) Подп.(15) Дата(10)
    StampLine ws, x0, y0, 10, 0, 10, 15, W_MAIN
    StampLine ws, x0, y0, 20, 0, 20, 15, W_MAIN
    StampLine ws, x0, y0, 30, 0, 30, 15, W_MAIN
    StampLine ws, x0, y0, 40, 0, 40, 15, W_MAIN
    StampLine ws, x0, y0, 55, 0, 55, 15, W_MAIN
    StampLine ws, x0, y0, 65, 0, 65, 15, W_MAIN
    ' горизонталь заголовок/данные внутри левого блока
    StampLine ws, x0, y0, 0, 5, 65, 5, W_THIN
    StampLine ws, x0, y0, 0, 10, 65, 10, W_THIN

    ' заголовки граф таблицы изменений (верхняя строка)
    StampText ws, x0, y0, 0, 10, 10, 5, "Изм.", 8
    StampText ws, x0, y0, 10, 10, 10, 5, "Кол.уч.", 8
    StampText ws, x0, y0, 20, 10, 10, 5, "Лист", 8
    StampText ws, x0, y0, 30, 10, 10, 5, "№ док.", 8
    StampText ws, x0, y0, 40, 10, 15, 5, "Подп.", 8
    StampText ws, x0, y0, 55, 10, 10, 5, "Дата", 8

    ' --- правый блок: графа "Лист" (10 мм) --------------------------
    StampLine ws, x0, y0, 175, 0, 175, 15, W_MAIN
    StampLine ws, x0, y0, 175, 7, 185, 7, W_THIN
    StampText ws, x0, y0, 175, 0, 10, 7, "Лист", 8
    StampText ws, x0, y0, 175, 7, 10, 8, CStr(CntPgs), 8

    ' центр (65..175) - обозначение документа, заполняется вручную
    StampText ws, x0, y0, 65, 0, 110, 15, cipher, 14
    ' --- группировка фигур штампа в одну ----------------------------
    GroupShapes ws, startIdx, True
End Sub

'---------------------------------------------------------------------
' Основная надпись: форма 5 по ГОСТ 21.101-2026 (185x15 мм).
' Левый блок - таблица изменений (6 граф), правый блок - графа "Лист",
' центр - обозначение документа (заполняется вручную).
' x0, y0 - левый верхний угол штампа, пт.
Private Sub DrawStampForm5(ws As Worksheet, _
    ByVal x0 As Double, ByVal y0 As Double, _
    ByVal CntPgsAll As Double)
    Dim startIdx As Long
    startIdx = ws.Shapes.Count

    ' наружный контур (белая заливка маскирует сетку под штампом)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        x0, y0, MM(STAMP_W), MM(40))
    With shp
        .Name = NextName()
        .Fill.ForeColor.RGB = vbWhite
        .Fill.Visible = msoTrue
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With

    ' --- левый блок 65 мм: таблица изменений ------------------------
    ' колонки: Изм.(10) Кол.уч.(10) Лист(10) № док.(10) Подп.(15) Дата(10)
    StampLine ws, x0, y0, 10, 0, 10, 15, W_MAIN
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
    StampText ws, x0, y0, 0, 10, 10, 5, "Изм.", 8
    StampText ws, x0, y0, 10, 10, 10, 5, "Кол.уч.", 8
    StampText ws, x0, y0, 20, 10, 10, 5, "Лист", 8
    StampText ws, x0, y0, 30, 10, 10, 5, "№ док.", 8
    StampText ws, x0, y0, 40, 10, 15, 5, "Подп.", 8
    StampText ws, x0, y0, 55, 10, 10, 5, "Дата", 8
    ' подписи (графы 10-13 - должности по усмотрению организации)
    StampTextLeft ws, x0, y0, 0, 15, 20, 5, "Разраб.", 10
    StampTextLeft ws, x0, y0, 0, 20, 20, 5, "Пров.", 10
    StampTextLeft ws, x0, y0, 0, 25, 20, 5, "Т.контр.", 10
    StampTextLeft ws, x0, y0, 0, 30, 20, 5, "Н.контр.", 10
    StampTextLeft ws, x0, y0, 0, 35, 20, 5, "Утв.", 10
    
    ' --- правый блок 50 мм: Стадия / Лист / Листов, организация -------
    StampLine ws, x0, y0, 135, 15, 135, 40, W_MAIN
    StampLine ws, x0, y0, 150, 15, 150, 25, W_MAIN
    StampLine ws, x0, y0, 165, 15, 165, 25, W_MAIN
    StampLine ws, x0, y0, 135, 20, 185, 20, W_MAIN
    StampLine ws, x0, y0, 135, 25, 185, 25, W_MAIN
    StampText ws, x0, y0, 135, 15, 15, 5, "Стадия", 8
    StampText ws, x0, y0, 135, 20, 15, 5, "Р", 8
    StampText ws, x0, y0, 150, 15, 15, 5, "Лист", 8
    StampText ws, x0, y0, 150, 20, 15, 5, "2", 8
    StampText ws, x0, y0, 165, 15, 20, 5, "Листов", 8
    
    ' --- пользовательские данные -------
    StampText ws, x0, y0, 65, 15, 70, 25, "Журнал кабельный", 14 ' наименование документа (журнал кабельный)
    StampText ws, x0, y0, 135, 25, 50, 15, "ООО «УралТЭП»" & vbCrLf & "ООО «УралОРГРЭС»", 12 ' наименование организации
    StampText ws, x0, y0, 165, 20, 20, 5, CntPgsAll, 8 'общее количество листов
    StampText ws, x0, y0, 65, 0, 120, 15, cipher, 14 ' шифр документа
    StampTextLeft ws, x0, y0, 20, 15, 20, 5, "Кокин", 10 ' фамилия разраб.
    StampTextLeft ws, x0, y0, 20, 20, 20, 5, "Феденев", 10 ' фамилия пров.
    StampTextLeft ws, x0, y0, 20, 25, 20, 5, "Феденев", 10 ' фамилия т.контр.
    StampTextLeft ws, x0, y0, 20, 30, 20, 5, "Сосновских", 10 ' фамилия н.контр.
    StampTextLeft ws, x0, y0, 20, 35, 20, 5, "Уляшов", 10 ' фамилия утв.
    StampTextLeft ws, x0, y0, 55, 15, 10, 5, Date0, 7 ' дата в штампе
    StampTextLeft ws, x0, y0, 55, 20, 10, 5, Date0, 7 ' дата в штампе
    StampTextLeft ws, x0, y0, 55, 25, 10, 5, Date0, 7 ' дата в штампе
    StampTextLeft ws, x0, y0, 55, 30, 10, 5, Date0, 7 ' дата в штампе
    StampTextLeft ws, x0, y0, 55, 35, 10, 5, Date0, 7 ' дата в штампе
    
    
    ' --- группировка фигур штампа в одну ----------------------------
    GroupShapes ws, startIdx, True
End Sub

'---------------------------------------------------------------------
' Дополнительные графы слева (последующие листы) по ГОСТ 21.101
' x0, y0 - левый верхний угол штампа в пунктах.
Private Sub DrawStampDopS(ws As Worksheet, ByVal x0 As Double, _
    ByVal y0 As Double)
    Dim startIdx As Long
    startIdx = ws.Shapes.Count

    ' наружный контур
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x0, y0, MM(12), MM(85))
    With shp
        .Name = NextName()
        .Fill.ForeColor.RGB = vbWhite           ' маскирует линии сетки под штампом
        .Fill.Visible = msoTrue
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With

    ' --- внутренние линии ---------------
    StampLine ws, x0, y0, 5, 0, 5, 85, W_MAIN
    StampLine ws, x0, y0, 0, 25, 12, 25, W_MAIN
    StampLine ws, x0, y0, 0, 60, 12, 60, W_MAIN
    
    ' заголовки граф таблицы
    StampTextVert ws, x0, y0, 0, 0, 5, 25, "Взам. инв. №", 8
    StampTextVert ws, x0, y0, 0, 25, 5, 35, "Подп. и дата", 8
    StampTextVert ws, x0, y0, 0, 60, 5, 25, "Инв. № подл.", 8
    
    ' --- группировка фигур штампа в одну ------------------------------
    GroupShapes ws, startIdx, False
End Sub
'---------------------------------------------------------------------
' Дополнительные графы слева (заглавный лист) по ГОСТ 21.101
' x0, y0 - левый верхний угол штампа в пунктах.
Private Sub DrawStampDopB(ws As Worksheet, ByVal x0 As Double, _
    ByVal y0 As Double)
    Dim startIdx As Long
    startIdx = ws.Shapes.Count

    ' наружный контур
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x0, y0, MM(15), MM(65))
    With shp
        .Name = NextName()
        .Fill.ForeColor.RGB = vbWhite           ' маскирует линии сетки под штампом
        .Fill.Visible = msoTrue
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With

    ' --- внутренние линии ---------------
    StampLine ws, x0, y0, 5, 0, 5, 65, W_MAIN
    StampLine ws, x0, y0, 10, 0, 10, 65, W_THIN
    StampLine ws, x0, y0, 5, 10, 15, 10, W_MAIN
    StampLine ws, x0, y0, 5, 25, 15, 25, W_MAIN
    StampLine ws, x0, y0, 5, 45, 15, 45, W_MAIN
    
    ' заголовки граф таблицы
    StampTextVertLeft ws, x0, y0, 0, 0, 5, 65, "Согласовано", 8
    
    ' --- группировка фигур штампа в одну ------------------------------
    GroupShapes ws, startIdx, False
End Sub

'---------------------------------------------------------------------
' Линия штампа: координаты в мм относительно угла штампа
Private Sub StampLine(ws As Worksheet, _
    ByVal x0 As Double, ByVal y0 As Double, _
    ByVal x1 As Double, ByVal y1 As Double, _
    ByVal x2 As Double, ByVal y2 As Double, ByVal wt As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddLine(x0 + MM(x1), y0 + MM(y1), _
        x0 + MM(x2), y0 + MM(y2))
    With shp
        .Name = NextName()
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = wt
        .Placement = xlFreeFloating
    End With
End Sub

'---------------------------------------------------------------------
' Надпись штампа: координаты и размеры в мм относительно угла штампа
Private Sub StampText(ws As Worksheet, _
    ByVal x0 As Double, ByVal y0 As Double, _
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
                .Name = "Times New Roman"
                .Size = sz
                .Color = vbBlack
                .Bold = False
            End With
        End With
    End With
End Sub
'---------------------------------------------------------------------
' Надпись штампа (вертикальная): координаты и размеры в мм относительно угла штампа
Private Sub StampTextVert(ws As Worksheet, _
    ByVal x0 As Double, ByVal y0 As Double, _
    ByVal x1 As Double, ByVal y1 As Double, _
    ByVal wd As Double, ByVal ht As Double, _
    ByVal txt As String, ByVal sz As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationVertical, _
        x0 + MM(x1), y0 + MM(y1), MM(wd), MM(ht))
    With shp
        .Name = NextName()
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        .Rotation = -180
        With .TextFrame
            .MarginLeft = 0: .MarginRight = 0
            .MarginTop = 0: .MarginBottom = 0
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .Characters.Text = txt
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
            With .Characters.Font
                .Name = "Times New Roman"
                .Size = sz
                .Color = vbBlack
                .Bold = False
            End With
        End With
    End With
End Sub
'---------------------------------------------------------------------
' Надпись штампа (Форматирование по левому краю):
' координаты и размеры в мм относительно угла штампа
Private Sub StampTextLeft(ws As Worksheet, _
    ByVal x0 As Double, ByVal y0 As Double, _
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
            .MarginLeft = 3: .MarginRight = 0
            .MarginTop = 0: .MarginBottom = 0
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
            .Characters.Text = txt
            With .Characters.Font
                .Name = "Times New Roman"
                .Size = sz
                .Color = vbBlack
                .Bold = False
            End With
        End With
    End With
End Sub
'---------------------------------------------------------------------
' Надпись штампа (вертикальная): координаты и размеры в мм относительно угла штампа
Private Sub StampTextVertLeft(ws As Worksheet, _
    ByVal x0 As Double, ByVal y0 As Double, _
    ByVal x1 As Double, ByVal y1 As Double, _
    ByVal wd As Double, ByVal ht As Double, _
    ByVal txt As String, ByVal sz As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationVertical, _
        x0 + MM(x1), y0 + MM(y1), MM(wd), MM(ht))
    With shp
        .Name = NextName()
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        .Rotation = -180
        With .TextFrame
            .MarginLeft = 0: .MarginRight = 0
            .MarginTop = 3: .MarginBottom = 0
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
            .Characters.Text = txt
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
            With .Characters.Font
                .Name = "Times New Roman"
                .Size = sz
                .Color = vbBlack
                .Bold = False
            End With
        End With
    End With
End Sub


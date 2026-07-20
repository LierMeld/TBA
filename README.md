# Чертёжная рамка для печати из Excel (VBA, Excel 2016)

Макрос VBA для Microsoft Excel 2016: выводит таблицу на печать с
чертёжной рамкой на листах формата **A3 (альбомная)** с основной
надписью по **форме 6 ГОСТ 21.101-2026** на каждом листе.

## Возможности

- Формат печати всегда **A3, альбомная ориентация** (420×297 мм).
- Поля листа по ГОСТ: слева 20 мм, сверху/справа/снизу 5 мм. Рамка
  чертится от края рабочей области листа, поэтому таблица,
  начинающаяся с A1, сразу попадает внутрь рамки.
- Основная надпись по **форме 6 ГОСТ 21.101-2026** (185×15 мм) на
  **каждом** листе, вплотную к нижней линии рамки: таблица изменений
  (Изм. / Кол.уч. / Лист / № док. / Подп. / Дата), графа обозначения
  документа и графа «Лист».
- **Область печати — столбцы A…Q**; число печатаемых строк
  определяется по содержимому листа.
- **Многостраничная печать**: длинная таблица разбивается на
  страницы, каждая получает свою рамку и свой штамп. Строки, которые
  попали бы под штамп, переносятся на следующий лист.
- Исходный лист не изменяется — печать идёт через временную копию
  листа, которая удаляется после печати.
- Настройка рамки по ширине одним числом (см. ниже про поправку).

## Установка

**Вариант 1 — импорт готового модуля:**

1. Откройте книгу Excel и нажмите `Alt+F11` (редактор VBA).
2. Меню **File → Import File…** и выберите файл `GostFrame.bas`
   из этого репозитория (кодировка Windows-1251 — на русской Windows
   импортируется корректно).

**Вариант 2 — копирование кода:**

1. Нажмите `Alt+F11`, затем **Insert → Module**.
2. Скопируйте в модуль код из раздела «Полный код» ниже.

Сохраните книгу в формате **.xlsm** (книга с поддержкой макросов).

## Использование

1. Откройте лист с таблицей (может начинаться с A1).
2. Нажмите `Alt+F8`, выберите **PrintWithFrame** → «Выполнить».
3. Макрос сообщит число листов и предложит печать или предпросмотр.

Формат, ориентация и штамп настраивать не нужно — они фиксированы
(A3 альбомная, форма 6 на каждом листе). Макрос строит временную
копию листа `GOST_PRINT_TMP`, печатает её и удаляет.

## Настройка ширины рамки (важно)

Если рамка вылезает за правый край листа или, наоборот, не достаёт до
него, отредактируйте **одну** константу в начале модуля:

```vba
Private Const FRAME_ADJUST_MM As Double = 28
```

Это на сколько миллиметров рамка сужается справа; штамп привязан к
правому краю рамки и смещается на столько же. Увеличьте число, если
рамка всё ещё широка; уменьшите (можно до 0), если стала узкой.
Похожим образом рядом лежат `FRAME_BOTTOM_MM` (нижний отступ рамки)
и размеры полей `FLD_LEFT_MM` / `FLD_OTHER_MM`.

## Про кодировку файла GostFrame.bas

Редактор VBA (VBE) **не поддерживает UTF-8**: `.bas`-файлы он читает
только в ANSI-кодировке системной локали Windows (для русской Windows —
Windows-1251). Поэтому файл `GostFrame.bas` сохранён в Windows-1251 —
при просмотре на GitHub комментарии выглядят «кракозябрами», но при
импорте в VBA на русской Windows текст отображается корректно.

- В VS Code: «Reopen with Encoding → Cyrillic (Windows 1251)».
- Если кириллица ломается **после импорта в Excel**, значит в Windows
  язык для программ без поддержки Юникода не русский (Панель
  управления → Региональные стандарты → Дополнительно). Тогда
  скопируйте код из раздела «Полный код» — вставка через буфер обмена
  работает в любой локали.

## Примечания

- Штамп — форма 6 ГОСТ 21.101-2026 (основная надпись для последующих
  листов). Размеры и графы заданы константами в начале модуля.
- Пустые графы штампа (обозначение, номер листа, подписи) заполняются
  вручную поверх штампа или на распечатке.
- Печать в масштабе 100 %. У большинства принтеров есть непечатаемая
  зона ~3–5 мм по краям — при необходимости подстройте `FRAME_ADJUST_MM`.
- Все строки кода не длиннее 72 символов, чтобы при копировании из
  «обёрнутого» вида ничего не разрывалось.
- Макрос `RemoveFrame` удаляет фигуры рамки/штампа с активного листа.

## Полный код

> **Важно:** ниже приведён код без служебной строки
> `Attribute VB_Name = "GostFrame"` — она есть только в файле
> `GostFrame.bas` и нужна лишь для импорта файла целиком.
> Вставлять её в модуль вручную нельзя: VBA выдаст
> «Compile error: Syntax error».

```vba
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
Private Const TMP_SHEET As String = "GOST_PRINT_TMP"

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
Private Const STAMP_GAP As Double = 0           ' зазор таблица-штамп

' столбцы области печати: A (1) .. Q (17)
Private Const FIRST_COL As Long = 1
Private Const LAST_COL As Long = 17

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
        MsgBox "Активный лист не является рабочим листом.", _
            vbExclamation
        Exit Sub
    End If
    Dim src As Worksheet
    Set src = ActiveSheet
    If src.Name = TMP_SHEET Then
        MsgBox "Активен временный лист печати." & vbCrLf & _
            "Откройте лист с таблицей.", vbExclamation
        Exit Sub
    End If

    ' рабочая область страницы (внутри полей), пт
    Dim cW As Double, cH As Double
    cW = MM(SHEET_W_MM - FLD_LEFT_MM - FLD_OTHER_MM)
    cH = MM(SHEET_H_MM - 2 * FLD_OTHER_MM)

    ' ширина рамки с поправкой; штамп привязан к её правому краю
    Dim frameW As Double
    frameW = cW - MM(FRAME_ADJUST_MM)
    Dim frameH As Double
    frameH = cH - MM(FRAME_BOTTOM_MM)

    On Error GoTo errH
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

    ' --- последняя строка содержимого -------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

    ' --- разбивка на страницы (проход 1: строки-заполнители) --------
    ' На каждой странице резервируется нижняя зона под штамп, а хвост
    ' страницы добивается пустой строкой до полной высоты, чтобы рамки
    ' совпадали с границами бумаги.
    Dim stampReserve As Double
    stampReserve = MM(STAMP_H + STAMP_GAP)

    Dim fills As Collection
    Set fills = New Collection      ' Array(строка, высота вставки)
    Dim r As Long, acc As Double, limit As Double, rowH As Double
    limit = cH - stampReserve
    acc = 0: r = 1
    Do While r <= lastRow
        rowH = ws.Rows(r).Height
        If acc + rowH > limit + 0.3 And acc > 0 Then
            fills.Add Array(r - 1, cH - acc)
            acc = 0
        Else
            acc = acc + rowH
            r = r + 1
        End If
    Loop

    ' вставка заполнителей снизу вверх
    Dim i As Long, itm As Variant, inserted As Long
    For i = fills.Count To 1 Step -1
        itm = fills(i)
        inserted = inserted + _
            InsertFiller(ws, CLng(itm(0)), CDbl(itm(1)))
    Next i
    lastRow = lastRow + inserted

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

    Dim k As Long, thisH As Double, sTop As Double
    For k = 1 To pages.Count
        pTop = pages(k)
        ' нижняя линия чуть выше края, чтобы не ушла на след. лист
        If k = pages.Count Then
            thisH = frameH
        Else
            thisH = frameH - 0.75
        End If
        DrawFrameRect ws, pTop, frameW, thisH
        ' штамп формы 6 у нижней линии рамки
        sTop = pTop + thisH - MM(STAMP_H)
        DrawStampForm6 ws, frameW - MM(STAMP_W), sTop
    Next k

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

    Dim printLastCol As Long, accCol As Double
    accCol = 0
    For i = 1 To LAST_COL
        accCol = accCol + ws.Columns(i).Width
    Next i
    printLastCol = LAST_COL
    Do While accCol < frameW + MM(3) And printLastCol < 16384
        printLastCol = printLastCol + 1
        accCol = accCol + ws.Columns(printLastCol).Width
    Loop

    ' --- параметры страницы -----------------------------------------
    With ws.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .LeftMargin = MM(FLD_LEFT_MM)
        .RightMargin = MM(FLD_OTHER_MM)
        .TopMargin = MM(FLD_OTHER_MM)
        .BottomMargin = MM(FLD_OTHER_MM)
        .HeaderMargin = 0: .FooterMargin = 0
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "": .CenterFooter = "": .RightFooter = ""
        .CenterHorizontally = False
        .CenterVertically = False
        .Zoom = 100
        .PrintArea = ws.Range( _
            ws.Cells(1, FIRST_COL), _
            ws.Cells(printLastRow, printLastCol)).Address
    End With

    Application.ScreenUpdating = True

    ' --- печать / предварительный просмотр --------------------------
    Dim msg As String
    msg = "Формат: A3 альбомная. Листов к печати: " & _
        pages.Count & vbCrLf & vbCrLf & _
        "Отправить на печать сразу?" & vbCrLf & _
        "Нет - открыть предварительный просмотр."
    If MsgBox(msg, vbYesNo + vbQuestion, "Печать") = vbYes Then
        ws.PrintOut
    Else
        ws.PrintPreview
    End If

    KillTmpSheet
    src.Activate
    Exit Sub

errH:
    Application.ScreenUpdating = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation, _
        "Чертёжная рамка"
    KillTmpSheet
    src.Activate
End Sub

'---------------------------------------------------------------------
Public Sub RemoveFrame()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    DeleteFrameShapes ActiveSheet
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
    Dim remain As Double, h As Double, n As Long
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
        n = n + 1
    Loop
    InsertFiller = n
End Function

'---------------------------------------------------------------------
' Рамка одной страницы: от края рабочей области листа
Private Sub DrawFrameRect(ws As Worksheet, ByVal topPt As Double, _
    ByVal wPt As Double, ByVal hPt As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        0, topPt, wPt, hPt)
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
    ByVal x0 As Double, ByVal y0 As Double)
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

    ' --- левый блок 62 мм: таблица изменений ------------------------
    ' колонки: Изм.(7) Кол.уч.(10) Лист(10) № док.(10) Подп.(15) Дата(10)
    StampLine ws, x0, y0, 7, 0, 7, 15, W_MAIN
    StampLine ws, x0, y0, 17, 0, 17, 15, W_MAIN
    StampLine ws, x0, y0, 27, 0, 27, 15, W_MAIN
    StampLine ws, x0, y0, 37, 0, 37, 15, W_MAIN
    StampLine ws, x0, y0, 52, 0, 52, 15, W_MAIN
    StampLine ws, x0, y0, 62, 0, 62, 15, W_MAIN
    ' горизонталь заголовок/данные внутри левого блока
    StampLine ws, x0, y0, 0, 8, 62, 8, W_THIN

    ' заголовки граф таблицы изменений (верхняя строка)
    StampText ws, x0, y0, 0, 0, 7, 8, "Изм.", 5
    StampText ws, x0, y0, 7, 0, 10, 8, "Кол.уч.", 5
    StampText ws, x0, y0, 17, 0, 10, 8, "Лист", 5
    StampText ws, x0, y0, 27, 0, 10, 8, "№ док.", 5
    StampText ws, x0, y0, 37, 0, 15, 8, "Подп.", 5
    StampText ws, x0, y0, 52, 0, 10, 8, "Дата", 5

    ' --- правый блок: графа "Лист" (10 мм) --------------------------
    StampLine ws, x0, y0, 175, 0, 175, 15, W_MAIN
    StampText ws, x0, y0, 175, 0, 10, 8, "Лист", 5

    ' центр (62..175) - обозначение документа, заполняется вручную

    ' --- группировка фигур штампа в одну ----------------------------
    Dim cnt As Long
    cnt = ws.Shapes.Count - startIdx
    If cnt > 1 Then
        Dim arr() As Variant, i As Long
        ReDim arr(0 To cnt - 1)
        For i = 1 To cnt
            arr(i - 1) = ws.Shapes(startIdx + i).Name
        Next i
        With ws.Shapes.Range(arr).Group
            .Name = NextName()
            .Placement = xlFreeFloating
        End With
    End If
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
                .Name = "Arial"
                .Size = sz
                .Color = vbBlack
                .Bold = False
            End With
        End With
    End With
End Sub
```

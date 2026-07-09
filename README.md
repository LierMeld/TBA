# Чертёжная рамка для печати из Excel (VBA, Excel 2016)

Макрос VBA для Microsoft Excel 2016, который выводит таблицу на печать
с чертёжной рамкой по ГОСТ 2.104 для формата на выбор: **A4** или **A3**.

## Возможности

- Выбор формата листа: **A4** (210×297 мм) или **A3** (297×420 мм).
- Выбор ориентации: книжная или альбомная.
- Рамка с полями по ГОСТ 2.104: слева 20 мм, сверху/справа/снизу 5 мм.
- По желанию — основная надпись (штамп, упрощённая форма 1, 185×55 мм)
  в правом нижнем углу рамки.
- Автоматическая настройка страницы: формат бумаги, ориентация, нулевые
  поля, масштаб строго 100 % (размеры рамки на бумаге соответствуют мм).
- Вывод сразу на печать или в предварительный просмотр — по выбору.
- Макрос `RemoveFrame` удаляет рамку и штамп с листа.

## Установка

**Вариант 1 — импорт готового модуля:**

1. Откройте книгу Excel и нажмите `Alt+F11` (редактор VBA).
2. Меню **File → Import File…** и выберите файл `GostFrame.bas`
   из этого репозитория (файл в кодировке Windows-1251 — на русской
   Windows импортируется корректно).

**Вариант 2 — копирование кода:**

1. Нажмите `Alt+F11`, затем **Insert → Module**.
2. Скопируйте в модуль код из раздела «Полный код» ниже.

Сохраните книгу в формате **.xlsm** (книга с поддержкой макросов).

## Использование

1. Разместите таблицу на листе так, чтобы она находилась внутри будущей
   рамки: отступ слева ~20 мм, сверху ~5 мм от края листа
   (после первого запуска рамка видна на листе — таблицу удобно
   подогнать под неё и запустить макрос повторно).
2. Нажмите `Alt+F8`, выберите **PrintWithFrame** и нажмите «Выполнить».
3. Ответьте на вопросы макроса: формат (A4/A3), ориентация, нужен ли
   штамп, печатать сразу или открыть предварительный просмотр.

Чтобы убрать рамку с листа — запустите макрос **RemoveFrame**.

## Про кодировку файла GostFrame.bas

Редактор VBA (VBE) **не поддерживает UTF-8**: `.bas`-файлы он читает
только в ANSI-кодировке системной локали Windows (для русской Windows —
Windows-1251). Поэтому файл `GostFrame.bas` сохранён именно в
Windows-1251 — при просмотре на GitHub или в редакторах, ожидающих
UTF-8, комментарии выглядят «кракозябрами», но это нормально: при
импорте в VBA на русской Windows текст отображается корректно.

- Чтобы посмотреть файл в VS Code: «Reopen with Encoding →
  Cyrillic (Windows 1251)».
- Если кириллица ломается **после импорта в Excel**, значит в Windows
  язык для программ без поддержки Юникода не русский (Панель
  управления → Региональные стандарты → Дополнительно). В этом случае
  скопируйте код из раздела «Полный код» ниже — вставка через буфер
  обмена работает в любой локали.

## Примечания

- Печать идёт в масштабе 100 % с нулевыми полями. У большинства
  принтеров есть аппаратная непечатаемая зона ~3–5 мм по краям, поэтому
  линии рамки (отступ 5 мм от края) обычно печатаются полностью, но на
  некоторых принтерах края могут слегка обрезаться.
- Рамка и штамп — это фигуры (Shapes) с именами, начинающимися на
  `GOST_`; они не привязаны к ячейкам и не двигаются при изменении
  ширины столбцов.
- Пустые графы штампа (обозначение документа, наименование, фамилии,
  организация) заполняются вручную — например, надписями
  (Вставка → Надпись) поверх штампа.

## Полный код

```vba
Attribute VB_Name = "GostFrame"
Option Explicit

'=====================================================================
'  Печать таблицы с чертёжной рамкой (по ГОСТ 2.104) — Excel 2016
'
'  Макросы для запуска (Alt+F8):
'    PrintWithFrame - выбор формата (A4/A3) и ориентации, построение
'                     рамки (и, по желанию, основной надписи - штампа),
'                     настройка страницы и вывод на печать / просмотр
'    RemoveFrame    - удаление рамки и штампа с активного листа
'
'  Поля рамки по ГОСТ 2.104: слева 20 мм, сверху/справа/снизу 5 мм.
'  Печать выполняется строго в масштабе 100%, поэтому таблица должна
'  располагаться внутри рамки на листе Excel.
'=====================================================================

Private Const PFX As String = "GOST_"       ' префикс имён фигур рамки

' поля рамки, мм
Private Const FLD_LEFT_MM As Double = 20
Private Const FLD_OTHER_MM As Double = 5

' основная надпись (форма 1), мм
Private Const STAMP_W As Double = 185
Private Const STAMP_H As Double = 55

' толщина линий, пт (основная / тонкая)
Private Const W_MAIN As Double = 2#
Private Const W_THIN As Double = 0.5

Private mSeq As Long                        ' счётчик имён фигур

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
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- выбор формата ---------------------------------------------
    Dim ans As String
    ans = UCase$(Trim$(InputBox( _
        "Укажите формат листа для печати: A4 или A3", _
        "Чертёжная рамка", "A4")))
    If Len(ans) = 0 Then Exit Sub           ' нажата "Отмена"

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
    withStamp = (MsgBox("Добавить основную надпись (штамп, форма 1, 185х55 мм)?", _
        vbYesNo + vbQuestion, "Основная надпись") = vbYes)

    Application.ScreenUpdating = False
    mSeq = 0

    RemoveFrame                              ' убрать старую рамку
    DrawFrame ws, wMM, hMM
    If withStamp Then DrawStamp ws, wMM, hMM
    SetupPage ws, paper, landscape, wMM, hMM

    Application.ScreenUpdating = True

    ' --- печать / предварительный просмотр --------------------------
    If MsgBox("Отправить на печать сразу?" & vbCrLf & _
        "Нет - открыть предварительный просмотр.", _
        vbYesNo + vbQuestion, "Печать") = vbYes Then
        ws.PrintOut
    Else
        ws.PrintPreview
    End If
End Sub

'---------------------------------------------------------------------
' Удаление всех фигур рамки/штампа с активного листа
Public Sub RemoveFrame()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Dim ws As Worksheet, i As Long
    Set ws = ActiveSheet
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PFX)) = PFX Then ws.Shapes(i).Delete
    Next i
End Sub

'---------------------------------------------------------------------
' Внутренняя рамка листа
Private Sub DrawFrame(ws As Worksheet, ByVal wMM As Double, ByVal hMM As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        MM(FLD_LEFT_MM), MM(FLD_OTHER_MM), _
        MM(wMM - FLD_LEFT_MM - FLD_OTHER_MM), MM(hMM - 2 * FLD_OTHER_MM))
    With shp
        .Name = PFX & "Frame"
        .Fill.Visible = msoFalse
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With
End Sub

'---------------------------------------------------------------------
' Основная надпись (упрощённая форма 1 по ГОСТ 2.104, 185х55 мм)
' в правом нижнем углу рамки
Private Sub DrawStamp(ws As Worksheet, ByVal wMM As Double, ByVal hMM As Double)
    Dim x0 As Double, y0 As Double
    x0 = MM(wMM - FLD_OTHER_MM - STAMP_W)
    y0 = MM(hMM - FLD_OTHER_MM - STAMP_H)

    Dim startIdx As Long
    startIdx = ws.Shapes.Count

    ' --- наружный контур штампа -------------------------------------
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x0, y0, MM(STAMP_W), MM(STAMP_H))
    With shp
        .Name = NextName()
        .Fill.Visible = msoFalse
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = W_MAIN
        .Placement = xlFreeFloating
    End With

    ' --- левый блок (таблица изменений и подписи), x = 0..65 --------
    ' вертикали (колонки 7 / 10 / 23 / 15 / 10 мм)
    StampLine ws, x0, y0, 7, 0, 7, 25, W_MAIN         ' только зона изменений
    StampLine ws, x0, y0, 17, 0, 17, 55, W_MAIN
    StampLine ws, x0, y0, 40, 0, 40, 55, W_MAIN
    StampLine ws, x0, y0, 55, 0, 55, 55, W_MAIN
    StampLine ws, x0, y0, 65, 0, 65, 55, W_MAIN
    ' горизонтали зоны изменений
    StampLine ws, x0, y0, 0, 5, 65, 5, W_THIN
    StampLine ws, x0, y0, 0, 10, 65, 10, W_THIN
    StampLine ws, x0, y0, 0, 15, 65, 15, W_THIN
    StampLine ws, x0, y0, 0, 20, 65, 20, W_THIN
    StampLine ws, x0, y0, 0, 25, 65, 25, W_MAIN
    ' горизонтали строк подписей
    StampLine ws, x0, y0, 0, 30, 65, 30, W_THIN
    StampLine ws, x0, y0, 0, 35, 65, 35, W_THIN
    StampLine ws, x0, y0, 0, 40, 65, 40, W_THIN
    StampLine ws, x0, y0, 0, 45, 65, 45, W_THIN
    StampLine ws, x0, y0, 0, 50, 65, 50, W_THIN
    ' заголовки граф таблицы изменений
    StampText ws, x0, y0, 0, 20, 7, 5, "Изм.", 6
    StampText ws, x0, y0, 7, 20, 10, 5, "Лист", 6
    StampText ws, x0, y0, 17, 20, 23, 5, "№ докум.", 6
    StampText ws, x0, y0, 40, 20, 15, 5, "Подп.", 6
    StampText ws, x0, y0, 55, 20, 10, 5, "Дата", 6
    ' подписи
    StampText ws, x0, y0, 0, 25, 17, 5, "Разраб.", 7
    StampText ws, x0, y0, 0, 30, 17, 5, "Пров.", 7
    StampText ws, x0, y0, 0, 35, 17, 5, "Т.контр.", 7
    StampText ws, x0, y0, 0, 45, 17, 5, "Н.контр.", 7
    StampText ws, x0, y0, 0, 50, 17, 5, "Утв.", 7

    ' --- средняя и правая части -------------------------------------
    ' низ графы "Обозначение документа" (65..185, высота 15)
    StampLine ws, x0, y0, 65, 15, 185, 15, W_MAIN
    ' правый блок 50 мм (Лит. / Масса / Масштаб, Лист / Листов)
    StampLine ws, x0, y0, 135, 15, 135, 55, W_MAIN
    StampLine ws, x0, y0, 135, 20, 185, 20, W_THIN
    StampLine ws, x0, y0, 150, 15, 150, 30, W_MAIN
    StampLine ws, x0, y0, 165, 15, 165, 30, W_MAIN
    StampLine ws, x0, y0, 135, 30, 185, 30, W_MAIN
    StampLine ws, x0, y0, 160, 30, 160, 35, W_MAIN
    StampLine ws, x0, y0, 135, 35, 185, 35, W_MAIN
    ' подписи правого блока
    StampText ws, x0, y0, 135, 15, 15, 5, "Лит.", 6
    StampText ws, x0, y0, 150, 15, 15, 5, "Масса", 6
    StampText ws, x0, y0, 165, 15, 20, 5, "Масштаб", 6
    StampText ws, x0, y0, 135, 30, 25, 5, "Лист", 6
    StampText ws, x0, y0, 160, 30, 25, 5, "Листов", 6

    ' --- группировка всех фигур штампа в одну -----------------------
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

'---------------------------------------------------------------------
' Параметры страницы: формат, ориентация, нулевые поля, масштаб 100%,
' область печати по размеру листа
Private Sub SetupPage(ws As Worksheet, ByVal paper As XlPaperSize, _
    ByVal landscape As Boolean, ByVal wMM As Double, ByVal hMM As Double)

    ' последний столбец/строка, целиком помещающиеся на странице
    Dim lastCol As Long, lastRow As Long, acc As Double
    acc = 0: lastCol = 0
    Do While lastCol < 500
        If acc + ws.Columns(lastCol + 1).Width > MM(wMM) + 0.5 Then Exit Do
        lastCol = lastCol + 1
        acc = acc + ws.Columns(lastCol).Width
    Loop
    If lastCol = 0 Then lastCol = 1

    acc = 0: lastRow = 0
    Do While lastRow < 2000
        If acc + ws.Rows(lastRow + 1).Height > MM(hMM) + 0.5 Then Exit Do
        lastRow = lastRow + 1
        acc = acc + ws.Rows(lastRow).Height
    Loop
    If lastRow = 0 Then lastRow = 1

    With ws.PageSetup
        .PaperSize = paper
        .Orientation = IIf(landscape, xlLandscape, xlPortrait)
        .LeftMargin = 0: .RightMargin = 0
        .TopMargin = 0: .BottomMargin = 0
        .HeaderMargin = 0: .FooterMargin = 0
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "": .CenterFooter = "": .RightFooter = ""
        .CenterHorizontally = False
        .CenterVertically = False
        .Zoom = 100                      ' печать строго в масштабе 1:1
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
    End With
End Sub
```

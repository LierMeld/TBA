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
- **Повтор заголовка**: строки 1–4 листа «Таблица кабелей»
  (константа `HEADER_ROWS`) дублируются в начале каждого печатного
  листа, чтобы шапка таблицы была на всех страницах.
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

## Формирование спецификации (BuildSpecification)

Макрос **BuildSpecification** считает суммарную длину кабелей по типам
и выводит таблицу на лист «Спецификация».

- Источник данных — лист **«Таблица кабелей»**: тип кабеля в столбце
  **E**, длина в столбце **Q**.
- Строки без типа (пустой E) и без числовой длины (Q не число)
  пропускаются — так автоматически отсекаются заголовки и итоги.
- Одинаковые типы объединяются (без учёта регистра), длины
  суммируются, список сортируется по алфавиту.
- В таблице кабелей столбцы: **Тип кабеля | Кол-во позиций | Экран |
  Длина, м**. «Кол-во позиций» — сколько строк данного типа; «Экран»
  показывает наличие экранирования по букве **Э** в названии типа.
- В таблице труб и металлорукавов только **наименование и суммарная
  длина** (без количества и экрана). Столбец «Длина, м» — тот же, что
  и в таблице кабелей, поэтому колонки совпадают на всех листах.
- Лист «Спецификация» формируется заново при каждом запуске (прежнее
  содержимое очищается) и оформляется рамками формата **A4 книжная**
  с основной надписью формы 6 и левым штампом на каждом листе.
  Таблица кабелей и таблица труб — на разных листах.
- Листы спецификации добавляются в общий счётчик листов: запускайте в
  порядке **PrintWithFrame → BuildSpecification → PrintWithFrameAnn**,
  тогда номера листов и общее число в штампах будут сквозными.

Запуск: `Alt+F8` → **BuildSpecification**.

**Трубы и металлорукава.** Тот же макрос дополнительно разбирает
текстовую строку в столбце **P** листа «Таблица кабелей»: обозначение
`Тр` (труба) или `МР` (металлорукав), затем диаметр, тире и длина в
метрах (например `Тр25-3м`; «Dn» перед диаметром необязательно; в
одной ячейке может быть несколько записей через запятую — прочие
обозначения, например лоток `Лк`, игнорируются). Длины суммируются по типу и диаметру, и вторая
таблица «Труба / металлорукав, Dn | Суммарная длина, м» выводится на
листе «Спецификация» **с разрывом страницы после таблицы кабелей**.
Столбец P задаётся константой `SPEC_PIPE_SRC_COL`.

Расположение и столбцы задаются константами в начале модуля
(`SPEC_SRC_SHEET`, `SPEC_DST_SHEET`, `SPEC_TYPE_SRC_COL`,
`SPEC_LEN_SRC_COL`, `SPEC_START_ROW`, `SPEC_TYPE_DST_COL`,
`SPEC_LEN_DST_COL`) — если таблицу нужно выводить в другое место или
брать из других столбцов, поменяйте их.

> Модуль в этом репозитории адаптирован под конкретный проект (листы
> «Таблица кабелей», «Спецификация», «Аннотация»). Имена листов,
> данные штампа (шифр, организация, фамилии) и т.п. заданы
> константами и жёсткими значениями в начале модуля — при переносе на
> другой проект поправьте их.

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
Private Const TMP_SHEET As String = "КЖ"

' Данные для основной надписи
Private Const cipher As String = "1.2.1.NT307R.00.001.AK04.GK01"
Private Const Date0 As String = "30.07.26"
Private CountPageAll As Double ' листов до спецификации
Private PagesSpec As Double    ' листов спецификации
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
Private Const SPEC_PIPE_SRC_COL As Long = 16  ' P - строка труб/МР
Private Const SPEC_START_ROW As Long = 2      ' первая строка вывода
' A - левое поле (гутер), таблица с B
Private Const SPEC_TYPE_DST_COL As Long = 2   ' B - наименование
Private Const SPEC_CNT_DST_COL As Long = 3    ' C - кол-во позиций
Private Const SPEC_SCR_DST_COL As Long = 4    ' D - экран (Э)
Private Const SPEC_LEN_DST_COL As Long = 5    ' E - суммарная длина, м

' листы спецификации: A4 книжная, мм
Private Const SPEC_SHEET_W_MM As Double = 210
Private Const SPEC_SHEET_H_MM As Double = 297
' поправка ширины рамки спецификации, мм (0 = по геометрии листа)
Private Const SPEC_FRAME_ADJUST_MM As Double = 0

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
    DrawStampForm5 ws, MM(20), sTopa, CountPageAll + PagesSpec
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

    ' ===== 1. Кабели: длина и число позиций по типам ================
    Dim cabLen As Object, cabCnt As Object
    Set cabLen = CreateObject("Scripting.Dictionary")
    Set cabCnt = CreateObject("Scripting.Dictionary")
    cabLen.CompareMode = vbTextCompare
    cabCnt.CompareMode = vbTextCompare
    Dim lastRow As Long, r As Long, t As String, v As Variant
    lastRow = src.Cells(src.Rows.Count, SPEC_TYPE_SRC_COL).End(xlUp).Row
    For r = 1 To lastRow
        t = Trim$(CStr(src.Cells(r, SPEC_TYPE_SRC_COL).Value))
        v = src.Cells(r, SPEC_LEN_SRC_COL).Value
        If Len(t) > 0 And IsNumeric(v) Then
            cabLen(t) = cabLen(t) + CDbl(v)
            cabCnt(t) = cabCnt(t) + 1
        End If
    Next r

    ' ===== 2. Трубы и металлорукава: разбор строки в столбце P ======
    ' Формат: Тр|МР + диаметр + тире + длина(+"м"), напр. "Тр25-3м".
    ' "Dn" необязательно; несколько записей через запятую; прочие
    ' обозначения (лоток Лк и т.п.) игнорируются.
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "(Тр|МР)\s*\.?\s*(?:[Dd][Nn])?\s*" & _
        "(\d+(?:[.,]\d+)?)\s*[-–—]\s*(\d+(?:[.,]\d+)?)"

    Dim pLen As Object
    Set pLen = CreateObject("Scripting.Dictionary")
    pLen.CompareMode = vbTextCompare
    Dim lastP As Long, cellTxt As String, mt As Object
    Dim kind As String, dia As String, pkey As String, lng As Double
    lastP = src.Cells(src.Rows.Count, SPEC_PIPE_SRC_COL).End(xlUp).Row
    For r = 1 To lastP
        cellTxt = CStr(src.Cells(r, SPEC_PIPE_SRC_COL).Value)
        If Len(cellTxt) > 0 Then
            For Each mt In rx.Execute(cellTxt)
                kind = mt.SubMatches(0)
                dia = mt.SubMatches(1)
                lng = Val(Replace(mt.SubMatches(2), ",", "."))
                pkey = kind & "|" & dia
                pLen(pkey) = pLen(pkey) + lng
            Next mt
        End If
    Next r

    If cabLen.Count = 0 And pLen.Count = 0 Then
        MsgBox "На листе '" & SPEC_SRC_SHEET & "' не найдено " & _
            "данных для спецификации.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' --- полная очистка листа спецификации (формируется заново) -----
    Dim lu As Long
    lu = dst.Cells.SpecialCells(xlCellTypeLastCell).Row
    DeleteFrameShapes dst
    dst.ResetAllPageBreaks
    If lu >= 1 Then dst.Rows("1:" & (lu + 5)).Delete

    ' ширины: A - гутер (левое поле), B..E - колонки таблицы
    SetColMM dst, 1, 20
    SetColMM dst, SPEC_TYPE_DST_COL, 95
    SetColMM dst, SPEC_CNT_DST_COL, 22
    SetColMM dst, SPEC_SCR_DST_COL, 22
    SetColMM dst, SPEC_LEN_DST_COL, 30

    ' --- таблица кабелей (B: тип, C: кол-во, D: экран, E: длина) -----
    Dim i As Long, cabLast As Long
    cabLast = SPEC_START_ROW - 1
    If cabLen.Count > 0 Then
        Dim ck() As String
        SortDictKeysText cabLen, ck
        Dim nm() As String, cn() As Long, sc() As String, ln() As Double
        ReDim nm(0 To UBound(ck)): ReDim cn(0 To UBound(ck))
        ReDim sc(0 To UBound(ck)): ReDim ln(0 To UBound(ck))
        For i = 0 To UBound(ck)
            nm(i) = ck(i)
            cn(i) = cabCnt(ck(i))
            sc(i) = IIf(InStr(1, ck(i), "Э", vbTextCompare) > 0, "Э", "—")
            ln(i) = cabLen(ck(i))
        Next i
        cabLast = WriteSpecTable(dst, SPEC_START_ROW, "Тип кабеля", _
            nm, cn, sc, ln, True, True)
    End If

    ' --- таблица труб и металлорукавов ------------------------------
    Dim pipeLast As Long, pipeStart As Long
    pipeLast = cabLast
    pipeStart = 0
    If pLen.Count > 0 Then
        pipeStart = cabLast + 2
        Dim pk() As String
        SortPipeKeys pLen, pk
        Dim pnm() As String, pcn() As Long
        Dim psc() As String, pln() As Double
        ReDim pnm(0 To UBound(pk)): ReDim pcn(0 To UBound(pk))
        ReDim psc(0 To UBound(pk)): ReDim pln(0 To UBound(pk))
        For i = 0 To UBound(pk)
            pnm(i) = PipeLabel(pk(i))
            psc(i) = ""
            pln(i) = pLen(pk(i))
        Next i
        pipeLast = WriteSpecTable(dst, pipeStart, _
            "Труба / металлорукав", pnm, pcn, psc, pln, _
            False, False)
    End If

    ' --- рамки, штампы, разбивка на листы A4 книжной ----------------
    Dim specBase As Long
    If CountPageAll > 0 Then
        specBase = CountPageAll
    Else
        specBase = PAGE_OFFSET
    End If
    PagesSpec = FrameSpecSheet(dst, specBase, pipeStart)

    Application.ScreenUpdating = True
    dst.Activate
    MsgBox "Спецификация обновлена." & vbCrLf & _
        "Типов кабеля: " & cabLen.Count & vbCrLf & _
        "Позиций труб/металлорукавов: " & pLen.Count & vbCrLf & _
        "Листов спецификации: " & PagesSpec, _
        vbInformation, "Спецификация"
    Exit Sub

errH:
    Application.ScreenUpdating = True
    Application.PrintCommunication = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Спецификация"
End Sub

'---------------------------------------------------------------------
' Ключи словаря -> отсортированный по алфавиту строковый массив
Private Sub SortDictKeysText(dict As Object, ByRef arr() As String)
    ReDim arr(0 To dict.Count - 1)
    Dim key As Variant, idx As Long
    idx = 0
    For Each key In dict.Keys
        arr(idx) = key: idx = idx + 1
    Next key
    Dim i As Long, j As Long, tmp As String
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

'---------------------------------------------------------------------
' Ключи труб ("Тр|20") -> трубы раньше МР, затем по диаметру
Private Sub SortPipeKeys(dict As Object, ByRef arr() As String)
    ReDim arr(0 To dict.Count - 1)
    Dim key As Variant, idx As Long
    idx = 0
    For Each key In dict.Keys
        arr(idx) = key: idx = idx + 1
    Next key
    Dim i As Long, j As Long, tmp As String
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If PipeSortVal(arr(j)) < PipeSortVal(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

'---------------------------------------------------------------------
' Числовой ключ сортировки: тип*100000 + диаметр
Private Function PipeSortVal(ByVal key As String) As Double
    Dim p As Long, kind As String, dia As Double
    p = InStr(key, "|")
    kind = Left$(key, p - 1)
    dia = Val(Replace(Mid$(key, p + 1), ",", "."))
    PipeSortVal = IIf(kind = "Тр", 0, 100000) + dia
End Function

'---------------------------------------------------------------------
' Ключ "Тр|20" -> подпись "Труба Dn20" / "Металлорукав Dn20"
Private Function PipeLabel(ByVal key As String) As String
    Dim p As Long, kind As String, dia As String, nm As String
    p = InStr(key, "|")
    kind = Left$(key, p - 1)
    dia = Mid$(key, p + 1)
    If kind = "Тр" Then nm = "Труба" Else nm = "Металлорукав"
    PipeLabel = nm & " Dn" & dia
End Function


'---------------------------------------------------------------------
' Запись таблицы спецификации: B - наименование, C - кол-во позиций,
' D - экран (для кабелей), E - длина. Возвращает строку итога.
' showScr=False оставляет столбец экрана пустым (трубы).
Private Function WriteSpecTable(dst As Worksheet, _
    ByVal startRow As Long, ByVal nameHead As String, _
    names() As String, cnts() As Long, scrs() As String, _
    lens() As Double, ByVal showScr As Boolean, _
    ByVal showCnt As Boolean) As Long
    Dim i As Long, rr As Long
    Dim totalCnt As Long, totalLen As Double
    ' длина всегда в одном столбце - колонки совпадают на всех
    ' листах спецификации
    dst.Cells(startRow, SPEC_TYPE_DST_COL).Value = nameHead
    If showCnt Then
        dst.Cells(startRow, SPEC_CNT_DST_COL).Value = "Кол-во"
    End If
    If showScr Then dst.Cells(startRow, SPEC_SCR_DST_COL).Value = "Экран"
    dst.Cells(startRow, SPEC_LEN_DST_COL).Value = "Длина, м"
    rr = startRow + 1
    For i = LBound(names) To UBound(names)
        dst.Cells(rr, SPEC_TYPE_DST_COL).Value = names(i)
        If showCnt Then
            dst.Cells(rr, SPEC_CNT_DST_COL).Value = cnts(i)
            totalCnt = totalCnt + cnts(i)
        End If
        If showScr Then dst.Cells(rr, SPEC_SCR_DST_COL).Value = scrs(i)
        dst.Cells(rr, SPEC_LEN_DST_COL).Value = lens(i)
        totalLen = totalLen + lens(i)
        rr = rr + 1
    Next i
    dst.Cells(rr, SPEC_TYPE_DST_COL).Value = "Итого"
    If showCnt Then dst.Cells(rr, SPEC_CNT_DST_COL).Value = totalCnt
    dst.Cells(rr, SPEC_LEN_DST_COL).Value = totalLen
    WriteSpecTable = rr
End Function

'---------------------------------------------------------------------
' Ширина столбца col примерно mmVal миллиметров
Private Sub SetColMM(ws As Worksheet, ByVal col As Long, _
    ByVal mmVal As Double)
    Dim target As Double
    target = MM(mmVal)
    ws.Columns(col).ColumnWidth = 1
    Do While ws.Columns(col).Width < target
        ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth + 0.5
        If ws.Columns(col).ColumnWidth > 255 Then Exit Do
    Loop
End Sub

'---------------------------------------------------------------------
' Рамка + форма 6 + левый штамп на листе спецификации (A4 книжная).
' forcedRow - строка начала таблицы труб (принудительный разрыв
' страницы), 0 - нет. pageBase - число листов до спецификации.
' Возвращает число листов спецификации.
Private Function FrameSpecSheet(ws As Worksheet, _
    ByVal pageBase As Long, ByVal forcedRow As Long) As Long
    Dim cW As Double, cH As Double, frameW As Double, frameH As Double
    cW = MM(SPEC_SHEET_W_MM - FLD_LEFT_MM - FLD_OTHER_MM)
    cH = MM(SPEC_SHEET_H_MM)
    frameW = cW - MM(SPEC_FRAME_ADJUST_MM)
    frameH = cH - MM(FRAME_BOTTOM_MM)

    DeleteFrameShapes ws
    ws.ResetAllPageBreaks

    Dim lastRow As Long
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

    ' проход 1: границы страниц (с учётом принудительного разрыва)
    Dim stampReserve As Double
    stampReserve = MM(STAMP_H + STAMP_GAP)
    Dim brks As Collection
    Set brks = New Collection
    Dim r As Long, acc As Double, limit As Double, rowH As Double
    limit = cH - stampReserve
    acc = 0: r = 1
    Do While r <= lastRow
        rowH = ws.Rows(r).Height
        If ((r = forcedRow) Or (acc + rowH > limit + 0.3)) _
            And acc > 0 Then
            brks.Add Array(r, cH - acc)
            acc = 0
        Else
            acc = acc + rowH
            r = r + 1
        End If
    Loop

    Dim i As Long, itm As Variant
    For i = brks.Count To 1 Step -1
        itm = brks(i)
        InsertFiller ws, CLng(itm(0)) - 1, CDbl(itm(1))
    Next i
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

    ' проход 2: разрывы, рамки, штампы
    Dim pages As Collection
    Set pages = New Collection
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

    Dim k As Long, thisH As Double, sTopa As Double
    Dim sTopb As Double, thisHDop As Double, cnum As Long
    For k = 1 To pages.Count
        pTop = pages(k)
        If k = pages.Count Then
            thisH = frameH - 2
            thisHDop = thisH
        Else
            thisH = frameH
            thisHDop = frameH - 2.5
        End If
        DrawFrameRect ws, pTop, frameW, thisH
        sTopa = pTop + thisH - MM(STAMP_H)
        sTopb = pTop + thisHDop - MM(85)
        cnum = pageBase + k
        DrawStampForm6 ws, frameW - MM(STAMP_W - 20), sTopa, cnum
        DrawStampDopS ws, MM(8), sTopb
    Next k

    ' область печати: до низа последнего листа и до правого края рамки
    Dim lastTop As Double, printLastRow As Long
    lastTop = pages(pages.Count)
    printLastRow = lastRow
    Do While printLastRow < 1000000
        If ws.Rows(printLastRow + 1).Top _
            >= lastTop + frameH + MM(3) Then Exit Do
        printLastRow = printLastRow + 1
    Loop
    Dim pcol As Long, accCol As Double
    accCol = 0
    For i = 1 To SPEC_LEN_DST_COL
        accCol = accCol + ws.Columns(i).Width
    Next i
    pcol = SPEC_LEN_DST_COL
    Do While accCol < MM(20) + frameW + MM(3) And pcol < 16384
        pcol = pcol + 1
        accCol = accCol + ws.Columns(pcol).Width
    Loop

    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = MM(FLD_OTHER_MM)
        .BottomMargin = 0
        .HeaderMargin = 0: .FooterMargin = 0
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "&""Times New Roman""&10&F"
        .RightFooter = "&""Times New Roman""&10Формат А4" & Space(20)
        .CenterHorizontally = False
        .CenterVertically = False
        .Zoom = 100
        .PrintArea = ws.Range(ws.Cells(1, 1), _
            ws.Cells(printLastRow, pcol)).Address
    End With
    Application.PrintCommunication = True

    FrameSpecSheet = pages.Count
End Function

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

```

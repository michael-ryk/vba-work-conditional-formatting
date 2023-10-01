'==================
Public Const CondFormtaPackageVersion As String = "V5"
'==================
Sub CondFormatDailyAnalysis()
'
' Clear all conditional formats and build them from scratch to fix separated formats after manipulation with rows
'

' === Set Constants ===
    Const WeekColumn = "C:C"
    Const DayColumn = "D:D"
    Const ItemColumns = "F:L"
    Const EstimateSumToday = "J4"
    Const ActualSumToday = "K4"
    Const ActualSumTodayCorner = "B2"
    Const BalanceIndicator = "K3"

' === Clear All Conditional Formats ===
    Range("A:R").Cells.FormatConditions.Delete

' === Week Number Yellow bold when = current week ===
    With Range(WeekColumn).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TEXT(WEEKNUM(TODAY())-1,""0"")")
        .Font.ColorIndex = 30
        .Interior.ColorIndex = 36
        .StopIfTrue = False
    End With
    
' === Color Day Number Yellow bold if it is today ===
    With Range(DayColumn).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TODAY()")
        .Font.ColorIndex = 30
        .Interior.ColorIndex = 36
        .StopIfTrue = False
    End With
    
' === Strikeout and Grey out all rows with Done status ===
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, Formula1:= _
        "=$L1=""Done""")
        .Font.Strikethrough = True
        .Font.ColorIndex = 16
        .Interior.ColorIndex = 15
        .StopIfTrue = False
    End With
    
' === Book hours for current day - Color yellow with red if larger that hours left for today ===
    With Range(EstimateSumToday).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=$F$3")
        .Font.Color = -16777024
        .Interior.Color = 65535
        .StopIfTrue = False
    End With
    
' === Actual hours i spent
    'Color green if same as passed hours from moment come to work  ===
    With Range(ActualSumTodayCorner).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=$C$2")
        .Font.Color = -16752384
        .Interior.Color = 13561798
        .StopIfTrue = False
    End With
    
    ' Color red if passed more hours than i reported  ===
    With Range(ActualSumTodayCorner).FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=$E$2")
        .Font.Color = -16383844
        .Interior.Color = 13551615
        .StopIfTrue = False
    End With
    
' === Balance Indicator
    'Color green if balance = 0
    With Range(BalanceIndicator).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="0")
        .Font.Color = -16752384
        .Interior.Color = 13561798
        .StopIfTrue = False
    End With
    
    ' Color red if balance != 0
    With Range(BalanceIndicator).FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="0")
        .Font.Color = -16383844
        .Interior.Color = 13551615
        .StopIfTrue = False
    End With
    
' === Color rows based on Goal  ===
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Wasted""")
        .Interior.Color = 13551615
        .StopIfTrue = False
    End With
    
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Projects""")
        .Interior.Color = 15917529
        .StopIfTrue = False
    End With
        
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Administrative""")
        .Interior.Color = 16777164
        .StopIfTrue = False
    End With
    
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Self_Improve""")
        .Interior.Color = 14083324
        .StopIfTrue = False
    End With
    
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Help_Others""")
        .Interior.Color = 13431551
        .StopIfTrue = False
    End With
    
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Company_Events""")
        .Interior.Color = 13434828
        .StopIfTrue = False
    End With
        
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Lunch""")
        .Interior.Color = 15395562
        .StopIfTrue = False
    End With
    
    With Range(ItemColumns).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F1=""Troubleshooting""")
        .Interior.Color = 11654649
        .StopIfTrue = False
    End With
    
End Sub

'==================
Sub myTasksExcel()
'
' Clear all conditional formats and build them from scratch to fix separated formats after manipulation with rows
'

' === Columns Declaration ===
    Const ageDaysCountColumn As String = "E:E"
    Const overdueDaysCountColumn As String = "F:F"
    Const indicatorColumn As String = "G:G"
    Const statusColumn As String = "K:K"
    Const crosstrikeColumns As String = "A:L"
    Const priorityColumns As String = "D:D"

' === Clear All Conditional Formats ===
    Range("A1").Select
    Cells.FormatConditions.Delete
    
' =======================
' Format Overdue column
' =======================
    Columns(overdueDaysCountColumn).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=-0.1", Formula2:="=-9999"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Application.CutCopyMode = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0.1", Formula2:="=$F$5"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Application.CutCopyMode = False

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=$F$5", Formula2:="=99999"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("F:F").Select
    Application.CutCopyMode = False
    
' =======================
' Format Age Column
' =======================
    Columns(ageDaysCountColumn).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=$E$4", Formula2:="=99999"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Application.CutCopyMode = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=$E$5", Formula2:="=$E$4"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
' =======================
' Format Status column
' =======================
    Columns(statusColumn).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""1-Not Started"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""2-In-Progress"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""3-Wait"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -10209504
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 15917529
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

' =======================
' Format Indicator
' =======================
    Columns(indicatorColumn).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Overdue"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Fresh"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Old"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False


' =======================
' Format underline
' =======================
    Columns(crosstrikeColumns).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$A1=""No"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
        .Color = -5855578
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
' =======================
' Format Priority
' =======================
    Columns(priorityColumns).Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D1=""1-Crit"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D1=""2-Norm"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -10209504
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 15917529
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D1=""3-Low"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16751204
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$A1=""No"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -5855578
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 14277081
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("A1").Select
    
End Sub

Sub MyTasksProjectView()

    'Clear filter if applied
    On Error Resume Next
    ActiveSheet.ShowAllData

    ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort.SortFields. _
        Add2 Key:=Range("TaskTable[Project]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort.SortFields. _
        Add2 Key:=Range("TaskTable[Start]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub MyTasksPriorityView()

    'Clear filter if applied
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort.SortFields. _
        Add2 Key:=Range("TaskTable[Relevant]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort.SortFields. _
        Add2 Key:=Range("TaskTable[Priority]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort.SortFields. _
        Add2 Key:=Range("TaskTable[Age]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tasks").ListObjects("TaskTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub


'==================
Sub RegressionResultShiftRight()
'
' Shift 4 cells from selected right to the right 1 cell
'
Dim mySelectedCells As Range
Set mySelectedCells = Selection.Resize(, 5)
Dim newCells As Range
Set newCells = mySelectedCells.Offset(0, 1)

mySelectedCells.Copy newCells
Selection.Value = "-"

Application.CutCopyMode = False
' Debug.Print (mySelectedCells.Address)

End Sub
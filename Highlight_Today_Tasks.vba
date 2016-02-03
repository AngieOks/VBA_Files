Sub Highlight_Today_Tasks()
'
' Highlight_Today_Tasks Macro
'

' sort_by_due_date Macro
'
    Dim EndRow As Long
    DEndRow = Range("D2").End(xlDown).Row
    FEndRow = Range("F2").End(xlDown).Row
    
    Range("B8").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear   '& stands for string concatenation
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("D2:D" & DEndRow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        '.SetRange Range(Range("A1"), Range("F1").End(xlDown))
        .SetRange Range("A1:F19" & FEndRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


'Select Column B colour settings to default
Range(Range("B2"), Range("B2").End(xlDown)).Select
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
With Selection.Font
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
End With

'Loop through cells until empty cell and see if date is today
Range("B2").Select
Do Until IsEmpty(ActiveCell)
    'remove border first (so only today's tasks have borders)
    Range(ActiveCell, ActiveCell.Offset(0, -1)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    ActiveCell.Offset(0, 1).Select
    
    'if it is today's task add a border, make task red
    If ActiveCell.Offset(0, 2) = Range("H2") Then  'Range("H2") is today's date
     'Highlight Task and Status as Red'
     Range(ActiveCell, ActiveCell.Offset(0, -1)).Select
     Selection.Style = "Bad"
     'Add thick border
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        
        'Highlight done tasks as green
        If (ActiveCell.Value = "Done") Or (ActiveCell.Value = "done") Then
           ActiveCell.Select
           Selection.Style = "Good"
        End If
     ActiveCell.Offset(0, 1).Select
    End If
ActiveCell.Offset(1, 0).Select
'ActiveCell.Offset(1, 1).Select
Loop
End Sub
Sub remove_border()
'
' remove_border Macro
'

'
    Range("B5").Select
    ActiveWindow.SmallScroll Down:=9
    Range("B15").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Sub thick_border()
'
' thick_border Macro
'

'
    Range("B18").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub



Sub Highlight_Rows()
'
' Highlight_Rows Macro
'

'
    Range("B12").Select
    Selection.Style = "Bad"
End Sub
Sub Change_Color_settings()
'
' Change_Color_settings Macro
'

'
    Range("B12:B14").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

Sub Highlight_today_red()

'Loop through cells until empty cell and see if date is today
Range("B2").Select
Do Until IsEmpty(ActiveCell)
    If ActiveCell.Offset(0, 2) = Range("H2") Then  'Range("H2") is today's date
     'Highlight Task and Status as Red'
     Range(ActiveCell, ActiveCell.Offset(0, -1)).Select
     Selection.Style = "Bad"
     'Highlight done tasks as green
     If (ActiveCell.Value = "Done") Or (ActiveCell.Value = "done") Then
        ActiveCell.Select
        Selection.Style = "Good"
     End If
     ActiveCell.Offset(0, 1).Select
    End If
ActiveCell.Offset(1, 0).Select
Loop
End Sub


Sub sort_by_due_date()
'
' sort_by_due_date Macro
'

'
    Dim EndRow As Long
    DEndRow = Range("D2").End(xlDown).Row
    FEndRow = Range("F2").End(xlDown).Row
    Range("B8").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear   '& stands for string concatenation
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("D2:D" & DEndRow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        '.SetRange Range(Range("A1"), Range("F1").End(xlDown))
        .SetRange Range("A1:F19" & FEndRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


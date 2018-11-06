VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateForm 
   Caption         =   "Cycle time"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   OleObjectBlob   =   "DateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Button_Click()

Unload Me

End Sub

Private Sub Clear_Button_Click()

Call UserForm_Initialize

End Sub

Private Sub Month1_Click()

Dim onemonth As Date

If StartDate.Value = "" Then
    Month1.Value = False
    MsgBox ("Start date required")
End If

If Month1.Value = True Then

    onemonth = WorksheetFunction.EoMonth(StartDate.Value, 2)
    onemonth = Str(onemonth)
    EndDate.Value = onemonth

End If

End Sub

Private Sub Month2_Click()

Dim twomonth As Date

If StartDate.Value = "" Then
    Month2.Value = False
    MsgBox ("Start date required")
End If

If Month2.Value = True Then

    twomonth = WorksheetFunction.EoMonth(StartDate.Value, 5)
    onemonth = Str(twomonth)
    EndDate.Value = twomonth

End If


End Sub

Private Sub Month3_Click()

Dim threemonth As Date

If StartDate.Value = "" Then
    Month3.Value = False
    MsgBox ("Start date required")
End If

If Month3.Value = True Then

    threemonth = WorksheetFunction.EoMonth(StartDate.Value, 11)
    threemonth = Str(threemonth)
    EndDate.Value = threemonth

End If

End Sub

Private Sub OK_Button_Click()

Dim next_date As Date

'Ship row is base row
'What is row and column number for the cell in the build row corresponding
'to the 1st date for the 1st part?
startRow = 7
startCol = 9

'Enter the row/column value for the cell indicating the number of days in the
'cycle
daysRow = 2
daysCol = 12

'Enter the row/column value for the cell indicating the start of the cycle
cycleStartRow = 2
cycleStartCol = 10

'Enter the row/column value for the cell indicating the end of the cycle
cycleEndRow = 3
cycleEndCol = 10

'Which row number is the dates row in?
datesRow = 5

'Which row number is the months row in?
monthRow = 4

'Which row number is the weeks row in?
weekRow = 6

'Which column number is the Part ID column in?
partCol = 1

'Which column number is the Total column in?
totalCol = 3

'Which column number is the Level Load column in?
LT_Col = 6

'Which column number is the Level Load column in?
LL_Col = 4

'Which column number is the Safety Stock column in?
SS_Col = 7

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    LastCol = .Cells(datesRow, .Columns.Count).End(xlToLeft).Column
    
    'Unmerges the months
    .Range(.Cells(monthRow, startCol), .Cells(monthRow, LastCol)).Select
    Selection.UnMerge
    
    'Unmerges the weeks
    .Range(.Cells(weekRow, startCol), .Cells(weekRow, LastCol)).Select
    Selection.UnMerge
End With

'Deletes months, dates, weeks, and quantity field
Call delete_all

'Determine length of loop with Networkdays
day_length = WorksheetFunction.NetworkDays(StartDate.Value, EndDate.Value, Worksheets("Exclusion").Range("A2:A71"))
Cells(datesRow, startCol).Select
Call format_dates

'Insert workdays and number of work days
Cells(datesRow, startCol).Value = StartDate.Value
Cells(cycleStartRow, cycleStartCol).Value = StartDate.Value
Cells(cycleEndRow, cycleEndCol).Value = EndDate.Value
Cells(daysRow, daysCol).Value = day_length

Cells(weekRow, startCol).Select
Call format_weeks

'Enter the month in a cell directly above a cell containing a corresponding date
For j = startCol To day_length + 8

Cells(monthRow, j).Value = WorksheetFunction.Text(Cells(datesRow, j), "mmmm")

    If j > startCol Then
        next_date = WorksheetFunction.WorkDay(Cells(5, j - 1).Value, 1, Worksheets("Exclusion").Range("A2:A71"))
        Cells(datesRow, j).Value = next_date
        Cells(monthRow, j).Value = WorksheetFunction.Text(Cells(5, j), "mmmm")
        Cells(datesRow, j).Select
        Call format_dates
        Cells(weekRow, j).Select
        Call format_weeks
    End If
Next j

'Enter the week in a cell directly below a cell containing a corresponding date
n = 1

For j = startCol To day_length + 8

If j = startCol Then
    Cells(weekRow, j).Value = "Week " & n
End If

If j <> day_length + 8 Then
    If Cells(datesRow, j + 1).Value = Cells(datesRow, j).Value + 1 Then
        Cells(weekRow, j + 1).Value = "Week " & n
    ElseIf Cells(datesRow, j + 1).Value <> Cells(datesRow, j).Value + 1 Then
        n = n + 1
        Cells(weekRow, j + 1).Value = "Week " & n
    End If
End If

Next j

k = 0

'Identify and merge months
For j = startCol To day_length + 8

If k > startCol And k < day_length + 8 Then
    j = k
End If

k = j + 1

    If Cells(monthRow, j).Value = Cells(monthRow, k).Value And Cells(monthRow, k).Value <> "" Then
        Do While Cells(monthRow, j).Value = Cells(monthRow, k).Value
            k = k + 1
        Loop
    End If

    If Cells(monthRow, j).Value = "" Then
        GoTo EndWeekLoop
    End If

    With ActiveSheet
        .Range(.Cells(monthRow, j), .Cells(monthRow, k - 1)).Select
        .Cells(monthRow, j).Interior.Color = RGB(38, 38, 38)
    End With

    Selection.Merge

EndWeekLoop:
Next j

k = startCol

'Identify and merge weeks
For j = startCol To day_length + 8

If k > startCol And k < day_length + 8 Then
    j = k
End If

k = j + 1

    'If the date in the adjacent cell is only 1 day greater than the date in the
    'current cell, it is the same week; otherwise, merge range and restart counter

    If Cells(weekRow, k).Value = Cells(weekRow, j).Value And Cells(weekRow, k).Value <> "" Then
        Do While Cells(weekRow, k).Value = Cells(weekRow, j).Value
            k = k + 1
        Loop
    End If

    If Cells(weekRow, j).Value = "" Then
        GoTo EndMonthLoop
    End If

    With ActiveSheet
        .Range(.Cells(weekRow, j), .Cells(weekRow, k - 1)).Select
        .Cells(weekRow, j).Interior.Color = RGB(204, 192, 218)
    End With

    Selection.Merge

EndMonthLoop:
Next j

With ActiveSheet
    .Range(.Cells(weekRow, 19), .Cells(weekRow, day_length + 8)).Select
    Selection.EntireColumn.AutoFit
    .Range(.Cells(monthRow, 9), .Cells(weekRow, day_length + 8)).Select
    Call bold_border
End With

Unload Me

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

StartDate.Value = ""
EndDate.Value = ""

Month1.Value = False
Month2.Value = False
Month3.Value = False

End Sub

Sub format_dates()

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

Sub format_sheet()
    Range("I5:ZZ5500").Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("H7:H526").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
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
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
End Sub

Sub delete_all()

    Range("D7:D2000").Select
    Range("D7").Activate
    Selection.ClearContents
    Range("G7:G2000").Select
    Selection.ClearContents
    
    Range("I4:MZ2000").Select
    
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    With ActiveSheet
         LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range(.Cells(7, 8), .Cells(LastRow, 8)).Select
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
    Range("E4").Select
    
End Sub
Sub format_weeks()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub bold_border()
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
End Sub

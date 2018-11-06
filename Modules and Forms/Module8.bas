Attribute VB_Name = "Module8"
Sub UpdateWeekly()

Application.ScreenUpdating = False

Worksheets("Summary").Activate

With ActiveSheet
    LastSumRow = .Cells(.Rows.Count, "B").End(xlUp).Row
    LastSumCol = .Cells(4, .Columns.Count).End(xlToLeft).Column
End With

'The first row that needs to be written to on the Summary sheet
sumRow = 6

Worksheets(1).Activate

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    LastCol = .Cells(5, .Columns.Count).End(xlToLeft).Column
End With


'Cycle through all rows
For i = 7 To LastRow
sumCol = 5

'k is a counter for the number of weeks identified
k = 0

    'Cycle through all columns until the 4th week has been identified
    For j = 9 To LastCol
    
        startCol = j
        
        If Cells(5, j).Value = "" Or k >= 4 Then
            GoTo Next_Iteration
        End If

        'Is the next cell only 1 day greater than the current cell?
        'If so, then designate the next cell as the last day
        
        Do While Cells(5, j + 1).Value = Cells(5, j).Value + 1
            lastCell = j + 1
            j = j + 1
        Loop
        
        'A week has been found
        k = k + 1
        endCol = j
        
        'Last day of week identified; sum values and insert into summary page
        Sum = Application.WorksheetFunction.Sum(Range(Cells(i + 2, startCol), Cells(i + 2, endCol)))
        StartDate = Cells(5, startCol).Value
        EndDate = Cells(5, endCol).Value
        
        Worksheets("Summary").Activate
        
        'If this is the first pass then copy the date ranges for each week
        If i = 7 Then
            Cells(sumRow - 1, sumCol).Value = StartDate
        End If
        
        'Write sum value for this week to Summary sheet
        Cells(sumRow, sumCol).Value = Sum

        'Increment the row to write to by 1
        sumCol = sumCol + 1
        Cells(sumRow, 2).Value = Worksheets(1).Cells(i, 1).Value
        Cells(sumRow, 4).Value = Worksheets(1).Cells(i, 7).Value
        Worksheets(1).Activate
        
Next_Iteration:
    Next j

sumRow = sumRow + 1
i = i + 3
Next i

Worksheets("Summary").Activate

Call RefreshAllPivotTables

End Sub

Private Sub RefreshAllPivotTables()

Dim PT As PivotTable
Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    
        For Each PT In ws.PivotTables
          PT.RefreshTable
        Next PT
    Next ws
    
End Sub

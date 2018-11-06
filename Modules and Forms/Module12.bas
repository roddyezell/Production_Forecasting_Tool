Attribute VB_Name = "Module12"
Sub SS_Calculation()

'Build row is base row
'What is row and column number for the cell in the build row corresponding
'to the 1st date for the 1st part?
startRow = 9
startCol = 9

'Enter the row/column value for the cell indicating the number of days in the
'cycle
daysRow = 2
daysCol = 12

'Which column number is the Total column in?
totalCol = 3

'Which column number is the Level Load column in?
LL_Col = 4

'Which column number is the Safety Stock column in?
SS_Col = 7

Application.ScreenUpdating = True

Dim x As Single, have As Single, SS_add As Single, need As Single
Dim LL As Single, i As Integer, j As Integer, k As Integer
Dim pctCompl As Single, status As String, valFound As Single
Dim rowVal As Single, colVal As Single, pctCompl_actual As Single, sumFuture As Integer

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    LastCol = .Cells(5, .Columns.Count).End(xlToLeft).Column
End With

'Build row is base row
For i = startRow To LastRow + 2
    
    rowVal = i
    have = 0
    SS_add = 0
    k = 0
    LL = (Worksheets(1).Cells(i - 2, totalCol).Value) / (Worksheets(1).Cells(daysRow, daysCol).Value)
    Worksheets(1).Cells(i - 2, LL_Col).Value = Application.WorksheetFunction.RoundUp(LL, 0)
    LL = Worksheets(1).Cells(i - 2, LL_Col).Value
    
Next_j:

    For j = startCol To LastCol

        colVal = j
        
        If k = 1 And j = startCol Then
            have = Worksheets(1).Cells(i - 2, SS_Col).Value
        End If
        
         'Checks the sum of future orders
        sumFuture = Application.Sum(Range(Cells(i - 1, j), Cells(i - 1, LastCol)))
    
        need = Worksheets(1).Cells(i - 1, j).Value
        x = have + LL - need
        
        If k = 1 And have >= sumFuture Then
            x = have - need
        End If
        
        'Checks whether the sum of future orders exceeds current stock amount
        'If it does, then stop building that part.
        
        If have < sumFuture Then
            Worksheets(1).Cells(i, j).Value = LL
            Worksheets(1).Cells(i, j).Interior.Color = RGB(197, 217, 241)
        
        ElseIf have >= sumFuture Then
            Worksheets(1).Cells(i, j).Value = 0
            Worksheets(1).Cells(i, j).Interior.Color = RGB(197, 217, 241)
        End If
        
        'If have+LL-need is less than zero, increase safety stock
        If x < 0 And k <> 2 Then
            SS_add = SS_add + (x * -1)
            have = 0
        End If

        'If (have + LL - need) is greater than or equal to zero
        'then we will have inventory on hand for the next order
        If x >= 0 Then
            have = x
        End If
        
        'Insert stock quantities
            If j = startCol Then
                Worksheets(1).Cells(i + 1, j).Value = Worksheets(1).Cells(i - 2, SS_Col).Value
                Worksheets(1).Cells(i + 1, j).Interior.Color = RGB(153, 255, 153)
                Worksheets(1).Cells(i + 1, j + 1).Value = have
            End If
            
            If j > startCol And j <> LastCol Then
                Worksheets(1).Cells(i + 1, j + 1).Value = have
                Worksheets(1).Cells(i + 1, j).Interior.Color = RGB(153, 255, 153)
            End If
            
            If j = LastCol Then
                Worksheets(1).Cells(i + 1, j).Interior.Color = RGB(153, 255, 153)
            End If
    
Skip_Col:
    Next j
    
    Worksheets(1).Cells(i - 2, SS_Col).Value = SS_add
    
    'Go through all columns again with calculated SS value
    k = k + 1
    If k = 1 Then
        GoTo Next_j
    End If
        
i = i + 3
Next i

Call Delinquent

End Sub

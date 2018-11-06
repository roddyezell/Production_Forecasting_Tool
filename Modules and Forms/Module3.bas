Attribute VB_Name = "Module3"
Sub Delinquent()

'Ship row is base row
'What is row and column number for the cell in the build row corresponding
'to the 1st date for the 1st part?
startRow = 7
startCol = 9

'Enter the row/column value for the cell indicating the number of days in the
'cycle
daysRow = 2
daysCol = 12

'Which column number is the Delinquent column in?
delQty = 5
delDays = 5

'Which column number is the Total column in?
totalCol = 3

'Which column number is the Level Load column in?
LL_Col = 4

'Which column number is the Safety Stock column in?
SS_Col = 7

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    LastCol = .Cells(5, .Columns.Count).End(xlToLeft).Column
End With

'Base row is Ship;
For i = startRow To LastRow

        qty = Cells(i + 1, delQty).Value
        orig_qty = qty
        days = Cells(i + 3, delDays).Value
       
    If qty <> "" And qty > 0 Then

        LL_add = qty / days
        round_LL_add = Application.WorksheetFunction.RoundUp(LL_add, 0)
        LL_new = LL + round_LL_add
        endAdd = days + 8
        
        For j = startCol To endAdd
            'Increase level load amount temporarily
            If qty > 0 Then
                Cells(i + 2, j).Value = Cells(i + 2, j).Value + round_LL_add
                qty = qty - round_LL_add
            End If
        Next j
        
        'Arithmetic
        For j = startCol To LastCol

        If j > startCol Then
            Cells(i + 3, j).Value = Cells(i + 3, j - 1).Value + Cells(i + 2, j - 1).Value - Cells(i + 1, j - 1).Value
        End If

        Next j
    
    End If

i = i + 3
Next i

End Sub

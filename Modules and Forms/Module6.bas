Attribute VB_Name = "Module6"
Sub SS_Calculation2()

Dim x As Single, have As Single, SS_add As Single, need As Single
Dim LL As Single, i As Integer, j As Integer, k As Integer, sumFuture As Integer

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    LastCol = .Cells(5, .Columns.Count).End(xlToLeft).Column
End With

'Build row is base row
For i = 9 To LastRow + 2
    
    have = 0
    SS_add = 0
    k = 0
    LL = Cells(i - 2, 4).Value
    
Next_j:

    For j = 9 To LastCol

If Cells(5, j).Value > Cells(i - 2, 2).Value Then
    GoTo Skip_Col
End If
        
        If k = 1 And j = 9 Then
            have = Worksheets(1).Cells(i - 2, 7).Value
        End If
        
        'Checks the sum of future orders
        sumFuture = Application.Sum(Range(Cells(i - 1, j), Cells(i - 1, LastCol)))
        need = Cells(i - 1, j).Value
        x = have + LL - need
        
        If k = 1 And have >= sumFuture Then
            x = have - need
        End If
        
        If have < sumFuture Then
            Cells(i, j).Value = LL
            Cells(i, j).Interior.Color = RGB(197, 217, 241)
        
        ElseIf have >= sumFuture Then
            Cells(i, j).Value = 0
            Cells(i, j).Interior.Color = RGB(197, 217, 241)
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
            If j = 9 Then
                Cells(i + 1, j).Value = Worksheets(1).Cells(i - 2, 7).Value
                Cells(i + 1, j).Interior.Color = RGB(153, 255, 153)
                Cells(i + 1, j + 1).Value = have
            End If
            
            If j > 9 And j <> LastCol Then
                Cells(i + 1, j + 1).Value = have
                Cells(i + 1, j).Interior.Color = RGB(153, 255, 153)
            End If
            
            If j = LastCol Then
                Cells(i + 1, j).Interior.Color = RGB(153, 255, 153)
            End If
    
Skip_Col:
    Next j
    
    Cells(i - 2, 7).Value = SS_add
    
    'Go through all columns again with calculated SS value
    k = k + 1
    If k = 1 Then
        GoTo Next_j
    End If
        
        
i = i + 3
Next i

Call Delinquent

End Sub


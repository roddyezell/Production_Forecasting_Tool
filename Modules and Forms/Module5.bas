Attribute VB_Name = "Module5"
Sub Delete_40m()

Application.ScreenUpdating = False

Worksheets("Data").Activate

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
End With

x = LastRow

Do While x > 2
   
    If Cells(x, 2).Value < 50000000 Then
        Worksheets("Data").Rows(x).Delete
    End If
    
    x = x - 1
    
Loop

Call shipdates

Worksheets(1).Activate

End Sub


Sub shipdates()
Attribute shipdates.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A2000"), Type:=xlFillDefault
    Range("A2:A2000").Select

End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "In progress"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()

code

End Sub

Private Sub UserForm_Activate()

Application.ScreenUpdating = False
code

End Sub

Sub code()

'Ship row is base row
'What is row and column number for the cell in the build row corresponding
'to the 1st date for the 1st part?
startRow = 7
startCol = 9

'Enter the row/column value for the cell indicating the number of days in the
'cycle
daysRow = 2
daysCol = 12

'Which row number is the dates row in?
datesRow = 5

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

Dim i As Integer, j As Integer, k As Integer, p As Integer, n As Integer
Dim pctCompl As Single, status As String, valFound As Single
Dim rowVal As Single, colVal As Single, pctCompl_actual As Single

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    LastCol = .Cells(datesRow, .Columns.Count).End(xlToLeft).Column
End With

Worksheets("Pivot").Activate

With ActiveSheet
        LastPivotRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

Worksheets("Analysis").Activate

valFound = 0
status = "MRP Data updating..."
progress pctCompl, status, rowVal, colVal, valFound

'Clear contents of qty field

With ActiveSheet
    .Range(.Cells(startRow, startCol), .Cells(500, 78)).Select
    Selection.ClearContents
    Selection.Interior.Color = xlNone
End With

Range("A1").Select

'Refresh MRP Data
Call refresh_data

'Remove orders below 50,000,000 from the Data sheet
Call Delete_40m

pctCompl_actual = 10
pctCompl = pctCompl_actual
status = "Refreshing pivot table..."
progress pctCompl, status, rowVal, colVal, valFound

'Refresh Pivot Table
Call RefreshAllPivotTables

pctCompl_actual = 20
pctCompl = pctCompl_actual
status = "Populating Qty field..."
progress pctCompl, status, rowVal, colVal, valFound

'Cycle through columns on main sheet
For i = startCol To LastCol

    'Cycle through rows on main sheet
        For j = startRow To LastRow
            
        With ActiveSheet
            .Range(.Cells(j + 3, 1), .Cells(j + 3, LastCol)).Select
        End With
        
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        Worksheets(1).Cells(j, i).Value = 0
        Worksheets(1).Cells(j, i).Interior.Color = RGB(255, 153, 153)
        Worksheets(1).Cells(j + 1, i).Value = 0
        Worksheets(1).Cells(j + 1, i).Interior.Color = RGB(252, 213, 180)
        
            rowVal = j
            colVal = i
            progress pctCompl, status, rowVal, colVal, valFound
                        
            'Cycle through rows on pivot sheet
            
                For k = 5 To LastPivotRow
                
                    'Does a date on the ship date column from the pivot table match a date
                    'on the 3-month forecast?

                    If Worksheets(2).Cells(k, 6).Value = Worksheets(1).Cells(datesRow, i).Value Then
    
                    'Does a part on the parts ID column from the pivot table match the part
                    'on the 3-month forecast for the given date?
    
                        If Worksheets(2).Cells(k, 7).Value = Worksheets(1).Cells(j, partCol).Value Then
    
                        Worksheets(1).Cells(j, i).Value = Worksheets(2).Cells(k, 8).Value
                        Worksheets(1).Cells(j, i).Interior.Color = RGB(255, 153, 153)
                                                
                        valFound = valFound + 1
                        progress pctCompl, status, rowVal, colVal, valFound
                        
                        'Place braze finish date
                        
                            If i >= startCol + Cells(j, LT_Col).Value Then
                        
                            n = i - Cells(j, LT_Col).Value
                                Worksheets(1).Cells(j + 1, n).Value = Worksheets(2).Cells(k, 8).Value
                                Worksheets(1).Cells(j + 1, n).Interior.Color = RGB(252, 213, 180)
            
                            End If
                        
                            If i < (startCol + Cells(j, LT_Col).Value) Then
                            
                                Worksheets(1).Cells(j + 1, startCol).Value = Worksheets(2).Cells(k, 8).Value
                                Worksheets(1).Cells(j + 1, startCol).Interior.Color = RGB(252, 213, 180)
                            
                            End If
                        
                            End If

                    End If
                    
                Next k
                
Skip_Row:
        
        j = j + 3
        Next j
    
        pctCompl_actual = pctCompl_actual + 1
        pctCompl = Application.WorksheetFunction.Round(pctCompl_actual, 0)
        status = "Populating Qty field..."
        progress pctCompl, status, rowVal, colVal, valFound

Next i

pctCompl = 100
rowVal = 0
colVal = 0
status = ""
progress pctCompl, status, rowVal, colVal, valFound

Unload Me

Range("Q2").Value = (Now)
Range("Q3").Value = TimeValue(Now)

End Sub

Sub progress(pctCompl As Single, status As String, rowVal As Single, colVal As Single, valFound As Single)

ProgressBar.Col_Val.Caption = "Column: " & colVal
ProgressBar.Row_Val.Caption = "     Row: " & rowVal
ProgressBar.Val_found.Caption = "Items found:  " & valFound

ProgressBar.State.Caption = status
ProgressBar.Text.Caption = pctCompl & "% Completed"
ProgressBar.Bar.Width = pctCompl * 2

DoEvents

End Sub
Sub refresh_data()

    ActiveWorkbook.XmlMaps("Order_Navigator_Map").DataBinding.Refresh
    
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

Attribute VB_Name = "Module2"
Sub bonus():
For Each ws In Worksheets
    'define last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'BONUS
    '--------------------------------------------
    'Set greatest % increase variable
    Dim Percent_Max As Double
    'Set initial max of 0 and loop through rows
    Percent_Max = 0
    'Greatest increase ticker
    Dim Increase_Ticker As String
    'Set greatest % decrease variable
    Dim Percent_Min As Double
    'Set initial max of 0 and loop through rows
    Percent_Min = 0
    'Greatest decrease ticker
    Dim Decrease_Ticker As String
    'Set greatest total stock volume variable
    Dim Volume_Max As LongLong
    'Set initial max of 0 and loop through rows
    Volume_Max = 0
    'Greatest volume ticker
    Dim Volume_Ticker As String
        
    For i = 2 To lastrow
    
        If ws.Cells(i, 11).Value > Percent_Max Then
            Percent_Max = ws.Cells(i, 11).Value
            ws.Cells(2, 17).Value = Percent_Max
            ws.Cells(2, 17).NumberFormat = "0.00%"
            Increase_Ticker = ws.Cells(i, 9).Value
            ws.Cells(2, 16).Value = Increase_Ticker
        'Set actual greatest decrease in each sheet
        ElseIf ws.Cells(i, 11).Value < Percent_Min Then
            Percent_Min = ws.Cells(i, 11).Value
            ws.Cells(3, 17).Value = Percent_Min
            ws.Cells(3, 17).NumberFormat = "0.00%"
            Decrease_Ticker = ws.Cells(i, 9).Value
            ws.Cells(3, 16).Value = Decrease_Ticker
        Else
        End If
        
        'Set actual greatest total stock volume in each sheet
        If ws.Cells(i, 12).Value > Volume_Max Then
            Volume_Max = ws.Cells(i, 12).Value
            ws.Cells(4, 17).Value = Volume_Max
            Volume_Ticker = ws.Cells(i, 9).Value
            ws.Cells(4, 16).Value = Volume_Ticker
        Else
        End If
    Next i
Next ws
End Sub


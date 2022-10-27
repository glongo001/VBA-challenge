Attribute VB_Name = "Module1"
Sub alphabetical_testing():

'Loop through all worksheets in book
For Each ws In Worksheets
    'Variable that holds file name, and last row
    Dim WorksheetsName As String
    'define last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Store worksheet name
    WorksheetName = ws.Name
    'NAME ALL COLUMN AND ROW HEADERS FOR OBTAINED DATA AND SUMMARY
    '----------------------------------------------------------
    'Add Ticker Header to Column I
    ws.Cells(1, 9).Value = "Ticker Symbol"
    'Add Yearly Change header to Column J
    ws.Cells(1, 10).Value = "Yearly Change ($)"
    'Add Yearly Change header to Column K
    ws.Cells(1, 11).Value = "Percent Change"
    'Add Yearly Change header to Column L
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'Add Yearly Change header to Column J
    ws.Cells(2, 15).Value = "Greatest % Increase"
    'Add Yearly Change header to Column J
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    'Add Yearly Change header to Column J
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    'Add Yearly Change header to Column J
    ws.Cells(1, 16).Value = "Ticker Symbol"
    'Add Yearly Change header to Column J
    ws.Cells(1, 17).Value = "Value"
    '----------------------------------------------------------
    
    'Set initial varible for ticker symbol
    Dim ticker As String
    
    'Track the location of each ticker symbol in worksheet
    Dim New_Table_Row As Integer
    New_Table_Row = 2
    
    'Set variable for price change within a given year
    Dim Price_Change As Double
    
    'Set variable for percent change within a given year
    Dim Percent_Change As Double
    
    'Set variable for total stock volume
    Dim Stock_Volume As LongLong
    'Set stock volume equal to 0 before starting loop
    Stock_Volume = 0
    
    Dim open_date As Long
    Dim close_date As Long
    
    Dim open_value As Double
    Dim close_value As Double

    'Start loop
    For i = 2 To lastrow
        'Check if we are within the same ticker symbol, if we are not then:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set ticker symbol
            ticker = ws.Cells(i, 1).Value
            'Copy ticker symbol to table with obtained information
            ws.Range("I" & New_Table_Row).Value = ticker
            
            'Define the stock closing date
            close_date = ws.Cells(i, 2).Value
            'Set the stock closing value
            close_value = ws.Cells(i, 6).Value
            'Calculate price change
            Price_Change = close_value - open_value
            'Insert price change in obtained data table
            ws.Range("J" & New_Table_Row).Value = Price_Change
            'Calculate percent change
            Percent_Change = Price_Change / open_value
            'Insert percent change in obtained data table
            ws.Range("K" & New_Table_Row).Value = Percent_Change
            'Change number format to percentage
            ws.Range("K" & New_Table_Row).NumberFormat = "0.00%"
            
            'Add to total stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            'Insert total stock volume in obtained data table
            ws.Range("L" & New_Table_Row).Value = Stock_Volume
            'Add one to obtained data table row
            New_Table_Row = New_Table_Row + 1
            'Reset stock volume
            Stock_Volume = 0
        
        'If cell before a row has a different ticker symbol
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Obtain stock opening date
            open_date = ws.Cells(i, 2).Value
            'Set the opening stock value
            open_value = ws.Cells(i, 3).Value
        'If a cell following a row has the same ticker symbol
        Else
            'Add to stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        End If
        
        'Conditional formatting looking at price change cells
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
Next ws
End Sub


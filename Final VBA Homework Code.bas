Attribute VB_Name = "Module1"
Sub stockAnalysis()

'loop through all the worksheets in the work book
For Each ws In Worksheets

'declare variables
Dim ticker As String
Dim totalVolume As Double
Dim lastRow As Long
Dim yearOpen As Double
Dim yearClose As Double
Dim yearChange As Double
Dim percentChange As Double
Dim i As Long
Dim summaryTableRow As Integer


'column headers for ticker, yearly change, percent change, total stock volume
'greatest % increase, greatest % decrease, greatest total volume, ticker, value
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'start the summary fields in row 2
summaryTableRow = 2

'find last row of worksheet
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'set volume counter to 0
totalVolume = 0

'first year open value
yearOpen = ws.Cells(2, 3).Value

'Loop through from starting row to last row
For i = 2 To lastRow

    'check if we are in the same ticker name, if not then
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'find the ticker name
        ticker = ws.Cells(i, 1).Value
        
        'output ticker name into summary table
        ws.Cells(summaryTableRow, 9).Value = ws.Cells(i, 1).Value
        
        'output the total volume of ticker into summary table
        ws.Cells(summaryTableRow, 12).Value = totalVolume + ws.Cells(i, 7).Value
        
        'reset volume total back to 0 to start counting the next ticker
        totalVolume = 0
        
        'if the ticker values are equal then
        Else
    
        'continue adding total volume until ticker name is not the same
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
    End If
    
    'check to see if ticker is the same as above and different from below then
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'find year close value at current row
    yearClose = ws.Cells(i, 6).Value
    
    'calculate year change value
    yearChange = yearClose - yearOpen
    
    
        'if year open is zero, then percent change will be 0
        If yearOpen = 0 And yearChange <> 0 Then
            
            percentChange = yearClose / yearClose
            
        ElseIf yearOpen = 0 And yearClose = 0 Then
        
        percentChange = 0
        
        'calculate percent change if year open is not 0
        Else
        
        percentChange = yearChange / yearOpen
        
        End If
        
    'move to next year open value if ticker is the same
    yearOpen = ws.Cells(i + 1, 3).Value
        
    'output year change in the summary table
    ws.Cells(summaryTableRow, 10).Value = yearChange
        
        'conditional formatting for year change, positive will be green, negative will be red
        If ws.Cells(summaryTableRow, 10).Value < 0 Then
        
            'set negative values to red
            ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
            
        Else
        
            'set positive values to green
            ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
            
        End If
        
        
    'output percentChange into summary table
    ws.Cells(summaryTableRow, 11).Value = percentChange
    
    'format percent change as percentage in cell to two decimal places
    ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
    
    'move to next row in summary table for next ticker result
        summaryTableRow = summaryTableRow + 1
        
    End If
    
        
Next i

Next ws


End Sub


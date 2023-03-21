Attribute VB_Name = "Module11"
Sub Module2_script()
'establish variables
Dim rowcount As Long
Dim ticker As Long
Dim volume As Long
Dim rowindex As Integer
Dim count As Long
Dim current As Worksheet

'starting the loop to go through all worksheets
For Each ws In Worksheets



'establish headings
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly_Change"
    ws.Range("L1").Value = "Percent_Change"
    ws.Range("M1").Value = "Total_stock_volume"

'Summary headings
    ws.Range("O2").Value = "Greatest Percent Increase"
    ws.Range("O3").Value = "Greatest Percent Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"




'start from the bottom of the sheet and work up until it finds the
'final row that is not empty. then sets this row as the final count of filled rows in sheet
    rowcount = ws.Cells(Rows.count, 1).End(xlUp).row

'setting start point for variables once in current sheet
    column = 1
    ticker = 2
    volume = 0
    count = 0
    
    'looping through rows until the end of current sheet
    For i = 2 To rowcount
        
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
           count = count + 1
  'separate ifs create, closing price variabe and opening price variable
        
        'checking tio see if there is a new ticker
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
        'place ticker string into appropriate cell
           ws.Cells(ticker, 10).Value = ws.Cells(i, 1).Value
        'now, when ticker is different, find (closing price - opening price)
           ws.Cells(ticker, 11).Value = ws.Cells(i, 6).Value - ws.Cells(i - count, 3).Value
        'additionally, calculate yearly %change and input appropriately
           ws.Cells(ticker, 12).Value = (ws.Cells(i, 6).Value / ws.Cells(i - count, 3).Value) - 1
        'sum the volume of sales
           ws.Cells(ticker, 13).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(i, 7), ws.Cells(i - count, 7)))
        
        count = 0
        
        'establish color formatting
        color_rule = ws.Cells(ticker, 11).Value
        Select Case color_rule
         'if value is positive, color green
        Case Is > 0
          ws.Cells(ticker, 11).Interior.Color = vbGreen
        'if value is negative, color red
        Case Is < 0
          ws.Cells(ticker, 11).Interior.ColorIndex = 3
        'otherwise no color change
        Case Else
          ws.Cells(ticker, 11).Interior.ColorIndex = 0
        End Select
        
    'format to make percent column as such
    ws.Cells(ticker, 12).Value = FormatPercent(ws.Cells(ticker, 12))

' move down 1 row in the yearly summary columns
    ticker = ticker + 1

'complete actions for 1 ticker symbol
    End If

'search until ticker symbols change again
    Next i

'Summary Data Values
    ' get max percent increase
    ws.Range("Q2").Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("L:L")))
    ' get max percent decrease
    ws.Range("Q3").Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("L:L")))
    ' get greatest volume
    ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("M:M"))
    
    'format to percent again
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    
    
'Summary Data - corresponding ticker
    maxtick = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0)
    ws.Range("P2") = ws.Cells(maxtick, 10)
    mintick = WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("L:L")), ws.Range("L:L"), 0)
    ws.Range("P3") = ws.Cells(mintick, 10)
    voltick = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("M:M")), ws.Range("M:M"), 0)
    ws.Range("P4") = ws.Cells(voltick, 10)

    


Next ws

End Sub



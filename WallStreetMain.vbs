Sub WallStreetMain()

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------

Dim ws As Worksheet

For Each ws In Worksheets

 
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Set Col 1, the column has the ticker letter

'Header
ws.Range("H" & 1).Value = "Ticker"
ws.Range("I" & 1).Value = "YearlyChange"
ws.Range("J" & 1).Value = "PercentChange"
ws.Range("K" & 1).Value = "TotalStockVolumn"

  
  Dim column As Integer
  column = 1
  
'Ticker_Symbol_Row is where I put the Ticker Symbol letter when the value changes.
'Excluding Header,all data values start from row 2

  Dim Ticker_Symbol_Row As Integer
  Ticker_Symbol_Row = 2
  
  Dim Year_Open As Double
  Year_Open = ws.Cells(2, 3).Value
  
  Dim Year_Close As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  Dim Total_Vol As Double
  Total_Vol = 0
  
  ws.Cells(Ticker_Symbol_Row, 8) = ws.Cells(2, column).Value
  

  ' Loop through rows
  For i = 2 To LastRow
  
  Total_Vol = Total_Vol + ws.Cells(i, 7)
    
    ' Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
    
    'Total_Vol = Total_Vol + ws.Cells(i, 7)
    
    Year_Close = ws.Cells(i, 6).Value
    
    Yearly_Change = Year_Close - Year_Open
    
    If Yearly_Change >= 0 Then
    ws.Cells(Ticker_Symbol_Row, 9).Interior.ColorIndex = 4
    Else
    ws.Cells(Ticker_Symbol_Row, 9).Interior.ColorIndex = 3
    End If
    
    ws.Cells(Ticker_Symbol_Row, 9) = Yearly_Change
    
    If Year_Open = 0 Then
    Percent_Change = 1
    Else
    'Percent_Change = Round((Yearly_Change / Year_Open) * 100, 2)
    Percent_Change = Yearly_Change / Year_Open
    End If
    ws.Cells(Ticker_Symbol_Row, 10).NumberFormat = "0.00%"
    ws.Cells(Ticker_Symbol_Row, 10).Value = Percent_Change
    
    ws.Cells(Ticker_Symbol_Row, 11) = Total_Vol
    
    Total_Vol = 0

    Ticker_Symbol_Row = Ticker_Symbol_Row + 1
    
    ws.Cells(Ticker_Symbol_Row, 8) = ws.Cells(i + 1, column).Value
      
    Year_Open = ws.Cells(i + 1, 3).Value
      
    End If
 
  Next i

Next ws

End Sub





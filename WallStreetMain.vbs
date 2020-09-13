Sub WallStreetMain()

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------

Dim ws As Worksheet

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Header Columns
ws.Range("I" & 1).Value = "Ticker"
ws.Range("J" & 1).Value = "YearlyChange"
ws.Range("K" & 1).Value = "PercentChange"
ws.Range("L" & 1).Value = "TotalStockVolumn"
ws.Range("O" & 2).Value = "Greatest % Increase"
ws.Range("O" & 3).Value = "Greatest % Decrease"
ws.Range("O" & 4).Value = "Greatest Total Volumn"
ws.Range("P" & 1).Value = "Ticker"
ws.Range("Q" & 1).Value = "Value"

ws.Range("Q2:Q4").NumberFormat = "0.00%"

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
  Dim Greatest_Increase As Double
  Greatest_Increase = 0
  Dim Greatest_Decrease As Double
  Greatest_Descrease = 0
  Dim Greatest_Total As Double
  Greatest_Total = 0
  Dim Ticker_I As String
  Dim Ticker_D As String
  Dim Ticker_T As String
  
'First Ticker Symbol Row Value before loop
  ws.Cells(Ticker_Symbol_Row, 9) = ws.Cells(2, column).Value
  
' Loop through rows
  For i = 2 To LastRow
  
  Total_Vol = Total_Vol + ws.Cells(i, 7)
    
    ' Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        
    Year_Close = ws.Cells(i, 6).Value
    
    Yearly_Change = Year_Close - Year_Open
    
    If Yearly_Change >= 0 Then
        ws.Cells(Ticker_Symbol_Row, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(Ticker_Symbol_Row, 10).Interior.ColorIndex = 3
    End If
    
    ws.Cells(Ticker_Symbol_Row, 10) = Yearly_Change
    
    If Year_Open = 0 Then
        Percent_Change = 1
    Else
        Percent_Change = Yearly_Change / Year_Open
    End If
    
    'Compare Percent chage increase or decrease
    If Percent_Change >= 0 Then
        If Percent_Change >= Greatest_Increase Then
            Greatest_Increase = Percent_Change
            Ticker_I = ws.Cells(Ticker_Symbol_Row, 9).Value
        End If
    End If
    ws.Range("P" & 2).Value = Ticker_I
    ws.Range("Q" & 2).Value = Greatest_Increase

    If Percent_Change < 0 Then
        If Percent_Change < Greatest_Decrease Then
            Greatest_Decrease = Percent_Change
            Ticker_D = ws.Cells(Ticker_Symbol_Row, 9).Value
        End If
    End If
    ws.Range("P" & 3).Value = Ticker_D
    ws.Range("Q" & 3).Value = Greatest_Decrease
    
    If Total_Vol >= 0 Then
        If Total_Vol >= Greatest_Total Then
            Greatest_Total = Total_Vol
            Ticker_T = ws.Cells(Ticker_Symbol_Row, 9).Value
        End If
    End If
    ws.Range("P" & 4).Value = Ticker_T
    ws.Range("Q" & 4).Value = Greatest_Total
'
    ws.Cells(Ticker_Symbol_Row, 11).NumberFormat = "0.00%"
    ws.Cells(Ticker_Symbol_Row, 11).Value = Percent_Change
    
    ws.Cells(Ticker_Symbol_Row, 12) = Total_Vol
    
'Reset Total Volume to 0 for next Ticker
    Total_Vol = 0

    Ticker_Symbol_Row = Ticker_Symbol_Row + 1
    
    ws.Cells(Ticker_Symbol_Row, 9) = ws.Cells(i + 1, column).Value
      
    Year_Open = ws.Cells(i + 1, 3).Value
      
    End If
 
  Next i

Next ws

End Sub





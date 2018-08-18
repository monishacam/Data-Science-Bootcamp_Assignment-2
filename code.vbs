Sub Wall_Street()

For Each ws In Worksheets

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Vol"
    
  Dim Ticker As String
  Dim Stock_Vol As Double
  Stock_Vol = 0
  Dim Yearly_Change As Double
  Yearly_Change = 0
  Dim open_price As Double
  Dim close_price As Double
  Dim Percent_Change As Double
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  open_price = ws.Cells(2, 3).Value

  For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      Ticker = ws.Cells(i, 1).Value
      Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
      close_price = ws.Cells(i, 6).Value
      Yearly_Change = close_price - open_price
      If (open_price = 0 And close_price = 0) Then
        Percent_Change = 0
        ElseIf (open_price = 0 And close_price <> 0) Then
            Percent_Change = 1
        Else
            Percent_Change = Yearly_Change / open_price
            ws.Range("L" & Summary_Table_Row).Value = Percent_Change
            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
        End If
      ws.Range("J" & Summary_Table_Row).Value = Ticker
      ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
      ws.Range("M" & Summary_Table_Row).Value = Stock_Vol
      Summary_Table_Row = Summary_Table_Row + 1
      Stock_Vol = 0
      open_price = ws.Cells(i + 1, 3).Value
    Else
      Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
    End If

  Next i
  
  For j = 2 To lastrow
    If ws.Cells(j, 11).Value < 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 3
    ElseIf ws.Cells(j, 11).Value > 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 4
    End If
    
  Next j
  
  Dim max_val As Double
  max_val = ws.Cells(2, 12).Value
  Dim min_val As Double
  min_val = ws.Cells(2, 12).Value
  Dim max_total As Double
  max_total = ws.Cells(2, 12).Value
  Dim Ticker2 As String
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  
  
  For k = 3 To lastrow
    If (ws.Cells(k, 12).Value > max_val) Then
        max_val = ws.Cells(k, 12).Value
        Ticker2 = ws.Cells(k, 10).Value
    End If
  Next k
  ws.Cells(2, 16).Value = Ticker2
  ws.Cells(2, 17).Value = max_val
  ws.Cells(2, 17).NumberFormat = "0.00%"
  
  For l = 3 To lastrow
    If ws.Cells(l, 12).Value < min_val Then
        min_val = ws.Cells(l, 12).Value
        Ticker2 = ws.Cells(l, 10).Value
    End If
  Next l
  ws.Cells(3, 16).Value = Ticker2
  ws.Cells(3, 17).Value = min_val
  ws.Cells(3, 17).NumberFormat = "0.00%"

  
  For m = 3 To lastrow
    If ws.Cells(m, 13).Value > max_total Then
        max_total = ws.Cells(m, 13).Value
        Ticker2 = ws.Cells(m, 10).Value
    End If
  Next m
  ws.Cells(4, 16).Value = Ticker2
  ws.Cells(4, 17).Value = max_total
  
  
Next ws

End Sub





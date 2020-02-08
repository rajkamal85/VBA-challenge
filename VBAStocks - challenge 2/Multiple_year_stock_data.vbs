Sub Assignment()

    Dim startrow As Double
    Dim Ticker_Last As Double
    Dim Ticker_First As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Variant
    Dim Total_Stock_Vol As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        startrow = 2
        Ticker_First = ws.Cells(2, 3).Value
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "percent Change"
        ws.Range("L1") = "Total Stock Vol"
        
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                Ticker_Last = ws.Cells(i, 6).Value
                Yearly_Change = Ticker_Last - Ticker_First
                
                If Ticker_First > 0 Then
                    Percent_Change = (Ticker_Last / Ticker_First) - 1
                Else
                    Percent_Change = Null
                End If

                ws.Range("I" & startrow) = ws.Cells(i, 1).Value
                ws.Range("J" & startrow) = Yearly_Change
                ws.Range("J" & startrow).NumberFormat = "0.00"
                ws.Range("K" & startrow) = Percent_Change
                ws.Range("L" & startrow) = Total_Stock_Vol
                
                startrow = startrow + 1
                Ticker_First = 0
                Ticker_Last = 0
                Yearly_Change = 0
                Percent_Change = 0
                Total_Stock_Vol = 0
                Ticker_First = ws.Cells(i + 1, 3).Value
            
            Else
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
            End If
        
        Next i
        
        lastrow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row
        ws.Range("K2:K" & lastrow1).NumberFormat = "0.00%"
        
        ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow1))
        ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow1))
        ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow1))
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(2, 15).Value = "Gretest % Increase"
        ws.Cells(3, 15).Value = "Gretest % Decrease"
        ws.Cells(4, 15).Value = "Gretest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
                
        For j = 2 To lastrow1
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            If ws.Cells(j, 11) = ws.Cells(2, 17) Then
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            ElseIf ws.Cells(j, 11) = ws.Cells(3, 17) Then
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            ElseIf ws.Cells(j, 12) = ws.Cells(4, 17) Then
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            End If
        Next j

        lastrow1 = 0
        
    Next ws

End Sub
Sub Assignment()

    Dim startrow As Double
    Dim Ticker_Last As Double
    Dim Ticker_First As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Variant
    Dim Total_Stock_Vol As Double
'    Dim ws As Worksheet
    
 '   For Each ws In Worksheets
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        startrow = 2
        Ticker_First = Cells(2, 3).Value
        
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "percent Change"
        Range("L1") = "Total Stock Vol"
        
        
        For i = 2 To lastrow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Total_Stock_Vol = Total_Stock_Vol + Cells(i, 7).Value
                Ticker_Last = Cells(i, 6).Value
                Yearly_Change = Ticker_Last - Ticker_First
                
                If Ticker_First > 0 Then
                    Percent_Change = (Ticker_Last / Ticker_First) - 1
                Else
                    Percent_Change = Null
                End If

                Range("I" & startrow) = Cells(i, 1).Value
                Range("J" & startrow) = Yearly_Change
                Range("J" & startrow).NumberFormat = "0.00"
                Range("K" & startrow) = Percent_Change
                Range("L" & startrow) = Total_Stock_Vol
                
                startrow = startrow + 1
                Ticker_First = 0
                Ticker_Last = 0
                Yearly_Change = 0
                Percent_Change = 0
                Total_Stock_Vol = 0
                Ticker_First = Cells(i + 1, 3).Value
            
            Else
                Total_Stock_Vol = Total_Stock_Vol + Cells(i, 7).Value
            End If
        
        Next i
        
        lastrow1 = Cells(Rows.Count, 10).End(xlUp).Row
        Range("K2:K" & lastrow1).NumberFormat = "0.00%"
        
        Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow1))
        Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow1))
        Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow1))
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(2, 15).Value = "Gretest % Increase"
        Cells(3, 15).Value = "Gretest % Decrease"
        Cells(4, 15).Value = "Gretest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
                
        For j = 2 To lastrow1
            If Cells(j, 10).Value >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            If Cells(j, 11) = Cells(2, 17) Then
                Cells(2, 16).Value = Cells(j, 9).Value
            ElseIf Cells(j, 11) = Cells(3, 17) Then
                Cells(3, 16).Value = Cells(j, 9).Value
            ElseIf Cells(j, 12) = Cells(4, 17) Then
                Cells(4, 16).Value = Cells(j, 9).Value
            End If
        Next j

        lastrow1 = 0
        
'    Next ws

End Sub
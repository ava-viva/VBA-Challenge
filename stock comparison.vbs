Sub Stockcomparison()
    
    For Each ws In Worksheets

    
        Dim StockName As String
        Dim StockTotal As LongLong
        Dim StockStart As Double
        Dim StockEnd As Double
        Dim YearlyChange As Double
        Dim lastrowYearlyChange As Double
        Dim Table_Row As Integer
        Dim PercentChange As Double
    
        
        StockTotal = 0
        StockEnd = 0
        YearlyChange = 0
        Table_Row = 2
        PercentChange = 0
        
        'defining the titles for the table
        
        ws.Cells(1, 16).Value = "ticker"
        ws.Cells(1, 17).Value = "value"
        ws.Cells(2, 15).Value = "greatest % increase"
        ws.Cells(3, 15).Value = "greatest % decrease"
        ws.Cells(4, 15).Value = "greatest volume total"
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        StockStart = ws.Cells(2, 3).Value
        
        'summarizing all data for each stock for the year
        For i = 2 To LastRow
        
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                StockName = ws.Cells(i, 1).Value
                StockEnd = ws.Cells(i, 6).Value
                StockTotal = StockTotal + ws.Cells(i, 7).Value
                YearlyChange = StockEnd - StockStart
                If (StockStart = 0) Then
                    PercentChange = 0
                Else
                
                    PercentChange = YearlyChange / StockStart
                End If
                
                ws.Range("I" & Table_Row).Value = StockName
                
                ws.Range("J" & Table_Row).Value = YearlyChange
                
                If (YearlyChange > 0) Then
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                End If
                
                ws.Range("K" & Table_Row).Value = PercentChange
                ws.Range("K" & Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Table_Row).Value = StockTotal
                Table_Row = Table_Row + 1
                StockStart = ws.Cells(i + 1, 3).Value
                StockEnd = 0
                StockTotal = 0
            
            Else
            
                    StockTotal = StockTotal + ws.Cells(i, 7).Value
            End If
            
    
        Next i
        
        'adding more information about the comparison between all stocks of the year:
        
        lastrowYearlyChange = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For j = 2 To lastrowYearlyChange
                        
                       If ws.Cells(j, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrowYearlyChange)) Then
                                ws.Cells(2, 17).Value = ws.Cells(j, 11).Value
                                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                                ws.Cells(2, 17).NumberFormat = "0.00%"
                        ElseIf ws.Cells(j, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrowYearlyChange)) Then
                                ws.Cells(3, 17).Value = Cells(j, 11).Value
                                ws.Cells(3, 16).Value = Cells(j, 9).Value
                                ws.Cells(3, 17).NumberFormat = "0.00%"
                        End If
                        
                        
                        If ws.Cells(j, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrowYearlyChange)) Then
                                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
                                ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
                        End If
        Next j
          
    Next ws

End Sub




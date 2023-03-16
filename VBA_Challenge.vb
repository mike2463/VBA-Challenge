Sub CalculateStockData()

    'Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolumeTicker As String
    Dim greatestTotalVolume As Double
    Dim i As Long
    
    'Loop through all worksheets
    For Each ws In Worksheets
        
        'Initialize variables
        ticker = ""
        openingPrice = 0
        closingPrice = 0
        yearlyChange = 0
        percentChange = 0
        totalVolume = 0
        greatestPercentIncreaseTicker = ""
        greatestPercentIncrease = 0
        greatestPercentDecreaseTicker = ""
        greatestPercentDecrease = 0
        greatestTotalVolumeTicker = ""
        greatestTotalVolume = 0
        
        'Find last row of data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all rows of data
        For i = 2 To lastRow
            
            'Check if new ticker symbol
            If ws.Cells(i, 1).Value <> ticker Then
                
                'Calculate yearly change and percent change if not first ticker
                If ticker <> "" Then
                    yearlyChange = closingPrice - openingPrice
                    percentChange = yearlyChange / openingPrice
                    
                    'Update greatest percent increase/decrease and greatest total volume
                    If percentChange > greatestPercentIncrease Then
                        greatestPercentIncrease = percentChange
                        greatestPercentIncreaseTicker = ticker
                    ElseIf percentChange < greatestPercentDecrease Then
                        greatestPercentDecrease = percentChange
                        greatestPercentDecreaseTicker = ticker
                    End If
                    
                   
                    If totalVolume > greatestTotalVolume Then
                        greatestTotalVolume = totalVolume
                        greatestTotalVolumeTicker = ticker
                    End If
                    
                    If yearlyChange > 0 Then
                    ws.Cells(i, 10).Interior.Color = vbGreen
                ElseIf yearlyChange < 0 Then
                    ws.Cells(i, 10).Interior.Color = vbRed
                End If
                
                    'Output results for previous ticker
                    ws.Cells(1, 9).Value = "Ticker"
                    ws.Cells(1, 10).Value = "Yearly Change"
                    ws.Cells(1, 11).Value = "Percent Change"
                    ws.Cells(1, 12).Value = "Total Stock Volume"
                    ws.Cells(ws.Cells(Rows.Count, 9).End(xlUp).Row + 1, 9).Value = ticker
                    ws.Cells(ws.Cells(Rows.Count, 10).End(xlUp).Row + 1, 10).Value = yearlyChange
                    ws.Cells(ws.Cells(Rows.Count, 11).End(xlUp).Row + 1, 11).Value = percentChange
                    ws.Cells(ws.Cells(Rows.Count, 12).End(xlUp).Row + 1, 12).Value = totalVolume
                    ws.Cells(1, 14).Value = "Greatest % Increase Ticker"
                    ws.Cells(1, 15).Value = "Greatest % Increase"
                    ws.Cells(1, 16).Value = "Greatest % Decrease Ticker"
                    ws.Cells(1, 17).Value = "Greatest % Decrease"
                    ws.Cells(1, 18).Value = "Greatest Total Volume Ticker"
                    ws.Cells(1, 19).Value = "Greatest Total Volume"
                    
                    ws.Cells(2, 14).Value = greatestPercentIncreaseTicker
                    ws.Cells(2, 15).Value = greatestPercentIncrease
                    ws.Cells(2, 16).Value = greatestPercentDecreaseTicker
                    ws.Cells(2, 17).Value = greatestPercentDecrease
                    ws.Cells(2, 18).Value = greatestTotalVolumeTicker
                    ws.Cells(2, 19).Value = greatestTotalVolume
                End If
                
                
                
                    
                
                'Set variables for new ticker
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                totalVolume = ws.Cells(i, 7).Value
                
            Else
                
                'Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
            End If
            
            'Set closing price for current row
            closingPrice = ws.Cells(i, 6).Value
            
        Next i
        
        'Calculate and output
Next

End Sub



Sub stockmarketdata()

'Looping through all sheets organized by year
For Each ws In Worksheets

    'Creating all variables
    Dim ticker, greattickerinc, greattickerdec, greattickertv As String
    Dim openingprice, closingprice, yearlychange, percentchange, maxpercent, minpercent As Double
    Dim totalstockvolume, greattotalvolume, lastrow, lastrow2, i, j, k As Long

    'Caluculating lastrow and lastrow2
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    lastrow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row

    'Setting all initial values
    maxpercent = 0
    minpercent = 0
    greattotalvolume = 0
    openingprice = 0
    closingprice = 0
    totalstockvolume = 0
    j = 2
    k = 0

    'Creating all fixed headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    '-----------------------------------------------------------------------------------------------
    'Looping through all the stocks for one year and returning for each Ticker the following values:
    'Yearly change
    'Percentage of change
    'Total stock volume
    '-----------------------------------------------------------------------------------------------
    For i = 2 To lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Calculting values of Yearly Change, Percent Change and Total Stock Volume
            ticker = ws.Cells(i, 1).Value
            openingprice = ws.Cells(i - k, 3).Value
            closingprice = ws.Cells(i, 6).Value
            yearlychange = closingprice - openingprice
            percentchange = yearlychange / openingprice
            totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            
            'Assigning the values calculated to the appropriate cell
            ws.Cells(j, 9).Value = ticker
            ws.Cells(j, 10).Value = yearlychange
            ws.Cells(j, 11).Value = FormatPercent(percentchange)
            ws.Cells(j, 12).Value = totalstockvolume
            
            'Conditional formatting: Red for all negative values and Green for all positive values
            If yearlychange < 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 3
                ws.Cells(j, 11).Interior.ColorIndex = 3
                
            Else:
            
                ws.Cells(j, 10).Interior.ColorIndex = 4
                ws.Cells(j, 11).Interior.ColorIndex = 4
                
            End If
            
            'Resetting values of totalstockvolume and counter k to 0
            totalstockvolume = 0
            k = 0
            
            'Incrementing the counter j to create a row for each ticker
            j = j + 1
        
        Else:
        
            'Calculating the sum of all stock volumes for each ticker
            totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            
            'incrementing the counter k to keep the very first opening price for each ticker stocked in opening price
            k = k + 1
        
        End If

    Next i

    
    '-------------------------------------------------------------------------
    'Looping through the new created table and returning the following values:
    'The stock with the greatest % increase
    'The stock with the greatest % decrease
    'The stock with the greatest total volume
    '-------------------------------------------------------------------------
    For i = 2 To lastrow2

        If ws.Cells(i, 11).Value >= maxpercent Then

        maxpercent = ws.Cells(i, 11).Value
        greattickerinc = ws.Cells(i, 9)
        ws.Range("P2") = greattickerinc
        ws.Range("Q2") = FormatPercent(maxpercent)
        
        End If
        
        If ws.Cells(i, 11).Value < minpercent Then
        
        minpercent = ws.Cells(i, 11).Value
        greattickerdec = ws.Cells(i, 9)
        ws.Range("P3") = greattickerdec
        ws.Range("Q3") = FormatPercent(minpercent)
        
        End If
        
        If ws.Cells(i, 12).Value >= greattotalvolume Then
        
        greattotalvolume = ws.Cells(i, 12).Value
        greattickertv = ws.Cells(i, 9)
        ws.Range("P4") = greattickertv
        ws.Range("Q4") = greattotalvolume
        
        End If

    Next i

'Autofitting all columns to displaythe entire data
ws.Columns("A:Q").AutoFit

Next ws

End Sub
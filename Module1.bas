Attribute VB_Name = "Module1"
Sub DataProcessing()
    tickerCol = 9
    yearlyChangeCol = 10
    presentCharCol = 11
    totalStockVolCol = 12
    tickerIterator = 2
    
    Cells(1, tickerCol) = "Ticker"
    Cells(1, tickerCol).Font.Bold = True
    
    Cells(1, yearlyChangeCol) = "Yearly Change"
    Cells(1, yearlyChangeCol).Font.Bold = True
    
    Cells(1, presentCharCol) = "Percent Change"
    Cells(1, presentCharCol).Font.Bold = True
    
    Cells(1, totalStockVolCol) = "Total Stock Volume"
    Cells(1, totalStockVolCol).Font.Bold = True
    
    For Each ws In Sheets
        Dim NumRows As Long
        NumRows = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count
        firstPrice = 0#
        lastPrice = 0#
        Dim volumeCount As Long
        tickerName = ""
        
        For iterator = 2 To NumRows
            currentDate = ws.Cells(iterator, 2)
            processedDate = Mid(currentDate, 5, 4)

            volumeCount = ws.Cells(iterator, 7).Value
            
            If processedDate = "0101" Or tickerName <> ws.Cells(iterator, 1).Value Then
                firstPrice = ws.Cells(iterator, 3).Value
                tickerName = ws.Cells(iterator, 1).Value
            ElseIf processedDate = "1230" Then
                lastPrice = ws.Cells(iterator, 6).Value
                
                Cells(tickerIterator, tickerCol) = tickerName
                Cells(tickerIterator, yearlyChangeCol) = lastPrice - firstPrice
                
                If (lastPrice - firstPrice) >= 0 Then
                    Cells(tickerIterator, yearlyChangeCol).Interior.ColorIndex = 4
                Else
                    Cells(tickerIterator, yearlyChangeCol).Interior.ColorIndex = 3
                End If
                
                If lastPrice = 0 And firstPrice = 0 Then
                    Cells(tickerIterator, presentCharCol) = 0
                ElseIf firstPrice = 0 Then
                    Cells(tickerIterator, presentCharCol) = lastPrice
                Else
                    Cells(tickerIterator, presentCharCol) = ((lastPrice - firstPrice) / firstPrice) * 100
                End If
                
                Cells(tickerIterator, totalStockVolCol) = volumeCount
                Cells(tickerIterator, tickerCol) = tickerName
                Cells(tickerIterator, presentCharCol).NumberFormat = "0.00%"
                
                firstPrice = 0#
                lastPrice = 0#
                volumeCount = 0
                tickerName = ""
                
                tickerIterator = tickerIterator + 1
            End If
        Next
    Next
End Sub

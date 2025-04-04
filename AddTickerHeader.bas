Attribute VB_Name = "AddTickerHeader"
Sub AddTickerHeader()
    ' This macro adds "Ticker", "Quarterly Change", "Percent Change", and "Total Stock Volume" headers
    ' two columns to the right of the last used column
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim tickerCol As Long
    
    ' Open the workbook
    Set wb = Workbooks.Open("/Users/kp/Git/Mod_2_VBA/Multiple_year_stock_data.xlsx")
    
    ' Loop through each worksheet in therkbook
    For Each ws In wb.Worksheets
        ' Find the last used column in row 1 (header row)
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Set the ticker column to be two columns to the right of the last column
        tickerCol = lastCol + 2
        
        ' Add the headers
        ws.Cells(1, tickerCol).Value = "Ticker"
        ws.Cells(1, tickerCol + 1).Value = "Quarterly Change"
        ws.Cells(1, tickerCol + 2).Value = "Percent Change"
        ws.Cells(1, tickerCol + 3).Value = "Total Stock Volume"
        
        ' Format the headers to match other headers (optional)
        ws.Range(ws.Cells(1, tickerCol), ws.Cells(1, tickerCol + 3)).Font.Bold = True
    Next ws
    
    ' Save and close the workbook
    wb.Save
    wb.Close True
    
    MsgBox "Headers added successfully to all worksheets.", vbInformation
End Sub

Sub ProcessStockData()
    ' This macro processes stock data and calculates quarterly metrics
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentTicker As String
    Dim openPrice As Double, closePrice As Double
    Dim quarterlyChange As Double, percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim tickerCol As Long
    
    ' Open the workbook
    Set wb = Workbooks.Open("/Users/kp/Git/Mod_2_VBA/Multiple_year_stock_data.xlsx")
    
    ' Loop through each worksheet
    For Each ws In wb.Worksheets
        ' Find the last row with data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Find the Ticker header column
        tickerCol = 1
        For i = 1 To 100  ' Reasonably large number to search across columns
            If ws.Cells(1, i).Value = "Ticker" Then
                tickerCol = i
                Exit For
            End If
        Next i
        
        ' If ticker column not found, show error and skip this worksheet
        If tickerCol = 1 Then
            MsgBox "Ticker header not found in worksheet: " & ws.Name & ". Run AddTickerHeader first.", vbExclamation
            GoTo NextWorksheet
        End If
        
        ' Initialize output row
        outputRow = 2
        
        ' Assuming data structure: A=Ticker, B=Date, C=Open, F=Close, G=Volume
        ' (Adjust these column indices if your data structure is different)
        
        ' Initialize tracking variables
        currentTicker = ""
        totalVolume = 0
        
        ' Loop through all rows with data
        For i = 2 To lastRow
            ' Check if we're starting a new ticker
            If ws.Cells(i, 1).Value <> currentTicker Then
                ' If not the first ticker, output the previous ticker's data
                If currentTicker <> "" Then
                    ' Write data to output columns
                    ws.Cells(outputRow, tickerCol).Value = currentTicker
                    ws.Cells(outputRow, tickerCol + 1).Value = quarterlyChange
                    ws.Cells(outputRow, tickerCol + 2).Value = percentChange
                    ws.Cells(outputRow, tickerCol + 3).Value = totalVolume
                    
                    ' Format percent change as percentage
                    ws.Cells(outputRow, tickerCol + 2).NumberFormat = "0.00%"
                    
                    ' Format quarterly change with colors (green for positive, red for negative)
                    If quarterlyChange > 0 Then
                        ws.Cells(outputRow, tickerCol + 1).Interior.Color = RGB(0, 255, 0)
                    ElseIf quarterlyChange < 0 Then
                        ws.Cells(outputRow, tickerCol + 1).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    ' Increment output row
                    outputRow = outputRow + 1
                End If
                
                ' Start tracking new ticker
                currentTicker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value  ' Opening price at start of period
                totalVolume = 0
            End If
            
            ' Add to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' If this is the last row for this ticker, capture closing price
            If i = lastRow Or ws.Cells(i + 1, 1).Value <> currentTicker Then
                closePrice = ws.Cells(i, 6).Value  ' Closing price at end of period
                
                ' Calculate changes
                quarterlyChange = closePrice - openPrice
                
                ' Prevent division by zero
                If openPrice <> 0 Then
                    percentChange = quarterlyChange / openPrice
                Else
                    percentChange = 0
                End If
            End If
        Next i
        
        ' Output the last ticker's data
        If currentTicker <> "" Then
            ws.Cells(outputRow, tickerCol).Value = currentTicker
            ws.Cells(outputRow, tickerCol + 1).Value = quarterlyChange
            ws.Cells(outputRow, tickerCol + 2).Value = percentChange
            ws.Cells(outputRow, tickerCol + 3).Value = totalVolume
            
            ' Format percent change as percentage
            ws.Cells(outputRow, tickerCol + 2).NumberFormat = "0.00%"
            
            ' Format quarterly change with colors
            If quarterlyChange > 0 Then
                ws.Cells(outputRow, tickerCol + 1).Interior.Color = RGB(0, 255, 0)
            ElseIf quarterlyChange < 0 Then
                ws.Cells(outputRow, tickerCol + 1).Interior.Color = RGB(255, 0, 0)
            End If
        End If
        
        ' Auto-fit columns for better readability
        ws.Columns(tickerCol).AutoFit
        ws.Columns(tickerCol + 1).AutoFit
        ws.Columns(tickerCol + 2).AutoFit
        ws.Columns(tickerCol + 3).AutoFit
        
NextWorksheet:
    Next ws
    
    ' Save and close the workbook
    wb.Save
    wb.Close True
    
    MsgBox "Stock data processed successfully.", vbInformation
End Sub

Sub CalculateExtremes()
    ' This macro calculates and identifies tickers with greatest % increase, decrease, and volume
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tickerCol As Long, percentCol As Long, volumeCol As Long
    Dim lastRow As Long, i As Long
    Dim maxPercentIncrease As Double, maxPercentDecrease As Double, maxVolume As Double
    Dim maxIncTicker As String, maxDecTicker As String, maxVolTicker As String
    
    ' Open the workbook
    Set wb = Workbooks.Open("/Users/kp/Git/Mod_2_VBA/Multiple_year_stock_data.xlsx")
    
    ' Loop through each worksheet
    For Each ws In wb.Worksheets
        ' Find the Ticker header column
        tickerCol = 0
        For i = 1 To 100  ' Reasonably large number to search across columns
            If ws.Cells(1, i).Value = "Ticker" Then
                tickerCol = i
                Exit For
            End If
        Next i
        
        ' If ticker column not found, show error and skip this worksheet
        If tickerCol = 0 Then
            MsgBox "Ticker header not found in worksheet: " & ws.Name & ". Run AddTickerHeader and ProcessStockData first.", vbExclamation
            GoTo NextWorksheet
        End If
        
        ' The other columns relative to Ticker column
        percentCol = tickerCol + 2  ' "Percent Change" column
        volumeCol = tickerCol + 3   ' "Total Stock Volume" column
        
        ' Find the last row with data in the Ticker column
        lastRow = ws.Cells(ws.Rows.Count, tickerCol).End(xlUp).Row
        
        ' Initialize variables to track extreme values
        maxPercentIncrease = -9999999  ' Very low starting value
        maxPercentDecrease = 9999999   ' Very high starting value
        maxVolume = 0
        maxIncTicker = ""
        maxDecTicker = ""
        maxVolTicker = ""
        
        ' Loop through all rows with data
        For i = 2 To lastRow
            ' Check for greatest % increase
            If ws.Cells(i, percentCol).Value > maxPercentIncrease Then
                maxPercentIncrease = ws.Cells(i, percentCol).Value
                maxIncTicker = ws.Cells(i, tickerCol).Value
            End If
            
            ' Check for greatest % decrease
            If ws.Cells(i, percentCol).Value < maxPercentDecrease Then
                maxPercentDecrease = ws.Cells(i, percentCol).Value
                maxDecTicker = ws.Cells(i, tickerCol).Value
            End If
            
            ' Check for greatest total volume
            If ws.Cells(i, volumeCol).Value > maxVolume Then
                maxVolume = ws.Cells(i, volumeCol).Value
                maxVolTicker = ws.Cells(i, volumeCol - 3).Value
            End If
        Next i
        
        ' Set up the results table headers
        ws.Cells(1, 15).Value = "Category"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Format the headers
        ws.Range(ws.Cells(1, 15), ws.Cells(1, 17)).Font.Bold = True
        
        ' Set up category names
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Output results
        ws.Cells(2, 16).Value = maxIncTicker
        ws.Cells(3, 16).Value = maxDecTicker
        ws.Cells(4, 16).Value = maxVolTicker
        
        ' Output values with appropriate formatting
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 17).Value = maxPercentDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 17).Value = maxVolume
        ws.Cells(4, 17).NumberFormat = "#,##0"
        
        ' Auto-fit columns for better readability
        ws.Columns(15).AutoFit
        ws.Columns(16).AutoFit
        ws.Columns(17).AutoFit
        
NextWorksheet:
    Next ws
    
    ' Save and close the workbook
    wb.Save
    wb.Close True
    
    MsgBox "Extreme values calculated successfully.", vbInformation
End Sub 
Attribute VB_Name = "Module1"
Sub stock_analysis()
'identify your variables- there is a lot - so i'm giving those to you
    
    Dim tickerList As Object
Set tickerList = CreateObject("Scripting.Dictionary")

    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        j = 0
        total = 0
        change = 0
        start = 2
        
        'rest of the row titles
        'find the row number of the last row with data
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        

        'Loop through each row of data
        For i = 2 To rowCount
            
            Ticker = ws.Cells(i, 1).Value
            
         If Not tickerList.Exists(Ticker) Then
    outputRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row + 1
    tickerList.Add Ticker, outputRow
    startPrice = ws.Cells(i, 3).Value
    totalVolume = ws.Cells(i, 7).Value
                
                start = start + 1 ' increment next empty row for this ticker
            
            'Calculate yearly change and percent change if it's the last row for the current ticker
            ElseIf i = rowCount Or ws.Cells(i + 1, 1).Value <> Ticker Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                endPrice = ws.Cells(i, 6).Value
                yearlyChange = endPrice - startPrice
                percentChange = yearlyChange / startPrice
                
                 ws.Range("I" & tickerList(Ticker)).Value = Ticker
                ws.Range("J" & tickerList(Ticker)).Value = yearlyChange
                ws.Range("K" & tickerList(Ticker)).Value = percentChange
                ws.Range("K" & tickerList(Ticker)).Value = percentChange
ws.Range("K" & tickerList(Ticker)).NumberFormat = "0.0%"
                ws.Range("L" & tickerList(Ticker)).Value = totalVolume
                
                If yearlyChange >= 0 Then
                    ws.Range("J" & tickerList(Ticker)).Interior.ColorIndex = 4 'Green
                Else
                    ws.Range("J" & tickerList(Ticker)).Interior.ColorIndex = 3 'Red
                End If
                
                start = tickerList(Ticker) + 1 ' update next empty row for this ticker
                
                'Reset variables for next ticker
                tickerList.Remove Ticker
                startPrice = 0
                totalVolume = 0
                yearlyChange = 0
                percentChange = 0
            
            'Add volume to total volume for current ticker
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        
       Next i
    
       ' take the max and min and place them in a separate part in the worksheet
        'examples of max function. you need a Min too, which works similarly
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

                ' returns one less because header row not a factor
        'Another function - Match
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
'
ws.Range("O1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
     
' final ticker symbol for  total, greatest % of increase and decrease, and average
ws.Range("P2").Value = ws.Range("I" & increase_number + 1).Value
ws.Range("P3").Value = ws.Range("I" & decrease_number + 1).Value
ws.Range("P4").Value = ws.Range("I" & volume_number + 1).Value
    
    Next ws
End Sub



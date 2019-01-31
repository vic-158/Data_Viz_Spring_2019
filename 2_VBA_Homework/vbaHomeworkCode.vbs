Sub stockCalc()
'Declare Variables
'i and j as indexing variables
Dim i As Long
Dim j As Long
Dim totalRows As Long
Dim returnCol As Long

'TickerSymbols() as an array of all unique ticker symbols and declar number of unique ticker symbols
Dim TickerSymbols() As String
Dim numUniqueTickers As Long

'Total Volume of Each Stock
Dim totalVol As Double
Dim totalVolArray() As Double

'First Open and Last Close value of Each stock, per year
Dim firstOpen As Double
Dim lastClose As Double

'Greatest Percent Increase, Greatest Percent Decrease, Greatest Total Volume and Row Numbers associated with each
Dim gPcInc As Double
Dim gPcDec As Double
Dim gTotVol As Double
Dim gPcIncRow As Double
Dim gPcDecRow As Double
Dim gTotVolRow As Double

    'Find the last column data set and set return column to be two after that
    returnCol = Cells(1, 1).End(xlToRight).Column + 2
    
    'Find the total number of rows in table
    totalRows = Cells(1, 1).End(xlDown).Row
    
    'Write Return Column Headers
    Cells(1, returnCol).Value = "Ticker Symbol"
    Cells(1, returnCol).Columns.AutoFit
    Cells(1, returnCol + 1).Value = "Yearly Change"
    Cells(1, returnCol + 1).Columns.AutoFit
    Cells(1, returnCol + 2).Value = "Percentage Change"
    Cells(1, returnCol + 2).Columns.AutoFit
    Cells(1, returnCol + 3).Value = "Total Stock Volume"
    Cells(1, returnCol + 3).Columns.AutoFit

    'Find the number of unique ticker symbols in the list
    numUniqueTickers = 0
    For j = 2 To totalRows
        If Cells(j, 1).Value <> Cells(j - 1, 1) Then
            numUniqueTickers = numUniqueTickers + 1
        End If
    Next j

    'Define an array with length equal to the number of unique ticker symbols
    ReDim TickerSymbols(numUniqueTickers - 1)
    ReDim totalVolArray(numUniqueTickers - 1)
    'Populate TickerSymbols array with all unique ticker symbols
    i = 0
    For j = 2 To totalRows
        If Cells(j, 1).Value <> Cells(j - 1, 1) Then
          TickerSymbols(i) = Cells(j, 1).Value
          i = i + 1
        End If
    Next j

    'Populate Return Column with Ticker Symbols
    For i = 0 To numUniqueTickers - 1
        Cells(2 + i, returnCol).Value = TickerSymbols(i)
    Next i

    'For each Unique Ticker Symbol, query the entire list of stock.
    'If the given stock matches the current ticker symbol, then recalc total stock vol and re-write to appropriate cell.
    totalVol = 0
    For i = 0 To numUniqueTickers - 1
        For j = 2 To totalRows
            If Cells(j, 1).Value = TickerSymbols(i) Then
                totalVol = totalVol + Cells(j, 7)
                Cells(i + 2, returnCol + 3).Value = totalVol
            End If
        Next j
       totalVol = 0
    Next i

    'Calculate the Yearly Change by finding the first open value and the last close value for each individual ticker symbol
    For i = 0 To numUniqueTickers - 1
        For j = 2 To totalRows
            If Cells(j, 1).Value = TickerSymbols(i) And (Cells(j, 1).Value <> Cells(j - 1, 1).Value) Then
                firstOpen = Cells(j, 3).Value
            End If
            If Cells(j, 1).Value = TickerSymbols(i) And (Cells(j, 1).Value <> Cells(j + 1, 1).Value) Then
                lastClose = Cells(j, 3).Value
            End If
        Next j
        'Write Yearly Change to Yearly Change Column
        Cells(i + 2, returnCol + 1).Value = lastClose - firstOpen
        If Cells(i + 2, returnCol + 1).Value < 0 Then
            Cells(i + 2, returnCol + 1).Interior.Color = RGB(255, 0, 0)
        Else
            Cells(i + 2, returnCol + 1).Interior.Color = RGB(0, 255, 0)
        End If


        'If firstOpen = 0 then write exception string. Otherwise, write Percentage Change to Percentage Change Column
        If (firstOpen = 0) Then
            Cells(i + 2, returnCol + 2).Value = 0
        Else
            Cells(i + 2, returnCol + 2).Value = (lastClose - firstOpen) / firstOpen
        End If
    Next i

    'Write Headers for Greatest % Increase, Decrease, Total Volume
    Cells(1, returnCol + 7).Value = "Ticker"
    Cells(1, returnCol + 8).Value = "Value"
    Cells(2, returnCol + 6).Value = "Greatest % Increase"
    Cells(3, returnCol + 6).Value = "Greatest % Decrease"
    Cells(4, returnCol + 6).Value = "Greatest Total Vol"
    Cells(4, returnCol + 6).Columns.AutoFit

    gPcInc = 0
    gPcDec = 0
    gTotVol = 0
    
    For i = 2 To numUniqueTickers + 2
        'If the current percent change is greater than the current gPcInc, then replace the value
        If (Cells(i, returnCol + 2).Value > gPcInc) Then
            gPcInc = Cells(i, returnCol + 2).Value
            gPcIncRow = i
        End If
        
        'If the current percent change is less than the current gPcDec, then replace the value
        If (Cells(i, returnCol + 2).Value < gPcDec) Then
            gPcDec = Cells(i, returnCol + 2).Value
            gPcDecRow = i
        End If

        If (Cells(i, returnCol + 3).Value > gTotVol) Then
            gTotVol = Cells(i, returnCol + 3).Value
            gTotVolRow = i
        End If
    Next i
    
    'Write Greatest Percent Increase, Decrease, Total Vol, and Ticker Symbols.
    Cells(2, returnCol + 8).Value = gPcInc
    Cells(3, returnCol + 8).Value = gPcDec
    Cells(4, returnCol + 8).Value = gTotVol
    Cells(2, returnCol + 7).Value = Cells(gPcIncRow, returnCol).Value
    Cells(3, returnCol + 7).Value = Cells(gPcDecRow, returnCol).Value
    Cells(4, returnCol + 7).Value = Cells(gTotVolRow, returnCol).Value
    
    Cells(4, returnCol + 7).Columns.AutoFit
    Cells(4, returnCol + 8).Columns.AutoFit

End Sub



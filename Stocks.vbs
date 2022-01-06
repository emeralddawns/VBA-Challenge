Sub TickerTape():

Dim i As Double ''for counter
Dim j As Double ''for counter

Dim tickerALL As Double ''column containing ticker symbols
Dim opening As Double ''column containing opening price
Dim closing As Double ''column containing closing price
Dim volume As Double ''column containing volume

Dim SUMMline As Double ''first row of summary
Dim tickerSUMM As Double ''column location of ticker summary
Dim YearChgCol As Double ''column containing Yearly Change
Dim PerChgCol As Double ''column containing Percent Change
Dim TotalStockCol As Double ''column containing Total Stock Volume

Dim TickerAn As Double ''column with Ticker Analysis row titles
Dim TickerSym As Double ''column with Ticker Symbol for the called out values
Dim Value As Double ''column with special (max incr, etc) values


Dim yearchange As Double ''yearly change amount
Dim perchange As Double ''percent change amount
Dim tot_ticker As Double ''line the data summary starts on
Dim totalstock As Double ''total stock volume
Dim lastrow As Double ''the last row...
Dim firstline As Double ''1st line for ticker
Dim lastline As Double ''Last line for ticker

Dim MaxInc As Double ''Greatest % Increase
Dim MaxDec As Double ''Greatest % Decrease
Dim MaxTotVol As Double ''Greatest Total Volume

''locations
tickerALL = 1
opening = 3
closing = 6
volume = 7

SUMMline = 1
tickerSUMM = 9
YearChgCol = 10
PerChgCol = 11
TotalStockCol = 12

TickerAn = 15
TickerSym = TickerAn + 1
Value = TickerAn + 2

''initial values
tot_tickers = 1
totalstock = 0
firstline = 2

''equations
lastrow = Cells(Rows.Count, tickerALL).End(xlUp).Row


''create summary table
Cells(SUMMline + 1, tickerSUMM).Value = Cells(2, tickerALL)

For i = 2 To lastrow
    totalstock = totalstock + Cells(i, volume)
    If Cells(i + 1, tickerALL).Value <> Cells(i, tickerALL).Value Then
        lastline = i
        
        yearchange = Cells(lastline, closing).Value - Cells(firstline, opening).Value
        Cells(tot_tickers + 1, YearChgCol) = yearchange
        
        ''if opened at 0...
        If Cells(firstline, opening).Value = 0 Then
            Cells(tot_tickers + 1, PerChgCol) = "NaN"
        Else
            perchange = yearchange / Cells(firstline, opening).Value
            Cells(tot_tickers + 1, PerChgCol) = perchange
        End If
        
          
        Cells(tot_tickers + 1, TotalStockCol) = totalstock
        
        ''prep for next ticker symbol
        totalstock = 0
        firstline = i + 1
        Cells(tot_tickers + 2, tickerSUMM).Value = Cells(i + 1, tickerALL) ''add symbol
        tot_tickers = tot_tickers + 1
    End If
Next i

        ''MsgBox ("there are " & tot_tickers - 1 & " total tickers") ''-1 bc of the blank cell at the end
        
        
''Bonus Summary
MaxTotVol = Application.WorksheetFunction.Max(Columns(TotalStockCol))
MsgBox ("Greatest Total Volume:" & MaxTotVol)


MaxInc = Application.WorksheetFunction.Max(Columns(PerChgCol))
MsgBox ("Max % Chg: " & MaxInc)










''Format Sheet
Cells(SUMMline, tickerSUMM).Value = "Ticker"

Cells(SUMMline, YearChgCol).Value = "Yearly Change"
    ''begin macro i recorded (and edited)
    Cells(SUMMline + 1, YearChgCol).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 3407718
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ''end macro i recorded

Cells(SUMMline, PerChgCol).Value = "Percent Change"
    Columns(PerChgCol).NumberFormat = "0.00%"

Cells(SUMMline, TotalStockCol).Value = "Total Stock Volume"










End Sub


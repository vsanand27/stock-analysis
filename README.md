# stock-analysis
# Challenge 2 - Creating Stock Price Performance
The analysis outlines 2017 and 2018 Stock Performance to help individuals pick stocks with high volume and performance.  The program outlines 12 Stock Stickers that were chosen to track in 2017 and 2018.  The program can easily be modified to including thousand of stock tickers that needed to be tracked given the values are already assigned in the tickers array ahead to time.  The program is designed such that it will go through the worksheet only once and display all values.

# Below is VBA code for the Analysis

# Sub AllStockAnalysis()

    Worksheets("AllStocksAnalysis").Activate

    YearValue = InputBox("Which Year would you like to run the analysis for?")


    Range("A1").Value = "All Stocks (" + YearValue + ")"

## '1) Create a header row: For Challenge Assignement - Action Created 2 more Columns


Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"
Cells(3, 4).Value = "Start Price"
Cells(3, 5).Value = "End Price"


## '2) Initialize array of all tickers
Dim tickers(12) As String
tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

Dim TotalVolume(12) As Long
TotalVolume(0) = 0
TotalVolume(1) = 0
TotalVolume(2) = 0
TotalVolume(3) = 0
TotalVolume(4) = 0
TotalVolume(5) = 0
TotalVolume(6) = 0
TotalVolume(7) = 0
TotalVolume(8) = 0
TotalVolume(9) = 0
TotalVolume(10) = 0
TotalVolume(11) = 0

Dim StartingPrice(12) As Double
StartingPrice(0) = 0
StartingPrice(1) = 0
StartingPrice(2) = 0
StartingPrice(3) = 0
StartingPrice(4) = 0
StartingPrice(5) = 0
StartingPrice(6) = 0
StartingPrice(7) = 0
StartingPrice(8) = 0
StartingPrice(9) = 0
StartingPrice(10) = 0
StartingPrice(11) = 0

Dim EndingPrice(12) As Double
EndingPrice(0) = 0
EndingPrice(1) = 0
EndingPrice(2) = 0
EndingPrice(3) = 0
EndingPrice(4) = 0
EndingPrice(5) = 0
EndingPrice(6) = 0
EndingPrice(7) = 0
EndingPrice(8) = 0
EndingPrice(9) = 0
EndingPrice(10) = 0
EndingPrice(11) = 0

## '3b) Activiate Data Worksheet
Worksheets(YearValue).Activate

## '3c) Get Row Count
Dim RowCount As Long
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
MsgBox (RowCount)

    Dim TickerIndex As Integer
    TickerIndex = 0
    
## '4) only 1 loop no outer loop to go through the loop once; Set up ticker index

    For J = 2 To RowCount

            ### '5a Get Total Volume for each Ticker
            If Cells(J, 1).Value = tickers(TickerIndex) Then
            TotalVolume(TickerIndex) = TotalVolume(TickerIndex) + Cells(J, 8).Value
            End If
            
            
            ### 'calculate starting price for Return calc - Adding Starting Price 0 but doesn't need it
            If Cells(J - 1, 1).Value <> tickers(TickerIndex) And Cells(J, 1).Value = tickers(TickerIndex) Then
            StartingPrice(TickerIndex) = StartingPrice(TickerIndex) + Cells(J, 6).Value
            End If
            
            ### 'calculate end price for Return Calc - Adding Starting Price to 0 but it doesn't need it
            
            If Cells(J + 1, 1).Value <> tickers(TickerIndex) And Cells(J, 1).Value = tickers(TickerIndex) Then
            EndingPrice(TickerIndex) = EndingPrice(TickerIndex) + Cells(J, 6).Value
            End If
            
       ### 'Increment Ticker Index
       
       If Cells(J + 1, 1).Value <> tickers(TickerIndex) And Cells(J, 1).Value = tickers(TickerIndex) Then
       TickerIndex = TickerIndex + 1
       End If

Next J
     

## '6) Output for current ticker - Challenge 2 Output for Starting and End Price
Worksheets("AllStocksAnalysis").Activate

For T = 0 To 11
Cells(4 + T, 1).Value = tickers(T)
Cells(4 + T, 2).Value = TotalVolume(T)
Cells(4 + T, 3).Value = EndingPrice(T) / StartingPrice(T) - 1
Cells(4 + T, 4).Value = StartingPrice(T)
Cells(4 + T, 5).Value = EndingPrice(T)

Next T

' Formatting - Challenge Assignment 2: Formatting New columns Starting and Ending Price

Worksheets("AllStocksAnalysis").Activate
Range("A3:E3").Font.Bold = True
Range("A3:E3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"


datastartRow = 4
DataendRow = 15

## For i = datastartRow To DataendRow

If Cells(i, 3) > 0 Then
    'Format Color Green
    Cells(i, 3).Interior.Color = vbGreen

    ElseIf Cells(i, 3) < 0 Then
    'Format Color Red
    Cells(i, 3).Interior.Color = vbRed

    Else: Cells(i, 3).Interior.Color = xlNone
End If
Next i

## End Sub

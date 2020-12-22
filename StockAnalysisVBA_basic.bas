Attribute VB_Name = "Module11"
Sub StockAnalysis()

'Excel Commands to enable faster processing
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Variable Declarations for processing and summary
Dim ticker As String
Dim volume As Long
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentageChange As Double

Dim firstTicker As Boolean
firstTicker = True

'Create Summary Table headers and format
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Count Number of populated rows in sheet
RowCount = Cells(Rows.Count, 1).End(xlUp).Row


'For loop for all the ticker symbols in the sheet
For i = 2 To RowCount

'Capture the opening price for the ticker
    If firstTicker = True Then
      openingPrice = Cells(i, 3).Value
    End If

' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker symbol
      ticker = Cells(i, 1).Value
      
      'establish the closing price for the ticker
      closingPrice = Cells(i, 6).Value

      ' Add final volume to the total stock volume
      'volume = volume + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker
      
      ' Calculate and Print the yearly change in the Summary Table
      yearlyChange = closingPrice - openingPrice
      Range("J" & Summary_Table_Row).Value = yearlyChange
      
      ' Color code positive and negative changes (green and red)
      If yearlyChange < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      Else
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
      
      ' Calculate and Print the percentage change in the Summary Table
      percentageChange = yearlyChange / openingPrice
      Range("K" & Summary_Table_Row).Value = percentageChange
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

      ' Print the stock volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = volume

      ' Set the next summary table row value
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the opening price, closing price, and total stock volume
      openingPrice = 0
      closingPrice = 0
      'volume = 0
      yearlyChange = 0
      firstTicker = True

    ' If the cell immediately following a row is the same ticker...
    Else
      
      firstTicker = False
      
      ' Add to the volume for the stock
      'volume = volume + Cells(i, 7)

    End If

  Next i

End Sub



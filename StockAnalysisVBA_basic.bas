Attribute VB_Name = "Module11"
Sub StockAnalysis()

'Excel Commands to enable faster processing
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Variable Declarations for processing and summary
Dim ticker As String
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentageChange As Double

'Volume start index to be used in SUM function since total volume overflows Long var
Dim volStart As Long

'New ticker status variable to mark when opening price needs to be captured
Dim firstTicker As Boolean
firstTicker = True

'Create Summary Table headers and format
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Keep track of summary table row index
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Count Number of populated rows in sheet
RowCount = Cells(Rows.Count, 1).End(xlUp).Row


'Start loop
'For loop for all the ticker symbols in the sheet
For i = 2 To RowCount

'Capture the opening price for the ticker
    If firstTicker = True Then
      openingPrice = Cells(i, 3).Value
      volStart = i
    End If

' Check if we are still within the same ticker symbol, if it is not capture summary row data
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the summary ticker symbol
      ticker = Cells(i, 1).Value
      
      'establish the closing price for the ticker
      closingPrice = Cells(i, 6).Value

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
      If openingPrice <> 0 Then
        percentageChange = yearlyChange / openingPrice
        Range("K" & Summary_Table_Row).Value = percentageChange
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      Else
        Range("K" & Summary_Table_Row).Value = "ERR /0"
      End If
      

      ' Print the stock volume to the Summary Table by formula
      Range("L" & Summary_Table_Row).Formula = "=SUM(" & "G" & volStart & ":" & "G" & i & ")"
          
      ' Set/increment the next summary table row value
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the opening price, closing price, and mark that a new ticker is expected
      openingPrice = 0
      closingPrice = 0
      firstTicker = True

    ' If the cell immediately following a row is the same ticker continue down the list
    Else
      
      ' Not a new ticker
      firstTicker = False
      
    End If

  Next i

End Sub



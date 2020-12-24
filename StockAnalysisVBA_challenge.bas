Attribute VB_Name = "Module2"
Sub StockAnalysis_Challenge()

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
'volStart = 2

'New ticker status variable to mark when opening price needs to be captured
Dim firstTicker As Boolean
firstTicker = True
    
'Keep track of summary table row index
Dim Summary_Table_Row As Integer

' Loop through all sheets
For Each ws In Worksheets
        
    'MsgBox (ws.Name)

    'Create Summary Table headers and format
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    
    Summary_Table_Row = 2

    'Count Number of populated rows in sheet
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'Start ticker loop
    'For loop for all the ticker symbols in the current worksheet
    For i = 2 To RowCount

    'Capture the opening price for the ticker
    If firstTicker = True Then
      openingPrice = ws.Cells(i, 3).Value
      volStart = i
    End If

    ' Check if we are still within the same ticker symbol, if it is not capture summary row data
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the summary ticker symbol
      ticker = ws.Cells(i, 1).Value
      
      'establish the closing price for the ticker
      closingPrice = ws.Cells(i, 6).Value

      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker
      
      ' Calculate and Print the yearly change in the Summary Table
      yearlyChange = closingPrice - openingPrice
      ws.Range("J" & Summary_Table_Row).Value = yearlyChange
      ' Color code positive and negative changes (green and red)
      If yearlyChange < 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
      
      ' Calculate and Print the percentage change in the Summary Table
      If openingPrice <> 0 Then
        percentageChange = yearlyChange / openingPrice
        ws.Range("K" & Summary_Table_Row).Value = percentageChange
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      Else
        ws.Range("K" & Summary_Table_Row).Value = "ERR /0"
      End If
      

      ' Print the stock volume to the Summary Table by formula
      ws.Range("L" & Summary_Table_Row).Formula = "=SUM(" & "G" & volStart & ":" & "G" & i & ")"
          
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
  
  'End loop
  
  'Start further summary analysis
  
  'Row counter for summary table
  Dim summaryRowCount As Long
        
       'Instert Calculation Table Headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Row Count to determine min, max change and max volume
        summaryRowCount = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Find biggest increase and decrease changes in the summary along with largest total volume
        ws.Range("Q2").Formula = "=MAX(" & "K2" & ":" & "K" & summaryRowCount & ")"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Formula = "=MIN(" & "K2" & ":" & "K" & summaryRowCount & ")"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Formula = "=MAX(" & "L2" & ":" & "L" & summaryRowCount & ")"
        ws.Range("Q4").NumberFormat = "General"
        
        'For loop through summary table to find tickers for players
        For j = 2 To summaryRowCount
        If ws.Cells(j, 11).Value = ws.Range("Q2").Value Then
            ws.Range("P2").Value = ws.Cells(j, 9).Value
        ElseIf ws.Cells(j, 11).Value = ws.Range("Q3").Value Then
            ws.Range("P3").Value = ws.Cells(j, 9).Value
        ElseIf CStr(ws.Cells(j, 12).Value) = CStr(ws.Range("Q4").Value) Then
            ws.Range("P4").Value = ws.Cells(j, 9).Value
        End If
        Next j
        
        ' Autofit to summary data
        ws.Columns("J:L").AutoFit
        ws.Columns("O:Q").AutoFit
     
'All complete.  Move on to next worksheet
Next ws
     
    
End Sub


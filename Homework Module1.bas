Attribute VB_Name = "Module1"
Sub Stockticker()

Dim ws As Worksheet
Dim WorksheetName As String

For Each ws In Worksheets
    WorksheetName = ws.Name
    MsgBox WorksheetName

'Set up the summary table
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"


'Create summary table variables - values will change by stock ticker via the loop below
Dim SummaryTicker As String
Dim SummaryVolume As Double
  SummaryVolume = 0
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
Dim SummaryYrChange As Double
Dim SummaryPercentChange As Double

'Create variables for calculations performed on specific values pulled from the ticker data
Dim Lastrow As Long
  Lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TickerRowCount As Double
  TickerRowCount = 0

  
  ' Loop through all rows in the sheet to analyze the data
  For i = 2 To Lastrow

    ' Check if we are still within the same stock ticker, if we are not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      SummaryTicker = Cells(i, 1).Value
      
      ' Set the value of Closeprice to the closing price of the last row in the ticker data
      ClosePrice = Cells(i, 6).Value

      ' Add to the Summary volume counter
      SummaryVolume = SummaryVolume + Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = SummaryTicker
      
      'Calculate the value of the year's stock price change & add to the summary table
      SummaryYrChange = ClosePrice - OpenPrice
      Range("J" & Summary_Table_Row).Value = SummaryYrChange
        If SummaryYrChange > 0 Then Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        If SummaryYrChange < 0 Then Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      
      'Calculate the year's % change & add to the suummary table
      SummaryPercentChange = (ClosePrice - OpenPrice) / OpenPrice
      Range("K" & Summary_Table_Row).Value = SummaryPercentChange
      
      ' Print the Volume total to the Summary Table
      Range("L" & Summary_Table_Row).Value = SummaryVolume
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      SummaryVolume = 0

    ' If the cell immediately following a row is the same ticker...
        Else
    
          ' Add to the Volume total counter
          SummaryVolume = SummaryVolume + Cells(i, 7).Value
          
          ' Find the value of the opening price
          If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
          OpenPrice = Cells(i, 3).Value
          
          End If
      
        End If
    
    Next i

   'Set up the summary table2
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Volume"
    Columns("I:Q").AutoFit
    Range("K:K").NumberFormat = "0.00%"

    'Determine the last row of the summary table
    Dim LastTableRow As Long
    LastTableRow = ActiveSheet.Range("I" & Rows.Count).End(xlUp).Row
    smallest = Range("J" & 2).Value
    largest = Range("J" & 2).Value
    largestvol = Range("L" & 2).Value
    
    'Find the min value of column K - % Change & related ticker
    
    For J = 2 To LastTableRow
      If Cells(J, 11).Value < smallest Then
        smallest = Cells(J, 11).Value
        Rowposition = J
      End If
    Next J
      
    Range("Q" & 3).Value = smallest
    Range("P" & 3).Value = Range("I" & Rowposition)
 
    'Find the max value of column K - % Change - & related ticker
    
    For k = 2 To LastTableRow
      If Cells(k, 11).Value > largest Then
        largest = Cells(k, 11).Value
        Rowloc = k
      End If
    Next k
    
    Range("Q" & 2).Value = largest
    Range("P" & 2).Value = Range("I" & Rowloc)
    
    Range("Q2", "Q3").NumberFormat = "0.00%"
    
     'Find the max value of column L - Volume - & related ticker
    
    For l = 2 To LastTableRow
      If Cells(l, 12).Value > largestvol Then
        largestvol = Cells(l, 12).Value
        Rowloc = l
      End If
    Next l
    
    Range("Q" & 4).Value = largestvol
    Range("P" & 4).Value = Range("I" & Rowloc)
    
    
Next ws

End Sub


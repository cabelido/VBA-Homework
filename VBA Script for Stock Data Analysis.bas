Attribute VB_Name = "Module1"
Sub Stockdatanalytics()
For Each ws In Worksheets

'Initialize variables
Dim Ticker As String
Dim OpenPrYear As Double
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Total_Volume As LongLong
Dim SummaryTick_Row As Integer
Dim yearlastRow As Long
Dim groupfirstrow As LongLong

'Place column headers in the Summary Table
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percentage Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"


'Set Initial Value for Yearly_Change, Total_Volume,Summary Tick_Row,startrow, and equation for total number of rows for each ws
Yearly_Change = 0
Total_Volume = 0
SummaryTick_Row = 2
yearlastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
startrow = 2

'Loop thru all data in each worksheet to arrive at analytical data required

For i = 2 To yearlastRow
    

    If ws.Cells((i + 1), 1).Value <> ws.Cells(i, 1).Value Then
    
      'Set Ticker Name
      Ticker = ws.Cells(i, 1).Value
               
      'Set Total Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
      'Set Yearly Change
      Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(startrow, 3).Value
                 
      'Set Percentage Change
      OpenPrYear = ws.Cells(startrow, 3).Value
        If OpenPrYear = 0 Then
        Percentage_Change = 0
      
        Else
        Percentage_Change = Yearly_Change / OpenPrYear
         End If
         
      'Add Ticker to the Summary Table
      ws.Range("I" & SummaryTick_Row).Value = Ticker
      
      'Add Yearly_Change to the Summary Table
      ws.Range("J" & SummaryTick_Row).Value = Yearly_Change
      
      'Add Percentage Change to Summary Table
      ws.Range("K" & SummaryTick_Row).Value = Percentage_Change
      ws.Range("K" & SummaryTick_Row).NumberFormat = "0.00%"
      
      
      'Add Total Stock Volume to the Summary Table
      ws.Range("L" & SummaryTick_Row).Value = Total_Volume
      
      Total_Volume = 0
      Yearly_Change = 0
      Percentage_Change = 0
      SummaryTick_Row = SummaryTick_Row + 1
      startrow = i + 1
    Else
    
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
    End If

Next i

'Declare variables to find greatest% inc and dec;greatest total volume
Dim SumtablelastRow As Integer
Dim percentmax As Double
Dim percentmin As Double
Dim volumemax As LongLong
Dim rp As Range
Dim rv As Range
Dim Ticksum As String

'Set equation to arrive at last row of each Summary Table
SumtablelastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Loop thru all rows of the summary table to format Yearly_Change Column for each ws.

For i = 2 To SumtablelastRow

    If ws.Cells(i, 10) > 0 Then
    ws.Cells(i, 10).Interior.Color = vbGreen

    Else
    ws.Cells(i, 10).Interior.Color = vbRed
    
    End If

Next i

percentmax = 0
percentmin = 0
volumemax = 0

    'Set column range to find values for percentmax, percentmin, and volumemax
    Set rp = ws.Range("K2:K" & SumtablelastRow)
    Set rv = ws.Range("L2:L" & SumtablelastRow)
    
    'Find value for percentmax,percentmin and volumemax
    percentmax = Application.WorksheetFunction.Max(rp)
    percentmin = Application.WorksheetFunction.Min(rp)
    volumemax = Application.WorksheetFunction.Max(rv)
    
    'Add percentmax,percentmin,and volumemax on second summary table. If value is 0, nothing is displayed.
    If percentmax > 0 Then
    ws.Range("Q2") = percentmax
    ws.Range("Q2").NumberFormat = "0.00%"
    
    End If
    
    If percentmin < 0 Then
    ws.Range("Q3") = percentmin
    ws.Range("Q3").NumberFormat = "0.00%"
    
    End If
    
    ws.Range("Q4") = volumemax

For i = 2 To SumtablelastRow
    
    'Set conditionals to arrive at Ticksum (Ticker symbol) for percentmax,percentmin and volume max
    'Add Tickersum to second summary table
    'If value for percentmax and percentmin is 0, then no Ticker symbol is displayed.
    
    If ws.Cells(i, 11).Value = percentmax And percentmax > 0 Then
    Ticksum = ws.Cells(i, 9).Value
    ws.Cells(2, 16) = Ticksum
    
    End If
      
    If ws.Cells(i, 11).Value = percentmin And percentmin < 0 Then
    Ticksum = ws.Cells(i, 9).Value
    ws.Cells(3, 16).Value = Ticksum
    
    End If
       
    If ws.Cells(i, 12) = volumemax Then
    Ticksum = ws.Cells(i, 9).Value
    ws.Cells(4, 16).Value = Ticksum
    
    End If

Next i

Next ws

End Sub


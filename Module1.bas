Attribute VB_Name = "Module1"
Sub Ticker_Setup()

 ' this is to help the macro run quicker
 Application.ScreenUpdating = False
 Application.Calculation = xlCalculationManual
 Application.EnableEvents = False
 
 ' denoting the ws and adding a function to select and run through each tab
 Dim ws As Worksheet
 For Each ws In ThisWorkbook.Worksheets
 ws.Select
 
  ' help create the conditional formatting for values with positive and negative growth
  Dim rng As Range
  Dim greater As FormatCondition, lessthan As FormatCondition


  ' Set to identify the name of the Ticker
  Dim Ticker_Name As String
  
  ' Set to identify the closing value of the Ticker
  Dim Ticker_Close As String
    
  ' Set to identify the opening value of the Ticker
  Dim Ticker_Open As String

  ' Set to identify the total value of the Ticker
  Dim Ticker_Total As Double
  Ticker_Total = 0



  ' labels for all sections being added
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"

  ' this is where we create the summary for the data
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' holds the value of the end of the row to help tie up all format and formulas
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' again for format help on the positive or negative growth
    Set rng = Range("J2:J" & lastrow)
    rng.FormatConditions.Delete
   Set greater = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
   Set lessthan = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")

  ' Loop through all ticker data
  For i = 2 To lastrow
  
    ' condition to help find the opening value of the first day since it is sorted first by ticker than by date
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
    ' opening ticker value
    Ticker_Open = Cells(i, 3).Value
    
    
    End If

    ' value used to track the ticker for each line, once it changes it will skip this section and start a new one
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' record ticker name
      Ticker_Name = Cells(i, 1).Value
      
      ' record ticker closing value
      Ticker_Close = Cells(i, 6).Value
      
      ' this is to caluclate the yearly change from open day one to last day close (hence year change)
      Ticker_Year_Change = Ticker_Close - Ticker_Open
      
      ' caluclation to find the percent change
      Ticker_Percent_Change = (Ticker_Year_Change / Ticker_Open)

      ' totals the ticker volume
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

      ' record all values in the new summary
      Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      Range("J" & Summary_Table_Row).Value = Ticker_Year_Change
   
      Range("K" & Summary_Table_Row).Value = Ticker_Percent_Change

      Range("L" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Total = 0

    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      
    End If

  Next i

  ' this starts after the i iterations because this new value depends on the last loop to fully complete
  ' this now finds 3 new values, max percentage, min percentage and the minimum value
  Max_Percent = WorksheetFunction.Max(Range("K2:K" & lastrow))
  Min_Percent = WorksheetFunction.Min(Range("K2:K" & lastrow))
  Max_Value = WorksheetFunction.Max(Range("L2:L" & lastrow))
  
  ' these xlookup against the min and max percents and values previously to pull the correct ticker that has the min and max
  TickerMax = WorksheetFunction.XLookup(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), (Range("I2:I" & lastrow)))
  TickerMin = WorksheetFunction.XLookup(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), (Range("I2:I" & lastrow)))
  TickerMaxV = WorksheetFunction.XLookup(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), (Range("I2:I" & lastrow)))


  ' places all the values in the second summary box that was created after the loop
  Range("Q2").Value = Max_Percent
  Range("Q3").Value = Min_Percent
  Range("Q4").Value = Max_Value
  Range("P2").Value = TickerMax
  Range("P3").Value = TickerMin
  Range("P4").Value = TickerMaxV
  
  ' final formatting to match pictures in direction
  Range("J2:J" & lastrow).NumberFormat = "0.00"
  Range("K2:K" & lastrow).NumberFormat = "0.00%"
  Range("Q2").NumberFormat = "0.00%"
  Range("Q3").NumberFormat = "0.00%"
  Range("I:L").EntireColumn.AutoFit
  Range("O:Q").EntireColumn.AutoFit
    
    With greater
    .Interior.Color = vbGreen
   End With

   With lessthan
     .Interior.Color = vbRed
   End With

 Next ws
 Application.ScreenUpdating = True
 Application.Calculation = xlCalculationAutomatic
 Application.EnableEvents = True
End Sub

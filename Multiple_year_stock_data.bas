Sub Stonks()

Dim LR As Long

 LR = Range("A:A").SpecialCells(xlCellTypeLastCell).Row

         Dim Current As Worksheet

         For Each Current In Worksheets

            Stock_Volume = 0

            
     Dim Stock_Open As Double
     
     Stock_Open = Cells(2, 3).Value

  Dim lastRow As Long
  
  Summary_Table_Row = 2

  ' Loop through all credit card purchases
  For input_row = 2 To LR

    If Cells(input_row + 1, 1).Value <> Cells(input_row, 1).Value Then

      Stock_Name = Cells(input_row, 1).Value

      Stock_Volume = Stock_Volume + Cells(input_row, 7).Value

      Range("I" & Summary_Table_Row).Value = Stock_Name
      
      Range("J" & Summary_Table_Row).Value = Cells(input_row, 6).Value - Stock_Open
      
      Range("K" & Summary_Table_Row).Value = (Cells(input_row, 6).Value - Stock_Open) / Stock_Open * 100
      
      Stock_Open = Cells(input_row + 1, 3).Value

      Range("L" & Summary_Table_Row).Value = Stock_Volume

      Summary_Table_Row = Summary_Table_Row + 1
      
      Stock_Volume = 0

    Else

      Stock_Volume = Stock_Volume + Cells(input_row, 7).Value

    End If

  Next input_row


Dim Rng As Range

Set Rng = Range("J2:J3001")

Cells(2, 16).Value = WorksheetFunction.Max(Rng)
Cells(3, 16).Value = WorksheetFunction.Min(Rng)


Dim RngV As Range

Set RngV = Range("L2:J3001")

Cells(4, 16).Value = WorksheetFunction.Max(RngV)


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"
Cells(1, 16).Value = "Value"
Cells(2, 15).Value = "Greatest Increase"
Cells(3, 15).Value = "Greatest Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

            
         Next

For input_row = 2 To LR


If Cells(input_row + 1, 10).Value >= 0 Then Cells(input_row + 1, 10).Interior.ColorIndex = 4

If Cells(input_row + 1, 10).Value < 0 Then Cells(input_row + 1, 10).Interior.ColorIndex = 3

Next input_row

      End Sub


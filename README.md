Sub year_data()


Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    

  ' Set a variable for the ticker name

  Dim ticker As String
  



  ' Set an variable for stock volume

  Dim total_stock_volume As Double

  total_stock_volume = 0



  ' Keep track of the location for each ticker volume

  Dim Summary_Table_Row As Integer

  Summary_Table_Row = 2



  ' Loop through all tickers

  For i = 2 To 800000



    ' Check if we are still within the same ticker symbol, if it is not...

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then



      ' Set the ticker symbol

      ticker = Cells(i, 1).Value



      ' Add to the ticker Total

     total_stock_volume = total_stock_volume + Cells(i, 7).Value



      ' Print the Ticker Symbol in the Summary Table

      Range("I" & Summary_Table_Row).Value = ticker



      ' Print the ticker volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = total_stock_volume



      ' Add one to the summary table row

      Summary_Table_Row = Summary_Table_Row + 1

      

      ' Reset the ticker Total

      total_stock_volume = 0



    ' If the cell immediately following a row is the same ticker...

    Else



      ' Add to the ticker Total

      total_stock_volume = total_stock_volume + Cells(i, 7).Value



    End If




Next i

Next ws



End Sub


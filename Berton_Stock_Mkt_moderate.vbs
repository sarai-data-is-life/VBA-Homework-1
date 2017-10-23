 ' Moderate with challenge:

    ' 1. Create a script that will loop through all the stocks and take the following info.
        ' a. Yearly change from what the stock opened the year at to what the closing price was.
        ' b. The percent change from the what it opened the year at to what it closed.
        ' c. The total Volume of the stock and ticker symbol
        ' d. Apply conditional formatting that will highlight positive change in green and negative change in red.
    ' 2. Make the appropriate adjustments to the script that will allow it to run on every worksheet just by running it once.


Sub moderate_Hmwk2()

 ' Loop through all the sheets
    
    For Each ws In Worksheets
 
        ' Insert the labels for cell I1 through L1 and change the width of columns J, K & L.
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("J:L").ColumnWidth = 15
        ws.Columns("J").NumberFormat = "0.00"
        ws.Columns("K").NumberFormat = "0.0000%"
        
        ' Find the last row of the each (year) worksheet and find each stock symbol in column A and put it in column I
        ' Find the total volume of the each stock symbol and put it in column L so that it coincides with the its stock symbol.

        Dim yearLastRow As Long
        Dim i As Long
        Dim ticker_location As Integer
        Dim volume As Double
        Dim open_val As Double
        Dim close_val As Double
        Dim stock_start As Long
        ticker_location = 2
        stock_start = 2
        volume = 0
        
        yearLastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To yearLastRow


            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the stock symbol
                ws.Range("I" & ticker_location).Value = ws.Range("A" & i).Value

                ' Note: You may want to loop here to find the first non-zero
                ' opening price and increment stock_start each loop
                open_val = ws.Cells(stock_start, 3).Value
                stock_start = i + 1 ' Location of the next stock start
                

             ' Add the last row to volume total
                volume = volume + ws.Range("G" & i).Value
            
            ' Place the total volume in column L that coincides with its ticker symbol
                ws.Range("L" & ticker_location).Value = volume

            ' Get the closing value for the year  for the stock symbol.

                close_val = ws.Cells(i, 6).Value

            ' Place the  change in value and the percent change in column J
           
                ws.Range("J" & ticker_location).Value = (close_val - open_val)

             
                If open_val = 0 Then
                    ws.Range("K" & ticker_location).Value = 0
                Else
                    ws.Range("K" & ticker_location).Value = ((close_val - open_val) / open_val)
                End If

            ' Add to the ticker  & volume location
               ticker_location = ticker_location + 1

             ' Reset the volume

                volume = 0

            Else
            ' If the cell immediately following a row is the same stock
            ' Add to the volume for the stock symbol
            
               volume = volume + ws.Range("G" & i).Value


            End If
        
        Next i

        Dim value_lastRow As Double
  
        ' Find the last non-blank cell in column K where the new data is stored
        value_lastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
               
        For i = 2 To value_lastRow
            If ws.Range("J" & i).Value > 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & i).Value < 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 3
            End If
             
        Next i

    Next ws
         

MsgBox ("Changes Applied")

End Sub



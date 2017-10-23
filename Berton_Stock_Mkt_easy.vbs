' Easy plus the Challenge

' 1. Place "Ticker" in Cell "I1" and "Total Stock Volume" in Cell "J1" for each sheet.
' 2. Loop through each sheet (year) and place the stock symbol in Column A in Row I.
' 3. Loop through each sheet (year) and add the volume in colimn G for each stock and place in Row J.
' 4. Make the appropriate adjustments to the script that will allow it to run on every worksheet just by running it once.


Sub Easy_Hmwk2()

    ' Loop through all the sheets
    
    For Each ws in Worksheets
 
        ' Insert the labels for cell I1 & cell J1 and change the width of column J. 
        ws.Range("I1").value = "Ticker"
        ws.Range("J1").value = "Total Stock Volume"
        ws.Columns("J").ColumnWidth = 20
        
        ' Find the last row of the each (year) worksheet and find each stock symbol in column A and put it in column I
        ' Find the total volume of the each stock symbol and put it in column J so that it coincides with the its stock symbol.

        Dim yearLastRow as Long
        Dim i as Long
        Dim ticker_location as Integer
        Dim volume as Double
        ticker_location = 2
        volume = 0
        
        yearLastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To yearLastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the stock symbol
                ws.Range("I" & ticker_location).Value = ws.Range("A" & i).Value

            ' Add the last row to volume total
                volume = volume + ws.Range("G" & i).Value
            
            ' Place the total volume in column J that coincides with its ticker symbol
                ws.Range("J" & ticker_location).Value = volume
            
            ' Add to the ticker  & volume location
                ticker_location = ticker_location + 1

            ' Reset the volume
                volume = 0

            ' If the cell immediately following a row is the same stock
            Else
                
                ' Add to the volume for the stock symbol
                volume = volume + ws.Range("G" & i).Value
               

            End If
        
        Next i

  
    Next ws
MsgBox ("Changes Applied")

End Sub
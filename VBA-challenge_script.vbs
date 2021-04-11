Sub VBA_challenge():

    'Establish your worksheet and loop varibales
    Dim ws As Worksheet
    Dim i As Long
   
    For Each ws In ActiveWorkbook.Worksheets

        'Find the value in the very last row of the data set
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Begin in row 2 of your data (because row 1 has headers), call this your "data_table_row"
        Dim data_table_row As Long
        data_table_row = 2

        'Establish a counter for Total Stock Volume, name it "total_stock_volume" and set it to 0
        Dim total_stock_volume As Double
        total_stock_volume = 0

        'Set your first value as the open price, name it "open_price" (this is the equivalent as clicking your mouse into the second row in column C)
        Dim open_price As Double
        open_price = ws.Range("C2").Value

        'Begin your for Loop to go through each row to find the first open price and the last close price
        For i = 2 To last_row

            'Use an if statement to check if the current ticker is different from next one
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Give your ticker a name and a value
                Dim Ticker As String
                Ticker = ws.Cells(i, 1).Value

                'Place your new ticker value in the summary table
                ws.Cells(data_table_row, 10).Value = Ticker

                'Add <vol> to total_stock_volume variable
                total_stock_volume = total_stock_volume + ws.Cells(i, 8).Value

                'Place the total_stock_volume in the summary table
                ws.Cells(data_table_row, 13).Value = total_stock_volume

                'Reset total_stock_volume to 0
                total_stock_volume = 0

                'Set close_price value
                Dim close_price As Double
                close_price = ws.Cells(i, 6).Value

                'Calculate the yearly change by subtracting your open_price value from your close_price value
                Dim yearly_change As Double
                yearly_change = close_price - open_price

                'Place your yearly_change value in the summary table
                ws.Cells(data_table_row, 11).Value = yearly_change

                'Create a variable for percentage change and calculate the percentage change by dividing yearly change and open price
                
                Dim percentage_change As Double

                If open_price = 0 Then
                    ws.Cells(data_table_row, 12).Value = "-"

                Else
                    percentage_change = yearly_change / open_price
                    ws.Cells(data_table_row, 12).Value = percentage_change
                    ws.Cells(data_table_row, 12).NumberFormat = "0.00%"

                End If

                'Format any cells that have a change of 0 to align right so that your data is clean
                ws.Cells(data_table_row, 12).HorizontalAlignment = xlRight

                'Use an if statement to format positive change in green and negative change in red (0 change is neither positive nor negative)
                If ws.Cells(data_table_row, 11).Value = 0 Then
                    ws.Cells(data_table_row, 11).Interior.ColorIndex = 0
                
                ElseIf ws.Cells(data_table_row, 11).Value > 0 Then
                    ws.Cells(data_table_row, 11).Interior.ColorIndex = 4

                ElseIf ws.Cells(data_table_row, 11).Value < 0 Then
                    ws.Cells(data_table_row, 11).Interior.ColorIndex = 3

                End If

                'Reset your open_price to beginning price of the next ticker type
                open_price = ws.Cells(i + 1, 3).Value
                data_table_row = data_table_row + 1

            'Else condition on what to do when current ticker and next ticker are same
            Else

                'Add <vol> to total_stock_volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

            End If
            
        Next i

        'Create and format your summary table
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percentage Change"
        ws.Range("M1").Value = "Total Stock Volume"

        'Bonus - create and format your bonus table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O2").Font.Bold = True
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O3").Font.Bold = True
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("O4").Font.Bold = True
        ws.Range("P1").Value = "Ticker"
        ws.Range("P1").Font.Bold = True
        ws.Range("Q1").Value = "Value"
        ws.Range("Q1").Font.Bold = True

        'Find max % and its corresponding ticker value
        Dim max_percentage As Double
        max_percentage = Application.WorksheetFunction.Max(ws.Columns("L:L"))
        ws.Range("Q2").Value = max_percentage
        ws.Range("Q2").Style = "Percent"
            ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = Application.WorksheetFunction.Index(ws.Columns("J:M"), Application.WorksheetFunction.Match(ws.Range("Q2").Value, ws.Columns("L:L"), 0), 1)
                
        'Find min % and its corresponding ticker value
        Dim min_percentage As Double
        min_percentage = Application.WorksheetFunction.Min(ws.Columns("L:L"))
        ws.Range("Q3").Value = min_percentage
        ws.Range("Q3").Style = "Percent"
            ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = Application.WorksheetFunction.Index(ws.Columns("J:M"), Application.WorksheetFunction.Match(ws.Range("Q3").Value, ws.Columns("L:L"), 0), 1)
        
        'Find max volume and its corresponding ticker value
        Dim max_volume As Double
        max_volume = Application.WorksheetFunction.Max(ws.Columns("M:M"))
        ws.Range("Q4").Value = max_volume
        ws.Range("P4").Value = Application.WorksheetFunction.Index(ws.Columns("J:M"), Application.WorksheetFunction.Match(ws.Range("Q4").Value, ws.Columns("M:M"), 0), 1)

        ws.Columns("J:M").EntireColumn.AutoFilter
        ws.Columns("A:Q").EntireColumn.AutoFit

    Next ws

End Sub
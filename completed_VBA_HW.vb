Option explicit

Sub stonks():
    'set variable to worksheets
    Dim ws As Worksheet
    'begin worksheet loop
    For Each ws In Sheets
        'create output table columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        'set up counter for ticker
        'set up counter for yearly change
        Dim Open_price, close_price, total_stock As Double
        Dim ticker_row, output_table As Long
            output_table = 2
        
        'save open price
        Open_price = ws.Range("C2").value

        'start loop and go through earch ticker row
        For ticker_row = 2 To ws.Range("A1").End(xlDown).Row

            'identify if next row of ticker is the same
            If ws.Cells(ticker_row, 1).value <> ws.Cells(ticker_row + 1 , 1) Then

                'add values form col G to Total Stock Volume and go to next row
                total_stock = total_stock + ws.Cells(ticker_row, 7).Value

                'identify last close year value
                close_price = ws.Cells(ticker_row, 6).Value

                'if statement to obtain percent change
                if Open_price = 0 then
                    ws.Cells(output_table, 11).value = 0
                Else
                    ws.Cells(output_table,11).value = (close_price - Open_price) / Open_price
                end if

                'provide Ticker name in column I
                ws.Cells(output_table, 9).Value = ws.Cells(ticker_row, 1).Value

               'subtract starting market value (X) from final market value (y) ----> y - x = yearly change
               'input result on columna J
                ws.Cells(output_table, 10).Value = close_price - Open_price
                                
                'provide Total Stock Volume in column L
                ws.Cells(output_table, 12).Value = total_stock

                'add 1 row to th output table
                output_table = output_table + 1

                'reset total stock volume counter
                total_stock = 0
                'update open to be the open price of the new ticker
                Open_price = ws.Cells(ticker_row + 1, 3).value       

            Else
                'add values form col G to Total Stock Volume and go to next row
                total_stock = total_stock + ws.Cells(ticker_row, 7).Value

            end if
            
            'color negative values as red and positive values as green
            If ws.Range("J" & output_table).Value < 0 Then
                ws.Range("j" & output_table).Interior.ColorIndex = 3
            elseif ws.Range("J" & output_table).Value > 0 Then
                ws.Range("j" & output_table).Interior.ColorIndex = 4
            End If

        'end ticker row count
        Next ticker_row
        'format percent change to Percent
        ws.Range("K:K").NumberFormat = "0.00%"

        'BONUS!!

        'asssign Headers
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest total volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "value"
        ws.Range("P2, P3").NumberFormat = "0.00%"
        'label variables
        Dim bonus As Integer
        Dim lastrow As Long
        lastrow = Cells(Rows.Count, 11).End(xlUp).Row
        
        'locate greatest % increase, greatest % decrease and Greatest total Volume
        '*obtained below code from Slack AskBCS Leanrning Assistant*
        For bonus = 2 To lastrow
            If ws.Range("K" & bonus).Value > ws.Range("P2").Value 'or ws.Range("P2").value = "NaN" Then
                ws.Range("P2").Value = ws.Range("K" & bonus).Value
                ws.Range("O2").Value = ws.Range("I" & bonus).Value
            End If
            If ws.Range("K" & bonus).Value < ws.Range("P3").Value Then
                ws.Range("P3").Value = ws.Range("K" & bonus).Value
                ws.Range("O3").Value = ws.Range("I" & bonus).Value
            End If
            If ws.Range("L" & bonus).Value > ws.Range("P4").Value Then
                ws.Range("P4").Value = ws.Range("L" & bonus).Value
                ws.Range("O4").Value = ws.Range("I" & bonus).Value
            End If
        Next bonus
        'fit all cokumns
        ws.Columns("A:P").AutoFit
    'end Ws loops    
     Next ws
    
End Sub
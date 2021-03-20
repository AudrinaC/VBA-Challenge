Attribute VB_Name = "Module1"
Sub StockChallenge()
    
    'Loop through each worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets

        'Declare Variables
        Dim vol_total As Double
        Dim summary_table_row As Integer
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim lastrow As Long
        
        'Set starting values
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        vol_total = 0
        summary_table_row = 2
        open_price = 0
        close_price = 0
        yearly_change = 0
        percent_change = 0
        
        
        'Create column labels for the analysis
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"


        'Loop to search through each row except for the headers
        For i = 2 To lastrow
            
            'Conditional to determine opening value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                open_price = ws.Cells(i, 3).Value
            End If

            'Add to the "Total Stock Value" amount
            vol_total = vol_total + ws.Cells(i, 7)

            'Conditional to determine when the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'Display unique ticker value and total volume
                ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summary_table_row, 12).Value = vol_total

                'Calculate yearly change
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                ws.Cells(summary_table_row, 10).Value = yearly_change

                'Conditional to formatting of positive change as green and negative change as red
                If yearly_change >= 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If
                
        'Find percent change
        If open_price = 0 And close_price = 0 Then
            percent_change = 0
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        ElseIf open_price = 0 Then
            'Cannot divide by 0, but must reflect percent change from 0
            Dim new_stock As String
            new_stock = "New Stock"
            ws.Cells(summary_table_row, 11).Value = new_stock
        Else
            percent_change = yearly_change / open_price
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
            End If
            

               'Increment summary_table_row and reset other values
                summary_table_row = summary_table_row + 1
                vol_total = 0
                open_price = 0
                close_price = 0
                yearly_change = 0
                percent_change = 0
                
             
            End If
            
        Next i

    Next ws

End Sub

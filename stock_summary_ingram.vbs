Attribute VB_Name = "Module1"
Sub stock_summary_info():

'Looping over worksheet
'Note: For ALL worksheets data should be in columns A:G, in the follwing order:
'<ticker>,<date>,<open>,<high>,<low>,<close>,<vol>

For Each ws In Worksheets

'This code requires sheets be sorted by ticker and date in ascending order.
'The following code was adapted from https://trumpexcel.com/sort-data-vba/
'to sort the data in the worksheet

'save last row
Dim last_row As Long
last_row = ws.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

With ws.Sort
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SetRange Range("A1:G" & last_row)
     .Header = xlYes
     .Apply
End With

'----------------------------------------------------------
'Creating the summary table

'Headers
ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

'Table values
Dim open_price As Double
Dim close_price As Double
Dim vol_sum As LongLong
Dim stock_count As Integer
Dim stock_symbol As String

stock_count = 0

For i = 2 To last_row

    'If new stock, update stock count,save name in table, save openning price, restart volume sum
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) Then
        stock_count = stock_count + 1
        stock_symbol = ws.Cells(i, 1).Value
        ws.Range("I" & (stock_count + 1)).Value = stock_symbol
        open_price = ws.Cells(i, 3).Value
        vol_sum = ws.Cells(i, 7).Value
     
     'If same stock, compute price change value, price change percent, and volume sum, and update table
     'Note that because we sorted stocks by ascending date the final price change recorded will be the
     'change from the opening price at the beginning of a given year to the closing price at the end of that year
    Else
        close_price = ws.Cells(i, 6).Value
        ws.Range("J" & (stock_count + 1)).Value = close_price - open_price
        ws.Range("K" & (stock_count + 1)).Value = (close_price - open_price) / open_price
        vol_sum = vol_sum + ws.Cells(i, 7).Value
        ws.Range("L" & (stock_count + 1)).Value = vol_sum
        
    End If
    
Next i

'Formatting summary table
    Dim table_last_row As Long
    table_last_row = stock_count + 1

'Percent
ws.Range("K2:K" & table_last_row).NumberFormat = "0.00%"

'Highlighting (There is conflicting info on whether we are supposed to highlight
'the raw and % change columns. The images only show raw change highlighted but
'the requirement sections mentions conditional formatting for the raw and
'% change columns... I've matched the images. Let me know if in the future
'I should match the requirements section instead

For i = 2 To table_last_row

    If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4   'Green
        
    ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3   'Red
        
    Else
        ws.Range("J" & i).Interior.ColorIndex = 0   'No fill
    
    End If
    
    'Autofit text
    ws.Range("I1:L" & table_last_row).Columns.AutoFit

Next i

'----------------------------------------------------------
'Pulling out certain stats

'Labels
ws.Range("P1:Q1").Value = Array("Ticker", "Value") 'Columns
ws.Range("O2:O4").Value = Application.Transpose(Array("Greatest % increase", "Greatest % decrease", "Greatest total volume")) 'Rows

Dim max_increase As Double
Dim max_decrease As Double
Dim max_vol As LongLong

max_increase = WorksheetFunction.Max(ws.Range("K2:K" & table_last_row))
max_decrease = WorksheetFunction.Min(ws.Range("K2:K" & table_last_row))
max_vol = WorksheetFunction.Max(ws.Range("L2:L" & table_last_row))

For i = 2 To table_last_row
    If ws.Range("K" & i).Value = max_increase Then
        ws.Range("P2").Value = ws.Range("I" & i).Value
        ws.Range("Q2").Value = ws.Range("K" & i).Value
        ws.Range("Q2").NumberFormat = "0.00%"
        
    End If
    
    If ws.Range("K" & i).Value = max_decrease Then
        ws.Range("P3").Value = ws.Range("I" & i).Value
        ws.Range("Q3").Value = ws.Range("K" & i).Value
        ws.Range("Q3").NumberFormat = "0.00%"
    End If
    
    If ws.Range("L" & i).Value = max_vol Then
        ws.Range("P4").Value = ws.Range("I" & i).Value
        ws.Range("Q4").Value = ws.Range("L" & i).Value
    End If
    
    'Note if there is a tie for the max value this loop will only return
    'the last stock, alphabetically, not all stocks. With continuous
    'variables like these the risk of ties seems low but if it's a
    'concern the code could be modified.
    
Next i

'Autofit text
    ws.Range("O1:Q4").Columns.AutoFit

Next ws

End Sub

# VBA-challenge (UofM bootcamp challenge 2)

This repository contains a vbs script that generates summary tables for annual stock data. To add the vbs script to your excel file, go to the developer tab and open Visual Basic.  From the file menu, select "Import File" and select "stock_summary_ingram.vbs".  If you don't see it, make sure you have the "All files" option selected.

The summary table generated includes yearly change from the opening price at the beginning of a given year to the closing price at the end of that year, the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year, and the total stock volume of the stock for each stock.

The script also reports the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

Note: For the script to perform correctly, each worksheet's data must be in columns A:G and in the follwing order:
'ticker, date, open, high, low, close, vol.



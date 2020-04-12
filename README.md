# Stocks-Analysis
#Analyzing stock data to provide best stock
---
Generate a code to analize the stocks to provide the best options. We are computing the total volume and return (taking the price of closing at the beggining and the end of the year) per each stock in a year.  Return will be the ending price divided by the starting price minus 1.

To make this analysis we are using a nested loop.  The first for loop will be used to loop over the tickers, then the second for loop will be used to loop over the rows from 2nd to last row.

#Challenge2 Module2
---
Excel file name: green_stocks - Challenge Module 2 - Dahiana Mejia

Macro name for original code: AllStocksAnalysis

Macro name for challenge 2 code: AllStocksAnalysisChallenge2

Analysis
---
We are refactoring the code with the nested loop.  We are using only one for loop to loop over the tickers and compute the total volume and return.

We created arrays for tickers, total volume, starting prices and ending prices.

We set a tickerindex to call each ticker, then we will use a for loop to loop over the rows from 2nd to the last row.

We are computing the total volume and return more efficient.  We only need to open once the worksheets "year value", then compute the total volume and take the values for starting and ending price per ticker.  Since we created arrays these values will be storage on each array.  After finishing the for loop, we open the worksheet "AllStockAnalysis" once and print the results.

On the original code, we need to open the worksheets "year value" to take the ticker that we need to analyze, then loop over the rows, then open the worksheets "AllStockAnalysis" to print the results, then open again the worksheet "year value" to read the next ticker.  We open the worksheets 12 times, because we have 12 tickers, this make this code run slower than the refactoring one.

In conclusion we are making a code that runs faster, more efficient, and it is giving same results required.

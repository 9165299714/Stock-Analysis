# Stock Analysis

## Overview
### Purpose
The purpose of this project was to refactor a VBA Code that was implemented to analyze stock data in two years: 2017 and 2018. The goal behind refactoring this code was to increase efficiency and run the code faster.

### The Data
The data that is given includes two tables with information on 12 stocks. The information contains the ticker's value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

## Results
### Analysis
Prior to refactoring the code, I copied the code that was needed to create the chart headers, input box, ticker array, and to activate the worksheet that corresponds with the correct data. Then, the steps were listed in order to place the structure for the refactoring. Below is the instruction and code as written in the file. 

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

## Summary
### Pros and Cons of Refactoring Code
By refactoring code, we can structure and format our code in a more clean, concise, and straight forward manner. By having a simpler code structure, it is easier to catch bugs and errors, thereby creating a more efficient program. However, in the process of refactoring, new bugs may appear. Also, as refactoring does not change the nature, only the organization of code, it may be deemed as tedious or time-wasted by a client.  

### The Advantages of Refactoring Stock Analysis
By refactoring the code, the macro run time decreased greatly. The original analysis took more than half a second to run, whereas the new analysis after refactoring only took about a quarter of that time to run. Attached below are the screenshots that indicate the run time for our new analysis.
<img width="287" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/90435301/135735120-f50778cc-4a13-4783-9268-0f69cd3a9cc1.png">

<img width="282" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/90435301/135735121-869801e6-80b2-4332-bfbd-7b56e8c19c40.png">


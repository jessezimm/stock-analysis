# Stock Analysis with VBA
Click here to view the Excel file: [Module 2 Challenge - VBA (https://github.com/jessezimm/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project
The purpose of this Module 2 Challenge is to refactor Microsoft Excel VBA code so it runs more efficiently. The VBA code collects stock information such as stock volumes and returns in 2017 and 2018 to determine if these stocks are worth the investment.

## Results
### Process
To create the more efficient code, I copied the Module code and was provided commented steps to guide my construction of the refactorization. I created a tickerIndex and three output arrays. Then, I used for loops and if then statements to gather information stock by stock. Each stock output was produced on the **All Stocks Analysis** worksheet. Please see the refactored code below.

  '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
          
    Next i
    
### Analysis
For 2017 and 2018, the data is presented in two charts each containing the stock ticker (trading symbol), the total daily volume, and the annual return for each stock. The annual return is calculated by taking the stock's initial price at the beginning of the year and dividing it by the stock's year-ending price, and then, subtracting one. In 2017, only one of 12 stocks (Stock: TERP) had a negative return. In 2018, only two stocks had a positive return, ENPH and RUN. Overall, 2017 was a better performing year than 2018 for the 12 stocks evaluated. 

![Stock Analysis - 2017](https://github.com/jessezimm/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![Stock Analysis - 2018](https://github.com/jessezimm/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

## Summary
### Refactoring Code
Refactoring code makes it cleaner and more efficient. Additionally, organized code is easier to read and debug. If attempting to refactor code, one may accidently delete or bug the original code. This disadvantageous process can lead to a longer time spent recreating code. Further, spending time refactoring code can take valuable time away from attempting new coding projects. 

### Refactoring VBA Script
Refactoring VBA code decreases the code runtime and makes the code more cyclical, which is extremely beneficial if you want to create code for larger datasets. However, troubleshooting refactored code is much more intuitive and time consuming. 

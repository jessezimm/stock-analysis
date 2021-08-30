# Stock Analysis with VBA
Click here to view the Excel file: [Module 2 Challenge - VBA ()

## Overview of Project
The purpose of this Module 2 Challenge is to refactor Microsoft Excel VBA code so it runs more efficiently. The VBA code collects stock information such as stock volumes and returns in 2017 and 2018 to determine if these stocks are worth the investment.

## Results
### Process
To create the more efficient code, I copied the Module code and was provided commented steps to guide my construction of the refactorization. I created a tickerIndex and three output arrays. Then, I used for loops and if then statements to gather information stock by stock. Each stock output was produced on the **All Stocks Analysis** worksheet. Please see the refactored code below.

### Analysis
For 2017 and 2018, the data is presented in two charts each containing the stock ticker (trading symbol), the total daily volume, and the annual return for each stock. The annual return is calculated by taking the stock's initial price at the beginning of the year and dividing it by the stock's year-ending price, and then, subtracting one. In 2017, only one of 12 stocks (Stock: TERP) had a negative return. In 2018, only two stocks had a positive return, ENPH and RUN. Overall, 2017 was a better performing year than 2018 for the 12 stocks evaluated. 

![Stock Analysis - 2017]()
![Stock Analysis - 2018]()

## Summary
### Refactoring Code
Refactoring code makes it cleaner and more efficient. Additionally, organized code is easier to read and debug. If attempting to refactor code, one may accidently delete or bug the original code. This disadvantageous process can lead to a longer time spent recreating code. Further, spending time refactoring code can take valuable time away from attempting new coding projects. 

### Refactoring VBA Script
Refactoring VBA code decreases the code runtime and makes the code more cyclical, which is extremely beneficial if you want to create code for larger datasets. However, troubleshooting refactored code is much more intuitive and time consuming. 

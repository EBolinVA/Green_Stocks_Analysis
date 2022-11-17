# Analyzing Stocks for 2017 & 2018

## Overview of VBA Refactoring Challenge

A friend, Steve, asked for help analyzing a variety of green stocks for his parents. I started with an Excel file of the stock data he wants help analyzing. To get the most out of Excel, I used Visual Basic for Applications to automate some of the tasks instead of manually duplicating excel formulas over a very large spreadsheet. Along the way I learned several coding techniques, called macros or subroutines, to run the analysis more efficiently.

## Green Stocks Analysis Results

Steve provided an Excel file with two worksheets containing data for the 12 stocks. Each worksheet (named "2017" and "2018") had daily records over the course of a year for opening price, closing price, and volume traded.

### 2017 Results 

For the 12 stocks analyzed for 2017, all but one stock had a positive return. The best performing stocks were DQ with 199.4% return, and SEDG with 184.5% return. The lowest performing stock was TERP with a -7.2% return.

### 2018 Results

The story in 2018 was not as rosy. All but two stocks suffered losses. The only two stocks with positive returns in 2018 were ENPH at 81.9% return, and RUN with 84.0% return. DQ ended the year down 81.9% and SEDG lost 7.8%.

## How the returns were calculated
Indexes and for loops were used in the coding to loop over all the rows in the spreadsheet while collecting data for Starting Price, Ending Price, and Volume for each stock. The Starting Price was collected from the first row (first day) of each stock, and the Ending Price was collected from the last row (last day) of each stock. Volumes were collected daily for each stock and added together in order to be output into the 'Return' column on the analysis worksheet. 

```
''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker. In the cells range, i is the row, and 8 is the column holding values for daily stocks.
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + (Cells(i, 8).Value)
                
                    End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                    'Starting price is found in the "Closing Price" column on the stocks worksheet, column 6
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                           
            
        'End If
                    End If
            
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1) <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                    'Ending price is found in the "Closing Price" column on the stocks worksheet, column 6
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
                    tickerIndex = (tickerIndex + 1)
            
        'End If
    
                    End If
    Next i

### 2017 Coding Efficiencies

In the original script, the execution time of the code was 0.585 seconds for 2017 stocks. However, after refactoring the script, the code ran more efficiently in 0.148 seconds. 

![Image of pop-up message showing elapsed run time of original script for 2017 stocks](https://github.com/EBolinVA/Module_2_VBA_Challenge/blob/main/Original_time_2017.png)

![Image of pop-up message showing elapsed run time of refactored script for 2017 stocks](https://github.com/EBolinVA/Module_2_VBA_Challenge/blob/main/VBA_Challenge_2017.png)


In the refactored script, the execution time of the code was 0.148sec for 2017 stocks, and 0.141sec for 2018 stocks.

![Image of pop-up message showing elapsed run time of refactored script for 2017 stocks]


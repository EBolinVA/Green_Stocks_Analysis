# Analyzing Stocks for 2017 & 2018

## Overview of VBA Refactoring Challenge

A friend, Steve, asked for help analyzing a variety of green stocks for his parents. I started with an Excel file of the stock data he wants help analyzing. To get the most out of Excel, I used Visual Basic for Applications to automate some of the tasks instead of manually duplicating excel formulas over a very large spreadsheet. Along the way I learned several coding techniques, called macros or subroutines, to run the analysis more efficiently.

## Green Stocks Analysis Results

Steve provided an Excel file with two worksheets containing data for the 12 stocks. Each worksheet (named "2017" and "2018") had daily records over the course of a year for opening price, closing price, and volume traded.

### 2017 Results 

For the 12 stocks analyzed for 2017, all but one stock had a positive return. The best performing stocks were Daqo New Energy Corp (NYSE: DQ) with 199.4% return, and SolarEdge Technologies Inc (NASDAQ: SEDG) with 184.5% return. The lowest performing stock was TerraForm Power (NASDAQ: TERP) with a -7.2% return.

### 2018 Results

The story in 2018 was not as rosy. All but two stocks suffered losses. The only two stocks with positive returns in 2018 were Enphase Energy Inc (NASDAQ: ENPH) at 81.9% return, and Sunrun Inc (NASDAQ: RUN) with 84.0% return. Daqo New Energy Corp (DQ), the stock in which our client's parents were most interested, ended the year down 81.9%. Another loser was SolarEdge Technologies Inc (NASDAQ: SEDG) which lost 7.8% after their huge run in 2017.

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
```

### 2017 Coding Efficiencies

In the original script, the execution time of the code was 0.585 seconds for 2017 stocks. However, after refactoring the script, the code ran more efficiently in 0.148 seconds. 

Original VBA Code for 2017          |          Refactored Code for 2017
:-----------------------------------|-----------------------------------------:


![Image of pop-up message showing elapsed run time of original script for 2017 stocks](https://github.com/EBolinVA/Module_2_VBA_Challenge/blob/main/Original_time_2017.png) ![Image of pop-up message showing elapsed run time of refactored script for 2017 stocks](https://github.com/EBolinVA/Module_2_VBA_Challenge/blob/main/VBA_Challenge_2017.png)


Similar effects were observed when running the original and the refactored codes for 2018. 

Original VBA Code for 2018          |          Refactored Code for 2018
:-----------------------------------|-----------------------------------------:


![Image of pop-up message showing elapsed run time of original script for 2018 stocks](https://github.com/EBolinVA/Module_2_VBA_Challenge/blob/main/Original_time_2018.png) ![Image of pop-up message showing elapsed run time of refactored script for 2018 stocks](https://github.com/EBolinVA/Module_2_VBA_Challenge/blob/main/VBA_Challenge_2018.png)

## Summary: Pros and Cons of Refactoring

The advantages of refactoring code begin with the fact that you are not writing code from scratch. There may already some code out there which is a good start for the task. However, understanding and implementing more complex commands like nested for loops can significantly improve efficiency when running the code, especially over larger data sets. 

Another advantage to refactoring code is readability. As the code is easier to read and understand, there will be less energy put into maintaining the code in the future.

A disadvantage to refactoring code is that you could get yourself lost in the process. Once begun, it proved to be very time consuming. Each line of code must be checked and debugged before moving on to the next command. Will the runtime saved be worth the man hours it took to refactor?

In this project, the data analyst spent many hours reviewing manuals, writing and testing code, collaborating with colleagues, and calling the help desk in order to deliver a product that works for the client. The effort to refactor a coding project with only 12 stocks which gained less than half a second when running the refactored code does not seem worth the effort. However, this data analyst has ambitions for much larger projects with enormous sets of data. There is no con which can be applied to learning new skills. Refactoring code that saves the end user time when interacting with data may well be worth it for the long run. 
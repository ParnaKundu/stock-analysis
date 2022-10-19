# VBA of Wall Street

Using VBA to automate processing excel data

## Overview of Project

This project is to help Steve, a recent graduate with a Finance Degree, analyse the Stock Market and see the trends of several green Energy Stocks. With this analysis, he will be better equipped to advice his parents on which stocks to invest in.

### Purpose

The purpose of the project is to use VBA and automate the analysis of stock data. In addition refactoring of code is used to reduce execution time.

## Results

The original code used two "for loops" to iterate over the ticker array and the data rows to process and calculate 'total volume' and 'the return' as shown below :-

    'Loop through rows in the ticker array

    For i = 0 To 11
  
    ticker = tickers(i)
    totalVolume = 0
    'Loop through rows in the data
        
        Worksheets(yearValue).Activate
        For j = rowStart To rowEnd
            
           	'Get total volume for current ticker
           	If Cells(j, 1).Value = ticker Then
               	totalVolume = totalVolume + Cells(j, 8).Value
           	End If

In the refactored code only one "for loop" is used to process the data rows and the index of the ticker array is manupulated everytime a new ticker is being processed as shown below :-

	'Loop over all the rows in the spreadsheet.
    'Set tickerIndex to 0 before looping over the rows
    	tickerIndex = 0
    		    
    	For i = 2 To RowCount
    
        	'Increase volume for current ticker
        	tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        		.
			    .
			    .
			If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            	tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            	'Increase the tickerIndex.
            	tickerIndex = tickerIndex + 1
           			            
        	End If

This refactoring of code reduced the execution time significantly from **1.438 secs** using the original code [VBA_Challenge_2018](https://github.com/ParnaKundu/stock-analysis/blob/main/VBA_Challenge_2018.png) to **0.164 secs** using the refactored code [VBA_Challenge_2018_refactored](https://github.com/ParnaKundu/stock-analysis/blob/main/VBA_Challenge_2018_refactored.png) when analysing the 2018 stock data. 

In the same way, the execution time for 2017 stock data using the original code is **1.391 sec** [VBA_Challenge_2017](https://github.com/ParnaKundu/stock-analysis/blob/main/VBA_Challenge_2017.png) whereas it is **0.156 secs** in the refactored code [VBA_Challenge_2017_refactored](https://github.com/ParnaKundu/stock-analysis/blob/main/VBA_Challenge_2017_refactored.png).  

## Summary

1. What are the advantages or disadvantages of refactoring code?

    1. The advantages of refactoring are :-
	    - Improves the design of the software
	    - Makes the software easier to understand
	    - Helps find bugs
	    - Improves the run time of the software

    2. The disadvantage of refactoring is when done improperly it may introduce bugs to application. If the delivery schedule is tight, it may not be cost effective to refactor code.


2. How do these pros and cons apply to refactoring the original VBA script?

	Once the program was refactored, the run time reduced significantly. The use of a single 'for loop' as well as internal arrays for storing the calculated values and finally writting them in the excel helped reduce the run time. If care is not taken to handle the index of the arrays properly, it could lead to errorneous results.

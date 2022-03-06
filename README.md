# stocks-analysis

## Overview of Project
** Performance Improvement Analysis ** - Here we need to find out the time taken to generate requried output of a large stock dataset after refactoring the code.
 
### Purpose
 * Display Ticker wise total volume data byt processing the code with lease amount of time.
 * User should be able to enter the year for which he/she wants to view the data.


### Background
 * We need to provide a button for generating data
 * Ask user to input the year for which they need to generate data
 * Populate data in the "DQ Analysis" sheet for the selected year
 * Project File can be accessed here [Year based ticker volumne outcome - vba_challenge.xlsm]

## Results

### Analysis
 * For the year 2017 the ending CLOSE price is higher than the starting CLOSE price. Due to this the return is in profit for most of the tickers except TERP
 * TERP has a negative return .
 * Here is the snapshot of the data generated for 2017 [2017 DQ All Stock Analysis] ( )

 * For the year 2018 the ending CLOSE price lower than that starting CLOSE price for most of the tickers. 
 * That results in most of the ticker were in loss in 2018.
 * Here is the snapshot of the data generated for 2018 [2018 All Stock Analysis] ( ) 

### Challenges
 * Figuring out the startprice for each ticker within the loop had to be done by comparing previous value of the ticker rows 
 ***Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex)***
 * Figuring out the endprice for each ticker within the loop had to be done by comparing next value of the ticker rows 
 ***Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex)***
 * Identifying the code causing more time and refactoring it. Explained in Detail in Refactoring vba script(#Views on Refactoring the vba script)

 
    

## Summary
 ### Views on Refactoring in General
 ** Advantages **
	- Optimized code
 	- Better readability 
	- Reusable 
	- Improved Efficiency
 ** Disadvantages **
	- Time consuming

 ### Views on Refactoring the vba script 
  ** Advantages **
	- Instead of nested for loop, Initializing totalVolumn seperately before iterating through rows works faster. This is because when we initialize seperately it will execute only 11 times. Where as when its as part of nested loop, the overall iterations will be 11 * number of rows.
	- Switching between worksheets also increase the time - Best way to do is process the data from one worksheet (here the year worksheet) and store in the local array variables and then open the worksheet (All stock Analysis) to display the data 
	- Initializing Array prcoess faster as it knows the correct data type and no casting required.

 ** Disadvantages **
	- Additional time spent to optimize the code.
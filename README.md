# Stock Analysis

## Overview of Project

The purpose of our second challenge is to refactor, or rework, the code we constructed within the modules this week to be more efficient if it had to run with larger data files. Steve, our hypothetical client, wants a program that can run on larger datasets, or even the whole stock market. 

## Results

### Stock Market Analysis

The results of the stock anaysis itself are the same as in the module since we are using the same data set.  These selected "green stocks" had a strong year in 2017, with only one experiencing a negative return, and they had a bad year in 2018, with all but two of the stocks having a poor, sometimes terrifically bad year.   We should also note that there were two stocks, however, that had positive returns in both years.  "RUN" had a return of 5.5% in 2017 and 84.0% in 2018, so it had modestly good return in 2017 and amazingly good return in 2018.  "ENPH" performed extremely well in both years, with rates of return of 129.5% and 81.9% in 2017 and 2018 respectively.  If we were to recommend a single stock to purchase based only on this small data set of 12 stocks, the best stock is defintely "ENPH." At the same time, we should definitely caution Steve that the data indicates a large amount of volatility in these sorts of stocks in general, and we should always be careful before investing all our funds in a single place.  I would recomment a wider market analysis and more sorts of stocks to be included in their portfolio to balance out their return experience and insulate them a bit from the fluctuations inherent in investing in one single slice of the market. Below are the output tables with total volumes and annual rates of return for 2017 and 2018.

insert 2017 table
insert 2018 table

### VBA Code Analysis

The refactored code runs more slowly than the original code and I don't know why.  The original code has 2 for loops nested together, with the loop for iterating through the ticker values on the outside and the loop for iterating through the rows on the inside.  This means that in the original code, the program goes through all lines of the data for each ticker value for a total of 12 passes.  The refactored code has the two for loops separated, with an initial for loop to pass through the ticker volume array values to initialize them to zero. Then it has the longer for loop to collect the total volume and starting and ending prices as it iterated throgh the rows.  The if statements inside the loop that goes through the rows of data are similar in both situations, except that in the original code an if statement is evaluated in the inner loop each time through to decide if the ticker is the current ticker, whereas in the refactored code there is no if condition evaluated when the volume is incremented, it is handled through the array.  In the refactored code the  tickerIndex gets incremented in the final if statement if it is the last row of the current ticker. Maybe nested for loops run faster in VBA than running them in parallel, even though the refactored code looks at so mny fewer rows of data. Another possibility is that the use of arrays rather than single value variables slows down the code. In the original code we use single value variables to collect the data, while in the refactored code we use 12 element arrays to collect and store the values, so perhaps hanging on to all the additional data is using more memory and slowing down the code. 


## Summary

### Advantages/Disadvantages of Refactoring Code

### How Do these Pros and Cons Apply to Refactoring the Original Stock Analysis Script

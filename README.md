# Stock Analysis

## Overview of Project

The purpose of our second challenge is to refactor, or rework, the code we constructed within the modules this week to be more efficient if it had to run with larger data files. Steve, our hypothetical client, wants a program that can run on larger datasets, or even the whole stock market. 

## Results

### Stock Market Analysis

The results of the stock analysis itself are the same as in the module since we are using the same data set.  These selected "green stocks" had a strong year in 2017, with only one experiencing a negative return, and they had a bad year in 2018, with all but two of the stocks having a poor, sometimes terrifically bad year.   We should also note that there were two stocks, however, that had positive returns in both years.  "RUN" had a return of 5.5% in 2017 and 84.0% in 2018, so it had modestly good return in 2017 and amazingly good return in 2018.  "ENPH" performed extremely well in both years, with rates of return of 129.5% and 81.9% in 2017 and 2018 respectively.  If we were to recommend a single stock to purchase based only on this small data set of 12 stocks, the best stock is defintely "ENPH." At the same time, we should definitely caution Steve that the data indicates a large amount of volatility in these sorts of stocks in general, and we should always be careful before investing all our funds in a single place.  I would recomment a wider market analysis and more sorts of stocks to be included in their portfolio to balance out their return experience and insulate them a bit from the fluctuations inherent in investing in one single slice of the market. Below are the output tables with total volumes and annual rates of return for 2017 and 2018.

![2017 output table](https://github.com/mgsrichard/stock-analysis/blob/main/2017%20Data%20Output.png)
![2018 table](https://github.com/mgsrichard/stock-analysis/blob/main/2018%20Data%20Output.png)

### VBA Code Analysis

The refactored code runs in approximately 0.1 seconds for either year's data. The original code ran in about 0.7 seconds. Here are screenshots of the time reports from both the refactored code and the original code.


#### 2017 Code run time - Refactored code
![2017 refactored time](https://github.com/mgsrichard/stock-analysis/blob/main/VBA_Challenge_2017.png)
#### 2018 Code run time - Refactored code
![2018 refactored time](https://github.com/mgsrichard/stock-analysis/blob/main/VBA_Challenge_2018.png)
#### 2017 Code run time - Original code
![2017 original time](https://github.com/mgsrichard/stock-analysis/blob/main/2017%20original%20time.png)
#### 2018 Code run time - Original code
![2018 original time](https://github.com/mgsrichard/stock-analysis/blob/main/2018%20Original%20time%20new%20image.png)


The refactored code runs more quickly than the original code because un-nesting the for loops is so much more efficient. Instead of looping through all 3000+ lines of data for each of 12 tickers, the loop only goes through the data one time for a big savings in run time. Below are screenshots of the loops before and after refactoring. In the original code you can see how the for loops are nested.

#### Original code for loop
![original loop picture 1](https://github.com/mgsrichard/stock-analysis/blob/main/Original%20Code%20nested%20loop%201.png)
![original loop picture 2](https://github.com/mgsrichard/stock-analysis/blob/main/Orig_code_nested_loop2.png)

The refactored code has the two for loops separated, with an initial for loop to pass through the ticker volume array values to initialize them to zero. Then it has the longer for loop to collect the total volume and starting and ending prices as it iterates through the rows.  

#### Refactored code for loop
![refactored loop picture 1](https://github.com/mgsrichard/stock-analysis/blob/main/Refactored%20code%20loops%201.png)
![refactored loop picture 2](https://github.com/mgsrichard/stock-analysis/blob/main/Refactored%20Code%20Loops%202.png)


The if statements inside the loop that goes through the rows of data are similar in both situations, except that in the original code an if statement is evaluated in the inner loop each time through to decide if the ticker is the current ticker, whereas in the refactored code there is no if condition evaluated when the volume is incremented.  In the refactored code the  tickerIndex gets the current row volume amount added in the final if statement if it is the last row of the current ticker. In the original code we use single value variables to collect the data, while in the refactored code we use 12 element arrays to collect and store the values. Here you can see the way that the variables were declared in the refactored code and in the original code.

#### Refactored code array declarations
![refactored array declarations](https://github.com/mgsrichard/stock-analysis/blob/main/refactored%20code%20array%20declarations.png)

#### Original code variable declarations
![original variable declarations](https://github.com/mgsrichard/stock-analysis/blob/main/original%20code%20variable%20declarations.png)


### The code will not work for the whole market

As I worked through the coding assignment, I kept getting stuck on the idea from the description that Steve would like to have some code that would work for the whole market.  Because of the way the ticker names array is hard coded, it wouldn't work for the whole market. At least, if more stocks were included in the file, they would not turn up in the data output because the output loop is going to run 12 times and no more.  So, if we did want to include more data, we would have to go in and make all the arrays have more elements and hard code in all the new ticker names and adjust how many times the output loop and formatting loop run.  I considered trying to program this functionality into the code. If we put two new loops at the top, we could count how many tickers we have in the first loop and then capture all the ticker names in a second loop. After that, we would have to declare the arrays differently, with the number of tickers we just counted instead of the original 12, and make a few more tweaks through the code to make it all run smoothly.  I started doing this, but ran into a problem with the array declarations, where it seems like VBA doesn't like you to put a variable in an array declaration, and will only accept constants. So, I decided that this was probably outside the scope of this project, and that while I really wished I could have made it work, I was more interested in starting into Python than spending more time in VBA. 

### The code will only work for sorted files

The code as written will only work for files that are sorted by ticker and then by date. Adding this functionality into the code would be an easy fix, and possibly save the user from some errors.  Either way, with a macro like this, you have to be sure the user you are preparing it for knows exactly what kind of data they should use as input. 


## Summary

### Advantages/Disadvantages of Refactoring Code

The advantage of refactoring code is that you don't have to create everything again from scratch. You don't have to figure out all the logic and steps from nothing, and you avoid lots of the small bugs and typos you would get when you create something for the first time.  There's no reason to spend a long time imagining new code when you can adapt something that's already out there. The disadvantages of refactoring old code are that you will have to be very careful to understand exactly what it's doing, and make sure you make all the variable name changes and tweaks it will take to make it do precisely what you want it to do. You also have to be careful about getting the data input portion of it right, since you may or may not have a data set that looks like what the old code is looking for.  The code you are adapting probably was not written for your exact circumstance, so you need to make sure you have considered all the similarities and differences between old and new code.

### How Do these Pros and Cons Apply to Refactoring the Original Stock Analysis Script?

In this project, because we were using the exact same data and answering the same questions, refactoring meant that most of the work was already done.  We tweaked some variables and the way the loops were running, but most of the body of the code was already there.  Almost none of the disadvantages of reworking code apply here, as the input was identical, the processing was slightly modified, and the output only had to be modified to use the new arrays instead of the single value variables. The refactored code here makes a major improvement over the original in the amount of time it takes to run.  It is a better, more efficient way to process bigger files. The original code runs through all the data for each ticker, doing 12 times as much looking at the data file.  The new code accomplishes all the same data analysus as the original code with 1/12 as many passes through the data.

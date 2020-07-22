# VBA of Wall Street

## Overview of Project

### Purpose
I wrote a VBA script analyzing stock data and created buttons that run analysis and output the total daily volume and return of various companies. By refactoring my code I was able to create a more efficient version to help the customer analyze large sets of data in less time.

## Results
- Running analysis for all stocks in the year 2018 showed that only two companies, ENPH and RUN, had positive returns. DAQO's closing price decreased the most by about 63% from the beginning of the year to the end of the year. In the 'Total Daily Volume' column, it can be seen that DQ stock was traded the most and AY stock was traded the least. The refactored VBA script ran in about 0.19 seconds, while the original script took about 1.42 seconds.

![](https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2018_Output.png)

![](https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2018.png)

- In the year 2017 all companies except for TERP had a positive return. DAQO's closing price increased the most by about 199% from the beginning to the end of the year. Here 'Total Daily Volume' column tells us that SPWR stock was traded the most while DQ stock was traded the least. The refactored VBA script ran in about 0.18 seconds and the original code took about 1.42 seconds.

![](https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2017_Output.png)

![](https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2017.png)

## Summary

### Advantages and Disadvantages of Refactoring Code
- In general, refactoring code is a useful way of increasing efficiency and decreasing run time. It can also help make code easier to read for futuer users. A disadvantage of refactoring is the possibility of having to spend time debugging new errors that occur due to the changes in your code.

- For the data at hand, the refactored code decreased the runtime by about 1.23 seconds. Although this might not seem like a big improvement, if we ran the same program on a much larger data set then we would be saving ourselves a lot of time. One advantage of the original VBA script is it only uses one array and has nested for-loops. This makes the code easier to read and understand for some programmers, but with detailed and descriptive commenting this can also be avoided. 

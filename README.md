# VBA of Wall Street

## Overview of Project

### Purpose
Steve is looking into green energy stocks to help his parents diversify their investments and needs help analyzing the data. I wrote a VBA script to analyze the data and created buttons that run analytics and output the total daily volume and return of various companies. By refactoring my code I was able to create a more efficient version to help my customer analyze large sets of data in less time.

## Results

### 2018
- Running analysis for all stocks in the year 2018 showed that only ENPH and RUN had positive returns. Daqo's closing price decreased the most by 62.6% from the beginning of the year to the end of the year. In the 'Total Daily Volume' column, it can be seen that ENPH stock was traded the most and AY stock was traded the least. The refactored VBA script ran in about 0.19 seconds, while the original script took about 1.42 seconds.

<img src="https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2018_Output.png" width="350">

<img src="https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2018.png" width="350">

### 2017
- In the year 2017 all stocks except for TERP had positive returns. Daqo's closing price increased the most by 199.4% from the beginning to the end of the year. Here 'Total Daily Volume' column tells us that SPWR stock was traded the most while DQ stock was traded the least. The refactored VBA script ran in about 0.18 seconds and the original code took about 1.42 seconds.

<img src="https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2017_Output.png" width="350">

<img src="https://github.com/npantfoerder/stock-analysis/blob/master/resources/VBA_Challenge_2017.png" width="350">

### Conclusions
- In 2017 almost all of the 12 stocks had positive returns and DQ's closing price increased the most. However, DQ stock was also traded the least. Then in 2018 almost all of the stocks had negative returns and DQ's closing price decreased the most. Based on this and the fact that there was no clear company whose stock performed the best during both years, this highly suggests that Steve's parents should diversify their investments. 

## Summary

### Advantages and Disadvantages of Refactoring Code
- In general, refactoring code is a useful way of increasing efficiency and decreasing run time. It can also help make code easier to read for future users. A disadvantage of refactoring is the possibility of having to spend time debugging new errors that are created due to the changes you make in your code.

- For the data at hand, the refactored code decreased the runtime by about 1.23 seconds. Although this might not seem like a big improvement, if we ran the same program on a much larger data set then we would be saving ourselves a lot of time. Meanwhile, one advantage of the original VBA script is it only uses one array and has nested for-loops. This makes the code easier to read and understand for some programmers, but with detailed and descriptive commenting this can also be avoided. 

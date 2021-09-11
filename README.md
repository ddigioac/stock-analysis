# Stock Analysis

## Overview of Project

### Purpose

In this analysis I will compare the annual stock performances for 2017 and 2018, as well as examine execution times of the original code compared to my refactored script.

## Results

### Stock Performance 

To begin my analysis of the annual stock performances for 2017 and 2018, I first had to create a script that would comb through our dataset and determine the starting prices, ending prices and daily volumes for the stocks we were analyzing. In order to calculate the returns for the year I had to take the 'Ending Price' divided by the 'Starting Price' and subtract one, for each stock for that year. 

![This is an image](https://github.com/ddigioac/stock-analysis/blob/9903aa60937798ccb83fa2fdf5970d6dc56b0ffb/Return%20_Calc.png)

The entire set of stocks I analyzed in 2017, except for one, had a positive return. However, in 2018 each stock, except for two, had a negative return. The only two stocks that had a positive return for both 2017 and 2018 were ENPH and RUN. We can conclude that the overall market in 2017 outperformed 2018.

![This is an image](https://github.com/ddigioac/stock-analysis/blob/9648dcd084f16c169d893cc8f9e47e587c12a720/All_Stocks_%202017.png)
![This is an image](https://github.com/ddigioac/stock-analysis/blob/30c5b353ce1be96ab0df3e3e4ce8a35812dba09d/All_Stocks_2018.png)

### Execution Times of The Original Code vs. Refactored Script

By refactoring my script I saw a huge reduction in run time for Excel to execute my code. Originally my code took approximately 0.3 seconds to complete. Once I refactored the code to run more efficiently my script took approximately 0.09 seconds to execute. That equates to approximately 70% reduction in run time.

**Original Code Run Times** 

![This is an image](https://github.com/ddigioac/stock-analysis/blob/68f13cbcff8714b55e8c8330bdb5215ff20a61d9/OriginalRT_2017.png)
![This is an image](https://github.com/ddigioac/stock-analysis/blob/0d6564dc999f6931092876b42ea817eae0dcffaa/OriginalRT_2018.png)

**Refactored Script Run Times**
![This is an image]()
![This is an image]()
## Summary

**1. What are the advantages or disadvantages of refactoring code?**

The advantages of code refactoring are that one can increase the flexibility of the code by adding other functions to it, ultimately increasing the scripts overall capabilities. Also, one can decrease the complexity of the code, allowing others to understand and read it better.

The disadvantages of code refactoring are that it can be time consuming and cumbersome. One also increases the chances of making a mistake, depending on the complexity of the code.

**2.How do these pros and cons apply to refactoring the original VBA script?**



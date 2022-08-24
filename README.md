# Stock analysis with VBA

## Overview and purpose of the project

We need to help Steve analyse all the green energy stock data in addition to DAQO New Energy Corp to help his parents invest in green energy.
Using Excel is time consuming. 
Here we will use VBA to code to automate the tasks and we can also reuse the code to perform analysis on the stock data.

## Results and Analysis

Steve's parents are keen to invest in DAQO New Energy Corp so we will first analyse its stock performance for the years 2017 and 2018.
Now this DAQO New Energy Corp is indentified in the stock data with Ticker 'DQ'.
Using VBA, we will calculate the total daily volume by summing up the daliy volume of DQ for each year to know how actively DAQO was traded.
We will also calculate the yearly return for DQ to see if investment might grow by the end of the year.

We see that DQ traded well in 2017 and not in 2018. The yearly return of DQ dropped by 63% in 2018 which is not a good option for investment.

![image](https://user-images.githubusercontent.com/111020934/186285642-a13f51ab-3f80-4b52-b4d9-8aa3b9cdab04.png)
![image](https://user-images.githubusercontent.com/111020934/186285819-85070e9b-a1a5-4443-a485-54d1e507dd36.png)


In order to provide a better option, we will now analyse all the stocks in green energy.
We will loop through all the Tickers in the stock data to calculate the total daily volume and yearly return of all available Tickers for the years 
2017 and 2018.

First we run our analysis for the year 2017:

![image](https://user-images.githubusercontent.com/111020934/186286351-ace62da2-a348-434e-83fd-120e69fe9511.png)

The Tickers DQ,ENPH,SEDG and FSLR have been actively traded for the year 2017.

Next we run our analysis for the year 2018:

![image](https://user-images.githubusercontent.com/111020934/186286552-b4014076-7248-4413-97c7-680481be50e1.png)

The Tickers ENPH and RUN have been actively traded for the year 2018.

Based on this analysis we can say that ENPH is a safe option for Steve's parents to invest in order to see a growth.
Also RUN is a good option since there has been an increase in its yearly return for 2018. 
Rest of the stocks have performed poorly for the year 2018.

The execution times of the original script and the refactored script for year '2017':

The refactored code has less exceution time.

The execution of the original code.

![image](https://user-images.githubusercontent.com/111020934/186290778-5dddaaed-9ae2-438f-8663-c50dab12f708.png)
![image](https://user-images.githubusercontent.com/111020934/186290895-a264ffbb-4229-44b7-85e1-9347c5afdcf6.png)


The execution time of the refactored code.

![image](https://user-images.githubusercontent.com/111020934/186286956-f9d13315-ab5d-4106-92f5-74ca8624d609.png)
![image](https://user-images.githubusercontent.com/111020934/186287003-059e8d16-8604-44a3-a003-830af64faf9c.png)


## Summary

1. What are the advantages or disadvantages of refactoring code?

Advantages : Code refactoring is restructuring the source code by improving its structure without changing its functionality.
Refactoring makes the code clean and organized thereby improving the quality of the code providing a better understanding.
It helps in reducing errors in code.
Disadvantages : Stable code should not be refactored if not done early.
It is time consuming and hence cannot be done if the delivery schedule is tight.

2. How do these pros and cons apply to refactoring the original VBA script?

It is the easy way to rearrange, restructure, and clarify existing code at the same time even not changing the behavior of your code.
It also provides better outcomes of code performance.




 

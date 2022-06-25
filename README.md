# VBA Stock Challenge

## Overview of Analysis
### Background
#### This task was to use VBA along with excel to analyze stock yearly return and volume from the years 2017 and 2018. The code was written in VBA using the skills learned in Module 2. Those skills include arrays,variables, loops and nested loops, conditionals, and formatting. With those skills an analysis was produced to help Steve and his parents.
### Purpose
#### The purpose of this task was to create a workbook for Steve to analyze all stocks by tickers to help diversify his parents portfoilo. This analysis was to help determine if the DAQO (DQ) green stock was the best stock for Steve's parents.  The orignial code made to handle just the one stock. The code was then refactored to handle any stock added and at a faster speed. 

## Results
#### The original code worked well for the DQ stoce Steve was looking into. It gave him the information he needed to determine if this stock was right for his parents. Unfortunately,  the code ran slower when Steve asked us to include the other stock tickers. 

### ![Green_stock_2017](https://user-images.githubusercontent.com/105830665/175751771-f1080804-d4a5-4372-b62e-d8c0614c167a.jpg)
### ![Green_Stock_2018](https://user-images.githubusercontent.com/105830665/175751807-20d2f01e-c806-49ee-9132-0c92dde8c1c9.jpg) 

#### Here both photos you see that the original code took about 0.95 seconds to run each year.Once the code was refactored, the run time was cut to about 0.17 seconds for each year.

### ![VBA_Challenge_2017](https://user-images.githubusercontent.com/105830665/175751822-82d2567d-5444-4783-afb4-b2c2acc82fb9.png) 
### ![VBA_Challenge_2018](https://user-images.githubusercontent.com/105830665/175751835-99c40421-1918-4224-b620-659d267dd28a.png)

#### Based on the reports provided from both codes, results were the same, the Stocks in 2017 had a better return than in 2018. Steve does not want to let his parents solely invest in ticker DQ.  

## Summary
### Advantage
#### The refactored code is able to speed up the run time which will allow this code to analyze more stocks at one time. 
#### One advantae from the original code was the nested loops. As you stepped through the code using f8 you were able to follow along a little easier. This was because the code walked through every loop with you. 
### Disadvantage
#### The refactored code took longer to write because you needed to understand how arrays worked and the efficiency of the arrays. 
#### The obvious disadvantage of the original code was that it would slow down even more if you added more than the 12 tickers we had. 

# Stock Analysis - Module 2 : VBA Challenge

## Overview of Project

### Background
Steve's parents are interested in 'Green' or enviromentally friendly Stocks. We've created a vba code to review multiple tickers over multiple years using a AllStocksAnalysis module with VBA code. However, the code is not efficient and can be improved. This way if Steve would like to review a large number of stocks, the module will use less memory and still work without crashing. 

### Purpose
The purpose of this exercise is to use more efficient code. In using less steps and more abstract logic, code can become more efficient, use less memory, and decrease run time.

## Results & Analysis
In order to measure the performance of our two different approaches, we institute a timer to record the amount of time that elapses from the modules start to end. 


### Analysis of Performance: AllStockAnalysis() Module

The original approach using the AllStockAnalysis() Module, iterates over each cell to determine its Volume, StartingPrice and EndingPrice. Then uses a for loop to iterate over the list for each cell, repeating the code. 

The origial approach runs in 0.296875 seconds. While this seems fast, our Refactored analysis shows there is room for improvement. 

![Theater_Outcomes_vs_Launch](VBA_Challenge_2018_AllStocks.png "VBA_Challenge_2018_AllStocks")

### Analysis of Performance: AllStockAnalysisRefactored() Module

The improved approach using the AllStockAnalysisRefactored() Module instead stores the three variables (Volume, StartingPrice, and EndingPrice) into an array. By doing so we are able to create tickers within for loops, allowing us to perform calculations within loops. This makes the code more efficient and fully utilizes the object oriented tools provided in VBA. 

While the origial approach runs in 0.296875 seconds. The refactored approach runs in 0.09375 seconds, meaning the code runs more than 3 times faster than the original code. This is a signicant increase on efficiency. 


![Outcomes_vs_Goals](VBA_Challenge_2018_AllStocksRefactored.png "VBA_Challenge_2018_AllStocksRefactored")


### Summary

A challenge i

Source 1: See Link below for 

## Results

- What are 
    
   * May would 

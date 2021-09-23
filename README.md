# Stock Analysis by Year

## Overview of Project

Using VBA macros within Microsoft Excel, yearly summaries of daily volume and return percentage was calculated for each stock ticker. The goal of this project was to pull the appropriate data by creating a funtioning macro, then refactor it to run more efficiently should the data set increase in the future. Data used in the calculations was gathered over 2 years and split into tabs - 2017 and 2018. Details on the data set can be found in the following document. 

[VBA_Challenge.xlsx](https://github.com/brefrank/stock-analysis/files/7220227/VBA_Challenge.xlsx)


### Analysis Results

As shown below, the majority of stocks in our 2017 data set had a positive return while 2018 showed an overall negative return. By comparing return percentages to daily volumes, there does not seem to be a correlation between dollar amounts and success rate. If we were to use this data to determine which stock to invest in, there is only one option: ticker ENPH had the only positive return percentage over both years. Ideally, however, more data would be gathered and compared to find a more conclusive trend. 

<img width="147" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/90646961/134570764-9728db8d-8f1f-4d59-8186-ef8eeb371060.png"> <img width="147" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/90646961/134573985-6e743ede-c8d3-4e58-8bd7-2478a6a74a7d.png">


### Refactoring Results

The second portion of this project was to refactor our originally written macro code for efficiency. This goal was accomplished and can be seen below. The time it took to run our originally written code (photos 1) was roughly ten times longer than the refactored code(photo 2). 

1. <img width="157" alt="VBA_Original_2017_Time" src="https://user-images.githubusercontent.com/90646961/134580576-8822ca63-3885-4322-8ce4-6a524693721b.png"> <img width="162" alt="VBA_Original_2018_Time" src="https://user-images.githubusercontent.com/90646961/134580610-c0a8bbb7-655a-4c2c-9f53-e4c99ea45c2a.png">
2. <img width="162" alt="VBA_Challenge_2017_Time" src="https://user-images.githubusercontent.com/90646961/134580597-4d040974-a3a9-4694-9f2a-81de49d65a1b.png"> <img width="160" alt="VBA_Challenge_2018_Time" src="https://user-images.githubusercontent.com/90646961/134580615-618f043d-eee6-470c-be74-a6b856282837.png">

Both macro versions have been provided below for reference.

Original code draft: [Original_Macro.txt](https://github.com/brefrank/stock-analysis/files/7220729/Original_Macro.txt)
Refactored result: [VBA_Challenge_Macro.txt](https://github.com/brefrank/stock-analysis/files/7220272/VBA_Challenge_Macro.txt)

## Summary

There is huge benefit in the process of refactoring, from simplifying code to catching mistakes. A lot of brain power goes into writing code from scratch in order to obtain an intended result, and most people will type as they think things through. Because of this, code can come out functional without necessarily being fast or entirely accurate. Reviewing and refactoring code on your own or sending it to a peer is like a double check or second draft that can save time and money as a data set grows. 

Although unlikely, it is important to keep in mind that it is possible to worsen a code by refactoring. Potential outcomes could be adding too many insignificant steps by overthinking or making a change that creates a time consuming bug for all subsequent parties. For this reason, it is often best to get a second check by a coworker or peer.

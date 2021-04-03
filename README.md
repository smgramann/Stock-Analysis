# Stock Market Analysis with VBA

## Overview of Project

### Background
Steve was provided a workbook with VBA code to analyze a specific dataset of stock market tickers and he now sees value in expanding the analysis to the entire stock market.  The most efficient VBA code needs to be written to quickly accomplish this larger analysis.

### Purpose
Initial VBA code was written to analyze a specific dataset; however, that code may take too long to run for the entire stock market so it was refactored to increase efficiency.  The starter code was re-written as one macro to loop through the dataset one time.  In addition, a tickerIndex variable was created to avoid having a nested loop function and specific conditional code; therefore, also making the code more efficient.

## Results

### Refactoring of Code
The VBA code created to analyze a specific set of tickers included multiple macros to accomplish the task.  While that worked for the dataset focused on 12 tickers, this would likely run too slow if all stocks were required to be analyzed in the dataset as there would be more tickers and a much larger dataset.  To increase efficiency, the analysis is now run through one macro for a full stock market analysis, when needed.
The second add to this refactored code is the tickerIndex variable.  The add of the tickerIndex variable allows for one loop through the code when the module code required a nested loop and, therefore, multiple loops through the data to obtain the needed values.  This also resulted in no need for a conditional statement for gathering up tickerVolumes for the final output.

Visual 1 highlights these changes:
![Visual 1 - tickerIndex and no nested loop](https://user-images.githubusercontent.com/80165223/113482534-b9d02b00-9464-11eb-81a9-7695c582f24a.png)


### Analysis of Code Performance
The refactoring of code resulted in better run time performance between the code written for analyzing 12 tickers (module) vs the entire stock market (challenge).  Before refactoring the code, the run times for 2017 and 2018 were .784 and .624, respectively.  Below are the refactored code run times:

2017 refactored code run time:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/80165223/113482554-c6548380-9464-11eb-847b-107406692ee4.png)


2018 refactored code run time:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/80165223/113482563-cb193780-9464-11eb-977d-82644549e963.png)


### Stock Market Performance
In reviewing the stock market performance for the 12 identified tickers from 2017 to 2018, most stocks saw better returns, in general, for 2017 (with the exception of TERP). In 2018, most tickers experienced a decrease in returns except ENPH and RUN.  Given that ENPH and RUN tickers saw gains in both years, I’d recommend investing in those.

## Summary

### Advantages and Disadvantages of Refactoring Code
The purpose of refactoring code is to increase efficiency for the program to run, which is a clear advantage.  Personally, an advantage to refactoring the code is an increase in learning and retention of the VBA module course material.  This comes with some challenges and disadvantages, though.  First, debugging code can be a barrier to refactoring code quickly.  Secondly, while the code is more efficient, after refactoring it may be built for a very specific purpose and have less flexibility for gathering other data.  Depending on the goals of the analysis, especially as more questions arise, more code will need to be written.

Specifically related to the purposes of this stock market analysis, Steve needed to provide some basic answers to some stock market questions quickly.  The refactoring of the code will likely achieve that as there is one macro to achieve the work and there is no longer a nested loop.  The inclusion of the tickerIndex variable also allowed the macro to run more efficiently.  One clear disadvantage and an assumption that was made when writing the refactored code is that the tickers are organized together in the source data.  If they were randomly assigned over all the rows in the sheet, the conditional statements would no longer work and inaccurate results would be provided.

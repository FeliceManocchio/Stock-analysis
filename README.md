# Stock-analysis
Excel VBA for analyzing stock data

## Overview & Purpose

### Overview
In this Module we set out to determine the performance of 12 different stocks and how they performed in the years 2017 and 2018. The code we set prior to this challenge set out to go over the data entirely for each stock referenced and report to us the volume of the stock traded and the yearly return each stock produced over the course of the target year. 

### Purpose
Our purpose for Refactoring this macro that already runs perfect is to increase the efficiency of the macro run in our dataset for when we go to look at a larger pool of data say thousands of stocks as oppose to a mere 12. Prior to this module the macro ran the analysis running through the entire sheet over for each individual stock we set to analyze. Our goal in the module is to have the data run in a single run through and produce the same results with much more speed and efficiency.

## Results

### Analysis
By refactoring the code that we previously had running the stock analysis we decreased the runtime tenfold. Our new macro can run through larger data sets in less than half the time allowing for a larger analysis pool of thousands of stocks instead of just 12 and produce an accurate volume and stock yearly return over the year being analyzed. 

### Analysis of Data
In looking at the results of the stocks analyzed in the data set, we can see that most of the stocks under performed in 2018 not just the DQ stock which is what the client(steve) was originally interested in discovering. In the previous year (2017) almost all the stocks had an excellent return percentage which can bring us to infer that these stocks could have a type of industry connection where when the industry itself prospers so do the stocks presented. There is only one stock that had performed poorly in both years but even then it was a far less under performance than the others. the two stocks that prospered in 2018 whent he others had declined, could have a hand in an additional industry which offset their losses that the other stocks experienced, or have a superior market product or better company financial practices. 

## Summary

### Advantages and Disadvantages of Refactored VBA
One of the major advantages of refactoring code is that it would greatly increase the efficienc of how a macro runs if done properly. The output will remain the same but by properly refactoring code you can increase the data set size and reduce run speed.Some of the major disadvantages of refactoring code is depending on how complex you look to refactor the code there are higher chances of creating new errors that mmght have previously not been in the original coding. 

### Avantages and Disadvantage of Original Code vs refactored code
One of the advantages that are present for the original code is that it is simpler to write, and by that I mean there are fewer variable assignments, and has fewer lines writtenout to perfom the task at hand. The biggest disadvantage that this code also presents with is the runtime and execution of the process for the code presently. The original code loops through the entire data set for every new ticker presented making it go through over 3000 rows of data for all 12 tickers. This is where the advantage of refactoring chimes in making the code much more efficient. The refactored code runs through the entire data sheet once and produces the same values as the original cod did in a fraction of the time. The biggest disadvantage working on the refactoring is there was plenty of opportunity to introduce new bugs that were not present in the original code such as overflow errors which was the most common issue that arose while writing the new refactored macro. Overall if this analysis where to be taken to aa much larger data set over 100 or even 1000 stocks to analyze the refactored macro would be the best bet to get a relatively quick answer as oppose to the original code which could potentially take an hour or more to get the desired output. 


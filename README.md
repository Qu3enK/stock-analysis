# Stock Analysis

## Overview of Project

### Background
The Green Energy Stock effort is designed to assist Steve with analyzing stock data in order to diversify his parent's portfolio.  Steve is very happy with the initial workbook I prepared for him and wants to expand the dataset beyond the 12 stocks that were analyzed to include the entire stock market over the last few years. As part of the effort, I will need to refactor my code and ensure it executes in a prompt and timely fashion. [^1]

### Purpose
The purpose of the this analysis is to edit or refactor the Excel VBA code utilized in Module 2 to loop through all the data one time in order to collect and analyze the stock information.  Then determine whether refactoring my code made the VBA script run faster. [^2]

## Results  
*Note: The instructions stated to >Compare the stock performance between 2018 and 2018  as well as the execution times of the original script and the refactored script.*

To perform the comparison of the original and refactored Stocks portfolio as it relates to the execution time, portions of the original VBA code from Module 2 were refactored.  Per the requirements provided the following were refactored:
    - tickerIndex set to 0 [^a]
    - tickers, tickerVolumes, tickerStartingPrices & tickerEndingPrices arrays created [^b]
    - looping through the stock data to read and store the tickers, tickerVolumes, tickerStartingPrices & tickerEndingPrices row values [^c]

- INSERT Pictures - 

### Performance Results
After running the refactored code for 2018 stock analysis I confirmed that the output matched the 2018 snapshot provided in the challenge instructions.

**2018 Stocks Analysis- refactored code**
- INSERT PICTURES


### Execution Time Results
The execution time between the original and refactored 2018 Stock portfolio was approximately 1/100th of a second difference.

**Execution Time for 2018 All Stocks Analysis** (*based on original code*)

![VBA_Challenge_2018 original](https://user-images.githubusercontent.com/112449480/191653629-8f914490-b9bc-4ea3-9e40-749f07d732d1.png)


**Execution Time for 2018 Stocks VBA Challenge** (*based on refactored code*)


![VBA_Challenge_2018 refactored](https://user-images.githubusercontent.com/112449480/191653799-9ffba5fd-1485-44a6-8f91-1e39425554b2.png)


## Summary

- What are the advantages or disadvantages of refactoring code?

__Advantages__
- more efficient code
- improves the logic, easier to read
- shortens the development time
- easier maintenance
- ability to apply and/or adapt it to similiar projects

__Disadvantages__
- if the task is short and simple then refactoring code may not be practical and it may prove best to write a basic code
- if the code is large then may lead to more bugs and errors
- if you are working on code that someone else built then may need to adjust or rewrite the code which decreasing productivity time. In other words, increase development time

__How do these pros and cons apply to refactoring the original VBA script?__

Based on the fact that the original code included comments, I was able to utilize that information and the script to refactor the code for the challenge. I was able to align the original code to the ask for the refactored code to perform the analysis.

A con to refactoring the original VBA code is running the risk of misspelling a word(s) or not properly assign a varible which causes errors.




[^1]: Utilize some of description from Module 2 work to assist in writing my background for the Challenge 
[^2]: Utilize some of description from Module 2 Challenge- Background to assist with writing my purpose for the Challenge

# Stock Analysis using VBA

## Overview of Project
The main goal of this project is to analyze stock data of 11 different companies from the years 2017 and 2018 using VBA as a developing tool. We also compare time performance between original scripts and refactored code.

## Results
Stock Performance 2017

![VBA_Challenge_2017](https://user-images.githubusercontent.com/20058842/173478821-645f3504-def3-4b71-afb2-672c6aa313cb.png)


As seen above, we we can see that only "TERP" had a negavative retrun for the year 2017. The execution time for original code was 1.242188 seconds, while refactored code ran in 0.1484375 seconds.

Stock Performance 2018

![VBA_Challenge_2018](https://user-images.githubusercontent.com/20058842/173478641-1ef8ba7a-e53d-4a31-b057-0acdf65ddc33.png)


As seen above, only tickers "RUN" and "ENPH" showed a postive return for the year 2018. The execution time for original code was 1.210938 seconds while refactored code ran in 0.1484375 seconds.

It is clear that refactoring the original code evolved in a much more efficent and fast result. The refactored code utilized two seperate for loops as opposed to the original script which had 2 nested for-loops. On the original nested for loop script we used an inner loop that iterated 3012 rows of data with an outer loop of 11 tickers which would mean we did 3012 x 11 = 33132 iterations. On the refactored script we used one loop that iterated 11 times to initialize a ticker volume array and one final loop to fill that iterated the 3012 rows of data to calculate the ticker volumes and starting and ending prices that would then be used to calculate the final year outcome results. This means the refactored code ran 3023 iterations of code opposed to 33132 iterations of the original code! This clearly reduced computation time to obtain the same results.

## Summary

### Advantages of refactoring code
Reusing code avoids having to start from scratch saving valuable development time during a new project. Making changes to code and adjusting it to project needs can be helpful and contribute to existing problems that have been solved in the past. Code refactoring can improve the execution time of scripts, which was the case on this project.

### Disadvantages of refactoring code
It can sometimes be difficult to understand written code rather than making your own. If reused code isn't well documented, it can be painful and time consuming to adjust a script to your needs.


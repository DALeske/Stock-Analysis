# Module 2 Challenge: Refactoring VBA Code

## Overview of Project

The purpose of this analysis was to determine whether refactoring VBA code can improve the efficiency of VBA code. In this particular instance, the goal was to reduce the time required to calculate the performance of 12 sample stocks for the years 2017 and 2018. Execution times for the original VBA script (initially created as part of an introduction to VBA coding) and the refactored VBA script were compared.

## Results
Refactoring significantly reduced the execution times to calculate the performance of 12 sample stocks for both the years 2017 and 2018. For 2017, execution times were reduced from 0.63 seconds to 0.07 seconds (89% reduction). For 2018, execution times were similarly reduced, from 0.64 seconds to 0.07 seconds (89% reduction). For both 2017 and 2018, calculated total daily volume and returns were identical for both the original VBA script and the refactored VBA script.

Figure 1: 2017 Stocks Analysis: Original VBA Script

![2017 Stocks Analysis: Original VBA Script](Resources/GreenStocks_2017.png)

Figure 2: 2017 Stocks Analysis: Refactored VBA Script

![2017 Stocks Analysis: Refactored VBA Script](Resources/VBA_Challenge_2017.png)


Figure 3: 2018 Stocks Analysis: Original VBA Script

![2018 Stocks Analysis: Original VBA Script](Resources/GreenStocks_2018.png)

Figure 4: 2018 Stocks Analysis: Refactored VBA Script

![2018 Stocks Analysis: Refactored VBA Script](Resources/VBA_Challenge_2018.png)

## Summary
Refactoring the VBA code reduced execution times while providing the same results. 

### What are the advantages or disadvantages of refactoring code?
The advantage of refactoring code is that there is the potential to make the code run more efficiently. For small datasets, the increase in efficiency may not be noticeable, but for large datasets, improving efficiency of the code can greatly reduce the amount of time required to execute the code. Depending on your computer's hardware configuration, it is possible that refactoring may make the difference between running out of memory or not running out of memory when analyzing large sets of data. 

The disadvantage of refactoring code is that it is possible that you may actually slow the execution time.  Readability of the code may also become more difficult.  In addition, it's quite possible that the original programmer wrote the code specifically to handle unique challenges in their data, and refactoring the code may introduce unintended problems that may not readily become apparent. There is also the disadvantage of time required to refactor the code.  If it's not broke, don't fix it. 

### How do the pros and cons apply to refactoring the original VBA script?
By refactoring the original VBA script, I was able to reduce the execution times significantly, which was the main advantage I was trying to achieve. The process was not without its challenges. As I was refactoring, I reached a point where I was only getting the appropriate result for the first stock, and all subsequent stocks had a total daily volume of zero and an error for the return.  Through a series of troubleshooting steps and temporary message boxes, I was able to determine that I had a typo in a variable name at one referenced location ("tickersIndex" rather than "tickerIndex"). Once found and corrected, the script ran fine. 

Another challenge that I encountered with refactoring was with readability. While the nested loops used in the original script made it easier to visualize how the code was running, the refactored code was more difficult for me to understand. In particular, the purpose of incrementing the tickerIndex after each loop was not immediately evident to me. Once I understood how the refactored code worked, it made sense why the refactored code would run faster than original (required one single loop to pass through all stocks rather than requiring a series of inner loops for each pass of the outer loop)

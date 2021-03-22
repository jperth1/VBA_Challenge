# VBA_Challenge
Data Analysis of stocks using VBA
## Project Overview
The following analysis using VBA to evaluate stock data from 2017 and 2018 and highlight the total daily volume and yearly return for each stock. 

## Results

### Stock Performace
In 2017 11 out of the  12 stocks included in this analysis saw a positive yearly return, indicating that these companies saw an increase in their stock price. In 2017, only TERP saw a negative yearly return of -7.2%. 

![Alt_image title](/Resources/VBA_Challenge_Table_2017.png)

In 2018 however, 10 out of the 12 stocks surveyed saw a negative yearly return. Only ENPH and RUN were able to maintain a positive yearly return of 81.9% and 84.0%, respectively. 

![Alt_image title](/Resources/VBA_Challenge_Table_2018.png)

### Code Performance
After refactoring the VBA script the time it took the code to execute decreased. The original script took 1.835938 seconds to execute for the year 2017 and 1.363281 seconds for the year 2018. 

![Alt_image title](/Resources/VBA_Challenge_Original_2017.png)

![Alt_image title](/Resources/VBA_Challenge_Original_2018.png)

Refactoring the VBA scripted caused execution time to decrease. The refactored script took 0.2265625 seconds to execute the year 2017, and 0.2226562 seconds to execute the year 2018. 

![Alt_image title](/Resources/VBA_Challenge_2017.png)

![Alt_image title](/Resources/VBA_Challenge_2018.png)

## Summary

### Assessment of Refactored Code
In general refactoring code has advantages and disadvantages. In terms of advantages, refactoring betters the structure and design of the code, making it simpiler to read, especially when multiple individuals are working on the same script. Refactoring also makes it easier to debug, as the structure is simplier and more readable. Additionally, refactoring can improve the speed at which the code executes because refactoring elimiates any unnecessary or redundent cope.  Overall, refactoring is adventagous because the quality of the code is improved while maintaining the same desired output. 

Refactoring does have disadvantages. Refactoring can be time intensive so if time is of the essence refactoring is not adventagous. Moreover, refactoring may cause new bugs in the program, which will require even more time to fix. Finally, 

### Comparing the Original and Refactored Code
the advantages and disadvantages of the original and refactored VBA script

# VBA_Challenge
Data Analysis of stocks using VBA
## Project Overview
The following analysis using VBA to evaluate stock data from 2017 and 2018 and highlight the total daily volume and yearly return for each stock. 

## Results

### Stock Performance
In 2017 11 out of the  12 stocks included in this analysis saw a positive yearly return, indicating that these companies saw an increase in their stock price. In 2017, only TERP saw a negative yearly return of -7.2%. 

![Alt_image title](/Resources/VBA_Challenge_Table_2017.png)

In 2018 however, 10 out of the 12 stocks surveyed saw a negative yearly return. Only ENPH and RUN were able to maintain a positive yearly return of 81.9% and 84.0%, respectively. 

![Alt_image title](/Resources/VBA_Challenge_Table_2018.png)

### Code Performance
After refactoring the VBA script the time it took the code to execute decreased. The original script took 1.835938 seconds to execute for the year 2017 and 1.363281 seconds for the year 2018. 

![Alt_image title](/Resources/VBA_Challenge_Original_2017.png).        ![Alt_image title](/Resources/VBA_Challenge_Original_2018.png)

Refactoring the VBA scripted caused execution time to decrease. The refactored script took 0.2265625 seconds to execute the year 2017, and 0.2226562 seconds to execute the year 2018. 

![Alt_image title](/Resources/VBA_Challenge_2017.png)                 ![Alt_image title](/Resources/VBA_Challenge_2018.png)

## Summary

### Assessment of Refactored Code
In general refactoring code has advantages and disadvantages. In terms of advantages, refactoring betters the structure and design of the code and makes it easier to read, especially when multiple individuals are working on the same script. Refactoring also makes it easier to debug, as the structure is simpler and more readable. Additionally, refactoring can improve the speed at which the code executes because refactoring improves the structure and eliminates any unnecessary code.  Overall, refactoring is advantageous because the quality of the code is improved while maintaining the same desired output. 

Refactoring does have disadvantages. Refactoring can be time intensive so if time is of the essence refactoring is not advantageous. Moreover, refactoring may cause new bugs in the program, which will require even more time to fix. So if refactoring should be performed when there is plenty of time to review, debug, and test. 

### Comparing the Original and Refactored Code

The most significant advantage of the refactored VBA script compared to the original script is its speed in execution. In the original script the initial for loop "for i = 0 To 11 / ticker= tickers(i)/ totalVolume=0" has to execute each time the nested for loop "for j = 2 To RowCount is completed, so the nested for loop has to go through every single row before moving on to the next i, or ticker.  On the other hand, the refactored code uses the function "tickerIndex = tickerIndex + 1" within the for loop "for i= 2 to RowCount", causing the ticker to move to the next element in the array immediately after it has reached the final ticker in the row. In this way the original code goes over the 3013 rows 12 times while the refactored code goes over the 3013 rows once. Another advantage of the refactored code is its use of arrays: "Tickers", "Total Daily Volume", and "Return". The arrays allow the final outputs to be stored and if later analysis can take place if needed. This is a disadvantage for the original code as final outputs are not stored and further analysis cannot take place. 


One of the disadvantages of the refactored VBA code is the inclusion of the formatting section within the single subroutine "Sub AllStocksAnalysis()". In the original VBA code the formatting section for the output worksheet had a separate subroutine, "Sub FormatAllStocksAnalysisTable()". It improves organization and structure of the script when the formatting section has a separate subroutine because separate subroutines perform one specific tasks, not multiple tasks within one subroutine. Separating them in this way will also help debugging if problems arise. The images below show the original VBA code with its seperate subroutine and the refactored code that is within the single subroutine "Sub AllStocksAnalysis()".


![Alt_image title](Resources/VBA_Challenge_Original_Formatting.png)                     ![Alt_image title](/Resources/VBA_Challenge_Refactored_Formatting.png)


A disadvantage in both the original and the refactored script is the use of magic numbers. For example, the number 12 used for the elements in the array "tickers" in the original VBA script and the arrays "tickers", "tickerVolumes", "tickerStaringPrice", and "tickerEndingPrice" in the refactored script could be assigned to a variable such as "tickerAmount". This can prevent errors from a magic number that is being used multiple times. Additionally, replacing magic numbers with a variable can halt confusion if multiple people are working on the same project. The images below are examples of the magic number "12" that is used in both the original and refactored script.


![Alt_image title](/Resources/VBA_Challenge_Original_Magic_Number.png)                  ![Alt_image title](/Resources/VBA_Challenge_Refactored_Magic_Number.png)


Overall, the refactored code had an advantage over the original because it executes faster. On the other hand, the original code structure utilizes multiple subroutines to organize the different tasks being executed, while the refactored put all instructions into one subroutine, which can effect the debugging processes. Both codes use magic numbers, which can cause errors. Both scripts have advantages and disadvantages that can be addressed to optimize the script. 


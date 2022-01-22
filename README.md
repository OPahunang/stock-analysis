# Stock Performance Analysis

## Overview of Project:

The purpose of this analysis is to analyse a dataset of the entire stock market over the last few years. Presented a summarized comparison of ticker with total daily volume and corresponding return. 

There are 2 scripts developed through VBA -  run all stocks analysis original script and run all stocks analysis refactored script. Although the original code works well for a dozen of stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. So, a refactored script was developed. 

On the excel spreadsheet,  provided 3 button options  to click, 

    -	Run All Stock Analysis – Original Script  
    -	Run All Stocks Analysis – Refactored Script
    - 	Clear Cells

   ![Before_Running_Script_Screenshot.png](https://github.com/OPahunang/stock-analysis/blob/main/Resources/Before_Running_Script_Screenshot.png)
 

## Results:

Stock performance overall for 2017 are all green meaning positive returns except for Ticker TERP that lost -7.2%. Highest positive performer Ticker that year is DQ with a return of 199.4% and lowest positive performer Ticker RUN with return of  5.5%. 

Stock performance overall for 2018 are all red meaning negative returns except for two Tickers and RUN with a positive return of 84% and ENPH with a positive return of 81.9%. Lowest negative performer is Ticker DQ -62.6% return.

Overall, for 2017 and 2018 Ticker ENPH is the top performer. Ticker DQ even though it is the highest performer in 2017 but in 2018 it is the lowest performer. Ticker TERP both years where negative return.

Limitation of the analysis, data available are only according to 2017 and 2018 stock performance.  

# Shown below a big difference on run time of Original to Refactored Script. 

    Elapsed time of 2017 from  1.128906 seconds to 0.1982422 
    Elapsed time of 2018 from 1.15332 seconds to  0.1914063

# Shown below is the screen shot of the refactored script

   ![Refactored_Script.png](https://github.com/OPahunang/stock-analysis/blob/main/Resources/Refactored_Script.png)
    
   Major changes from the original to the refactored are:
    - Starting price and Ending price variable definition from double to single. Double is not needed for this case, and it takes more storage space
    - Search Ticker in one pass, before it was part of the loop


   1)	Results of Run All Stocks Analysis for 2017 original and refactored
 
 
   ![Original_Script_2017_Screenshot.png](https://github.com/OPahunang/stock-analysis/blob/main/Resources/Original_Script_2017_Screenshot.png)
            *Figure 1. File: Original Script 2017 Screenshot.png*


   ![Refactored_Script_2017_Screenshot.png](https://github.com/OPahunang/stock-analysis/blob/main/Resources/Refactored_Script_2017_Screenshot.png)
            *Figure 3. File Refactored Script 2017 Screenshot*

  

   2)	Results of Run All Stocks Analysis 2018 original and refactored


   ![Original_Script_2018_Screenshot.png](https://github.com/OPahunang/stock-analysis/blob/main/Resources/Original_Script_2018_Screenshot.png)
            *Figure 2. File: Original Script 2018 Screenshot.png*


   ![Refactored_Script_2018_Screenshot.png](https://github.com/OPahunang/stock-analysis/blob/main/Resources/Refactored_Script_2018_Screenshot.png)
            *Figure 4. File Refactored Script 2018 Screenshot*


## Summary:

   # Advantages and Disadvantages of Refactored Code
   
   Advantages:
	     - Code review to efficiency, using only needed definition variables
	     - Code logic review to remove redundant loops resulted to faster run time
	     - Rewriting existing running code, instead developing from scratch	

   Disadvantage:
            -	Refactored may resulted to issue or issues if code been modified to wrong logic interpretation

   # Original vs Refactored VBA Script
   
   Refactored VBA Script:
            - Running refactored script has shorter run time that the original script
            - More efficient when it comes to using definition variables 
            - More structured when added more comments

   Original Script 
            - Much faster to develop  and a starting point



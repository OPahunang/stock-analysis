# Stock Performance Analysis

## Overview of Project:

The purpose of this analysis is to analyse a dataset of the entire stock market over the last few years. Presented a summarized comparison of ticker with totally daily volume and corresponding return. 

There are 2 scripts developed thru VBA -  run all stocks analysis original script and run all stocks analysis refactored script. Although the original code works well for a dozen of stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. So, a refactored script was developed. 

On the excel spreadsheet,  provided 3 buttons option  to click, 

    -	Run All Stock Analysis – Original Script  
    -	Run All Stocks Analysis – Refactored Script
    - 	Clear Cells

![Before%20Running%20Script%20Screenshot.png](https://github.com/OPahunang/stock-analysis/blob/main/Resources/
 
## Results:

Stock performance overall for 2017 are all green meaning positive returns except for Ticker TERP that lost -7.2%. Highest positive performer Ticker that year is DQ with a return of 199.4% and lowest positive performer Ticker RUN with return of  5.5%. 

Stock performance overall for 2018 are all red meaning negative returns except for two Tickers and RUN with a positive return of 84% and ENPH with a positive return of 81.9%. Lowest negative performer is Ticker DQ -62.6% return.

Overall, for 2017 and 2018 Ticker ENPH is the top performer. Ticker DQ even though it is the highest performer in 2017 but in 2018 it is the lowest performer. Ticker TERP both years where negative return.

Limitation of the analysis, data are only according to 2017 and 2018 stock performance.  

Shown below a big difference on run time of Original to Refactored Script. 

    Elapsed time of 2017 from  1.128906 seconds to 0.1982422 
    Elapsed time of 2018 from 1.15332 seconds to  0.1914063


1)	Results of Run All Stocks Analysis 2017 and 2018– Original
 
 ./Resources/Original Script 2017 Screenshot.png
        Figure 1. File: Original Script 2017 Screenshot.png


 ./Resources/Original Script 2018 Screenshot
        Figure 2. File: Original Script 2018 Screenshot.png

2)	Results of Run All Stocks Analysis 2017 and 2018 – Refactored

./Resources/Refactored Script 2017 Screenshot.png
        Figure 3. File Refactored Script 2017 Screenshot

./Resources/Refactored Script 2018 Screenshot
        Figure 4. File Refactored Script 2018 Screenshot

##Summary:

In summary refactoring the original code makes a difference. 

  **Advantages of Refactoring Code:**
      -	Saves time executing the job
      -	Logic is reviewed and possibility to correct for any code inconsistencies
      -	Script is modified to become more structured and well-organized
      
  **Disadvantages of Refactoring Code:**
      -	If original script was not properly commented and documented it will be a nightmare to currently modifying it
      -	If not properly understood the logic of the original code, a possibility of wrong code modification to inaccurate  data result



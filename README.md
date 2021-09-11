# Stock Analysis

## Overview of Project
Our client (Steve) requested an analysis of the Stock Market for the years of 2017 and 2018.  Of particular interest was the stock for "Daqo", listed on the exchange as "DQ." Upon completion of this initial analysis, the program was enhanced to allow the client to further study any/all stocks during the year requested.

The client also requested formatting of the document to include "ease of read" features.  To accomodate this, conditional formatting was implemented to indicate positve and negative yearly performance.

After completing initial project, the code was refactored to allow for more data to be processed for future stock analysis.  The refactored code is more efficient in processing time, with the same results as the previous code (examples below).

### Results
- Stock Analysis Results
  - In 2017, Daqo (DQ) traded a volume of 35,796,200, and showed a return of 199.4% in the positive.
  - In 2018, Daqo (DQ) traded a volume of 107,873,900, and showed a return for 62.6% in the negative.  
  - While the trading volume showed a significant increase in 2017 compared to 2017, the return at close of 2018 was in a steady downward trend.  The only two Stocks that had a positive trend return for both 2017 and 2018 were (ENPH) and (RUN).

- Refactoring Code Rune Time Results

2017 Run Times

![Green_stocks_2017.png](https://github.com/nseddon/Stock-Analysis/blob/main/Module_Prep/green_stocks_2017.png) ![VBA_Challenge_2017.png](https://github.com/nseddon/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

2018 Run Times

![Green_Stocks_2018.png](https://github.com/nseddon/Stock-Analysis/blob/main/Module_Prep/green_stocks_2018.png) ![VBA_Challenge_2018.png](https://github.com/nseddon/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

As can be seen in the examples presented, the initial implementation of the code (left images) took considerably longer to run than the refactored code (right images).  While this time may seem minor with the limited data used for the requested analysis, larger datasets would suffer from exponential increases in runtime of the intial code.  The refactored code will allow for shorter run times, even with larger data sets.

### Attachments
1. [VBA_Challenge.xlsx](https://github.com/nseddon/Stock-Analysis/blob/main/VBA_Challenge.xlsm) - Spreadsheet containing raw data and analysis tabs.
2. [VBA_Challenge_2017.png](https://github.com/nseddon/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png) - Runtime report for VBA_Challenge for 2017 data.
3. [VBA_Challenge_2018.png](https://github.com/nseddon/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png) - Runtime report for VBA_Challenge for 2018 data.
4. [Green_Stocks_2017.png](https://github.com/nseddon/Stock-Analysis/blob/main/Module_Prep/green_stocks_2017.png) - Runtime report for previous code build for 2017 data.
5. [Green_Stocks_2018.png](https://github.com/nseddon/Stock-Analysis/blob/main/Module_Prep/green_stocks_2018.png) - Runtime report for previous code build for 2018 data.

#### Summary

- What are the advantages or disadvantages of refactoring code?
  - Advantages include:
    - Improved efficiency for speed and increase in size of data sets that can be accomodated.
    - Elimination of code duplication and existing bugs.
    - Better code readability allowing for teams to further refactor the code as needed for future projects. 
  - Disadvantages include
    - Increase in time needed to complete the project (financial and personnel availability implications)
    - Possibility of introduction of new bugs.
  
- How do these pros and cons apply to refactoring the original VBA script?

For this project, the time required for refactoring of the code was minimal.  The existing code was also able to be rewritten with few additions needed, preventing a large amount of bugs to be introduced.  For the client, the refactoring was completed the same day as the original code.  With the refactored code, the client now has the ability to process larger datasets, and in a shorter timeframe.  So the client has recieved a product that will be able to generate data for them into the foreseeable future.  Should the client want additional upgrades to the product in the future, another analyst will be able to perform the upgrades (based on coding refactoring and comments) should the original analyst be unavailable.

# Analyzing Stock with VBA
Click here to view the Excel file: [VBA Challenge](https://github.com/boggesstristyn/stock-analysis/blob/7c10a4c61adb514916a0e227887b9ce912354d97/VBA_Challenge.xlsm)
## Overview of Project
### Purpose
The purpose of this project was to refractor a Microsoft Excel VBA code in order to collect information on stocks from the years 2017 and 2018, taking the information and determining if a stock was worth investing in. The code had been previously written in pieces, by refractoring we were able to create one code with increased efficiency.

## Analysis and Challenges
Comparing the 2017 and 2018 stocks, we can see that 2017 performed better overall for most stocks; because tickers ENPH and RUN also produced positive returns in 2018, we can consider them good investments.

<img width="251" alt="StockPerformance_2017" src="https://user-images.githubusercontent.com/103851131/167204450-8d76ac7f-4e55-41ec-a62c-f0736650c4cc.png">
<img width="251" alt="StockPerformance_2018" src="https://user-images.githubusercontent.com/103851131/167204459-95541fa3-26b6-4adc-b067-4dccf4e31a45.png">



## Results
### 2017 Execution Time
The 2017 time using the original script was .7617188 seconds and for the refactored .140625 seconds.

AllStocksAnalysis

<img width="268" alt="AllStocks_Original_2017" src="https://user-images.githubusercontent.com/103851131/167201570-55f311fe-459c-42e8-91de-70d6996cc377.png">

AllStocksAnalysisRefractored

<img width="267" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/103851131/166128944-2f33a294-6cb2-4857-a8fb-c61926cb0ba8.png">


### 2018 Execution Time
The 2018 time using the original script was .7539062 seconds and for the refactored .1484375 seconds.

AllStocksAnalysis

<img width="276" alt="AllStocks_Original_2018" src="https://user-images.githubusercontent.com/103851131/167201898-272a1fc6-375a-47d3-8b50-4696d803ad3c.png">

AllStocksAnalysisRefractored

<img width="269" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/103851131/166128948-db4e9afe-f93c-4789-a970-1e1c4c17159e.png">


### Challenges and Difficulties Encountered
Writing the original code went mostly without trouble, I did find some difficulties while refractoring the code. I struggled a bit to know what to declare and understand the proper variables. In the end, practicing with fellow students and really taking the time to understand what was being declared and why helped me write a working code.


## Summary
Some advantages to refactoring code are that the program speed increases, the code itself is a bit eaiser to read, and it could be easier to find bugs by refractoring. Some disadvantages are that a pretty solid understanding of VBA code is necessary in order to make changes that will work and refactoring itself could create some bugs that might be difficult to fix.

Refractoring the original VBA script did ultimately increase execution speed, which is a pro; however, a con could be that the original script ran fine and produced the same results. If there was, say a time crunch, refractoring a working code might seem unnecessary.

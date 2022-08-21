# All-Stock-Analysis
# Refactor VBA Code and Measure Performance

## Overview of Project
The purpose of the project is to refactor VBA code of Stock Analysis workbook to reduce execution time and make code more efficient for Steve . This code automates the calculation of yearly volume and return for each stock for the selected year.
 ## Results
 ### Analysis

 In the beginning, a input box is used to select the year. A ticker index is created and set at zero and the output is set to be stored as arrays for starting price, ending price and total volume. A For loop is created to set ticker volume at zero. When the loop loops thru each row, the ticker volume is set to increase for the current ticker. The ticker starting prices and ending prices are set as same as the original code. Before the end of the row loop, ticker index is set to incerease by one. A For loop is used to loop the output arrays and print in the respective columns of the output spreadsheet. The last part of the code is formatting which is same as the original code. The run time is calculated by subtracting start time and end time and displaying it on the message box.

 In this analysis, the efficiancy of the refactored code is measured by execution time. A successful refactored code will run faster than the original code. When more data gets added, Steve will be able to run and analyze the data faster. Please refer to the screenshots below for execution time of Old vs Refactored code.

 For the Analysis of the stock itself,

 - Majority of the stocks did well in the year 2017 but had a rapid fall in the year 2018.
 - ENPH and RUN are the only two stocks that stayed on the positive returns for both years but RUN had tremendous increase in 2018 compared to 2017.
 - My reccomendation would be to invest on RUN as it shows positive return both years and had growth.

 ### The Old code execution time 

![StockAnalysis_2017_OldCode](https://user-images.githubusercontent.com/76926148/185770548-f156a7a4-1d9f-4014-9bcc-629386c2ca1f.PNG)

![StockAnalysis_2018_OldCode](https://user-images.githubusercontent.com/76926148/185770258-de0bdaf5-174d-4920-a3a0-f60891dce41e.PNG)

 ### Refactored code execution time
 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/76926148/185770282-e9258e27-ffa2-4cb9-938c-492142872b9b.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/76926148/185770288-0718f942-6722-4d28-8915-402cdd14f5d0.PNG)

 ### Code Comparison

![Code](https://user-images.githubusercontent.com/76926148/185770301-65d2254a-9494-4522-81ab-31c3ad7dafa4.png)

 ## Summary
 ### What are the advantages or disadvantages of refactoring code?
 #### Advantages
 - The code looks organized, clean and more readable
 - The code gets more efficient with fewer steps, uses less memory 
 - The code runs faster - beneficial when dealing huge data 
 
 #### Disadvantages
  - Refactoring can be time consuming
  - Accidental bug inclusions can happen

### How do these pros and cons apply to refactoring the original VBA script?

The refactored code has the same functionality as the originl code, the difference is the faster excution time after refactoring. 
Some of the key changes are 
 - Using ticker index
 - Removal of nested loop
 - looping output in arrays

Both original and refactored code has one limitations that I know, the number of ticker arrays are hardcoded and that has to be manually updated when the ticker data changes.


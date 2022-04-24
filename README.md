# Stock_Analysis
Performing analysis on Stock data for Module 2 challenge.

## Overview

Steve, who just graduated with his finance degree, wants to help his parents invest their money into green energy stocks. With little to no knowledge of green energy stocks, Steve's parents have decided to invest all of their money into DAQO New Energy Corp. Steve has asked for our help in analyzing a handful of green energy stocks in addition to DAQO stock to ensure his parents are making the right decision. Using VBA, we have created a code to automate the analysis for a dataset that includes the entire stock market over the last few years. Now, we will dive into refactoring the code in order to loop through all the data one time and determine whether refactoring our code had an impact on the run-time of the script.

## Results

### Stock Performance in 2017 vs 2018

The corresponding data is organized in a table consisting of a _Ticker_ column, which references the stock, a _Total Daily Volume_ column, which outputs how many shares were traded per day, and a _Return_ column, which produces the change in the profit of an investment over a period of time. The _Return_ of each stock was conditionally formatted to display a green cell for a positive return and a red cell for a negative return. The code for this can be seen below:

```
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 ```

Comparing the stock performance in 2017 versus 2018 reflects a very apparent difference. As seen in the tables below, in 2017, every stock had a positive return except for "TERP". The Stock with the highest return was "DQ" with a return of 199.4% followed by "SEDG" with a return of 184.5%. In 2018, every stock had a negative return except for "ENPH" and "RUN". The stocks with the highest return were "ENPH" with a return of 81.9% followed by "RUN" with a return of 84.0%. The stocks with the worst return were "DQ" with a return of -62.6% followed by "JKS" with a return of -60.5%. 

![Screen Shot 2022-04-24 at 1 42 20 PM](https://user-images.githubusercontent.com/101564349/164989337-2b089025-579a-4b6a-a7a6-9ab1685c8e44.png)

This analysis shows Steve that although the stock his parents plan to invest in, "DQ", was the most successful in 2017, it was also the least successful the following year. Therefore, this may seem like a risky stock to invest in. 

### Execution Times of Original Script vs Refactored Script

Comparing the execution time between the original script and the refactored script, we can see that the run time was slightly reduced. The run time for our original script in the year 2017 was 0.609375 seconds while our refactored script ran in 0.5273438 seconds. 

![Original Script Run Time (2017)](https://user-images.githubusercontent.com/101564349/164990687-ef449190-a7e9-4b24-aba8-628b81a05cc1.png)

<img width="415" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/101564349/164990791-ab3e6774-38eb-42e6-abe8-e1581929040b.png">

Similarly, the run time for our original script in 2018 was 0.59375 seconds while our refactored script ran in 0.521562 seconds.

![Original Script Run Time (2018)](https://user-images.githubusercontent.com/101564349/164990852-05f73091-2b3a-4bca-8c09-6107ffaf2473.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/101564349/164990878-63f476e8-6421-430d-b3bd-85ae9d36a42f.png)

We were able to reduce the processing time for both year's datasets by refactoring our code in a few different ways. First, we defined our primary variables as arrays and set a tickerindex variable to zero. The tickerindex variable is used to access the correct index across our various arrays. 

```
    '1a) Create a ticker Index
    tickerindex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
We were then able to use a for-loop to output the data into our table. This looped through the arrays we defined in order to give us an output. 

```
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

## Summary

### Advantages of Refactoring Code
The main purpose of refactoring code is to restructure and improve existing code without changing its base function. Refactoring code can provide many advantages when done with a sufficient amount of time such as:
* Improving efficiency which can save time and money in the future
* Simplifying the code for better understanding
* Makes code more maintainable and can help prevent bugs within the code during future use

### Disadvantages of Refactoring Code
Although there are many advantages to refactoring code, there can also be some disadvantages such as:
* May take a lot of time to work out how to optimize the code
* Can be risky in certain situations as you need to have sufficient amount of time to successfully refactor your code without any bugs

### How Pros / Cons Apply to Refactoring Original VBA Script
The advantages and disadvantages listed above directly apply to refactoring our original VBA script. In this case, we were able to organize the code to run more efficiently which allowed for the proper analysis of the larger dataset. The code was simplified to allow for maintainability and because we were allotted ample time, we were able to eliminate any bugs. 

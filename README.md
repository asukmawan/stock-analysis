# Green Energy Stock Analysis using VBA

## VBA generated automated reports to find Total Volume and Returns on multiple stocks

### Background

A client wants to invest into green energy stocks, where one stock in particular (DAQO:DQ) is of interest. The client wants to know the value of stock DQ, and how it compares to other green energy stocks. There are two main criteria the client wants to know in order to make his judgement on whether to invest or not: 

1. The stock's total traded volume 
2. Percentage yearly return

 The report will have to sum up all of the daily volume for stock DQ, as well as finding the years starting and ending prices to find the yearly return as a percentage difference in price.

 Knowing DQ's stock data alone won't be enough to decide whether DQ is a valuable energy stock. Which is why it needs to be compared with other green energy stocks available to invest in. The VBA script can be coded in a way we can store all the stock tickers we want to look at in an array, and output the same information we found with DQ's stock to all green energy stocks

 Lastly, if we want this report to be repeatable, fast generating and scalable, we need to refactor our code in a better way while accomplishing the same task.

### Refactoring the code
Once the script running without errors and the report is showing expected results, the code needs to be refactored to have it run more efficiently and improve readability. With the current provided stock data, it only provides data for years 2017 and 2018, as well as only providing data for 12 different stock. If this were to be scaled up to thousands of stocks across many years, the script will take much longer. Which is why its important for the future development of this code, that it be refactored.

## Results

### Stock performance between 2017 and 2018

By the client's criteria of a valuable stock, Total volume and percentage yearly return needed to be known. 

Before we can start calculating total volume and return, we need to initialize the tickers array with the ticker names and set up the loop so the tickers will cycle through them all:

    Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    For i = 0 To 11

    ticker = tickers(i)

To calculate the total volume, a `For` loop was used to go through all rows of data:

        For j = 2 To rowCount                
            If Cells(j, 1) = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        Next j

Where column 1 is where the ticker name is listed, and column 8 is where the daily volume is held. This code block allows us to add up all volume with the condition of the ticker, which will cycle through 0 to 11 with the previous for loop for the ticker array.

To determine return, the starting and end price were found using the following:

        'starting price for current ticker
            If Cells(j, 1) = ticker And Cells(j - 1, 1) <> ticker Then
                startingPrice = Cells(j, 6)
            End If

        'ending price for current ticker
            If Cells(j, 1) = ticker And Cells(j + 1, 1) <> ticker Then
                endingPrice = Cells(j, 6)
            End If

The code block for `startingPrice` is determining whether the cell above its current cell is the not the same, then we can set the start pricing for that ticker from that row under column 6. Similarily for `endingPrice` we are determining if the current cell and the cell below it are different. This will determine that this is the last row for that particular ticker, and so the end price will be set from that row.

The the current ticker is then output to cells on a separate worksheet and then the code will loop back again on the next ticker. 

The result of running the script is shown below:

### *Table 1 - Stock data for 2017*

<img src="resources/VBA_Challenge_2017.PNG"></img>

### *Table 2 - Stock data for 2018*

<img src="resources/VBA_Challenge_2018.PNG"></img>

Looking at DQ's stocks in 2017, DQ had the highest return out of all of the green energy stocks available. While this is good as one of the client criteria, the other criteria (stock volume) is showing as having the lowest volume. 

Also, though returns were high in 2017, boasting a 199.4% return from the beginning of the year to the end, DQ's stocks dropped 63% in the following year. In comparison to the other stocks that year, only two stocks ended up doing well - ENPH and RUN. 


### Execution times of the original script vs the refactored script

Looking at Table 1 [(Stock Data for 2017)](#table-1---stock-data-for-2017), we can see at the bottom of the image of a MsgBox showing the run time of the code. This duration refers to the code after we had already refactored our code to run more efficiently. To properly determine if our refactor was successful we need to compare its duration vs. the old code:

<img src="resources/VBA_Challenge_2017_before_after.PNG"></img>

<img src="resources/VBA_Challenge_2018_before_after.png"></img>

The refactored code runs much faster, and the differences in run times will become even more apparent as the data source gets larger. In the case we want to add code to have it run without userInput and run it for all years, we can have the current block code (that is going through the tickers array), refer to an index variable instead of using a magic number to loop the code. This will exponentially improve on the run times, as the added data will have to cycle through each loop in a more efficient and more flexible way.

To accomplish this, the `tickerIndex` was the variable added and used during the calculation of volume, starting price and ending price. This is in replacement of the original `For` loop which would have the code go through all the data for one ticker, and once done will check the next ticker and go through the data all over again.

The `tickerIndex` was increased as a conditional in the already existing loop:

            If Cells(j + 1, 1) <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If

The result of this change makes the code faster as the code will loop through all the data once, instead of going back through the code for each ticker.

## Summary
### Refactoring code

In general, refactoring code increases its repeatability, efficiency, readability for debugging. The code can be run faster, and as a result can be ran over and over with less time. It can benefit the developer as well during the debugging process, as the code will be tested and won't have to wait as long for the code to finish. It can also be easier to maintain as the code should be easier to understand and less complex. This makes the code more available to add other functions which increases the capability of the script.

But refactoring can be time consuming especially if its someone elses code and if there are not a lot of comments on what the code is doing. It is hard to gauge how much time it will take before you are finished, and the result may or may not be impactful. It can also be confusing and frustrating during the refactoring process, as attempts to change the code may cause errors and even more time commitment is needed to have the code work again.

Other issues with refactoring is that there may not be anything that can be done with the code, either due to the restraints of the language, the task required of the code was simple enough not to warrant refactoring, or the code was already written well no significant improvements can be made.

### Original vs. Refactored 

The original code outputs to the cell as it goes through the loop of tickers, whereas the refactored code has the ticker data stored in the array and then output all at the end. The issue with how this was designed has the assumption that the original tables are not filtered. If the table is filtered to one stock, trying to run the macro will result in a false output.

But this can be remedied with by entering either the below:

    Worksheets(yearValue).Activate
    ActiveSheet.AutoFilter.ShowAllData 

This will only work if there is a filter currently applied. So if there is no filter there is an error

Therefore we can use the below:

    Cells.AutoFilter

This acts more like a toggle to turn the filter on and off which clear the filter and avoid any errors in the output.

### Refactored code

The tickers array are currently hardcoded in both original and refactored scripts. This will be an issue when the client will want to add or examine new tickers. The script will not recognize new tickers unless we add it when the tickers array is initializing. not only that, we have to remember to increase the parameters of the array to account for how many new tickers there are. Leaving these hard coded as well can lead to mistakes when updating or debugging the code in the future.

The advantage of having the ticker volumes and starting/ending prices as arrays is that we will be able to update the parameter for the number of tickers in the `tickerIndex` variable, and it will update to the rest of the code. opposed to the original, where we would have to replace all For loop

We can also replace the hard coded tickers array with a conditional For loop to add ticker strings to the tickers array, as the script encounters new ticker strings.

As mentioned as a similar disadvantage with the original data that its code would only work if the data is not filtered, this refactored code would only work if the data is not filtered but even more importantly the data needs to be sorted. If not sorted, the code we use to track which ticker we are looking at relies on the change in ticker in column A: 

            If Cells(j + 1, 1) <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If

If the data is not sorted by ticker, the `tickerIndex` may change prematurely and will go to the next tickers() array leaving data untouched. Though this is a disadvantage if we leave the code as is, we could correct the script to make sure whatever year we are looking at needs to be sorted and unfiltered. 


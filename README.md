# Stock Analysis

## Overview of Project

### Purpose
In order to make our code more efficient and more user friendly down the road, we are refactoring the All Stocks Analysis VBA Script.   
Steve wants to expand the dataset to the entire stock market over the last few years.  Since this would incorporate thousands of more stocks, we would need to optimize this script in order for it to run faster.


### Results

In order to see how the 12 stocks did during the year 2017, we created our code to provide two metrics.  The first metric is the total daily volume.  The total daily volume is the number of all shares of a particular stock that was traded during the enire year.  The other metric we pulled from the data was the return.  The return percentage was the ending price divided by the starting price and then having 1 subtracted from the quotient.   Since the stock data was sorted in order by date, the starting price for each ticker would be the close price on the first instance of the ticker in the spreadsheet.  In order to get this through using our macro, we created the following conditional code

```
If Cells(j - 1, "A").Value <> tickers(tickerindex) Then
    
  tickerStartingPrices(tickerindex) = Cells(j, "F").Value
   
End If
```        
        

In order to get the ending price for each ticker, we used a similar code that instead checked to see if the row contained the last instance of a particular ticker.


``` 
If Cells(j + 1, "A").Value <> tickers(tickerindex) Then
        
  tickerEndingPrices(tickerindex) = Cells(j, "F").Value
    
End If
```

After we got the results from the macro, we can see that with the exception of the "TERP" stock, the other 11 stocks had a positive return percentage for the year 2017.  

Due to the input box function built into the beginning of our macro, we are also able to run this again for the 2018 stock data, by simply running it again and inputting 2018.  When we run the code for 2018, we see that all of stocks have a negative return percentage except for "ENPH" and "RUN".

In terms of the script's performance, we found that the refactored code actually ran faster than the original code.  For the 2017 dataset, the original code ran in .28125 seconds as shown below.

<img width="288" alt="Screen Shot 2021-10-23 at 11 36 20 PM" src="https://user-images.githubusercontent.com/87248687/138579483-cbc3effc-cb8f-4d51-a07d-ca3acd53bf9d.png">


With the refactored code, the same 2017 data set ran in .05078 seconds which is almost 5x times faster.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/87248687/138579492-1dee0a2e-65d9-44d9-985f-164bcdb5654a.png)

For the 2018 date set, we got similar performance results.  For the 2018 data set, the original code ran in .28125 seconds as shown below.

<img width="293" alt="Screen Shot 2021-10-23 at 11 41 05 PM" src="https://user-images.githubusercontent.com/87248687/138579579-0a355566-0a48-48fe-9ce4-80e1375666f1.png">

With the refactored code, it ran through the 2018 data in .05078 seconds.  A screenshot can be seen below.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/87248687/138579644-61bd0d87-9062-4364-b3a0-9e8b4ff2d914.png)

In both scenerios, the refactored script appears to be almost 5 times as fast as the original code.

## Summary

When it comes to refactoring code, it has advantages and disadvantages.  The most obvious advantages of refactoring your code is that it can provide immediate opportunities to make your code run faster.  There are also opportunities to simplify your code and also improve it's readability.  Another advantage of refactoring your code is that you can simplify it for future use through the use of things like arrays and variables so that instead of having to change the code completely for new data sets, you can more simply update an array or update the value in a particular variable.  The disadvantages of refactoing code can also be numerous.  One of the biggest disadvantages in refactoring code is that you can in fact introduce new errors to the code.  Things that may not have been a problem in the original code can be introduced accidentally through an attempt of refactoring.  Another disadvantage of refactoring code is that it can be very time consuming.  If the macro was created for a one time use, it may not be beneficial to spend time refactoring a code that you most likely will never use.

When it came to refactoring our original VBA script to do the stock analysis, it had many pros and cons.  One of the most noticeable pros was that we were able to improve the speed in which the code ran.  Although the a fraction of a second may not seem like a significant difference, when being ran on a data set 1000x times that size, it could potentially make a significant difference.  Another pro is that even though we changed a lot of things within the loop of the code, the implementation of variables and arrays have made it so that we only have to update those parts if we decide to apply this to a larger stock data set.

The cons of refactoring this VBA script is that it can be somewhat time consuming.  While the script did run faster, the script ultimately gave us the same exact output that the original script gave us.  Another con of refactoring this particular VBA script is that even though we have set this script up to be used on more stocks in the future, we still can not avoid manually inputting each new stock ticker into the ticker array.  In the event that we wanted to run this on a data set with 1000 new stock tickers, we would have to input each ticker one by one.  Lastly, this script would only work seamlessly with stock data that followed a smilar format to the data we currently have. If we received data with different heads or pehaps was not ordered in chronological order.  We would first have to sort that data before even being able to refactor our code.


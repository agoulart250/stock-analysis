# stock-analysis

## Overview of Project
  The purpose of this project is to design a stock market analysis model in VBA to help Steve on his investment advisory services to his parents. This new model is being refactored from a previous VBA code to include the entire stock market results for the years of 2017 and 2018. It is also being modified to increase its processing speed to help Steve analyze the data at light speed and with accurate findings. 

## Results
### 2017 vs. 2018 Stock Performance
    2017 overall stock performance was very positive with 11 out of the 12 selected stocks generating positive results. DQ was the highlight of the portfolio as it outperformed all other stocks by nearly doubling in price (199.4% return). TERP was the only stock with negative results, shrinking in price by -7.2%.
    While in some instances the story seems to repeat itself the same cannot be told about the stock market. 2018 was a bearish year for investors with 10 out of the 12 stocks in the portfolio turning negative. Only ENPH and RUN had positive results with 81.9% and 84.0% in returns respectively. 
     Over the two years ENPH was the best performer, accumulating 211.4% in total positive returns. SPWR was the worst multi-year performer with a cumulative -21.5% in total negative returns.  
###  Code
    The code for this analysis utilizes loops to run through the entire data set and is structured by the utilization of arrays for all 12 stocks, output arrays for ticker volumes, as well as for starting and ending prices (utilized to calculate annual returns). Below is quick explanation some of the codes used in accomplishing this task, along with the code example:
   -	To create a ticker index: tickerIndex = 0
   -	To loop over all rows in the spreadsheet to ensure coverage of the entire data set: ‘For i = 2 To RowCount
   -	Counting the total volume per ticker by utilizing an index for the tickers while looping through the entire data:    If Cells(i, 1).Value = tickers(tickerIndex) Then tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
   -	Looping through all the arrays to output Ticker, Total Daily Volume and Return: 
       For i = 0 To 11
          Worksheets("All Stocks Analysis").Activate
           Cells(4 + i, 1).Value = tickers(i)
           Cells(4 + i, 2).Value = tickerVolumes(i)
           Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
 ### Execution Time (Refactored Code vs. Previous Code)
    The refactored code showed a significant improvement compared to its previous version. For 2017 the refactored code ran in 0.2265625 vs. 0.2695313 in the old code, which equates to ~16% improvement in processing time. For 2018 the refactored code ran in 0.234375 seconds vs. 0.2734375 seconds in the old version which is 14% faster.

## Summary
  In sum, refactoring code is a good way to save time and effort in coding. It leverages from previous efforts in thinking through how to accomplish tasks through the use of coding which is a great advantage. The flipside of that is that one might miss updating certain parts of the old cold which might be detrimental to the quality and accuracy of what the new cold is trying to accomplish (disadvantage). In other words, refactoring is only as good as one’s ability to properly understand the code being re-written as well as his/her ability to effectively update the code to make it accurate and effective.
 In refactoring the code in this assignment, it is clear that its outcome was positive. The code shows the same results while running considerably faster. The disadvantage to this process, for me as the coder was that it was very time consuming to figure out how to improve the code, which is a task I was only able to pull with the thoughtful guidance of TA’s in the program. It was a big challenge!



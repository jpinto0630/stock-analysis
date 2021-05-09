# Unit 2 Challenge - Stock Analysis 

## Project Overview 
A friend of ours named Steve recently asked us to help him analyze a portion of his parent’s stock portfolio in excel utilizing VBA code. This latest iteration of the workbook is designed to run an analysis on the current data set and allows for future expansion as new data becomes available. 
## Results 
The analysis was performed on the data set provided by Steve which included stock performance for 12 stocks for the year 2017 and 2018.  
### 2017 
Our analysis shows that out of the two years 2017 was the better performing year. Out of the 12 stocks analyzed 11 of them had a positive return, with 4 stocks having over 100% returns for this year. 2017 had less total daily volume traded, however the discrepancy does not appear to be significant.

![image](https://user-images.githubusercontent.com/67031885/117557931-58dcd980-b046-11eb-93c8-b33b45ee6dac.png)


 
### 2018 
Inversely 2018 was the worse performing year out of our 2-year data set. Out of the 12 stocks analyzed 10 of them had negative returns for the year, we also observe that the two stocks that performed well both yielded over 80% return on investment. 2018 had more total daily volume traded, however the discrepancy does not appear to be significant.

![image](https://user-images.githubusercontent.com/67031885/117557934-62664180-b046-11eb-8447-f7afe56bd331.png)

### Code 
Our original code for this analysis was functional for out small data set but was structured in such a way that could pose issues in the future for larger data sets. Our original code looped through the whole spread sheet one row at a time checking column A for the ticker name and did not stop once it reached the last row that included the corresponding ticker. Our code would loop 3,000 times looking for “AY”  even after reaching the last row containing “AY” 

```   For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    Worksheets(yearValue).Activate
        For j = 2 To RowCount
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
                
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            
            End If
            
        Next j
 ```
        
We refactored the code to address this issue by decreasing the number of loops and recognizing that the data is neatly organized. After making the changes the code ran noticeably faster, previously the code ran in about 10 seconds, whereas after refactoring it is now under 1 second. 


```
for j = 0 to 11
        tickerVolumes(j) = 0
    Next j
    
    For i = 2 To RowCount
        
        
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        
     
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        
         
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            
            tickerIndex = tickerIndex + 1
            
        End If
    
            Next i
   ```
            
![VBA_Challenge_2017](https://user-images.githubusercontent.com/67031885/117558173-e5889700-b048-11eb-9d8c-e39a003af526.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/67031885/117558174-e7eaf100-b048-11eb-9a10-223e7098c871.PNG)

## Summary

### General Refactoring 
Refactoring code is generally a positive action that often results in better performing and more easily readable code. It also allows for easier debugging as refactoring removes redundancies or simplifies/improves processes. A disadvantage of refactoring code is that it can be time consuming and may be cost prohibitive in some legacy code situation. 

### Refactoring Our Code 
Refactoring our code appears to have yielded generally positive outcomes. Our friend Steve has provided input and we have modified our code to better suit his needs while providing a noticeable processing time improvement. However, a disadvantage is that considering this is a stock tracker the usage of this particular workbook may be limited. Thus, our improvement to the code may not be as impactful to Steve. 



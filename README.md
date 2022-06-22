# Refactoring "Green Stocks" Analysis Code Using VBA
## Project Overview 
An analysis on a group of "Green" stocks was performed using code in VBA. The code worked on a group of 12 stocks, but what would happen if huundreds of stocks were added to the data you wanted to analyze? Though your code may work, there could be a make your code run faster when it comes to running analysis on more data. Below is a comparison of run times of original code and refactored code for data from two different years and a summary of the changes made to speed up time of the code. 
 
## Results
The original script was modified to make the code more flexible to holding larger data sets and to run faster by reformatting certain sections of code with a new varible "tickerIndex" 
### Original Script 
I wanted a code that would analyze 2017 and 2018 volume and prices for 12 different stocks over a one year. The stocks were assigned ticker values that corresponded to the data in the 2017 and 2018 worksheets which contained a daily report of stock price and volume information. I wanted a code that would loop through volume and adjusted close for each stock, but to do that I needed to create a nested loop. In the first "for" loop I use the varible "i" to loop through an array of the 12 stock tickers. The varibles "startingPrice", "endingPrice", and "total volume" were created for the next loop.

```
Dim tickers(11) As String
 
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
    
'Prepare for the analysis of tickers.
  'Initialize variables for the starting price and ending price.
    
 Dim startingPrice As Double
 Dim endingPrice As Double
    
 'Activate the data worksheet.
    
 Worksheets(yearValue).Activate
    
 'Find the number of data rows to loop over in 2017 or 2018 Worksheet.
    
 RowCount = Cells(Rows.Count, "A").End(xlUp).Row
 
 'Loop through the tickers.
    
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
'Loop through rows in the data.

    Worksheets(yearValue).Activate
        For j = 2 To RowCount
 ```
While looping though ticker symbols, I created another "for" loop "j" that outputted "totalVolume", "startingPrice", and "endingPrice" for each stock by looping through daily closing price and volume data from the 2017 and 2018 worksheets. 
```
'Loop through rows in the data.

    Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
    'Find the total volume for the current ticker.
    'If "Ticker" cell in worksheet reads specified ticker value ("ticker(i)"), then add "Volume" cell value inw worksheet to "totalVolume"
        
        If Cells(j, 1).Value = ticker Then
            
            totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
    'Find the starting price for the current ticker.
    'If cell before "Ticker" cell in worksheet does not read specified ticker value ("ticker(i)") and reads  specified ticker value ("ticker(i)") then the   
     price recorded for that day in the "Close" column is the yearly starting price for that stock.
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            startingPrice = Cells(j, 6).Value
            
            End If
    'Find the ending price for the current ticker.
    'If cell after "Ticker" cell in worksheet does not read specified ticker value ("ticker(i)") and reads  specified ticker value ("ticker(i)") then the
     price recorded for that day in the "Close" column is the annual ending price for that stock.
         If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            endingPrice = Cells(j, 6).Value
            
            End If   
        Next j
```
The three "if" statements generated values for "totalVolume", "startingPrice" and "endingPrice" for each stock. To output the analysis I created a new workrsheet that contained a column for the 12 stocks, their annual total volume, and their annual return.
```
'Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    'Create column for ticker name
    Cells(4 + i, 1).Value = ticker
    'Create column for totalVoume
    Cells(4 + i, 2).Value = totalVolume
    'Create column for annual return
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
 Next i
```
The original code ran in 0.84375 seconds using the 2017 worksheet data.

![Original_Run_2017](https://user-images.githubusercontent.com/104794100/174687595-b8dd1075-7622-47de-a28b-ddf0a1e1229d.png)

The original code ran in 0.84375 seconds using the 2018 workheet data.

![Original_Run_2018](https://user-images.githubusercontent.com/104794100/174687599-53e5a46b-da6c-4b0f-82df-b77e732b9889.png)

### Refactored Script
The original code could output the data I needed for 12 stocks, but it could have been written better and run faster if I wanted to find data on hundreds or thousands of stocks. To make the code usable for more than 12 stocks I created an additional "tickerIndex" variable and created output arrays for volume and return analysis.
```
  'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    
     Dim tickerVolumes(12) As Long
     Dim tickerStartingPrices(12) As Double
     Dim tickerEndingPrices(12) As Double
 ```
I wanted to create a loop that would reset to the output value to zero every time it ran.
 ```
 '2a) Create a for loop to initialize the tickerVolumes to zero.
   For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
 ```
I wanted to create another for loop that similar to the original code, but contained a different value. The original "tickers" varible would not be "tickers(tickerIndex)". The "tickerIndex" variable is used in an additonal "if" statement to increase by one at the end of each loop.
```
'2b) Loop over all the rows in the spreadsheet.
         For j = 2 To RowCount
       
        '3a) Increase volume for current ticker
           If Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
             If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
             End If
            

            '3d Increase the tickerIndex.
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
        Next j
 ```
 The fourth "if" statement in the refactored code outputs values for the three output arrays by increaing "tickerIndex", and in essence tickers(i), by 1
 every time the loop repeats.
 ```
   '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
       For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
       Next i
 ```
 
The refactored code ran in 0.1953125 seconds using the 2017 worksheet data.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/104794100/174687524-afa5235e-f507-4422-8f42-23cd8c37ed3a.png)

The refactored code ran in 0.1953125 seconds using the 2018 worksheet data.

![VBA_Chellenge_2018](https://user-images.githubusercontent.com/104794100/174687531-3e69b4e0-c545-457e-91ab-da07b114bfc9.png)

## Summary

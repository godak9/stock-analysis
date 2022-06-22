# Refactoring "Green Stocks" Analysis Code Using VBA
## Project Overview 
An analysis on a group of "Green" stocks was performed using code in VBA. The code worked on a group of 12 stocks, but what would happen if huundreds of stocks were added to the data you wanted to analyze? Though your code may work, there could be a make your code run faster when it comes to running analysis on more data. Below is a comparison of run times of original code and refactored code for data from two different years and a summary of the changes made to speed up time of the code. 
 
## Results
The original script was modfied to make the code more flexible to holding larger data sets and to run faster by reformatting certain sections of code with a new varible "tickerIndex" 
### Original Script 
I wanted a code that would analyze 2017 and 2018 volume and prices for 12 different stocks over a one year. The stocks were assigned ticker values that corresponded to the data in the 2017 and 2018 worksheets which contained a daily report of stock price and volume information. I wanted a code that would loop through volume for and adjusted close for each stock, but to do that I needed to create a nested loop. In the first "for" loop I use the varible "i" to loop through an array of the 12 stock tickers. The varibles "startingPrice", "endingPrice", and "total volume" were created for the next loop.

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
    'If cell before "Ticker" cell in worksheet does not read specified ticker value ("ticker(i)") and reads  specified ticker value ("ticker(i)") then the  price recorded for that day in the "Close" column is the yearly starting price for that stock.
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            startingPrice = Cells(j, 6).Value
            
            End If
    'Find the ending price for the current ticker.
    'If cell after "Ticker" cell in worksheet does not read specified ticker value ("ticker(i)") and reads  specified ticker value ("ticker(i)") then the  price recorded for that day in the "Close" column is the annual ending price for that stock.
         If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            endingPrice = Cells(j, 6).Value
            
            End If   
        Next j
        Next i
```
The three "if" statments in generated values for "totalVolume", "startingPrice" and "endingPrice" for each stock. To output the analysis I created anew workrsheet that contained a column for the 12 stocks, their annual total volume, and their annual return.
```
'Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
```
The original code ran in 0.84375 seconds using the 2017 worksheet data.
![Original_Run_2017](https://user-images.githubusercontent.com/104794100/174687595-b8dd1075-7622-47de-a28b-ddf0a1e1229d.png)

The original code ran in 0.84375 seconds using the 2018 workheet data.
![Original_Run_2018](https://user-images.githubusercontent.com/104794100/174687599-53e5a46b-da6c-4b0f-82df-b77e732b9889.png)

### Refactored Script

The refactored code ran in 0.1953125 seconds using the 2017 worksheet data.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/104794100/174687524-afa5235e-f507-4422-8f42-23cd8c37ed3a.png)

The refactored code ran in 0.1953125 seconds using the 2018 worksheet data.
![VBA_Chellenge_2018](https://user-images.githubusercontent.com/104794100/174687531-3e69b4e0-c545-457e-91ab-da07b114bfc9.png)

## Summary

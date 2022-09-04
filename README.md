# VBA-of-Wall-Street

## Overview of Project

The purpose of this analysis is to figure out which green stock(s) to invest in based on 2017 and 2018 daily trading data of 12 companies. 2 measures are used: total daily volume of stocks traded, which is the sum of all transactions effectuated on each stock. The second measure is the % return on each stock, which is calculated using the change in starting and ending prices of stocks. After developing a code to perform this analysis, I refactored it to see if it would make the run time shorter. 

## Results

Comparing the stock returns in 2017 and 2018, we notice that most stocks double digit growth in 2017. This growth could be attributed to the fact that the companies may have just been created (start-ups), experiencing higher growth in the first years of creation and a slump right after, as they reach maturity. The only 2 stocks that experienced positive returns are stocks with ENPH and RUN tickers, with over 80% growth in 2018 (please refer to tab <a href="VBA_Challenge.xlsm">Yearly_Comparison</a> of excel file named VBA_Challenge.xlsm for the analysis). In terms of volumes, ENPH seems to have been the most traded stock out of all 12 companies. Thus, ENPH seems like the safest stock to invest in if we base our decision solely on trading volumes and returns.

Execution time of the original script was 0.67 s while execution time for the refactored script is 0.17 s when running the analysis for both years. The refactored script took 0.5 seconds less to execute than the first script.

The original run times can be found below:
<img src="Original_2017.png" width="500">
<img src="Original_2018.png" width="500">

The run times of the refactored code can be found below:
<img src="Refactored_2017.png" width="500">
<img src="Refactored_2018.png" width="500">

### Original code: 

```

Sub AllStocksAnalysis()
  'Adding a timer to measurecode performance
    Dim startTime As Single
    Dim endTime  As Single
  'Input year of analysis
   yearValue = InputBox("What year would you like to run the analysis on?")
   
   'Timer starts after user inputs answer
   startTime = Timer
   
   '1) Format the output sheet on All Stocks Analysis worksheet
   Sheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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
   
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   
   '3b) Activate data worksheet
   Sheets(yearValue).Activate
   
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
      Sheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
   'Timer ends here
   endTime = Timer
   MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```

### Refactored code 
```

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
    tickerVolumes(i) = 0
    
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            
               tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

        '3d Increase the tickerIndex.
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        'End If
        End If

    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```


## Summary

### The advantages or disadvantages of refactoring code
1.	Advantages of refactoring a script include lower execution time, and making the code more dynamic (including less magic numbers).  The code looks cleaner, shorter and more structured.

2.	Disadvantages of refactoring a script is having to use elaborate tricks and techniques which make understanding the code slightly more difficult. Faulty code/failure to write code properly to introduce new variables for instance can cause more errors and bugs, which makes the process of developing the code much longer. 

###	How do these pros and cons apply to refactoring the original VBA script?
Both the original and refactored codes yielded the same result, which would make this process seem unnecessary and time-consuming. 



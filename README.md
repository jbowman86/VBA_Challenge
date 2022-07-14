# VBA Coding and the Stock Market

## Overview of Project

The aim of this project is to analyse the yearly stock data in order to compare stock performance for the years 2017 and 2018.  To accomplish this task, the VBA script was refactored to determine total daily volume of trades for each stock as well as its rate of return.  Analysis was completed using Excel and VBA code.  The refactored code was then timed and compared against the original code to determine if newer code is more efficient.  The end goal of this analysis is to provide information of stocks from the years 2017 and 2018 to thereby identify stocks that have the best return and should be invested in future.

## Results

The original code used for this analysis was obtained via the following link.  

(https://github.com/jbowman86/VBA_Challenge/blob/de861d00223c13744dd9c4b273a28ba0124726af/Resources/VBA_Challenge.vbs)

The data that the code is applied to can be accessed via the following link:

(https://github.com/jbowman86/VBA_Challenge/blob/c7db4407f0107a9a54f122179cac31fdf4939b03/VBA_Challenge2.xlsm)

The following is the steps completed to refactor the original VBA code:

1. Set up the macro to allow analysis to be completed in multiple years by creating an input box command.  Additionally, create a timer to measure length of time it takes to complete the tabulations.  This code is as follows:

        Sub AllStocksAnalysisRefactored()
    
            Dim startTime As Single
    
            Dim endTime  As Single

            yearValue = InputBox("What year would you like to run the analysis on?")

            startTime = Timer


2. Format the output sheet on All Stocks Analysis worksheet.


        Worksheets("All Stocks Analysis").Activate
    
        Range("A1").Value = "All Stocks (" + yearValue + ")"


3. Create a header row.

        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

4. Initialise array of all tickers.

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

5. Activate the worksheet.
    
        Worksheets(yearValue).Activate

6. Get the number of rows to loop over.


        RowCount = Cells(Rows.Count, "A").End(xlUp).Row


7. Create a ticker Index

        tickerIndex = 0

8. Create output arrays for tickerVolumes, tickerStartingPrices and tickerEndingPrices.  The tickerVolumes array was defined as a Long data type and the tickerStartingPrices and tickerEndingPrices were defined as Single data types.  The revised code was as follows:

        Dim tickerVolumes(12) As Long
    
        Dim tickerStartingPrices(12) As Single
    
        Dim tickerEndingPrices(12) As Single

9. Create a for loop to initialise the tickerVolumes to zero.

        For i = 0 To 11
    
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        
        Next i

10. Create a for loop that will loop over all rows of the spreadsheet.

        For i = 2 To RowCount

11. Write script that increases the current tickerVolumes variable and adds the ticker volume for the current stock.

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

12. Check if the current row is the first row with the selected tickerIndex. If row is identified as the first row, assign it as the ticker starting price.

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If

13. Check if the current row is the last row with the selected tickerIndex. If row is identified as the last row, assign it as the ticker ending price.

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If

14. Write script that increases the tickerIndex if the next row's ticker doesn't match the previous row's ticker.

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
                
        End If

15. Close loop with the following code:

        Next i

16. Create a for loop to loop through array (tickers, tickerVolumes, tickerStartingPrices, tickerEndingPrices) to output the "Ticker", "Total Daily Volume", and "Return" results in the spreadsheet.


        For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i


17. Format the worksheet

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

18. Code Timer to stop after analysis was completed.  Create a message box to state the total time needed to run the code.


        endTime = Timer
    
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

19. Close the subroutine.

        End Sub

20. Run the refactored code.

21. Record and the present the stock markets results for each stock's total daily volume traded and return on investment for 2017 and 2018.  Make note of time needed to run the code.  The results of these analyses are included in the links below:

- 2017 Results (https://github.com/jbowman86/VBA_Challenge/blob/de828b4b55fac96e530b737d3c06c2c101d7f0b3/Resources/VBA_Challenge_2017%202.png)
- 2018 Results (https://github.com/jbowman86/VBA_Challenge/blob/82ae1b4fabeb3ebcc61fdad075849b3ea9733963/Resources/VBA_Challenge_2018.png)


22. Compare results with original code.  The original results are included below:

- Original 2017 Results (https://github.com/jbowman86/VBA_Challenge/blob/9dbadc10561f0e82aca8929fad4ba4ade6d681a8/Resources/VBA_Challenge_Original_2017_stock_analysis.png)
- Original 2018 Results
(https://github.com/jbowman86/VBA_Challenge/blob/31d3ed8a112d89bfe2a9471734084ef34477a9c8/Resources/VBA_Challenge_Original_2018_stock_analysis.png)

## Summary

### Advantages of Refactoring Code

The advantages of refactoring code include:

- It can be faster than the original code.
- It can reveal patterns that may not be easily observed through the original code.
- If the code is well-structured, it can be easier to identify errors.

### Disadvantages of Refactoring Code

The disadvantages of refactoring code include:

- It can take a long time to write the code.
- A long procedure may have the same lines of code in multiple locations, you can change the logic to remove the duplicate lines.
- A complex unstructured code is usually best split into several functions.
- The refactoring process can affect testing outcomes.

### Application of Refactoring the Original VBA Script

Refactoring the original VBA script is helpful as it creates a good foundational code that is more useable and easier to read.  It is a clearer way to communicate the steps involved in the code; thereby, making it more user friendly in terms of understandability and maintenance.  Further to these points, it easier to share the code with other programmers due to its clarity and readability.  Additionally, utilising refactored code early in the process can help avoid frustration and difficulty later as making changes in refactored code is easier than if the code was unstructured.

# Stock Analaysis Script - Refactored

## Overview:   
   - Steve original sought to analyze the performance of a select few green energy stocks. His VBA worked well, but it was not equipped to handle large sets of data. Since Steve wanted to expand his analysis to the entire stock market, he needed to refactor his code to streamline it and enhance its processing speeds to allow for more data manipulation. As such Steve's original code was refactored and the results and analysis are as follows.

## Results:
   - 2018 was clearly a down year for these stocks with only two stocks having positive returns, while all but one had positive returns in 2017. The code was successfully refactored, with the images below evidencing the stock performance in 2017 & 2018 of both the refactored and original codes. You can see the output is the exact same (we will get to the time efficiencies later):
    
        - 2017 Return Refactored: 
   
        ![2017 Refactored](https://user-images.githubusercontent.com/111612130/188525708-7be04510-21d3-4c47-a683-c1d825383934.png)
   
        - 2018 Return Refactored:
   
        ![2018 Refactored](https://user-images.githubusercontent.com/111612130/188525734-78be198a-c313-48f9-8485-eb096575f279.png)
    
        - 2017 Return Original:
        
  
        <img width="230" alt="2017 Original" src="https://user-images.githubusercontent.com/111612130/188525945-a8c91cea-68db-4990-8935-cbeed547571e.png">
  
        - 2018 Return Original:
        
  
       <img width="234" alt="2018 Original" src="https://user-images.githubusercontent.com/111612130/188525952-97e3ab20-4e64-43c8-bbaa-a931fecc7406.png">

### Runtime

   - The execution times were significantly better however on the refactored code vs. the original. The refactored code had a runtime of 0.07 seconds while the original had a run time of 0.54 seconds, clearly evidencing the benefits of the refactored code vs. the original. This enhanced runtime was also evident on when running the refactored script for year 2017. Examples as follows: 

  2018 Original Runtime: 
  
  <img width="263" alt="Original Run Time" src="https://user-images.githubusercontent.com/111612130/188526272-8b572e09-b8c1-4ade-95ee-3277e406168f.png">
  
  2018 Refactored Runtime:
  
  ![Refactored Run Time](https://user-images.githubusercontent.com/111612130/188526299-07d8d150-45b9-410d-b480-d2d1df5604fe.png)
  
  2017 Refactored Runtime: 
  
  <img width="264" alt="Refactored 2017 run time" src="https://user-images.githubusercontent.com/111612130/188526326-e2b9f6d4-1d08-4eea-8546-1e6bf1f43930.png">

   - Runtime enhancements can mainly be attributed to creating a tickerIndex, which allowed the code to loop through all the variables and build up and index which was vital towards enhancing runtime. Code example as follows:
  
    '1a) Create a ticker Index

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
       If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              tickerIndex = tickerIndex + 1
    
        End If
       
     Next i
        

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysisRefactored").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

## Summary:
   - Refactoring code has its advantages and disadvantages. From an advantage standpoint it can save runtime and make the code more organized and efficient. Comparatively, it can also make the code more complicated from a troubleshooting perspective, should an issue arise, especially if you are creating a VBA script for a team with very few VBA experts. In refactoring, additional time is spent enhancing the code and one must weigh if the benefits (enhanced runtime & streamlined code) are worth the time it took to refactor the code and the complexity it added. Additionally, one must ensure that the current code, once refactored, won’t lead to the same issues as the old code.
   - In this example, the juice was worth the squeeze, the refactored code significantly cut down on code execution time and is more apt to handle an entire stock market worth of code vs. the original script. Steve should ensure he understands the refactored code should he need to make an adjustment or trouble shoot it in the future. Further, the refactored code propagated a potential issue in Steve's original code. Steve's base code relied on the assumption that the user would not sort the data at all and to find the beginning and ending price it just needed to find the first and last row for each ticker. One should look to also enhance code when refactoring and think of these user issues, a better way to do this would be to find the max and min date values for each ticker and populate the prices on those dates. Example of the flaw in the original code as follows.
                       
         '3b) Check if the current row is the first row with the selected tickerIndex.
         'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
          End If

         '3c) check if the current row is the last row with the selected ticker
          'If the next row’s ticker doesn’t match, increase the tickerIndex
            
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
           End If


## Refactored Code for Reference:
    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysisRefactored").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index

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
       If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              tickerIndex = tickerIndex + 1
    
        End If
       
     Next i
        

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysisRefactored").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("AllStocksAnalysisRefactored").Activate
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

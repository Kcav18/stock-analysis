# Green Stock Analysis


## Overview

### Background

This stock analysis was created for Steve to help his parents decide which stock options to invest in. The original analysis was created for a dataset that included only a dozen stocks. Steve loved the workbook that was created because it was interactive and easy to use - Showing performance at the click of a button. Steve liked it so much that he has requested the ability to expand the analysis of stocks to include more than the original dozen.

### Purpose

My original analysis code has been updated (also referred to as refactored) so that it will more efficiently work with additional stock data that Steve or his parents would like to have analyzed. This edited code has been simplified and is easier to follow should myself or another analyst work with the data at a later time.  The results also ran faster with the new code. This will help tremendously as the dataset expands! Let's now take a deep dive and look at the differences in code. 


## Results

To view the spreadsheet that includes the refactored code, click here: [VBA_Challenge](https://github.com/Kcav18/stock-analysis/blob/main/VBA_Challenge.xlsm)

In the updated / refactored version of the Green Stocks Analysis code, I created a loop that loops through the data one time and collects all the information. A ticker index variable was created to access the correct index across four different arrays. Three of these arrays are new in the refactored code. Those new output arrays are tickerVolumes, tickerStartingPrices, and ticketEndingPrices. The refactored code also includes the formatting in the analysis rather than having it as a second sub procedure as I had in the original code.

The original and refactored code are shown below for comparison. Be sure to continue past the code for further information.

Original Code:

```
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim Endtime As Single
    
    yearValue = InputBox("What year would you like to run the analysis for?")
    
    startTime = Timer
    
   '1) Format the output sheet on All Stocks Analysis worksheet
  
   Worksheets("All Stocks Analysis").Activate
   
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
   Worksheets(yearValue).Activate
   
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       
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
   
   Endtime = Timer
   
   MsgBox "This code ran in " & (Endtime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub formatallstockanalysistable()

Worksheets("All Stocks Analysis").Activate

Range("a3:c3").Font.Bold = True
Range("a3:c3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("c3").Font.Italic = True
Range("a3:c3").Font.Color = vbBlue
Range("a3:c3").Borders.Color = vbBlue
Range("a3:c3").HorizontalAlignment = xlCenter

Range("b:b").NumberFormat = "$#,##0"
Range("c:c").NumberFormat = "0.00%"
Columns("b").AutoFit

With Range("a1:c1")
.Merge
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.Font.Size = 18
End With

datarowstart = 4
Datarowwend = 15

For i = datarowstart To Datarowwend

    If Cells(i, 3) > 0 Then
    Cells(i, 3).Interior.Color = vbGreen
    
        ElseIf Cells(i, 3) < 0 Then
        Cells(i, 3).Interior.Color = vbRed
                 
             Else
             
             Cells(i, 3).Interior.Color = xlNone
        
End If
    
        
    Next i

End Sub

Sub clearformatting()

Cells.Clear

End Sub
```

Refactored Code:
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
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
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerindex As Integer
    tickerindex = 0
    

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerindex = 0 To 11
    tickerVolumes(tickerindex) = 0
    
    Worksheets(yearvalue).Activate
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
             If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerindex) = Cells(i, 6).Value
              
            
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesnt match, increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                'if it is the last row for current ticker, set ending price.
                tickerEndingPrices(tickerindex) = Cells(i, 6).Value
            End If
            

            '3d Increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                tickerindex = tickerindex + 1
                
                End If
                
            
    Next i
    
    Next tickerindex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
    Cells(i + 4, 1).Value = tickers(i)
    
    Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
       
        
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue) & " from the refactored code"

End Sub

Sub Clear_AllStocksAnalysis()

Cells.Clear

End Sub
```
As shown above, the code is a bit different! Even with its differences though, it does not change the actual content or conditional formatting of the data. A comparison of the data output and formatting before and after refactoring are shown below:

| 2017 Original Data Output and Formatting        |        2017 Refactored Data Output and Formatting |
| -------------                                   |        -------------                              | 
![2017 SockAnalysisData Original](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2017_SockAnalysisData_Original.png) | ![2017 StockAnalysisData Refactored](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2017_SockAnalysisData_Refactored.png) |


| 2018 Original Data Output and Formatting        |        2018 Refactored Data Output and Formatting |
| -------------                                   |        -------------                              | 
![2018 SockAnalysisData Original](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2018_SockAnalysisData_Original.png) | ![2018 StockAnalysisData Refactored](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2018_SockAnalysisData_Refactored.png) |



From the images above, it is obvious that the code did not change the output or the conditional formatting on the returns column. The only difference is in the headings on the original charts. That was because I was intentionally "playing" with some additional formatting code. :)

Another thing that really stands out from the data output graphics above is that Steve's parents really should investigate other stock options ...These stocks did great overall in 2017 but most completely tanked in 2018. I would be very hesitant about putting money in most of those options! The only two that I would take a second look at are tickers ENPH and RUN. Further research of other stock options is still recommended though! Hopefully this refactored code will help Steve and his parents analyze some better stock options...and the new code will analyze the data much faster than the original code! See the timing results below:

| 2017 Time with Original Code | 2017 Time with Refactored Code |
| -------------                |        -------------           | 
![2017 SockAnalysisData Original Timer](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2017_Timer_Original.png) | ![2017 StockAnalysisData Refactored Timer](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2017_Timer_Refactored.png) |

| 2018 Time with Original Code | 2018 Time with Refactored Code |
| -------------                |        -------------           | 
![2018 SockAnalysisData Original Timer](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2018_Timer_original.png) | ![2018 StockAnalysisData Refactored Timer](https://github.com/Kcav18/stock-analysis/blob/main/Resources/2018_Timer_Refactored.png) |

The timers above show that on average, the code will run 82% faster with the refactoring! That will be a huge deal with a larger dataset! 

I will be happy to complete further analysis once you provide me with more options. Otherwise, feel free to share this code with your next Data Analyst - it should help them get a jumpstart on your analysis.

## Summary

### What are the advantages or disadvantages of refactoring code?

### How do these pros and cons apply to refactoring the original VBA script?


# Stock Analysis With Excel VBA
https://github.com/J-tok/stock-analysis/blob/1423506c8decdfb9e50f12840f61fbc9c3aa0e95/VBA_Challenge.xlsm.xlsm?raw=true

## Overview of Project
### Purpose
The Purpose of this project was to refactor the Stock Analysis VBA code so that it can collect the 2017 and 2018 stock data and output the return values more efficiently. This information when activated with reflect the stock values performance by year.

### The Data
The data includes two years of stock data activity for 12 tickers (reflected by year in two separate worksheets). The stock information has the data issued, the opening, closing and adjusted closing price. It also contains the daily volume and return of the collected stock information.

## Results
### Analysis
The refactored code took the original stock analysis code to create the column and row headers for the important stock data values (Ticker, Total Daily Volume, & Return Values). The code formula for this challenge was used to gather the data for the respective values and when ran it inputted the data into a seperate worksheet.

    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
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
        
        '3c) check if the current row is the last row with the selected ticker
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
        
        Worksheets("All Stock Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stock Analysis").Activate
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

## Summary
### Advantages and Disadvantages
Refactoring code helps with processing a command function such as VBA more quickly and effectively. I believe this is beneficial because this allows all users that have access to certain code to run it more quickly depending on the processing power of their machine. Disadvantages may include overcomplicating the code, or not being able to finish the process of making the code more concise and clean. 
### Takeaway from Refactoring Stock Analysis
Not only did this code run more effectively, we were able to develop a new method to analyze this stock data and use this knowledge to create better code in the future.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/83428759/123555435-4f6d1500-d74b-11eb-8c19-46a8fc361fa9.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/83428759/123555437-5136d880-d74b-11eb-8251-5ed392d35916.png)



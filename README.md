# Refactoring VBA Code for Stock Analysis

### Overview of Project
Refactoring VBA code designed for analysis of a handful of green energy stocks so that it will work efficiently when performing the same analysis on the entire stock market.

#### Results

##### Initial VBA Code
The initial VBA code (Module 2) examined 12 green energy stocks, looking at their total daily volume and return on the chosen year. Here's a look at the chunk of code that performed this analysis:

![Initial Code](https://user-images.githubusercontent.com/107162310/174845059-afdb8db0-ff0c-4949-a31f-5dd76fccb472.png)

The worksheet included a button that allowed the user to type in the year they wanted to analyze (2017 or 2018), using the following code:

> yearValue = InputBox("What year would you like to run the analysis on?")

Also included in the code was some formatting, a button to run the analysis, and a button to clear the sheet to run it again. Finally, a timer was built to track the amount of time the code took to run:

> startTime = Timer
> endTime = Timer
> MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

##### Stock Performance
In 2017, all but one of the 12 green energy stocks had a positive rate of return:

![image](https://user-images.githubusercontent.com/107162310/175049856-dd8e72fa-d349-4158-9481-d93493ae44aa.png)

Things changed dramatically in 2018, with all but two stocks posting a negative return:

![image](https://user-images.githubusercontent.com/107162310/175050218-eb838281-9e85-4b27-9c6d-849ff5e6e6ec.png)

##### Initial VBA Code Run Time
Using this timer, the initial code took 0.890625 seconds to process the 12 stocks for the year 2017 & 0.9257813 seconds to process for the year 2018.

##### Refactoring for Improved Performance
The challenge was to refactor this code so that it ran faster and allowed for upscaling to include all the tickers on the stock market rather than just a dozen. Looking at the image above of the initial code, it really begins with a loop through the tickers:

> ticker = tickers(i)
> totalVolume = 0

From there a nested loop begins that loops through the rows in the data, first for volume, then starting price, then ending price. The nested loop ends once it's finished processing, then the data is output into our sheet, and the initial loop of the tickers moves to the next ticker until it too has finished processing.

To improve this performance, especially when applied to an even larger dataset, the tickers were first placed in an index. Then, the initial loop was expanded from simply the ticker to also include starting and ending prices. The nested loop expanded to include the tickerIndex in all aspects of its processing. And lastly, the formatting was shortened by a step to no longer color the return cell if no change was made on the year. In total, the new macro looked like this:

>'1a) Create a ticker Index
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.

    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.

        For i = 2 To RowCount
        
    '3a) Increase volume for current ticker

        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
    '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
    
    '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then

    If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
    
    '3d Increase the tickerIndex.
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
    
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
        Range("B4:B15").NumberFormat = "#,##0.00"
        Range("B4:B15").Style = "Currency"
        Range("C4:C15").NumberFormat = "0.00%"
        Columns("B").AutoFit
        
        dataRowStart = 4
        dataRowEnd = 15
        
For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        Else

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        End If

    Next i

These changes were designed to make the code able to process at a better efficiency. The results indicate that it worked. Here's a look at the code timers for both years with this improved code:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/107162310/175060663-35442a6b-dc46-42c6-9c09-5b339e4685b9.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107162310/175060682-3c543f49-295b-49a2-8ca6-6d47d5e1c8d4.png)

On average, the initial code took 0.90820315 seconds to process either year. The refactored code took 0.14453125 seconds to process on average.

**That means the refactored code processes over 6 times faster on average than the initial code.

#### Summary

##### Advantages or Disadvantages of Refactoring Code
There are several advantages to refactoring code. Whether the code is processing data for analysis or running a social media application, refactoring code can lead to faster processing times. Simply taking the time to clean up the formatting of the code (its use of white space, for instance) and including thoughtful notes explaining what each section is designed to do, this will improve the ability for others to utilize the code in the future.

The advantages outweigh the disadvantages when it comes to refactoring code. But certainly if one's code is so large and cumbersome, the amount of time, effort, and cost it could take to refactor it may make any gains from improvement not worthwhile. Another possible disadvantage could be that the refactoring might end up breaking the code's initial use.

##### Pros and Cons to Refactoring the Initial VBA Script
In this example of code refactoring, this initial VBA script was small enough to make refactoring worthwhile. Reduncancies such as the formatting in the initial VBA code were discovered and eliminated. The ability to build in a way for the code to upscale to greater datasets certainly makes it easier to use in the future. 

Luckily, this code was small enough that the disadvantages of refactoring were largely not a factor. If a mistake was made that broke the code, it was simple enough to start over.

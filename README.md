# **Analysis of Green Energy Stocks**
## **Overview of Project**
### Purpose
The purpose of this analysis was to evaluate different green energy stocks. The client was interested in better understanding the trends of green energy stocks so they could make knowledgeable investments while diversifying their portfolio. The goal was to calculate the total daily volume to show how often the stock was traded as well as the starting and ending prices. These indicators reflect how the stocks performed over the 2017 and 2018 fiscal year.
In deciding how to best code this project, the pros and cons of different patterns of “For Loops” and refactoring code were highlighted.


## **Analysis and Challenges**
### Analysis of Outcomes: Stock Performance
Stock performance was calculated using simple Excel arithmetic functions:
totalVolume = totalVolume + ‘new cell value
Return= endingPrice / startingPrice - 1

### Analysis of Outcomes: Refactored Code
The refactored code changed from a nested For Loop to separate For Loops that referenced a list of values and used a new variable “TickerIndex”.

### Challenges and Difficulties Encountered
Challenges included determining where and how to break up the one nested function into several parts. This in turn created the question of how to store the calculated values in a manner that created efficiency and reduced the number of times the program had to loop over the data.  

## **Results**
### Stock Performance Results
The majority of stocks yielded a higher return in 2017 than in 2018. Code was implemented to color code the value to indicate a growth (green) or a reduction in value (red).

    For l = dataRowStart To dataRowEnd
        If Cells(l, 3) > 0 Then
            Cells(l, 3).Interior.Color = vbGreen
        Else
            Cells(l, 3).Interior.Color = vbRed
        End If
    Next l

### Execution Times Results
Execution times of the refactored script were about 6x faster than the original code.
| Year | Original | Refactored |
| :----: | :----: | :----: |
| 2017 | 0.911 | 0.154 |
| 2018 | 0.949 | 0.153 |

Refactored Code:

    '1) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
        ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            ‘Code to gather relevant information.
        Next j
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For k = 0 To 11
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + k, 1).Value = tickers(k)
            Cells(4 + k, 2).Value = tickerVolumes(k)
            Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
        Next k

Run time results:
![ VBA_Challenge_2017_refactored](https://github.com/K10Huff/Stock-Analysis--Green-Energy/blob/main/Resources/VBA_Challenge_2017_refactored.png)
![ VBA_Challenge_2018_refactored](https://github.com/K10Huff/Stock-Analysis--Green-Energy/blob/main/Resources/VBA_Challenge_2018_refactored.png)


vs. Original Code

The subroutine was coded to gather the relevant data within the initial For Loop instead of its own function:

    '4. Nested Loops Through Tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
            '5 Loop through Rows within the data
            Worksheets("2018").Activate
            For j = rowStart To RowEnd
                'total volume
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
        
                'set startingPrice
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
        
                'set endingPrice
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
            Next j
            
            '6) Output data for current ticker
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
        Next i

Run time results:
![ VBA_Challenge_2017_original](https://github.com/K10Huff/Stock-Analysis--Green-Energy/blob/main/Resources/VBA_Challenge_2017_original.png)
![ VBA_Challenge_2018_original](https://github.com/K10Huff/Stock-Analysis--Green-Energy/blob/main/Resources/VBA_Challenge_2018_original.png)

## **Summary**
1.	What are the advantages or disadvantages of refactoring code? 
    Advantages to refactoring include making the code less complex and therefore easier to maintain. Simpler code is also more readable for the next developer who must work with the code. Disadvantages to refactoring include spending additional time/resources for coding and retesting. Refactoring code may also interrupt functionality of dependencies ifs they are not taken into consideration when reworking the code.
    
2.	How do these pros and cons apply to refactoring the original VBA script?
In this project, since the refactored code is simpler, the subroutine was much faster to execute. However, it also resulted in additional hours being spent on the project. 

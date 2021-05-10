# stocks-analysis

## Overview of Project

### Background

We have created a workbook with VBA macros capable of calculating the total daily volume and year end return for every unique stock ticker.  In this project our goal is to refactor our existing VBA macro so that it runs more efficiently. We will asses the different solutions and challenges we ran into when refactoring our code.

### Purpose

Refactoring code is an extremely important skill. Often times the first iteration of a solution is not always the most efficient or user friendly. Our goal is to streamline our program. Since currently we only are looking at dozen stock tickers we want to make our code more robust and able to handle hundreds if not thousands of different tickers. Throughout or refactoring process we will compare run times to determine if our refactoring is effective or not. These run times are store in the Resources folder.

### Results

#### Run Time Before Refactoring

![2017 before Refactor](https://github.com/rulma/stocks-analysis/blob/8e1a3c310ac69326d7cb0ae99b00516909559f54/Resources/2017%20before%20refactor.PNG)
![2018 before Refactor](https://github.com/rulma/stocks-analysis/blob/67148fee658fbf33757abe96c35e584def7baeff/Resources/2018%20before%20refactor.PNG)

As you can see our initial script is able to complete each request in just under a second. Now this may seem fast but we must consider that we are only working with 12 unique tickers and a couple thousand lines of data. A more realistic data set will have at least 100 times more data to sort and analyze. 

Our biggest time constraint is the nested for loops that fill and calculate our desired metrics
    
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        Worksheets(yearvalue).Activate
        For j = 2 To RowCount
            
            '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b) get starting price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '5c) get ending price for current ticker
            If Cells(j, 1).Value = ticker And Cells((j + 1), 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If

        Next j
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
  '''
  
We are calculating Total Volume and Return for each ticker and then writing those variables within each outer loop. Taking the time every iteration of i to write the cell values is cumbersome. Especially since we are switching back and forth between two excel sheets. In our refactor we will look at a more seamless way to get the same result by creating arrays for the Volume, Starting Price, and Ending Price for all of tickers then writing to the after we have filled our arrays.

#### Run Time After Refactoring
![2017 after Refactor](https://github.com/rulma/stocks-analysis/blob/main/Resources/2017%20refactored.PNG)
![2018 after Refactor](https://github.com/rulma/stocks-analysis/blob/29023f76f6e9b6a482e4fea462a00f66c8e30263/Resources/2018%20refactored.PNG)

After refactoring our script we were able to decrease our runtimes by aproximately 10%! This could prove to be very useful moving forward if we wanted to analyze multiple years at a time or more unique tickers. We were abelt to save time by changing the way we stored and wrote our metrics for each ticker. 
   
    '1a) Create a ticker Index
    tickIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(tickIndex) = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            If Cells(i, 1).Value = tickers(tickIndex) Then
            
                tickerVolumes(tickIndex) = tickerVolumes(tickIndex) + Cells(i, 8).Value
            
            End If
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickIndex) And Cells(i - 1, 1).Value <> tickers(tickIndex) Then
                
                tickerStartingPrices(tickIndex) = Cells(i, 6).Value
                
            End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickIndex) And Cells(i + 1, 1).Value <> tickers(tickIndex) Then
                
                tickerEndingPrices(tickIndex) = Cells(i, 6).Value

            End If
                   
        Next i
        
        '3d Increase the tickerIndex.
        tickIndex = tickIndex + 1
        
    Next j

'''


### Summary

Refactoring is something that every programmer will need to do at some point throughout their carreer. There can be both advantages and disadvantages when refactoring a piece or all of a code base. 

#### Advantages to Refactoring in General

When solving problems through code, our first solution is rarely if ever perfect. Refactoring provides us with an opportunity to move and develop quickly. We are able to release and push hot fixes to problems we are trying to solve while having the ability to go back and refine them after release. This level of flexibility allows organizations of all sizes to implement change fast and effectively. The faster a company can respond to the challenges it faces then the more productive and adaptive their product/services can be.

A company that is able to quickly change and adjust to the changing forces around them allows them to withstand changes that may have put them out of business otherwise. 

Refactoring is also an excellent way for an organization to train new engineers or developers on their respective code bases. Well documented code allows for new developers to familiarize themselves with the existing code base without requiring them to develop something from scratch.

#### Disadvantages to Refactoring in General

Refactoring can often cause unforseen problems. As organizations and their respective code bases grow it is important to maintain pristine documentation. Without proper documentation, changes made can have harmful consequences. For example if someone where to adjust the steps needed for a user to login without considering the dependecies in each step then they could in theory break the login process in production. When a companies service or product is down due to develop error they can lose thousands of dollars for every hour that it is down. 

The old saying "If it ain't broke, don't fix it"  can be a hard truth to accept. If we are just trying to make cosmetic changes that don't have a real effect on perfomance we may end up breaking something that once worked without issues. This can lead to hours of headaches and bug fixes. 

#### Advantages to Refactoring our Script

When we refacotered our VBA script, we were able to increase performance by nearly 10%. Now for a program that already took less than a second to run this may seem inconsequental but for a process that could time minutes if not hours this could potentially save days worth of computing power. This time saved translate to direct dollar savings for an organization. It will also allow for quicker analysis leading to more up to date decision making which may increase the likelihood of them making a more educated decision.

#### Disadvantages to Regactoring our Script

While we were able to refactor our script to run faster it, it did take a siginifigant amount of time to decrease our run time by 10%. If this program was never going to scale beyond the data set we were testing on then there may be little reason to spend the time refactoring it in the first place. Our refactor also still fails to account for a more dynamic data set. We had a hard coded number of stock tickers to work with in our data. If we were to look at a more robust market analysis we would need to be able to handle any given number of stock tickers. 

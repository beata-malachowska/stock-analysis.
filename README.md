# Title: Stock-analysis for Steve

##Overview of Project

Using stock data from 2017 and 2017 for eleven companies, revenue was calculated by collecting for each year starting and ending price. Those values were used to calculate yearly revenue in percentage. By summing all operation volume, total yearly volume operation was also calculated. Each calculation was performed separetly for each company. 

##Results

Refracotry code was based mostly on creating arrays to store output data from the loops.

        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
First loop was created in order to scan though each ticker and the variable with Volumes was set to zero. 
       
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
    
Then another loop was activated to scan through each row.  Worksheet with data was activated. 
  
            Worksheets(yearValue).Activate
            
            For j = 2 To RowCount

If the ticker in the database was the same as the one selected by tickerIndex then volume from this row was added to Volumes variable.

                If Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                End If
                
                
For each row it was checked if this is the first or the last price of the transaction and if yes then those prices were recorded in appropriate variables. 

            If Cells(J - 1, "A").Value <> tickers(tickerIndex) And Cells(J, "A").Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(J, "F").Value
            End If
                
        
            If Cells(J + 1, "A").Value <> tickers(tickerIndex) And Cells(J, "A").Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(J, "F").Value
            

Then, if the last condition was met tickerIndex is increased and and algoritm is performed for another J (row).                     
            
            tickerIndex = tickerIndex + 1
            End If
         Next J
    
Finally, the output of the search stored within variables is printed into the cells in appropriate worksheet with another loop. Apropriate worksheet is activated before the opperation. 


        Worksheets("All Stocks Analysis").Activate
        For i = 0 To 11
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        Next i
            

In 2017 most companies had positive revenue. Only one company had negative return: TERP. 

In 2018 most companies had negative revenue. Only two companies had positivve return: ENPH and RUN. 


Original execution times were provided in links below:

[https://github.com/beata-malachowska/stock-analysis./blob/main/Resources/VBA_Challenge_2017.png]

[https://github.com/beata-malachowska/stock-analysis./blob/main/Resources/VBA_Challenge_2018.png]


Refractord execution times were prodided in links below:

[https://github.com/beata-malachowska/stock-analysis./blob/main/Resources/VBA_Challenge_2017-new%20time.png]
[https://github.com/beata-malachowska/stock-analysis./blob/main/Resources/VBA_Challenge_2018-new%20time.png]

Refractored code was faster than original one but the differences were minimal.  

##Summary

###What are the advantages or disadvantages of refactoring code?

It was faster in execution of the refractoring code and it contained less lines of code (with loop in loop operation). But the original code was more easier to undersatnd so I think it would be less difficult for somebody from the outside to read it. However, refractoring was cumbersome so it is would not be worth doing for simple analysis like this. Better done fast than perfect! 

###How do these pros and cons apply to refactoring the original VBA script?

In this example, I spend too much time trying to improve the code than we will every saved on performing this calculations. However for more complicated operation it would be good to remember that loop in loop operations take so much more time.

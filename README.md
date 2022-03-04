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
    
Then another loop was activated to scan through each row.  Worksheet with data was activated. 
  
            Worksheets(yearValue).Activate
            
            For j = 2 To RowCount

If the ticker in the database was the same as the one selected by tickerIndex then volume from this row was added to Volumes variable.

                If Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                End If
                
                
Fro each row it was checked if this is the first of the last price of the transaction and if yes then those prices were recorded in appropriate variables. 

                If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                End If
                
                If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                End If

Then, the algorythm is going to the next row and search for new ticker.                    

            Next j

                     
            tickerIndex = tickerIndex + 1
    
Finally, the output of the search stored within variables is printed into the cells in appropriate worksheet. 


                Worksheets("All Stocks Analysis").Activate
                
                Cells(4 + i, 1).Value = tickers(i)
                Cells(4 + i, 2).Value = tickerVolumes(i)
                Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
            

In 2017 most companies had positive revenue. Only one company had negative return: TERP. 

In 2018 most companies had negative revenue. Only two companies had positivve return: ENPH and RUN. 


Execution times were provided in links below:

[https://github.com/beata-malachowska/stock-analysis./blob/main/VBA_Chellenge_2017.png]

[https://github.com/beata-malachowska/stock-analysis./blob/main/VBA_Challenge_2018.png]

Refractored code was faster than original one but the differences were minimal.  

##Summary

###What are the advantages or disadvantages of refactoring code?

It shoudl be faster in execution and should contain less lines of code so in theory it should be faster to run. But the original code was more easy to undersatnd so I think it would be easier for somebody from the outside to read it. However, it's cumbersome and not always speeding up the process, so it is not worth doing for simple analysis or the ones we are going to just do once or twice. Better done fast than perfect! 

###How do these pros and cons apply to refactoring the original VBA script?

Not worth it, I spend too much time trying to improve the code than we every saved on performing this calculations. 

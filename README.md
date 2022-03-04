# Title: Stock-analysis for Steve

##Overview of Project

Using stock data from 2017 and 2017 for eleven companies, revenue was calculated by collecting for each year starting and ending price. Those values were used to calculate yearly revenue in percentage. By summing all operation volume, total yearly volume operation was also calculated. Each calculation was performed separetly for each company. 

##Results

Refracotry code was based mostly on creating arrays to store output data from the loops.

        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    

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

# Stock Analysis using VBA

## Overview of the Project

Steve is trying to help his parents choose which Green Energy stocks they should invest in. Their initial leaning was towards investing in the ‘DQ’ stocks. Steve has approached us to help him. We start by looking at 

  *	The stock data of Green Energy stocks over the period of 2 years 2017 and 2018 
  *	DQ stock in particular to understand if their initial leaning is a good one
  *	All the stocks to check and see if there are options better than DQ for his parents to invest in 

We have two years (2017 & 2018) of stock performance data for 12 Green Energy stocks.  
We have decided to use VBA for this analysis.
The code we wrote provided Steve with the answers he was seeking, however, we felt that a more efficient code was possible that could help 
  * Reduce the run times
  * Avoid repetitive tasks thus reducing strain on the systems

We modified the existing working code (a process commonly called Refactoring) with the aim to make it more efficient. 
We called the first code All Stock Analysis and the second code All Stock Analysis Refactored.

## Results

Summarized into three parts  
  * The insights from Green Energy stocks and our advise to Steve and his parents based on that 
  * Steps involved and process followed to develop our original All Stocks Analysis code 
  * Steps followed for our Refactored code  

### Green Energy Stocks Analysis

#### DQ Stock Analysis

The ‘DQ’ stock performance declined by 63% by end of year 2018
We recommend to Steve and his parents that they need to look at better performing stocks to put their money

![DQ_Analysis_2018](https://user-images.githubusercontent.com/85518330/123638008-4714e880-d7e4-11eb-9f64-b54e898dad6c.png)


#### All Stocks Analysis 2017
Almost all stocks except TERP did well in 2017 

![All_Stocks_2017](https://user-images.githubusercontent.com/85518330/123638150-71ff3c80-d7e4-11eb-924a-286c3ef2d5ad.png)


#### All Stocks Analysis 2018

In 2018 the tables turned on most stocks including DQ. The two stocks that have been in the green over both years are ‘ENPH’ and ‘RUN’
![All_Stocks_2018](https://user-images.githubusercontent.com/85518330/123638257-96f3af80-d7e4-11eb-8da5-0cf09a8122a0.png)

#### Our Recommendation

Would be to invest in either ‘ENPH’ or ‘RUN’ stocks instead of DQ and to especially follow the stock ‘RUN’ as growth rates are on an increasing trend between 2017 & 2018 

### All Stocks Analysis Code 

Below is the stepwise process to build the All Stocks Analysis code. 
This documentation tells Steve or anyone else who uses this in future what was done and why.
   *	We created and activated a new worksheet called “AllStocksAnalysis” to tabulate our analysis results 
   *	Created the relevant title and headers for our analysis sheet to hold the relevant calculations
   *	We are interested in computing the totalVolume and returns for each stock
   *	Initialized the arrays for the 12 stocks in our dataset 
   *	Activated the sheet where our data for the analysis resides 
   *	Wrote a code to get the last row of data for our data set as we are dealing with 3000+ rows of data 
   *	We wrote a loop from 0 to 11 for all the 12 stocks in our dataset. We set the totalVolumes to zero initially  
   *	We next wrote a loop to go through each of the rows in our data from row 2 to the last row with data. Our code does the following
   *	For all 12 tickers in column 1. Goes through each row, checks to see if the next ticker is the same as the current ticker and if yes, then adds the volumes from               column 8 to get totalVolume for that ticker
            
    If Cells(j, 1).Value = ticker Then totalVolume = totalVolume + Cells(j, 8).Value
      
   * Check to make sure that the previous ticker in not the current ticker, then, the closing price from column 6 is taken to be the startingPrice of that ticker
            
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then startingPrice = Cells(j, 6).Value
 
   * Check to make sure that the next ticker is not the current ticker, then the closing price from column 6 is taken to be the endingPrice for that ticker 
           
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then endingPrice = Cells(j, 6).Value

   *  Next we output the data in the AllStocksAnalysis sheet tickers, totalVolume & return 
   
   *  Formatting  
   
   *  Color coded stocks with positive returns in green and those with negative returns in red so that Steve can know the performance in one glance 
          
     If Cells(4, 3) > 0 Then Cells(4, 3).Interior.Color = vbGreen
     ElseIf Cells(4, 3) < 0 Then Cells(4, 3).Interior.Color = vbRed
     Else Cells(4, 3).Interior.Color = xlNone

   *  We ensured that our code enabled Steve to pick the year on which he wanted to do his analysis 
      
     yearValue = InputBox("What year would you like to do the analysis for?")
      
   *  We added a timer to help us determine how long our code was taking to run.
   

![Original_Runtime_2017](https://user-images.githubusercontent.com/85518330/123559199-3753c080-d760-11eb-8f5b-cb48699bc24c.png)



![Original_Runtime_2018](https://user-images.githubusercontent.com/85518330/123559205-3f136500-d760-11eb-9652-5ee1684de04e.png)

       
### All Stocks Analysis Refactored Code 

Steve was quite happy with our analysis, but relooking at our code, we felt that there was a more efficient way to write it.  


#### Inefficiencies in our original code
 
 *	Our original code looped through each row in our data for each ticker to get our outputs
 	
 *	That means it went through 3000+ rows of data for each ticker. That is both time consuming and inefficient 
 
 * Our data is sorted by dates per ticker, this should tell us that once we complete going through a ticker and computing the data we need from it, there is no need to go        through it again for the next ticker.
 
 * Hence, we may now able to complete our data collection by looping our database just once instead of 12 times in our previous code. That should save us a bunch of time …. 

 *	We do this in our refactored code by following the below steps

   * Create a tickerIndex 
  
           tickerIndex = 0
  
   * Create 3 output arrays to hold our results
    
          Dim tickerVolumes(12) As Long
          Dim tickerStartingPrices(12) As Single
          Dim tickerEndingPrices(12) As Single 

   * A loop from 0 to 11 for all the 12 stocks in our data and initialize the tickerVolumes, tickerStartingPrices and tickerEndingPrices to 0 

   * Begin looping through each row of the spreadsheet to get the tickerVolumes using  
 
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

   *	If the current row is the first row with the selected tickerIndex. then set the corresponding price in column 6 as the tickerStartingPrices
 
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

   *	If the current row is the last row with the selected ticker,then set the corresponding price in column 6 as the tickerEndingPrice
  
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
 
   *	Increase the tickerIndex.If the next row’s ticker doesn’t match, increase the tickerIndex
 
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1

   *	We output the data in the AllStocksAnalysis sheet tickers, totalVolume & return    
 
   *  Formatting and adding color to our output sheet to make it readable 
 
   *	Allowing for inputting the year of analysis 

   *	Adding a timer to capture run time 
   
   
   ![VBA_Challenge_2017](https://user-images.githubusercontent.com/85518330/123559183-1db27900-d760-11eb-9654-887b93c76b24.png)
   
   
   
   ![VBA_Challenge_2018](https://user-images.githubusercontent.com/85518330/123559187-2440f080-d760-11eb-8a15-0ea5aacf1f6c.png)


## Results of Refactoring our code

Our refactored code reduced run time from about 0.66- 0.70 secs for the original code to between .12 and .14 secs for the refactored code. That is about a 500% improvement. 

# Advantages & Disadvantages of Refactoring

Coding as an art and science is ever evolving and improving, what was considered best practice a few years ago could be redundant today, so looking at our codes often to see how we can improve them is required 

**Refactoring of code is advantageous when** 
     
     *It makes the new code more robust
     *Increases flexibility
     *Simplifies it 
     *Provides additional clarity
     *Adds introspection 
     *Reduces deployment times
     *Boosts performance


**Refactoring of code is disadvantageous when**
   
     *It’s done purely to improve aesthetics of the code but not functionality
     *There is poor knowledge of the code 
     *When it's done on codes that are infrequently used 
     *When it takes so much time that it becomes counterproductive for the business  


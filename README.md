# Stocks_Analysis
In this project, we will explore green energy stock performance by analyzing financial data using VBA.
# Stock-analysis
In this project, we will explore green energy stock performance by analyzing financial data using VBA.

## Background

Green energy investments are very popular especially as the world increasingly moves toward a clean energy future. <br>
Many investors believe that as fossil fuels get used up, there will be more and more reliance on alternative energy production. <br>

There are many forms of green energy to invest in, including hydroelectricity, wind energy, geothermal energy, and bioenergy. 
However, our clients are decided to invest all their money into DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. <br>

Out of concern about diversifying their funds, they have requested an analysis of other green energy stocks, in addition to DAQO’s stock. 

The green energy stock data is provided to us by a financial advisor as an Excel file.
We will be using an extension to Excel built to automate tasks: Visual Basic for Applications, usually referred to as just VBA. 

VBA is a programming language that interacts with Excel. It can read and write to cells and worksheets, make calculations, and use complex logic to perform analyses. 
Using code to automate analyses, allows us to reuse it with any stock and reduces the chance of accidents and errors. 


### Purpose
Determine which green energy stock is a better investment for our clients. 

## Objectives
1. find the total daily volume and yearly return for the DAQO stock.



## Resources
- Data Sources: green_stocks.xlsx, green_stocks.xlsm
- Software: Microsoft Excel
- Libraries & Packages: Visual Basic for Applications (VBA)
- Online Tools: [Stocks_Analysis_VBA GitHub Repository](https://github.com/Magzzie/Stocks_Analysis_VBA)

## Methods & Code
Visual Basic for Applications, which is typically referred to as "VBA," is often used in the finance industry.
VBA provides essentially infinite extensibility to Excel. Using code to automate tasks decreases the chance of errors and reduces the time needed to run analyses, 
especially if they need to be done repeatedly. 

- We Calculated the total daily trading volumn of the DAQU stock using conditionals and a for loop to go through all the rows with the ticker DQ and add the daily volumn, 
then export the sum to a cell in a different sheet (DQ Analysis worksheet) using the right sheet activation command. 
- We calculated the difference between the starting price and ending price of DQ stock in 2018 using conditionals inside a for loop through all the rows with the DQ ticker in the first cell. 
- To run analyses on all of the stocks, we needed to create a program flow that loops through all of the tickers.



## Results
- Total number of records is 3,013 for 12 green energy stocks for each of the years 2017 and 2018.
- The records shows the ticker for each stock, date, daily opening, high, low, closing, and adjusted closing values in addition to the daily volume traded. 
- Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. 
- The yearly return is the percentage difference in price from the beginning of the year to the end of the year.
- We studied the daily volumn and yearly return of the DAQO Stock first. 
- The total daily volume of DQ stock is shown in the "DQ Analysis" worksheet: DQ traded 107,873,900 shares in 2018.
- Upon calculating the difference between the ending price and starting price of trading for the DQ stock we found that Daqo dropped over 63% in 2018. 
- Therefore, we started looking into other stocks that might be a better green energy investment. 





## Recommendations




---
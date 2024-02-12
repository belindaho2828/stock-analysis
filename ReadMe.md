# stock-analysis

Stock analysis for years 2018-2020, including the annual change between that year's open price and close price, % change, and total stock volume per Ticker per year.

Analysis also includes Ticker and value for the greatest % increase, greatest % decrease, and greatest stock volume for that year

VBA code for this analysis uses a For Loop that runs through each row and stores the following variables for calculation for each worksheet in the workbook:

#Ticker
	Reads the Ticker for each row and prints the current Ticker in column "I" when the statement If current row does not equal the previous row ticker evaluates as true. This is reading and printing the ticker at the first row of a new ticker. This is the opening price of the first day of the year for that ticker.
	Prints the Ticker in a new row in the Summary Table Row by adding 1 row to the existing SumTableRow variable and replacing the variable at each loop. The Variable was set at 2nd row under the header before the first loop.

#TickerOpen
	Stores the opening price in column C for the first row of a new ticker (i.e., when the statement If current row (i) does not equal the previous row (i -1) ticker evaluates as true).

#TickerClose
	Reads the closing price in column F for each row until it reaches the last row of the ticker (i.e., when the statement If current row (i) does not equal the next row (i + 1) ticker evaluates as true), at which point it will calculate the YEDelta.
	## YEDelta: calculated as the difference between the TickerClose at that last row and TickerOpen stored when we entered the first row of that ticker. 
		### Print and Format: YEDelta is then printed in column J and cell is formatted to red if value is negative, green otherwise
	## PercentChange: calculated as YEDelta / TickerOpen 
		### Print and Format: PercentChange is then printed in column K cell is formatted to red if value is negative, green otherwise
		### Max Min: Before the loop, HighestDelta and LowestDelta (i.e., Max and Min) are set to 0. At each last row of a ticker, the PercentChange is evaluated against HighestDelta and LowestDelta separately to determine if it is higher or lower respectively. If it is, it replaces that variable. At the end of the data set, it prints the HighestDelta (Max) and LowestDelta (Min) in Cells (2, Q) and (3, Q) respectively along with the Ticker in the corresponding row, column P.


#StockVolume
	Stores the opening price in column G for the first row of a new ticker (i.e., when the statement If current row (i) does not equal the previous row (i -1) ticker evaluates as true). For each row after that the Ticker is the same, adds the current row's stock volume to the previous stock volume and replace the stock volume variable for the next loop.
	At each last row of a ticker (i.e., when the statement If current row (i) does not equal the next row (i + 1) ticker evaluates as true), total stock volume will print in column L for that ticker
	## Max: Before the loop, HighestVolume (i.e., max) are set to 0. At each last row of a ticker, the StockVolume is evaluated against HighestVolume to determine if it is higher . If it is, it replaces that variable. At the end of the data set, it prints the HighestVolume (Max) in Cell (4, Q) along with the Ticker in Cell (4, P).
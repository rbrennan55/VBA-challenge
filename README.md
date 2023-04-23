# VBA-challenge

## Due: April 27, 2023
## Task:  Create a script that loops through all the stocks for one year and outputs the following information:
      •	The ticker symbol
      •	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
      •	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
      •	The total stock volume of the stock.
      •	The greatest persent increase
      •	The lowest perscnt increase
      •	The greatest total volume
      •	The VBA script can run on all sheets successfully.

## Solution files submitted:
- Multiple_year_stock_data.xlsm
- Multiple_year_stock_data_2018_screencap.PNG
- Multiple_year_stock_data_2019_screencap.PNG
- Multiple_year_stock_data_2020_screencap.PNG
- ReadMe

# Consideration Notes 1:

  - Needed help from AskBCS.  There was an error with: 

  - // id_row = ws.Range("L2:L" & lastrow).Find(greatest_stock_volume, , xlValues).Row

  - The above line caused an "Run-Time Error 91: Object Variable or With Block Variable Not Set"

  - It was deduced with AskBCS that there was a limitation with the 'Find' command and the length of the variable 'greatest_stock_volume'.  

  - With help from AskBCS, I replaced faulty code with: 

  - // id_row = ws.Application.WorksheetFunction.Match(greatest_stock_volume, ws.Range("L2:L" & lastrow), 0)
                        
# Consideration Notes 2:

  - Chose to use an Array to display the headers as I wanted to practise using Arrays.

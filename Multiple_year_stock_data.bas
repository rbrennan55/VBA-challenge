Attribute VB_Name = "Multiple_year_stock_data_Module"
'   This script loops through all the stocks for the year on all sheets and outputs the following information:
'   The ticker symbol
'   Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'   The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'   The total stock volume of the stock.
'   The greatest persent increase
'   The lowest perscnt increase
'   The greatest total volume


Sub getStockSummary()
    
  ' Set for each Worksheet
    For Each ws In Worksheets
    
  ' Set variable for holding the stock symbol
    Dim stock_symbol As String
  
  ' Set variable for holding the total stock volume
    Dim stock_volume As Double
    
  ' Set variable for holding the opening, closing, yearly change and percent change stock price
    Dim stock_opening_price As Double
    Dim stock_closing_price As Double
    Dim stock_yearly_change As Double
    Dim stock_percent_change As Double
    Dim greatest_stock_percent As Double
    Dim lowest_stock_percent As Double
    Dim greatest_stock_volume As Double
       
 
    Dim id_row As Double
       
      
    ' Set variable for holding the total total stock summary table
    Dim stock_summary_table_row As Integer
    
    ' Set variable for header titles
    Dim header_titles() As Variant
    Dim i_header As Integer
    
    ' Print header tiles
    header_titles = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value", "Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
        For i_header = 9 To 17
            ws.Cells(1, i_header).Value = header_titles(i_header - 9)
            Next i_header
        
        For i_header = (i_header - 16) To (i_header - 14)
            ws.Cells(i_header, 15).Value = header_titles(i_header + 7)
            Next i_header
                
  
  ' Set first location index for the summary table
    stock_summary_table_row = 2
        
  ' Set column headers
   
  
  ' Determine the Last Row
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  
  ' Loop through all stock data
    stock_opening_price = ws.Cells(2, 3).Value
    greatest_stock_percent = 0
    
    For i = 2 To lastrow
               
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       
            ' Set the Stock symbol name
              stock_symbol = ws.Cells(i, 1).Value
            
            ' Add to the stock Total
              stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            ' Print the Stock Symbol in the Summary Table
              ws.Range("I" & stock_summary_table_row).Value = stock_symbol
              
            ' Set year closing price
              
              stock_closing_price = ws.Cells(i, 6).Value
            
            ' Set Yearly Change
              stock_yearly_change = (stock_closing_price - stock_opening_price)
            
            ' Print Yearly Change + Colour
              ws.Range("J" & stock_summary_table_row).Value = stock_yearly_change
              
              If stock_yearly_change < 0 Then
                ws.Range("J" & stock_summary_table_row).Interior.ColorIndex = 3
               ElseIf stock_yearly_change > 0 Then
                ws.Range("J" & stock_summary_table_row).Interior.ColorIndex = 4
               Else
                ws.Range("J" & stock_summary_table_row).Interior.ColorIndex = 0
               End If
              
            ' Set + Print Percent Change
              stock_percent_change = Round((((stock_closing_price / stock_opening_price) * 100) - 100), 2)
              ws.Range("K" & stock_summary_table_row).Value = stock_percent_change
                          
            ' Print the stock total volumn into the Summary Table
              ws.Range("L" & stock_summary_table_row).Value = stock_volume
      
            ' Add one to the summary table row
              stock_summary_table_row = stock_summary_table_row + 1
                 
                    
            ' Set opening stock price for next symbol
              stock_opening_price = ws.Cells(i + 1, 3).Value
              
            ' Reset the stock volume, opening and closing price
              stock_volume = 0
                          
    
      ' If the cell immediately following a row is the same stock symbol
      Else
       ' Add to the stock volume total
         stock_volume = stock_volume + ws.Cells(i, 7).Value
           
                
     End If
    Next i
            ' Return greatest increase + decrease + volume percent change
              
              ' Code that causes issues with larger eponential data - caused an Error "91"
              ' greatest_stock_percent = ws.Application.Max(ws.Range("K2:K" & lastrow))
              ' id_row = ws.Range("K2:K" & lastrow).Find(greatest_stock_percent, , xlValues).Row
              
              
              ' Get greatest percent value and row number of data
              greatest_stock_percent = ws.Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
              id_row = ws.Application.WorksheetFunction.Match(greatest_stock_percent, ws.Range("K2:K" & lastrow), 0)
              
              ' Print greatest percent stock symbol and value
              stock_symbol = ws.Cells(id_row + 1, 9).Value
              ws.Cells(2, 16).Value = stock_symbol
              ws.Cells(2, 17).Value = greatest_stock_percent
             
              
              ' Code that causes issues with larger eponential data
              ' lowest_stock_percent = ws.Application.Min(ws.Range("K2:K" & lastrow))
              ' id_row = ws.Range("K2:K" & lastrow).Find(lowest_stock_percent, , xlValues).Row
              
              
              ' Get lowest percent value and row number of data
              lowest_stock_percent = ws.Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
              id_row = ws.Application.WorksheetFunction.Match(lowest_stock_percent, ws.Range("K2:K" & lastrow), 0)
              
              ' Print lowest percent stock symbol and value
              stock_symbol = ws.Cells(id_row + 1, 9).Value
              ws.Cells(3, 16).Value = stock_symbol
              ws.Cells(3, 17).Value = lowest_stock_percent
                        
              
              ' Code that causes issues with larger eponential data
              ' greatest_stock_volume = ws.Application.Max(ws.Range("L2:L" & lastrow))
              ' id_row = ws.Range("L2:L" & lastrow).Find(greatest_stock_volume, , xlValues).Row
              
              
              ' Get greatest stock colume value and row number of data
              greatest_stock_volume = ws.Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
              id_row = ws.Application.WorksheetFunction.Match(greatest_stock_volume, ws.Range("L2:L" & lastrow), 0)
              
              ' Print greatest stock volume symbol and value
              stock_symbol = ws.Cells(id_row + 1, 9).Value
              ws.Cells(4, 16).Value = stock_symbol
              ws.Cells(4, 17).Value = greatest_stock_volume
     
       
      
 Next ws

End Sub


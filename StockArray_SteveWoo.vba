Sub Stock():

'Setting variable for worksheet, to apply codes to each worksheet'

   Dim ws As Worksheet      

'Setting up Forloop through each worksheets'
  
       For Each ws In Worksheets:

'Setting variables and data types, inserting Titles for each data columns'
                      
           Dim ticker As String
           Dim year_change As Double
           Dim percent_change As Double
           Dim volume As Single
           Dim decvolume As Double
           decvolume = CDec(volume)

'To exclude error by including string titles, setting up the range from row 2'

           Dim datas As Integer
           datas = 2

'To obtain statistical values of the stock, set up variables for each value'
           
           Dim open_value As Single
           Dim close_value As Single
                      
           Dim Initialvalue As Integer
           Initialvalue = 2
           
           Dim head_ticker As Integer
           head_ticker = 9
           
           Dim head_yearly As Integer
           head_yearly = 10
           
           Dim head_percent As Integer
           head_percent = 11
           
           Dim head_volume As Integer
           head_volume = 12
           
           Dim head_ticker_gr As Integer
           head_ticker_gr = 16
           
           Dim head_value_gr As Integer
           head_value_gr = 17
           
           Dim great_inc As String
           great_inc = "Greatest % Increase"
           
           Dim great_dec As String
           great_dec = "Greatest % Decrease"
           
           Dim great_vol As String
           great_vol = "Greatest Total Volume"

'Setting up titles for each statistical values'
           
           ws.Cells(1, head_ticker) = "Ticker"
           ws.Cells(1, head_yearly) = "Yearly Change"
           ws.Cells(1, head_percent) = "Percent Change"
           ws.Cells(1, head_volume) = "Total Stock Volume"
           ws.Cells(1, head_ticker_gr) = "Ticker"
           ws.Cells(1, head_value_gr) = "Value"
           ws.Cells(2, 15) = great_inc
           ws.Cells(3, 15) = great_dec
           ws.Cells(4, 15) = great_vol

'Setting up the last row of the data by counting each rows with data'
           
           Dim Lastrow As Long
           Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
           
'Setting up For loop for each rows of data, setting i = 2 to exclude the first row with string titles'           

                For i = 2 To Lastrow:
               
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                   
                        ticker = Cells(i, 1).Value
                        
                        open_value = ws.Cells(Initialvalue, 3)
                        close_value = ws.Cells(i, 6)
                   
                        year_change = close_value - open_value
                        ws.Range("J" & datas).Value = year_change
                        
                        percent_change = year_change / open_value
                        ws.Range("K" & datas).Value = percent_change
                        ws.Range("K2:K999999").NumberFormat = "0.00%"
                        
                        ws.Range("I" & datas).Value = ticker
                        ws.Range("L" & datas).Value = volume
                   
                        decvolume = ws.Cells(i, 7).Value + 1
                        datas = datas + 1
                        volume = 0

'Resetting the cell selections of open_value and close_value, to react to the next choice of i'
                   
                        open_value = Cells(1, 1).Select
                        close_value = Cells(1, 1).Select

                    Else

'Setting the volume to be tracked while the first column value is the same as the proceeding ones'

                        volume = Cells(i, 7).Value + volume
                   
                    End If
                Next i

'Setting up statistical formulas'
                
        Dim greatest_inc As Single
        greatest_inc = Application.WorksheetFunction.Max(ws.Range("K2:K999999"))
        ws.Cells(2, 17) = greatest_inc
        ws.Range("Q2").NumberFormat = "0.00%"
        
        Dim greatest_dec As Single
        greatest_dec = Application.WorksheetFunction.Min(ws.Range("K2:K999999"))
        ws.Cells(3, 17) = greatest_dec
        ws.Range("Q3").NumberFormat = "0.00%"
        
        Dim greatest_vol As Double
	Dim decgreatest_vol As Double
	decgreatest_vol = CDec(greatest_vol)
        greatest_vol = Application.WorksheetFunction.Sum(ws.Range("L2:L999999"))
        ws.Cells(4, 17) = decgreatest_vol

'Setting up conditional coloring of the cell, by creating a variable rg to refer to J columns, which is our interest'
        
        Dim rg As Range
        Set rg = ws.Range("J2", ws.Range("J2").End(xlDown))
     
        Dim condition1 As FormatCondition
        Dim condition2 As FormatCondition
              
        Set condition1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
        Set condition2 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)

        condition1.Interior.Color = vbGreen
        condition2.Interior.Color = vbRed
                
        Next ws
End Sub






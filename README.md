# VBA-Challenge

#### See below link to obtain the original data sheet to run the script:
#### [SHARED LINK] https://drive.google.com/open?id=1mgT9x6cVZKlJ7A1nQLepDLdS-o9VDFk5

###### When the scripts are run with the shared data sheet, the followings are the results:

###### ![VBA Script Sheets 2014](/images/Multiple_year_stock_data_2014.jpg)

###### ![VBA Script Sheets 2015](/images/Multiple_year_stock_data_2015.jpg)

###### ![VBA Script Sheets 2016](/images/Multiple_year_stock_data_2016.png)

###### The following codes are used to set up the forloops in the script:

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
                        
###### The script automatically applies to all worksheets in active state.

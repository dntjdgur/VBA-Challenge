Attribute VB_Name = "Module1"
Sub Stock():
    
    Dim ws As Worksheet
    
        For Each ws In Worksheets:
                    
            Dim ticker As String
            Dim year_change As Long
            Dim percent_change As Double
            Dim volume As Single
        
            Dim datas As Integer
            datas = 2
                
            Dim open_value As Single
            Dim close_value As Single
            
            Dim op_v As Single
            Dim cl_v As Single
            
            Dim Lastrow As Long
            Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                            
            For i = 2 To Lastrow:
        
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                    ticker = Cells(i, 1).Value
                    
                    open_value = Range(Cells(i, 3), Cells(i + 1, 3)).End(xlUp).Select
                    close_value = Range(Cells(i, 6), Cells(i + 1, 6)).End(xlDown).Select
                                        
                    year_change = close_value - open_value
                    ws.Range("J" & datas).Value = year_change
                    
                    percent_change = close_value / open_value
                    ws.Range("K" & datas).Value = percent_change
                    
                    ws.Range("I" & datas).Value = ticker
                    ws.Range("L" & datas).Value = volume
                                
                    volume = Cells(i, 7).Value + volume
                    
                    datas = datas + 1
            
                    volume = 0
                    open_value = Cells(1, 1).Select
                    close_value = Cells(1, 1).Select
                    
                Else
                    
                    volume = Cells(i, 7).Value + volume
                    Range("C2:F2").Select
                    Range(Cells(i, 3), Cells(i, 3)).Activate
                    Range(Cells(i, 6), Cells(i, 6)).Activate

                    
                End If
        
            Next i
    
        Next ws
            
End Sub










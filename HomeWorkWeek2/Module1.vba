Attribute VB_Name = "Module1"
Sub LoopScript()
       
        Dim ws As Worksheet
    
        
        For Each ws In Worksheets
        
            
            Dim Ticker As String
            
            
            Dim Total As Double
            Total = 0
            
            Dim OpeningPrice As Double
            OpeningPrice = 0
            
            Dim ClosingPrice As Double
            ClosingPrice = 0
            
            Dim YearlyChange As Double
            YearlyChange = 0
            
            Dim Precent As Double
            Precent = 0
            
            
            Dim Table As Long
                Table = 2
            
          
            Dim Lastrow As Long
            
            Dim i As Long
            
            Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"
            
            
            OpeningPrice = ws.Cells(2, 3).Value
            
            For i = 2 To Lastrow
            
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    
                    Ticker = ws.Cells(i, 1).Value
                    
                    ClosingPrice = ws.Cells(i, 6).Value
                    
                    YearlyChange = ClosingPrice - OpeningPrice
                   
                    If OpeningPrice <> 0 Then
                        Percent = (YearlyChange / OpeningPrice) * 100#
                    Else
                    
                    End If
                    
                    
                    Total = Total + ws.Cells(i, 7).Value
                    ws.Range("I" & Table).Value = Ticker
                    ws.Range("J" & Table).Value = YearlyChange
                    
                     
                    ws.Range("K" & Table).Value = ((Percent) & "%")
                    
                    ws.Range("L" & Table).Value = Total
                    
                   
                    Table = Table + 1
                    
                    YearlyChange = 0
                    Percent = 0
                    ClosingPrice = 0
                  
                    OpeningPrice = ws.Cells(i + 1, 3).Value
                    
                    Total = 0
                    
                
                Else
                    
                    Total = Total + ws.Cells(i, 7).Value
                End If
                
          
            Next i
            
         Next ws
End Sub

# module2_VBA_Challenge
Module 2 VBA Challenge

    Sub stock_analysis()
        Dim ws As Worksheet
        
        
        
        For Each ws In Worksheets
            Dim Ticket_Table_Row As Long
            Ticker_Table_Row = 2
            Dim opening As Double
            Dim yearly_change As Double
            yearly_change = 0
            Dim total As Double
            Dim lastrow As Long
            Dim Yearly_change_percent As Double
            Yearly_change_percent = 0
            Dim Max_Percent_Ticker As String
            Max_Percent_Ticker = ""
            Dim Min_Percent_Ticker As String
            Min_Percent_Ticker = ""
            Dim Max_Percent As Double
            Max_Percent = 0
            Dim Min_Percent As Double
            Min_Percent = 0
            Dim Max_Volume_Ticker As String
            Max_Volume_Ticker = ""
            Dim Max_Volume As Double
            Max_Volume = 0
            
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("k1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            opening = ws.Cells(2, 3).Value
            total = 0
            
            For i = 2 To lastrow
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    
                    ws.Range("I" & Ticker_Table_Row).Value = ws.Cells(i, 1).Value
                    
                    yearly_change = ws.Cells(i, 6).Value - opening
                    ws.Range("J" & Ticker_Table_Row).Value = yearly_change
                    total = total + ws.Cells(i, 7).Value
                    ws.Range("L" & Ticker_Table_Row).Value = total
                    
                    
                    
                    
                    If opening <> 0 Then
                        Yearly_change_percent = (yearly_change / opening) * 100
                    End If
                    
                    If yearly_change > 0 Then
                    
                        ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 4
                        
                    ElseIf yearly_change < 0 Then
                        ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 3
                        
                    End If
                    
                    ws.Range("K" & Ticker_Table_Row).Value = (CStr(Yearly_change_percent) & "%")
                    
                    ws.Range("K" & Ticker_Table_Row).Interior.ColorIndex = 0
                    
                    Ticker_Table_Row = Ticker_Table_Row + 1
                    opening = ws.Cells(i + 1, 3).Value
                    
                    If (Yearly_change_percent > Max_Percent) Then
                        Max_Percent = Yearly_change_percent
                        Max_Percent_Ticker = ws.Cells(i, 1).Value
                        
                    ElseIf (Yearly_change_percent < Min_Percent) Then
                        Min_Percent = Yearly_change_percent
                        Min_Percent_Ticker = ws.Cells(i, 1).Value
                    End If
                    
                    If (total > Max_Volume) Then
                        Max_Volume = total
                        Max_Volume_Ticker = ws.Cells(i, 1).Value
                    End If
                    
                    Yearly_change_percent = 0
                    total = 0
                    
                Else
                    total = total + ws.Cells(i, 7).Value
                    
                    
                    
                End If
                
            Next i
                
            ws.Range("P2").Value = Max_Percent_Ticker
            ws.Range("Q2").Value = (CStr(Max_Percent) & "%")
            ws.Range("P3").Value = Max_Percent_Ticker
            ws.Range("Q3").Value = (CStr(Min_Percent) & "%")
            ws.Range("P4").Value = Max_Volume_Ticker
            ws.Range("Q4").Value = Max_Volume
            
        Next ws
            
            
            
    End Sub


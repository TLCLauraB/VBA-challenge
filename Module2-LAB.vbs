Attribute VB_Name = "Module2"
Sub StockInfoProcess2()
    
    Dim ws As Worksheet
    
    Dim Ticker As String
    
    Dim OpenPrice As Double
    
    Dim ClosePrice As Double
    
    Dim YearlyChange As Double
    
    Dim PercentChange As Double
    
    Dim LastRow As Long
    
    Dim SummaryRow As Long
    
    Dim Volume As Double
    
    'r for Row instead of i
    Dim r As Long
    
        'For Part Two
            Dim MaxIncreaseTicker As String
            
            Dim MaxIncreasePercent As Double
            
            Dim MaxDecreaseTicker As String
            
            Dim MaxDecreasePercent As Double
            
            Dim MaxVolumeTicker As String
            
            Dim MaxVolume As Double
        
    For Each ws In Worksheets
        
        SummaryRow = 2
        
        Volume = 0
        
        ws.Columns("J:L").AutoFit
        
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Yearly Change"
        
        ws.Range("K1").Value = "Percent Change"
        
        ws.Range("L1").Value = "Total Stock Volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For r = 2 To LastRow
            
            If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value Then
                
                Ticker = ws.Cells(r, 1).Value
                
                OpenPrice = ws.Cells(r, 3).Value
                
                Volume = 0
                
            End If
            
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Or r = LastRow Then
                
                ClosePrice = ws.Cells(r, 6).Value
                
                YearlyChange = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                
                    PercentChange = YearlyChange / OpenPrice
                    
                Else
                    PercentChange = 0
                    
                End If
                
                Volume = Volume + ws.Cells(r, 7).Value
                
                With ws.Range("I" & SummaryRow)
                
                    .Value = Ticker
                    
                    .Offset(0, 1).Value = YearlyChange
                    
                    .Offset(0, 2).Value = PercentChange
                    
                    .Offset(0, 3).Value = Volume
                    
                    .Offset(0, 2).NumberFormat = "0.00%"
                    
                    Select Case YearlyChange
                        
                        Case Is > 0 'In Green
                            .Offset(0, 1).Interior.ColorIndex = 4
                        
                        Case Is < 0 'In Red
                            .Offset(0, 1).Interior.ColorIndex = 3
                        
                        Case Else 'In Yellow
                            .Offset(0, 1).Interior.ColorIndex = 6
                        
                    End Select
                    
                End With
                
                SummaryRow = SummaryRow + 1
                
            Else
            
                Volume = Volume + ws.Cells(r, 7).Value
                
            End If
            
        Next r
        
        ws.Range("P1").Value = "Ticker"
        
        ws.Range("Q1").Value = "Value"
        
            ws.Columns("I:L").AutoFit
            
            ws.Range("O2").Value = "Greatest % Increase"
            
            ws.Range("O3").Value = "Greatest % Decrease"
            
            ws.Range("O4").Value = "Greatest Total Volume"
            
            MaxIncreaseTicker = ""
            
            MaxIncreasePercent = -9999#
            
            MaxDecreaseTicker = ""
            
            MaxDecreasePercent = 9999#
            
            MaxVolumeTicker = ""
            
            MaxVolume = -9999
        
    For r = 2 To SummaryRow - 1
  
            If ws.Range("K" & r).Value > MaxIncreasePercent Then
                
                MaxIncreaseTicker = ws.Range("I" & r).Value
                
                MaxIncreasePercent = ws.Range("K" & r).Value
                
            End If
            
            If ws.Range("K" & r).Value < MaxDecreasePercent Then
                
                MaxDecreaseTicker = ws.Range("I" & r).Value
                
                MaxDecreasePercent = ws.Range("K" & r).Value
                
            End If
            
            If ws.Range("L" & r).Value > MaxVolume Then
                
                MaxVolumeTicker = ws.Range("I" & r).Value
                
                MaxVolume = ws.Range("L" & r).Value
                
            End If
            
    Next r
        
        With ws.Range("P2")
        
            .Value = MaxIncreaseTicker
            
            .Offset(0, 1).Value = MaxIncreasePercent
            
            .Offset(0, 1).NumberFormat = "0.00%"
            
        End With
        
        With ws.Range("P3")
        
            .Value = MaxDecreaseTicker
            
            .Offset(0, 1).Value = MaxDecreasePercent
            
            .Offset(0, 1).NumberFormat = "0.00%"
            
        End With
        
        With ws.Range("P4")
        
            .Value = MaxVolumeTicker
            
            .Offset(0, 1).Value = MaxVolume
            
        End With
        
        ws.Columns("O:Q").AutoFit
        
    Next ws
    
        MsgBox "Done! This is really hard!"
    
End Sub


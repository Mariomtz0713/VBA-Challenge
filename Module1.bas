Attribute VB_Name = "Module1"
Sub Sheet_2018()

    For Each ws In Worksheets
            
        Dim i As Double
            
        Dim Ticker As String
        
        Dim Open_Price As Double
        
        Dim Close_Price As Double
        
        Dim Yearly_Change As Double
        
        Dim Volume As Double
        Volume = 0
        
        Dim Table_Row As Integer
        Table_Row = 2
        
        LastRow_Main = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'First Yearly Change
        Open_Price = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow_Main
            
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            
            'Ticker
            Ticker = ws.Cells(i, 1).Value
            
            ws.Range("K" & Table_Row) = Ticker
            
            'Yearly Change
            Close_Price = ws.Cells(i, 6).Value
            
            Yearly_Change = Close_Price - Open_Price
            
            ws.Range("L" & Table_Row) = Yearly_Change
            
                'Color
                If (Yearly_Change > 0) Then
                ws.Range("L" & Table_Row).Interior.ColorIndex = 4
                
                Else
                ws.Range("L" & Table_Row).Interior.ColorIndex = 3
                
                End If
            
            'Percent Change
            ws.Range("M" & Table_Row).Value = Yearly_Change / Open_Price
            
            'Percent Format pulled from (https://excelvbatutor.com/vba_lesson9.htm)
            ws.Range("M" & Table_Row).Value = Format(ws.Range("M" & Table_Row).Value, "0.00%")
                
            'Volume
            Volume = Volume + ws.Cells(i, 7).Value
            
            ws.Range("N" & Table_Row).Value = Volume
            
            'Next Ticker Prep
            Open_Price = ws.Cells(i + 1, 3)
            
            Table_Row = Table_Row + 1
            
            Volume = 0
    
            Else
            Volume = Volume + ws.Cells(i, 7).Value
            
            End If
            
        
        Next i
        
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
        
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = "Greatest % Decrease"
        ws.Cells(4, 17).Value = "Greatest Total Volume"
        ws.Cells(1, 18).Value = "Ticker   "
        ws.Cells(1, 19).Value = "Value"
            
        
        'Functionality
        Dim j As Double
        
        Dim Max_Percent As Double
        Max_Percent = 0
        
        Dim Max_Percent_Ticker As String
        
        LastRow_Percent = ws.Cells(Rows.Count, 13).End(xlUp).Row
        
        For j = 2 To LastRow_Percent
            
            If ws.Cells(j, 13).Value >= Max_Percent Then
            
            Max_Percent = ws.Cells(j, 13).Value
            Max_Percent_Ticker = ws.Cells(j, 11).Value
            
            End If
            
            
        Next j
               
        ws.Cells(2, 18).Value = Max_Percent_Ticker
        ws.Cells(2, 19).Value = Format(Max_Percent, "0.00%")
        
        
        Dim k As Integer
        
        Dim Min_Percent As Double
        Min_Percent = 0
        
        Dim Min_Percent_Ticker As String
        
        For j = 2 To LastRow_Percent
            
            If ws.Cells(j, 13).Value <= Min_Percent Then
            
             Min_Percent = ws.Cells(j, 13).Value
             Min_Percent_Ticker = ws.Cells(j, 11).Value
             
            End If
            
        Next j
        
        ws.Cells(3, 18).Value = Min_Percent_Ticker
        ws.Cells(3, 19).Value = Format(Min_Percent, "0.00%")
        
        
        Dim Max_Volume As Double
        Max_Volume = 0
        
        Dim Max_Volume_Ticker As String
        
        For j = 2 To LastRow_Percent
            
            If ws.Cells(j, 14).Value >= Max_Volume Then
            
            Max_Volume = ws.Cells(j, 14).Value
            Max_Volume_Ticker = ws.Cells(j, 11).Value
            
            End If
            
            
        Next j
        
        ws.Cells(4, 18).Value = Max_Volume_Ticker
        ws.Cells(4, 19).Value = Max_Volume
        
        'Autofit Colum width pulled from (https://excelchamps.com/vba/autofit/)
        ws.UsedRange.EntireColumn.AutoFit
        
    Next ws
    
    MsgBox ("           Done!")
    
End Sub


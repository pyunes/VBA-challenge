VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
'        Current row
        Dim i As Long
'        start row of ticker
        Dim j As Long
'        index counter to fill Ticker
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
'        following variables as Double due to decimals
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        WorksheetName = ws.Name
        
'        Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        TickCount = 2
        
'        start on row 2 since row 1 is a header
        j = 2
        
'       xlUP  Estamos indicando que desde la fila indicada , "suba" hasta que encuentre un valor
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRowA
            
'                double check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
'                show ticker in column I9
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
'                calculate and show Yearly Change in column J10
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
'                    conditional formating
                    If ws.Cells(TickCount, 10).Value < 0 Then
                
'                color red
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
'                    color green
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
'                    calculate and write percent change in column K11
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
'                  Conditional formating of percentage using VBA
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
'                calculate and write total volume in column L12
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                TickCount = TickCount + 1
                
'                start a new row for the ticker
                j = i + 1
                
                End If
            
            Next i
            
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
'        display the values on the respective cells
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
'           ForLoop for showing the results
            For i = 2 To LastRowI
            
'               if its not a higher value then it stays the same
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
'               if its not a higher value then it stays the same
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
'                chooses the smaller value, if not it stays the same
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
'           write the results in the worksheet
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            
'TEST it stresses me out the cell doesn't change, source: https://analysistabs.com/excel-vba/change-row-height-column-width/
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub

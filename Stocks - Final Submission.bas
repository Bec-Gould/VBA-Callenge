Attribute VB_Name = "Module3"
    
 Sub Stocks()

    'Set variables for worksheets
    Dim ws As Worksheet
    Dim WS_Count As Integer
    Dim J As Integer
    Set ws = ActiveSheet
     
    'Loop Worksheets
    For Each ws In Worksheets
    
     'Add column titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    'Set Last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set sheet variables
    Dim Ticker_Name As String
    Dim Volume_Total As LongLong
    Volume_Total = 0
    Dim Percentage_change As Double
    Percentage_change = 0
    Dim Summary_Table_Row As LongLong
    Summary_Table_Row = 2
    Dim y As Double
    y = 2

    
        'Loop through data
        For I = 2 To lastRow

            'Check if we are still within the same Ticker name, if not...
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
            'Calculate the Yearly Change
            OpenValue = ws.Cells(y, 3).Value
            CloseValue = ws.Cells(I, 6).Value
            YearlyChange = CloseValue - OpenValue
     
            'Add calculation to Yearly Change row
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
        
            'Calculate the Percentage change
            Percentage_change = YearlyChange / OpenValue
       
            'Add calculation to Percentage change row and round to 2 decimal places
            ws.Range("K" & Summary_Table_Row).Value = Percentage_change
           
            'Format cells as percentage
            ws.Columns("K").NumberFormat = "0.00%"
        
            'Set the ticker name
            Ticker_Name = ws.Cells(I, 1).Value
    
            'Print the Ticker Name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    

            'Add to the total volume
            Volume_Total = Volume_Total + ws.Cells(I, 7).Value
    
            'Print the Volume in the summary table
            ws.Range("L" & Summary_Table_Row).Value = Volume_Total
    
            'Add one to the Summary table
            Summary_Table_Row = Summary_Table_Row + 1

                'Set Last Row
                lastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
                
                For x = 2 To lastRow
    
                If ws.Cells(x, 10).Value < 0 Then
                    ws.Cells(x, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(x, 10).Interior.ColorIndex = 4
                
                End If
            Next x
 
            'Reset the Volume total
            Volume_Total = 0
            
                'If the cell immediately following a row is the same
                Else
        
                'Add to the total volume
                Volume_Total = Volume_Total + ws.Cells(I, 7).Value
            

            End If
    
        Next I

    Next ws
    
End Sub


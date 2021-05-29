Sub waves()

    For Each ws In Worksheets
    
    Dim wave, start, Row As Integer
        
    wave = 0
    start = 2
    
    lastRow = (ws.Cells(Rows.Count, 1).End(xlUp).Row)
     
    For Row = 3 To lastRow
        
        If (ws.Cells(Row, 2).Value = 1 And ws.Cells(Row, 2).Value <> ws.Cells((Row - 1), 2).Value) Or Row = lastRow Then
            
            wave = wave + 1
            
            Sheets.Add
            ActiveSheet.Name = ("Wave" + Str(wave))
            
            Sheets("Speed Dating Data").Range("A1").Copy
            Range("A1").Select
            ActiveSheet.Paste
            Sheets("Speed Dating Data").Range(ws.Cells(start, 1), ws.Cells((Row - 1), 1)).Copy
            Range("A2").Select
            ActiveSheet.Paste
                        
            Sheets("Speed Dating Data").Range("FC1").Copy
            Range("B1").Select
            ActiveSheet.Paste
            Sheets("Speed Dating Data").Range(ws.Cells(start, 159), ws.Cells((Row - 1), 159)).Copy
            Range("B2").Select
            ActiveSheet.Paste
                        
            Sheets("Speed Dating Data").Range("BQ1").Copy
            Range("C1").Select
            ActiveSheet.Paste
            Sheets("Speed Dating Data").Range(ws.Cells(start, 69), ws.Cells((Row - 1), 69)).Copy
            Range("C2").Select
            ActiveSheet.Paste
                        
            Sheets("Speed Dating Data").Range("BP1").Copy
            Range("D1").Select
            ActiveSheet.Paste
            Sheets("Speed Dating Data").Range(ws.Cells(start, 68), ws.Cells((Row - 1), 68)).Copy
            Range("D2").Select
            ActiveSheet.Paste
                        
            Sheets("Speed Dating Data").Range("AU1").Copy
            Range("E1").Select
            ActiveSheet.Paste
            Sheets("Speed Dating Data").Range(ws.Cells(start, 47), ws.Cells((Row - 1), 47)).Copy
            Range("E2").Select
            ActiveSheet.Paste
                        
            Sheets("Speed Dating Data").Range("AT1").Copy
            Range("F1").Select
            ActiveSheet.Paste
            Sheets("Speed Dating Data").Range(ws.Cells(start, 46), ws.Cells((Row - 1), 46)).Copy
            Range("F2").Select
            ActiveSheet.Paste
            
            Application.CutCopyMode = False
            
            start = Row
            
                         
        End If
            
    Next Row
    
    Next ws
    
End Sub
    


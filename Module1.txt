

'--------------------------------------------
'1列目で相応する行を、2列目で取得する

Function getSettingValue(titleName As String) As String
    Dim ws As Worksheet
    
    Set ws = Worksheets("設定")
    
    Dim rowMax As Long
    
    rowMax = ws.Cells(65536, 2).End(xlUp).row
    
    Debug.Print rowMax
    
    Dim i As Long
    
    For i = 1 To rowMax
        
        'Debug.Print ws.Cells(i, 1).Value
        
        If titleName = ws.Cells(i, 1).Value Then
            
            getSettingValue = ws.Cells(i, 2).Value
            Exit Function
        End If
        
        getSettingValue = "none"
        
    Next

End Function
'--------------------------------------------


Function getDataLine(row As Long, col As Long, rowtitle As Long) As Variant
    Dim ws As Worksheet
    Set ws = Worksheets(1)
    Dim i As Long
    
    
    Dim dataArray(20) As Variant
    
    
    
    For i = 1 To 20

        dataArray(i) = ws.Cells(row, i).Value
        
    Next
    
    
    
    dataArray(0) = ws.Cells(row, col).Value 'Value
    dataArray(6) = ws.Cells(rowtitle, col).Value 'Date
    
    
    getDataLine = dataArray
    

End Function
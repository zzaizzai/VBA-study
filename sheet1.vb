Option Explicit

Sub sample()
    Dim ws As Worksheet
    
    Set ws = Worksheets(1)
    Debug.Print ws.Name
    
    Debug.Print ws.Cells(6, 8)
    
End Sub


Sub test()
    Dim varFile As Variant
    Dim aa As Variant
    Dim i As Long
    Dim textTemp As String
    Dim titles(20) As String
    
    titles(0) = "title"
    titles(1) = "2nd"
    titles(2) = "3nd"
    
    varFile = Application.GetSaveAsFilename(InitialFileName:="test.csv", _
                                        FileFilter:="CSVｫﾕｫ｡ｫ､ｫ・*.csv),*.csv", _
                                        FilterIndex:=1, _
                                        Title:="ﾜﾁｫﾕｫ｡ｫ､ｫ・ﾎ・")
    If varFile = False Then
        Exit Sub
    End If
    
    Open varFile For Output As #1
        Worksheets(1).Select
        
        'Titles
        textTemp = titles(0)
        For i = 1 To UBound(titles)
            textTemp = textTemp + "," + titles(i)
        Next
        'textTemp = textTemp + vbCrLf
        Print #1, textTemp
        
        Dim dataTemp() As Variant
        
        'Values
        dataTemp() = getDataLine(6, 6)
        
        
        
        Debug.Print "datatemp"
        textTemp = ""
        textTemp = Str(dataTemp(0))
        For i = 1 To UBound(dataTemp)
            Debug.Print dataTemp(i)
            If WorksheetFunction.IsText(dataTemp(1)) Then
                textTemp = textTemp + "," + dataTemp(i)
            ElseIf IsNumeric(dataTemp(i)) Then
                textTemp = textTemp + "," + Str(dataTemp(i))
            Else
                textTemp = textTemp + "," + "-1"
            End If
        Next
        
        
        Print #1, textTemp
        
        
        
        
    Close #1
    
    
    
    
    'For i = 0 To 20
     '   aa = getDataLine(6, 6, 5)
      '  Debug.Print i, aa(i)
    'Next
End Sub

Function getDataLine(row As Long, col As Long, Optional rowtitle As Long = 5) As Variant
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






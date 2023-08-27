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
    
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")
    ws.Activate
    Debug.Print ws.Name
    
    Dim rowMAx As Long
    
    rowMAx = ws.Cells(65536, 2).End(xlUp).row
    
    'TODO:
    'Check Error
    
    
    
    'TODO: fill col titles with hard coding
    titles(0) = "title"
    titles(1) = "2nd"
    titles(2) = "3nd"
    
    'save file
    varFile = Application.GetSaveAsFilename(InitialFileName:="test.csv", _
                                         FileFilter:="CSV Files (*.csv), *.csv", _
                                        FilterIndex:=1, _
                                        Title:="saving")
    'when you cancled quit sub
    If varFile = False Then
        Exit Sub
    End If
    
    'saving process
    Open varFile For Output As #1
        Worksheets(1).Select
        
        'ouput col titles
        textTemp = titles(0)
        For i = 1 To UBound(titles)
            textTemp = textTemp + "," + titles(i)
        Next
        'textTemp = textTemp + vbCrLf
        Print #1, textTemp
        
        
        'output values
        Dim dataTemp() As Variant ' temp data row
        Dim iCol As Long ' col for output
        Dim iRow As Long ' row for output
        
        For iRow = 6 To 10
            
            For iCol = 6 To 10
                
                dataTemp() = getDataLine(iRow, iCol)
                
                ' it there is no value in data tmep, go to continue
                If IsEmpty(dataTemp(0)) Then
                    GoTo continue
                End If
                
                Debug.Print "datatemp"
                textTemp = ""
                textTemp = dataTemp(0)
                
                For i = 1 To UBound(dataTemp)
                
                    Debug.Print dataTemp(i)
                    
                    If IsEmpty(dataTemp(i)) Then
                        textTemp = textTemp & "," & ""
                    Else
                        textTemp = textTemp & "," & dataTemp(i)
                    End If
                Next i
        
                Print #1, textTemp
                
continue:
            Next iCol
        Next iRow
        
        
        'Values

        
    Close #1
    
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









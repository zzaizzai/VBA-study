Option Explicit

Sub test()
    Dim aa As String
    aa = getSettingValue("ﾞﾂ・ｪ・ｻ")
    Debug.Print aa

End Sub


Function getSettingValue(titleName As String) As String
    Dim ws As Worksheet
    
    Set ws = Worksheets("珞・")
    
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

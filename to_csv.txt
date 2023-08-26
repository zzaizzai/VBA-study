

'--------------------------------------------------------------------------
'保存する部分
Dim varFile As Variant
varFile = Application.GetSaveAsFilename(InitialFileName:="test.csv", _
                                    FileFilter:="CSVファイル(*.csv),*.csv", _
                                    FilterIndex:=1, _
                                    Title:="保存ファイルの指定")

'保存をキャンセルしたら、終了
If varFile = False Then
    Exit Sub
End If

Open varFile For Output As #1
    Worksheets(1).Select
    
    textTemp = titles(0)
    For i = 1 To UBound(titles)
        textTemp = textTemp + "," + titles(i)
        
    
    Next
    Debug.Print "textTemp"
    Debug.Print textTemp
    Print #1, textTemp
    
    
    
Close #1

'--------------------------------------------------------------------------

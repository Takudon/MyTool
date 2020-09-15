Public Sub Sheet3()

Dim i As Integer 'Counter
Dim searchStartLine As Integer
Dim insertLine As Integer
Dim measureDateFrom As Variant
Dim insertionInterval As Integer

searchStartLine = 17
insertionInterval = 32

For i = 0 To 200
    measureDateFrom = Cells(i + 1, 4).Value 'D1からスタート
    
    insertLine = searchStartLine + i * insertionInterval
    
    Cells(insertLine, 1) = Str(measureDateFrom)
    Cells(insertLine, 1) = Replace(Cells(insertLine, 1), " ", "") '半角スペースが混入するため、削除処理を追加
    
Next

End Sub

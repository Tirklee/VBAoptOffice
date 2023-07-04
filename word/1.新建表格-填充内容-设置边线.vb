Sub NewTable()
Dim docNew As Document
Dim tblNew As Table
Dim intX As Integer
Dim intY As Integer

Set docNew = Documents.Add
Set tblNew = docNew.Tables.Add(Selection.Range, 3, 5)
With tblNew
For intX = 1 To 3
For intY = 1 To 5
.Cell(intX, intY).Range.InsertAfter "Cell: R" & intX & ", C" & intY
Next intY
Next intX
.Columns.AutoFit
End With
With tblNew.Borders
    .InsideLineStyle = wdLineStyleSingle '设置内部线条样式
    .InsideLineWidth = wdLineWidth050pt '设置内部线条宽度
    .InsideColor = RGB(0, 0, 0) '设置内部线条颜色（此处为黑色）
    
    .OutsideLineStyle = wdLineStyleSingle '设置外部线条样式
    .OutsideLineWidth = wdLineWidth050pt '设置外部线条宽度
    .OutsideColor = RGB(0, 0, 0) '设置外部线条颜色（此处为黑色）
End With
End Sub

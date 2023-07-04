Sub insertTableCaption()
    ' 定义变量
    Dim doc As Document
    Set doc = ActiveDocument
    Dim rangeObj As Range
    Set rangeObj = Selection.Range
    ' 新建表格
    Dim tblNew As Table
    Set tblNew = doc.Tables.Add(Range:=rangeObj, NumRows:=7, NumColumns:=5, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
    With tblNew
        If .Style <> "网格型" Then
            .Style = "网格型"
        End If
        For intX = 1 To 7
            For intY = 1 To 5
            .Cell(intX, intY).Range.InsertAfter "Cell: R" & intX & ", C" & intY
            Next intY
        Next intX
        .columns.AutoFit
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    rangeObj.InsertCaption Label:="表", TitleAutoText:="阿松大123", Title _
        :="阿松大123", Position:=wdCaptionPositionAbove, ExcludeLabel:=0
End Sub


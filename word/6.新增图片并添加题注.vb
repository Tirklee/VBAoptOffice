Sub insertImageAndCaption()
    Selection.InsertBreak Type:=wdLineBreak
    Selection.InlineShapes.AddPicture FileName:= _
        "D:\360MoveData\Users\Administrator\Desktop\asd.png", LinkToFile:=False, _
        SaveWithDocument:=True
    Selection.Range.InsertCaption Label:="图", TitleAutoText:="阿松大", Title:="阿松大", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
End Sub
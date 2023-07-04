Sub styleSetApply()
    '创建自定义标题
    createStyle "C标题", wdOutlineLevelBodyText
    '创建自定义正文
    createStyle "C正文", wdOutlineLevelBodyText
End Sub

Function createStyle(styleName As String, outlineLevel As Integer)
    '检查样式是否已存在
    If StyleExists(styleName) Then
        '如果样式已存在，则删除该样式
        DeleteStyle styleName
    End If
    
    Dim objStyle As style
    Set objStyle = ActiveDocument.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
    '创建样式
    With objStyle
        '设置样式属性
        .BaseStyle = ActiveDocument.Styles("正文")
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = wdColorBlue
        .ParagraphFormat.SpaceAfter = 12
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .ParagraphFormat.outlineLevel = outlineLevel
        .QuickStyle = True
    End With
End Function

Function StyleExists(styleName As String) As Boolean
    Dim style As style
    On Error Resume Next
    Set style = ActiveDocument.Styles(styleName)
    On Error GoTo 0
    StyleExists = Not style Is Nothing
End Function

Sub DeleteStyle(styleName As String)
    On Error Resume Next
    ActiveDocument.Styles(styleName).Delete
    On Error GoTo 0
End Sub
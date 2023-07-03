Sub CustomizeAndApplyHeadingStyle()
    ' 定义一级标题样式的名称和样式类型
    Dim styleName As String
    styleName = "Heading 1"
    
    ' 定义应用样式的文本范围
    Dim selectedRange As Range
    Set selectedRange = Selection.Range
    
    ' 检查样式是否已存在
    If StyleExists(styleName) Then
        ' 如果样式已存在，则删除该样式
        DeleteStyle styleName
    End If
    
    Dim objStyle As style
    Set objStyle = ActiveDocument.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
    
    ' 创建样式
    With objStyle
        ' 设置样式属性
        .BaseStyle = ActiveDocument.Styles("正文")
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = wdColorBlue
        .ParagraphFormat.SpaceAfter = 12
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .ParagraphFormat.OutlineLevel = wdOutlineLevel1
    End With
    
    Set LT = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)
     With LT.ListLevels(1)
     .NumberStyle = wdListNumberStyleArabic
     .NumberPosition = InchesToPoints(0.25 * (1 - 1))
     .TextPosition = InchesToPoints(0.25 * 1)
     .NumberFormat = "%1"
     .LinkedStyle = styleName
     End With
    ' 将样式应用于所选文本范围
    selectedRange.style = ActiveDocument.Styles(styleName)
End Sub

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
Sub applyStyle()
   Selection.Range.style = ActiveDocument.Styles("Heading 1")
End Sub
 
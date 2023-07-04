Sub CustomizeAndApplyStyle()
   ' 创建列表样式
   createStyle "列表样式", wdOutlineLevelBodyText
   Dim listTemplateObj As listTemplate
   Set listTemplateObj = ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
   ' 编号一级标题
   LinkedMutilListStyleNumber "列表样式", 1, "%1", listTemplateObj
End Sub


Function createStyle(styleName As String, outlineLevel As Integer)
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
        .ParagraphFormat.outlineLevel = outlineLevel
    End With
End Function


Sub LinkedMutilListStyleNumber(styleName As String, level As Integer, styleFormat As String, listTemplateObj As listTemplate)
    
     With listTemplateObj.ListLevels(level)
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0.25 * 0)
        .TextPosition = InchesToPoints(0.25 * 0)
        .NumberFormat = styleFormat
        .ResetOnHigher = level - 1
        .StartAt = 1
        .LinkedStyle = styleName
     End With
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


Sub applyStyle1()
   Selection.Range.style = ActiveDocument.Styles("列表样式")
End Sub
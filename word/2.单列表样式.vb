
Sub CustomizeAndApplyHeadingStyle()
    ' 创建一到九级标题
    ' 创建一级标题
     createStyle "一级标题", wdOutlineLevel1
    ' 创建二级标题
     createStyle "二级标题", wdOutlineLevel2
    ' 创建三级标题
     createStyle "三级标题", wdOutlineLevel3
    ' 创建四级标题
     createStyle "四级标题", wdOutlineLevel4
    ' 创建五级标题
     createStyle "五级标题", wdOutlineLevel5
    ' 创建六级标题
     createStyle "六级标题", wdOutlineLevel6
    ' 创建七级标题
     createStyle "七级标题", wdOutlineLevel7
    ' 创建八级标题
     createStyle "八级标题", wdOutlineLevel8
    ' 创建九级标题
     createStyle "九级标题", wdOutlineLevel9

     ' 编号一级标题
     LinkedMutilListStyleNumber "一级标题", 1, "%1"
    ' 编号二级标题
     LinkedMutilListStyleNumber "二级标题", 2, "%1.%2"
    ' 编号三级标题
     LinkedMutilListStyleNumber "三级标题", 3, "%1.%2.%3"
    ' 编号四级标题
     LinkedMutilListStyleNumber "四级标题", 4, "%1.%2.%3.%4"
    ' 编号五级标题
     LinkedMutilListStyleNumber "五级标题", 5, "%1.%2.%3.%4.%5"
    ' 编号六级标题
     LinkedMutilListStyleNumber "六级标题", 6, "%1.%2.%3.%4.%5.%6"
    ' 编号七级标题
     LinkedMutilListStyleNumber "七级标题", 7, "%1.%2.%3.%4.%5.%6.%7"
    ' 编号八级标题
     LinkedMutilListStyleNumber "八级标题", 8, "%1.%2.%3.%4.%5.%6.%7.%8"
    ' 编号九级标题
     LinkedMutilListStyleNumber "九级标题", 9, "%1.%2.%3.%4.%5.%6.%7.%8.%9"
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


Sub LinkedMutilListStyleNumber(styleName As String, level As Integer, styleFormat As String)
    Set LT = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)
     With LT.ListLevels(level)
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
   Selection.Range.style = ActiveDocument.Styles("一级标题")
End Sub
Sub applyStyle2()
   Selection.Range.style = ActiveDocument.Styles("二级标题")
End Sub

Sub applyStyle3()
   Selection.Range.style = ActiveDocument.Styles("三级标题")
End Sub

Sub applyStyle4()
   Selection.Range.style = ActiveDocument.Styles("四级标题")
End Sub

Sub applyStyle5()
   Selection.Range.style = ActiveDocument.Styles("五级标题")
End Sub

Sub applyStyle6()
   Selection.Range.style = ActiveDocument.Styles("六级标题")
End Sub
Sub applyStyle7()
   Selection.Range.style = ActiveDocument.Styles("七级标题")
End Sub
Sub applyStyle8()
   Selection.Range.style = ActiveDocument.Styles("八级标题")
End Sub
Sub applyStyle9()
   Selection.Range.style = ActiveDocument.Styles("九级标题")
End Sub



Sub getDocOutLineTitle()
    For Each Paragraph In ActiveDocument.ListParagraphs
        If Paragraph.outlineLevel <> wdOutlineLevelBodyText Then
            Debug.Print Paragraph.Range.listFormat.ListString & Paragraph.Range.Text
        End If
    Next Paragraph
End Sub


'  to distinguish unnumbered paragraphs from paragraphs which genuinely have a List Level
If Selection.Paragraphs(1).Range.ListParagraphs.Count = 1 Then
    MsgBox Selection.Paragraphs(1).Range.ListFormat.ListLevelNumber
Else
   MsgBox "Not a numbered paragraph"
End If

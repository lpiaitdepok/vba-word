'  to distinguish unnumbered paragraphs from paragraphs which genuinely have a List Level
If Selection.Paragraphs(1).Range.ListParagraphs.Count = 1 Then
    MsgBox Selection.Paragraphs(1).Range.ListFormat.ListLevelNumber
Else
   MsgBox "Not a numbered paragraph"
End If

'return All built-in heading styles have an outline level, (Level 1 for a main heading, Level 2 for a subheading, and so on)
'If the OutlineLevel is “Body Text”, this code returns 10.
Selection.Paragraphs(1).OutlineLevel

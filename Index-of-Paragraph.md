' operate on the currently selected paragraph
With
Selection.Paragraphs(1).Range
  .Font.Bold = True 
End With

' get the index number, from first selected paragraph with the Count property
ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count 

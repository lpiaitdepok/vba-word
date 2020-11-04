Dim intCount As Integer
'for inline shapes or images
For intCount = 1 To ActiveDocument.InlineShapes.Count
ActiveDocument.InlineShapes(intCount).Select
Next intCount

'for select all shapes or images
ActiveDocument.Shapes.SelectAll

'for floating shapes or images
For intCount = 1 To ActiveDocument.Shapes.Count
ActiveDocument.Shapes(intCount).Select
Next intCount

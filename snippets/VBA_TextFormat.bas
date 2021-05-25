Sub VBA_TextFormat()
'PURPOSE: Format selected text similar to VBE appearance
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim MySelection As Range
Dim MySelection_Limit As Range
Dim BlueArray As Variant

BlueArray = Array("End Sub", "End If", "For ", "In ", "Do While", "Sub ", "Set ", "As ", _
  "As Single", "Dim ", "With ", "End With", "If ", "Else ", "ElseIf ", "On Error GoTo ", _
  "On Error GoTo 0", "Sub ", "Exit Sub", "For Each ", "Next ", "Private Sub ", " True", _
  " False", " To ", " LBound", " UBound", "Wend", "While ", " Then", "Preserve ")

Set MySelection = Selection.Range
Set MySelection_Limit = MySelection.Duplicate

'Adjust font type and font size
  MySelection.Font.Name = "Courier New"
  MySelection.Font.Size = 10

'Loop through words that should be blue
  For x = LBound(BlueArray) To UBound(BlueArray)

  Set MySelection = Selection.Range
  Set MySelection_Limit = MySelection.Duplicate

  With MySelection.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = BlueArray(x)
    .Replacement.Text = BlueArray(x)
    .Forward = True
    .Wrap = wdFindStop
    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Replacement.Font.Color = wdColorDarkBlue

    While .Execute
      If MySelection.InRange(MySelection_Limit) Then
        MySelection.Font.Color = wdColorDarkBlue
        MySelection.Collapse wdCollapseEnd
      End If
    Wend
  End With
    
Next x

'Look for a Quote and changes the rest of that line's font color to Green
  With Selection.Find
    .ClearFormatting
    .Text = "'[!^13""""]@[^13]"
    .Replacement.ClearFormatting
    .Replacement.Text = ""
    .Replacement.Font.Color = wdColorGreen
    .Format = True
    .Forward = True
    .Wrap = wdFindStop
    .MatchCase = False
    .MatchWholeWord = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  End With

End Sub

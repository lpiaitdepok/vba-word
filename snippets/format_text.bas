' to change the font properties of selected text...
    With Selection.Font
' the Latin text font name and size...    
        .Name = "Verdana"
        .Size = 12
' Bold and Italic are the Font Style...         
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
' The following are the font effects...        
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .SmallCaps = True
        .AllCaps = False
        .Color = wdColorAutomatic
        .Superscript = False
        .Subscript = False
    End With

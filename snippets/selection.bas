'https://www.automateexcel.com/
'Allen-Wyatt.Microsoft-Word-Guidebook.Sharon-Parq-Associates-Inc

'Backspace
Selection.TypeBackspace

'Select entire document
Selection.HomeKey Unit:=wdStory
Selection.Extend

'Copy
Selection.Copy

'Delete
Selection.Delete Unit:=wdCharacter, Count:=1

'Insert After
Selection.InsertAfter "text"

'Beginning of Line
Selection.HomeKey Unit:=wdLine

'End of Line
Selection.EndKey Unit:=wdLine

'Paste
Selection.Paste

'Select All
Selection.WholeStory

'Select Entire Line
Selection.EndKey Unit:=wdLine, Extend:=wdExtend

'Move Up Paragraph
Selection.MoveUp Unit:wdParagraph, Count:=1

'Move Right One Character
Selection.MoveRight Unit:=wdCharacter, Count:=1

'Move Right One Cell In Table
Selection.MoveRight Unit:=wdCell

'Go To Start of Doc
Selection.HomeKey Unit:=wdStory

'Go To end of Doc
Selection.EndKey Unit:=wdStory

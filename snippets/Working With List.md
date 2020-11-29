````VB
ActiveDocument.Lists 'returns the Lists collection

For Each li In ActiveDocument.Lists 'returns an individual list; lists are in reverse order, from the end of the document forward 

List.ListParagraph 'returns all the paragraphs in the list

List.SingleListTemplate 'returns True/False if the entire list uses the same list template

List.StyleName 'returns the name of the style
List.CanContinuePreviousList 'returns a WdContinue constant that indicates whether formatting from previous list can be continued:
'WdContinue constant :
'wdContinueDisabled
'wdContinueList
'wdResetList 

Selection.Paragraphs(1).Range.ListFormat.ListString 'get the number that has been applied to any list-numbered paragraph

````

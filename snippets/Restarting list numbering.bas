 https://wordmvp.com/
 
 'Restarting list numbering
 
 Public Sub MarkRestarts()

' Sample code to mark restarted List Number paragraphs

Dim ListItem As Paragraph

Dim n As Integer ' bookmark id

n = 1

For Each ListItem In ActiveDocument.ListParagraphs

If ListItem.Style = _

ActiveDocument.Styles(wdStyleListNumber).NameLocal Then

If ListItem.Range.ListFormat.ListValue = 1 Then

ActiveDocument.Bookmarks.Add "restart" & n, _ ListItem.Range

n = n + 1

End If

End If

Next ListItem

End Sub


Public Sub ReapplyLists()

' Sample macro to correct errors after paste

' to another document and style changes

 

Dim aPara As Paragraph

Dim aBookmark As Bookmark

Dim StyleName As String

Dim RestartPoint As Range

 

' reset the paragraph formatting of the style if the paragraph

' should be numbered in this document

For Each aPara In ActiveDocument.Paragraphs

StyleName = aPara.Style

If ActiveDocument.Styles(StyleName).ListLevelNumber > 0 Then

aPara.Reset

End If

Next aPara

 

' reapply restarts to bookmarked paragraphs and tidy up

For Each aBookmark In ActiveDocument.Bookmarks

If aBookmark.Name Like "restart*" Then ' it's a restart

Set RestartPoint = aBookmark.Range

RestartPoint.Collapse wdCollapseEnd

RestartPoint.MoveEnd unit:=wdCharacter, Count:=-1

With RestartPoint.ListFormat

.ApplyListTemplate .ListTemplate, _

 ContinuePreviousList:=False

End With

aBookmark.Delete

End If

Next aBookmark

End Sub

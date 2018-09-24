'select some text and then want quote marks added around the selected text. You can accomplish this task with the following macro:

Sub AddQuotes()
    Dim sBegQ As String
    Dim sEndQ As String

    If Options.AutoFormatAsYouTypeReplaceQuotes Then
        sBegQ = Chr(147)
        sEndQ = Chr(148)
    Else
        sBegQ = Chr(34)
        sEndQ = Chr(34)
    End If

    Selection.InsertBefore sBegQ
    Selection.InsertAfter sEndQ
End Sub

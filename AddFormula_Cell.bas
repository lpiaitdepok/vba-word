Sub AddFormula_Cell()
' This example inserts a formula at the insertion point that determines the largest number in the cells above the selected cell.
Selection.Collapse Direction:=wdCollapseStart
If Selection.Information(wdWithInTable) = True Then

Selection.Cells(1).Formula Formula:="=SUM(" & toAlphabet(Selection.Cells(1).ColumnIndex - 2) & Selection.Cells(1).RowIndex & ":" & toAlphabet(Selection.Cells(1).ColumnIndex - 1) & Selection.Cells(1).RowIndex & ") \#  ""# minute"""
Selection.MoveDown Unit:=wdLine, Count:=1
 Else
 MsgBox "The insertion point is not in a table."
End If

End Sub

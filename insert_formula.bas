' This example inserts a formula at the insertion point that determines the largest number in the cells above the selected cell.
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Formula Formula:="=Max(Above)" 
Else 
 MsgBox "The insertion point is not in a table." 
End If

Sub RebuildFormulas()
    On Error Resume Next
    Dim eq As Object
    Dim count As Long
    count = 0
    For Each eq In ActiveDocument.OMaths
        eq.Linearize
        eq.BuildUp
        count = count + 1
    Next
    MsgBox "Rebuilt " & count & " / " & ActiveDocument.OMaths.Count & " formulas"
End Sub

Sub ShowFormulaCount()
    On Error Resume Next
    MsgBox "Total formulas: " & ActiveDocument.OMaths.Count
End Sub

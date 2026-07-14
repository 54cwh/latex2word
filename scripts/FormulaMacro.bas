Sub RebuildAll()
    On Error Resume Next
    Dim eq As Object
    Dim i As Long
    For i = 1 To ActiveDocument.OMaths.Count
        Set eq = ActiveDocument.OMaths(i)
        eq.Linearize
        eq.BuildUp
    Next
    MsgBox "OK"
End Sub

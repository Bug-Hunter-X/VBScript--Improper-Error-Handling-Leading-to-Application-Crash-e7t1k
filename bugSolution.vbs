Function MyFunction(param1, param2)
  On Error Resume Next
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 1001, , "Parameters cannot be empty"
  End If
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description, vbCritical
    Err.Clear
    Exit Function ' or handle the error appropriately
  End If
  ' ... rest of the function ...
End Function
Function MyFunc(param1)
  If param1 = "" Then
    ' Explicitly check for empty string
    MyFunc = ""
    Exit Function
  ElseIf IsEmpty(param1) Then
    ' Handle truly empty (uninitialized) parameter
    MyFunc = "Parameter is empty"
    Exit Function
  End If
  ' ... rest of the function
End Function
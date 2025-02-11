Function MyFunc(param1)
  If IsEmpty(param1) Then
    ' Handle empty parameter
    MyFunc = ""
    Exit Function
  End If
  ' ... rest of the function
End Function
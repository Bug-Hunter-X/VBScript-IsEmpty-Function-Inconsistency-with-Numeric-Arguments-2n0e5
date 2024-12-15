Function f(a, b)
  If IsEmpty(a) Then
    WScript.Echo "a is empty"
    Exit Function
  End If
  If IsEmpty(b) Then
    WScript.Echo "b is empty"
    Exit Function
  End If
  If Not IsNumeric(a) Or Not IsNumeric(b) Then
    WScript.Echo "a or b is not numeric"
    Exit Function
  End If
  c = a + b
  WScript.Echo c
End Function

f(1, 2)
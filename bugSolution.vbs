Function f(a, b)
  If a = "" Then
    WScript.Echo "a is empty"
    Exit Function
  End If
  If b = "" Then
    WScript.Echo "b is empty"
    Exit Function
  End If
  If Not IsNumeric(a) Or Not IsNumeric(b) Then
    WScript.Echo "a or b is not numeric"
    Exit Function
  End If
  c = CDbl(a) + CDbl(b)  'Explicit type conversion
  WScript.Echo c
End Function

f(1, 2)
f("",2)
f(1,"")
f("","")
f("abc",2) 
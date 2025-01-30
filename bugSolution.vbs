Function MyFunction(param1)
  'Explicitly check for empty string using the Len function
  If Len(Trim(param1)) = 0 Then
    ' Handle empty string
  Else
    ' Code to execute if param1 is not empty
  End If
End Function
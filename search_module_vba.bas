Attribute VB_Name = "Module1"

Sub calculate_weight()
'For three'
For Each cell In Selection
    cad = cell.Value
    If Len(cad) > 1 Then
        cad = Split(cad, ",")
        result = 1 * cad(0) + 0.8 * cad(1) + 0.6 * cad(2)
        Debug.Print result
        cell.Offset(0, 1).Value = result
    Else
        cell.Offset(0, 1).Value = 0
    End If
    
    
Next cell

End Sub


Sub calculate_weight1()
'For Five'
For Each cell In Selection
    cad = cell.Value
    If Len(cad) > 1 Then
        cad = Split(cad, ",")
        result = 1 * cad(0) + 0.9 * cad(1) + 0.8 * cad(2) + 0.7 * cad(3) + 0.6 * cad(4)
        Debug.Print result
        cell.Offset(0, 1).Value = result
     Else
        cell.Offset(0, 1).Value = 0
    End If
Next cell

End Sub


Sub calculate_weight2()
'For Side Rail'
Dim Arr(1 To 9) As Integer

For Each cell In Selection
    cad1 = cell.Value
    If Len(cad1) > 1 Then
        cad = Split(cad1, ",")
        
        ArrayLen = UBound(cad) - LBound(cad) + 1
        For i = 1 To 9
          If i <= ArrayLen Then
           Arr(i) = cad(i - 1)
          Else
           Arr(i) = 0
          End If
        Next i
        
        result = 1 * Arr(1) + 0.95 * Arr(2) + 0.9 * Arr(3) + 0.8 * Arr(4) + 0.75 * Arr(5) + 0.7 * Arr(6) + 0.6 * Arr(7) + 0.55 * Arr(8) + 0.5 * Arr(9)
        Debug.Print result
        cell.Offset(0, 1).Value = result
     Else
        cell.Offset(0, 1).Value = 0
    End If
Next cell

End Sub

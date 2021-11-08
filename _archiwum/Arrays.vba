Function func(ByVal input)
'
' description.
'
' @since 1.0.0
' @param {type} [name] description.
' @return {type} [name] description.
' @see dependencies
'

    

End Function
'———————————————————————————————————————————————————————————————————————————————————————————
Function ArrayRemoveDups(MyArray As Variant) As Variant
    
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String
    
    Dim arrTemp() As String
    Dim Coll As New Collection
 
    'Get First and Last Array Positions
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)
 
    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = CStr(MyArray(i))
    Next i
    
    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0
 
    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i
    
    'Output Array
    ArrayRemoveDups = arrTemp
 
End Function
'———————————————————————————————————————————————————————————————————————————————————————————

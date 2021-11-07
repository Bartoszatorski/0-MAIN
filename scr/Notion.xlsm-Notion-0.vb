Option Explicit
    'tworzenie pustego pola do przechowywania popzedniej wartosci
    Private PoprzedniaWartosc
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    PoprzedniaWartosc = ActiveCell.value
    Debug.Print "zmiana pola"

    
End Sub
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Debug.Print Now, Sh.Name, Target.value
    
    Dim value As String: value = ""
    Dim col As Long: col = 1
    Dim i As Long: i = 0 ' Skoroszyt [sheet_id]
    
    value = Target.value
    i = 1
    col = 1
    
    Call mobileIn(value, col, i)
    
End Sub
Sub mobileOut()
    On Error Resume Next
    Dim Filename As String:     Filename = "C:\Users\smoka\OneDrive\0-MAIN\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(Filename:=Filename, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(i)
    Dim rowNext As Long:        rowNext = 21 'ws.Cells(1, col).CurrentRegion.Rows.Count + 1
    wb.Close 1
End Sub
Sub mobileNext(i As Long)
    On Error Resume Next
    Dim Filename As String:     Filename = "C:\Users\smoka\OneDrive\0-MAIN\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(Filename:=Filename, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(i)
    Dim rowNext As Long:        rowNext = 21 'ws.Cells(1, col).CurrentRegion.Rows.Count + 1
    
    

    

    ws.Range("A" & rowNext).EntireRow.Insert

End Sub

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
Sub pobierz2()
    Dim row As Long
    Dim col As Long
    Dim x
    
    Dim Sh As Worksheet
    Dim arr(3 To 890) As Variant
    
    Dim outputArray() As String
    
    Set Sh = ThisWorkbook.Worksheets("Arkusz1")
    
    For row = 3 To 890
        For col = 4 To 4
            If Len(Sh.Cells(row, col)) > 0 Then
            arr(row) = Sh.Cells(row, col)
            Debug.Print row, Now, arr(row)
            End If
        Next col
    Next row
    
    outputArray = ArrayRemoveDups(arr)
    For Each x In outputArray
    Debug.Print x
    Next x
    
End Sub
Public Sub pobierz()
Dim i As Long
Dim ii As String


For i = 1 To 100000 Step 3
    ii = CStr(i)
    'Call MobileIn("value", col, i)
    Call mobileIn(Now, 1, 3)
        Call mobileIn(ii, 2, 3)

    Call mobileOut
Next i
End Sub
Public Sub notionMobile()
    MsgBox ("działa")
    Call pobierz
    
End Sub


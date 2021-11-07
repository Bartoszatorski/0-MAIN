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
    Dim I As Long: I = 0 ' Skoroszyt [sheet_id]
    
    value = Target.value
    I = 1
    col = 1
    
    Call mobileIn(value, col, I)
    
End Sub
Sub mobileOut()
    On Error Resume Next
    Dim I As Long
    Dim fileName As String:     fileName = "C:\Users\smoka\OneDrive\0-MAIN\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(fileName:=fileName, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(I)
    Dim rowNext As Long:        rowNext = 21 'ws.Cells(1, col).CurrentRegion.Rows.Count + 1
    wb.Close 1
End Sub
Sub mobileNext(I As Long)
    On Error Resume Next
    Dim fileName As String:     fileName = "C:\Users\smoka\OneDrive\0-MAIN\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(fileName:=fileName, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(I)
    Dim rowNext As Long:        rowNext = 21 'ws.Cells(1, col).CurrentRegion.Rows.Count + 1
    
    

    

    ws.Range("A" & rowNext).EntireRow.Insert

End Sub

'———————————————————————————————————————————————————————————————————————————————————————————
Function ArrayRemoveDups(MyArray As Variant) As Variant
    
    Dim nFirst As Long, nLast As Long, I As Long
    Dim item As String
    
    Dim arrTemp() As String
    Dim Coll As New Collection
 
    'Get First and Last Array Positions
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)
 
    'Convert Array to String
    For I = nFirst To nLast
        arrTemp(I) = CStr(MyArray(I))
    Next I
    
    'Populate Temporary Collection
    On Error Resume Next
    For I = nFirst To nLast
        Coll.Add arrTemp(I), arrTemp(I)
    Next I
    Err.Clear
    On Error GoTo 0
 
    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For I = nFirst To nLast
        arrTemp(I) = Coll(I - nFirst + 1)
    Next I
    
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
   
    Dim I As Long
    Dim value As String
    Dim col As Long: col = 1

    Dim iMax As Long: iMax = 100
    Dim arr(999999) As String
    
    Dim outputArray() As String
    Dim x
    Dim colLast As Long: colLast = 0
    
    For I = iMax To 1 Step -1
        'Call MobileIn("value", col, i)
        value = I & "_v0"
        
        colLast = 1
        If I = iMax Then colLast = 0
        If I = 1 Then colLast = 2

        arr(I) = mobileIn(value, col, 1, colLast)
    Next I
    
    colLast = 0
    
    outputArray = ArrayRemoveDups(arr)
    For Each x In outputArray
    Debug.Print x
    Next x
    
End Sub
Function mobileIn(value As String, col As Long, I As Long, colLast As Long)
    On Error Resume Next
    Dim fileName As String:     fileName = "C:\Users\smoka\OneDrive\Documents\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(fileName:=fileName, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(I)
    Dim rowNext As Long:        rowNext = wb.Sheets(2).Cells(1, 1)
    
    
    ws.Cells(rowNext, col) = value
    
    rowNext = rowNext + 1
    With wb.Sheets(2)
        .Cells(1, 1) = rowNext
        .Cells(2, 1) = Now

        If colLast = 0 Then
            .Cells(4, 1) = Now
            .Cells(7, 1) = 0
        End If
        If colLast = 1 Then .Cells(5, 1) = Now
        If colLast = 2 Then .Cells(6, 1) = Now
        .Cells(7, 1) = .Cells(7, 1) + 1
    End With
    
    mobileIn = value
    Debug.Print Now, "<mobile_in>", I, col, colLast, rowNext, value

    'wb.Save
End Function
Function uDane(t)
End Function



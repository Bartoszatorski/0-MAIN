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
    Dim i As Long
    Dim fileName As String:     fileName = "C:\Users\smoka\OneDrive\0-MAIN\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(fileName:=fileName, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(i)
    Dim rowNext As Long:        rowNext = 21 'ws.Cells(1, col).CurrentRegion.Rows.Count + 1
    wb.Close 1
End Sub
Sub mobileNext(i As Long)
    On Error Resume Next
    Dim fileName As String:     fileName = "C:\Users\smoka\OneDrive\0-MAIN\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(fileName:=fileName, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(i)
    Dim rowNext As Long:        rowNext = 21 'ws.Cells(1, col).CurrentRegion.Rows.Count + 1
    
    

    

    ws.Range("A" & rowNext).EntireRow.Insert

End Sub

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
    Dim value As String
    Dim col As Long: col = 1

    Dim iMax As Long: iMax = 100
    Dim arr(999999) As String
    
    Dim outputArray() As String
    Dim x
    Dim colLast As Long: colLast = 0
    
    For i = iMax To 1 Step -1
        'Call MobileIn("value", col, i)
        value = i & "_v0"
        
        colLast = 1
        If i = iMax Then colLast = 0
        If i = 1 Then colLast = 2

        arr(i) = mobileIn(value, col, 1, colLast)
    Next i
    
    colLast = 0
    
    outputArray = ArrayRemoveDups(arr)
    For Each x In outputArray
    Debug.Print x
    Next x
    
End Sub
Function mobileIn(value As String, col As Long, i As Long, colLast As Long)
    On Error Resume Next
    Dim fileName As String:     fileName = "C:\Users\smoka\OneDrive\Documents\Notion_mobile.xlsx"
    Dim wb As Workbook:         Set wb = Workbooks.Open(fileName:=fileName, UpdateLinks:=0)
    Dim ws As Worksheet:        Set ws = wb.Sheets(i)
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
    Debug.Print Now, "<mobile_in>", i, col, colLast, rowNext, value

    'wb.Save
End Function
Sub txtfil()
    Dim fileName As String
    Dim i
    fileName = "C:\Users\smoka\OneDrive\0-MAIN\txt\d1.txt"
    
    Dim fso As Object, ts As Object
    'Need to define constants manually
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    'Need to define constants manually
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'The below will not Hello.txt if it does not exist and will open file for Unicode appending
    Set ts = fso.OpenTextFile(fileName, ForAppending, True, TristateFalse)
    
    For i = 1 To 100
    ts.Writeline "Hello"
    Next i
    ts.Close
     
    'Open same file for reading
    Set ts = fso.OpenTextFile(fileName, ForReading, True, TristateFalse)
     
    'Read till the end
    Do Until ts.AtEndOfStream
         Debug.Print "Printing line " & ts.Line; "  -  ";
         Debug.Print ts.ReadLine 'Print a line from the file
    Loop
    ts.Close
    
    End Sub
    


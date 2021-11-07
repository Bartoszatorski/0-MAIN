Sub LoopThroughFiles()
 
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    
    Dim oFolderName
    oFolderName = Cells(2, 1)
    Debug.Print oFolderName
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
     
    Set oFolder = oFSO.GetFolder(oFolderName & "/")
     
    For Each oFile In oFolder.Folders
        Cells(i + 2, 2) = oFile.Name
        i = i + 1
    Next oFile
     
    End Sub
    
Sub rename()
        Name "C:\Users\bzatorski\Desktop\Git\Programming-VBA\plik-zmiana_nazwy/nazwa1.xlsx" As "C:\Users\bzatorski\Desktop\Git\Programming-VBA\plik-zmiana_nazwy/nazwa2.xlsx"
    
    End Sub
    
Sub txtfil()
    Dim fileName As String
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
    
Sub GetFileNames()
    Dim MyFSO As FileSystemObject
    Dim MyFile As File
    Dim MyFolder As Folder
    
    Set MyFSO = New Scripting.FileSystemObject
    Set MyFolder = MyFSO.GetFolder("C:\Users\sumit\Desktop\Test")
    
    For Each MyFile In MyFolder.Files
        Debug.Print MyFile.Name
    Next MyFile

End Sub
Sub GetSubFolderNames()

    Dim MyFSO As FileSystemObject
    Dim MyFile As File
    Dim MyFolder As Folder
    Dim MySubFolder As Folder
    
    Set MyFSO = New Scripting.FileSystemObject
    Set MyFolder = MyFSO.GetFolder("C:\Users\sumit\Desktop\Test")
    
    For Each MySubFolder In MyFolder.SubFolders
        Debug.Print MySubFolder.Name
    Next MySubFolder 

End Sub
Sub CheckFolderExist()

    Dim MyFSO As FileSystemObject
    Set MyFSO = New FileSystemObject

    If MyFSO.FolderExists("C:\Users\sumit\Desktop\Test") Then
        MsgBox "The Folder Exists"
    Else
        MsgBox "The Folder Does Not Exist"
    End If

End Sub
Sub CheckFileExist()

    Dim MyFSO As FileSystemObject
    Set MyFSO = New FileSystemObject
    If MyFSO.FileExists("C:\Users\sumit\Desktop\Test\Test.xlsx") Then
        MsgBox "The File Exists"
    Else
        MsgBox "The File Does Not Exist"
    End If

End Sub
Sub CreateFolder()

    Dim MyFSO As FileSystemObject

    Set MyFSO = New FileSystemObject
    
    If MyFSO.FolderExists("C:\Users\sumit\Desktop\Test") Then
        MsgBox "The Folder Already Exist"
    Else
        MyFSO.CreateFolder ("C:\Users\sumit\Desktop\Test")
    End If
    
End Sub
Sub FSOGetFolder()

    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fld = FSO.GetFolder("C:\Src\")
    
    Debug.Print fld.DateCreated
    Debug.Print fld.Drive
    Debug.Print fld.Name
    Debug.Print fld.ParentFolder
    Debug.Print fld.Path
    Debug.Print fld.ShortPath
    Debug.Print fld.Size
    Debug.Print fld.Files.Count
    Debug.Print fld.Type
    
    For Each fold In fld.SubFolders
        Debug.Print fold.Name
    Next fold
    
    For Each fil In fld.Files
        Debug.Print fil.Name
        Set fil = FSO.GetFile("C:\Src\Test.xlsx")
            '
            'you can copy it:
            'fil.Copy "C:\Dst\"
            '
            'move it:
            'fil.Move "C:\Dst\"
            '
            'delete it:
            'fil.Delete
            '
            'or open it as a TextStream object:
            'fil.OpenAsTextStream
            '
    Next fil

    ParentFold= FSO.GetParentFolderName("C:\ParentTest\Test\")
    Debug.Print FSO.GetSpecialFolder(0)
 
    
End Sub
Sub FSOGetFolder_v2()

    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fld = FSO.GetFolder("C:\Users\smoka\OneDrive\0-MAIN")
    
    Debug.Print
    Debug.Print "FSO info"
    Debug.Print
    
    Debug.Print fld.DateCreated
    Debug.Print fld.Drive
    Debug.Print fld.Name
    Debug.Print fld.ParentFolder
    Debug.Print fld.Path
    Debug.Print fld.ShortPath
    Debug.Print fld.Size
    Debug.Print fld.Files.Count
    Debug.Print fld.Type
    
    Debug.Print
    Debug.Print "Folders"
    Debug.Print
    
    For Each fold In fld.SubFolders
        Debug.Print fold.Name,
        
    Next fold
    
    Debug.Print
    Debug.Print "Files"
    Debug.Print

    For Each fil In fld.Files
        Debug.Print fil.Name,
        Debug.Print fil.Type,
        Debug.Print fil.Path
        'Set fil = FSO.GetFile("C:\Src\Test.xlsx")
            '
            'you can copy it:
            'fil.Copy "C:\Dst\"
            '
            'move it:
            'fil.Move "C:\Dst\"
            '
            'delete it:
            'fil.Delete
            '
            'or open it as a TextStream object:
            'fil.OpenAsTextStream
            '
    Next fil

    'ParentFold = FSO.GetParentFolderName("C:\ParentTest\Test\")
    'Debug.Print FSO.GetSpecialFolder(0)
    
    
End Sub
    
Attribute VB_Name = "modFileSearch"
''' modFileSearch
''' Searches for files inside a directory.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Go through a directory and fills a list view with its filtered contents.
Public Sub GetDirectoryContents(strSource As String)
    Dim fso As New FileSystemObject
    Dim folder As folder
    Dim file As file
    Dim strFilename As String
    
    ' Decide if we should go through it recursively or not.
    If frmMain.chkRecursive.Value Then
        ' Go through the source directory looking for folders.
        For Each folder In fso.GetFolder(strSource).SubFolders
            GetDirectoryContents folder.Path
        Next
    End If
    
    ' Go through the source directory looking for files.
    For Each file In fso.GetFolder(strSource).Files
        frmMain.lstFound.AddItem file.ShortPath
    Next
End Sub

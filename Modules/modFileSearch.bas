Attribute VB_Name = "modFileSearch"
''' modFileSearch
''' Searches for files inside a directory.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Checks if a file is of an specific extension.
Public Function FileMatchesExtension(file As file, strExtension As String) As Boolean
    Dim strFileExt As String
    Dim posExtStr As Integer
    
    ' Get the extension of the file.
    posExtStr = InStrRev(file.Name, ".")
    If posExtStr > 0 Then
        strFileExt = Mid(file.Name, posExtStr + 1)
    End If
    
    ' Should we do this search case insensitively.
    If frmMain.chkCaseSensitive.Value = 0 Then
        strFileExt = UCase(strFileExt)
        strExtension = UCase(strExtension)
    End If
    
    ' Check if the extensions match.
    If strFileExt = strExtension Then
        FileMatchesExtension = True
    Else
        FileMatchesExtension = False
    End If
End Function

' Go through a directory and fills a list view with its filtered contents.
Public Sub GetDirectoryContents(strSource As String, Optional strBase As String)
    Dim fso As New FileSystemObject
    Dim folder As folder
    Dim file As file
    Dim strFilename As String
    Dim strExtension As String
    Dim idxExtension As Integer
    
    ' Check if the user supplied a base directory.
    If IsMissing(strBase) Then
        strBase = strSource
    End If
    
    ' Decide if we should go through it recursively or not.
    If IsRecursive Then
        ' Go through the source directory looking for folders.
        For Each folder In fso.GetFolder(strSource).SubFolders
            GetDirectoryContents folder.Path, strBase
        Next
    End If
    
    ' Go through the source directory looking for files.
    For Each file In fso.GetFolder(strSource).Files
        ' Check if we have no extension filters set.
        If frmMain.lstExtensions.ListCount <= 0 Then
            frmMain.lstFound.AddItem Mid(file.Path, Len(strBase) + 1)
        End If
        
        ' Check if the file has an extension that is listed in the extensions list.
        For idxExtension = 0 To frmMain.lstExtensions.ListCount - 1
            strExtension = frmMain.lstExtensions.List(idxExtension)
                
            If FileMatchesExtension(file, strExtension) Then
                frmMain.lstFound.AddItem Mid(file.Path, Len(strBase) + 1)
            End If
        Next
    Next
End Sub

' Checks if we should search for files recursively.
Public Function IsRecursive() As Boolean
    IsRecursive = frmMain.chkRecursive.Value
End Function

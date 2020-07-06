Attribute VB_Name = "modFileSearch"
''' modFileSearch
''' Searches for files inside a directory.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Go through a directory and fills a list view with its filtered contents.
Public Sub GetDirectoryContents(strSource As String)
    Dim fso As New FileSystemObject
    Dim atrFile As VbFileAttribute
    Dim strFilename As String
    
    ' Decide if we should go through it recursively or not.
    If frmMain.chkRecursive.Value Then
        atrFile = vbDirectory
    Else
        atrFile = vbNormal
    End If
    
    ' Go through the source directory looking for its contents.
    strFilename = Dir(strSource & "\*.*", atrFile)
    While strFilename <> ""
        ' Ignore the current and parent directory.
        If (strFilename <> ".") And (strFilename <> "..") Then
            If (GetAttr(strSource & "\" & strFilename) And vbDirectory) Then
                GetDirectoryContents strFilename
            End If
            
            frmMain.lstFound.AddItem strFilename
        End If
        
        ' Get next item.
        strFilename = Dir()
    Wend
End Sub

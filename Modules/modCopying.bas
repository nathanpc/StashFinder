Attribute VB_Name = "modCopying"
''' modCopying
''' Helps out on all the copying we're doing.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Copies a file from one place to the other, and maybe preserve it's folder structure.
Public Sub CopyFile(strFilePath As String, strSourceDir As String, strDestinationDir As String, blnPreserveStructure As Boolean)
    Dim strDestination As String
    
    ' Discard the whole folder structure?
    If Not blnPreserveStructure Then
        strDestination = strDestinationDir & Mid(strFilePath, InStrRev(strFilePath, "\"))
    Else
        strDestination = strDestinationDir & strFilePath
        CreateFolderStructure strDestination
    End If
    
    ' Actually copy the file.
    FileCopy strSourceDir & strFilePath, strDestination
End Sub

' Creates a whole folder structure to a file.
Public Sub CreateFolderStructure(ByVal strPath As String)
    Dim fso As New FileSystemObject
    Dim strCurFolder As String
    
    ' Go through each new folder creating it in the process.
    strCurFolder = Left(strPath, InStr(1, strPath, "\"))
    While strCurFolder <> ""
        ' Check if the folder doesn't exist to create it.
        If Not fso.FolderExists(strCurFolder) Then
            fso.CreateFolder strCurFolder
        End If
        
        ' Go to the next folder.
        strCurFolder = Left(strPath, InStr(Len(strCurFolder) + 1, strPath, "\"))
    Wend
End Sub

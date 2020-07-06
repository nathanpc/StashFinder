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
    End If
    
    ' Actually copy the file.
    FileCopy strSourceDir & strFilePath, strDestination
End Sub

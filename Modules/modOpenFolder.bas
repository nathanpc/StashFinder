Attribute VB_Name = "modOpenFolder"
''' modOpenFolder
''' Wrapper for Win32 API functions to let us open a folder instead of a file.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Define constants.
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

' BrowseInfo structure.
Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

' Import Win32 API functions.
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long

' Show the user a dialog to select a directory.
Public Function OpenDirectoryDialog(hwndParent As Long, strTitle As String) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo

    ' Setup the BrowserInfo structure.
    tBrowseInfo.hwndOwner = hwndParent
    tBrowseInfo.lpszTitle = lstrcat(strTitle, "")
    tBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS

    ' Get the directory from the user.
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        ' Allocate space for the path and get it.
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        
        ' Trim the excess and return it.
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        OpenDirectoryDialog = sBuffer
    End If
End Function

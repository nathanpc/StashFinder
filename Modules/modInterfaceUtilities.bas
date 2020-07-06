Attribute VB_Name = "modInterfaceUtilities"
''' modInterfaceUtilities
''' A collection of user interface utilities.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Import some API functions.
Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

' Define some API constants.
Const LB_SETHORIZONTALEXTENT = &H194

' Add a horizontal scrollbar to a ListBox.
Public Sub AddHorizontalScrollToListBox(lst As ListBox)
    Dim i As Long
    Dim lngNewLength As Long
    Dim lngMaxLength As Long
    
    For i = 0 To lst.ListCount - 1
        lngNewLength = 10 + lst.Parent.ScaleX( _
            lst.Parent.TextWidth(lst.List(i)), _
            lst.Parent.ScaleMode, vbPixels)
        
        If lngMaxLength < lngNewLength Then
            lngMaxLength = lngNewLength
        End If
    Next i
    
    SendMessage lst.hwnd, LB_SETHORIZONTALEXTENT, lngMaxLength, 0
End Sub

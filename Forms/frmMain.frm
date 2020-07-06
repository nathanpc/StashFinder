VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoarder Stash Finder"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "<<"
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdAddAll 
      Caption         =   ">>"
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   2760
      Width           =   495
   End
   Begin VB.ListBox lstToCopy 
      Height          =   4350
      Left            =   4680
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   1680
      Width           =   3855
   End
   Begin VB.ListBox lstFound 
      Height          =   4350
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton cmdCopyFiles 
      Caption         =   "Copy Files"
      Height          =   1095
      Left            =   7320
      TabIndex        =   16
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   1095
      Left            =   7320
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraDestination 
      Caption         =   "Destination"
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   6975
      Begin VB.CheckBox chkPreserveStructure 
         Caption         =   "Preserve Directory Structure"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtDestination 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   6255
      End
      Begin VB.CommandButton cmdBrowseDestination 
         Caption         =   "..."
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   3360
      Width           =   495
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CheckBox chkCaseSensitive 
         Caption         =   "Case Sensitive"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.ListBox lstExtensions 
         Height          =   645
         Left            =   6000
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox chkRecursive 
         Caption         =   "Recursive Search"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdBrowseSource 
         Caption         =   "..."
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtSource 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Extensions:"
         Height          =   255
         Left            =   6000
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label3 
      Caption         =   "To Copy:"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Found:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenSource 
         Caption         =   "&Open Source Directory..."
      End
      Begin VB.Menu mnuFileOpenDestination 
         Caption         =   "Open &Destination Directory..."
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuExtensions 
      Caption         =   "&Extensions"
      Begin VB.Menu mnuExtensionsAdd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuExtensionsDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmMain
''' Main form of the application.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Adds selected items from the Found list to the To Copy list.
Private Sub cmdAdd_Click()
    Dim i As Long
    Dim idxLast As Long
    
    ' Check if we only have one item selected.
    If lstFound.SelCount = 1 Then
        lstToCopy.AddItem lstFound.Text
        lstFound.RemoveItem lstFound.ListIndex
        Exit Sub
    End If
    
    ' Go through the list looking for the selected items.
    i = 0
    idxLast = lstFound.ListCount
    While i < idxLast
        If lstFound.Selected(i) Then
            lstToCopy.AddItem lstFound.List(i)
            lstFound.RemoveItem i
            
            i = i - 1
            idxLast = idxLast - 1
        End If
        
        i = i + 1
    Wend
    
    ' Make sure the ListBox can display its new contents.
    AddHorizontalScrollToListBox lstFound
    AddHorizontalScrollToListBox lstToCopy
End Sub

' Add all the files from the Found list to the To Copy list.
Private Sub cmdAddAll_Click()
    Dim i As Integer
    
    ' Push all the Found items to the To Copy list.
    For i = 0 To lstFound.ListCount - 1
        lstToCopy.AddItem lstFound.List(i)
    Next
    
    ' Clear the Found list.
    lstFound.Clear
    
    ' Make sure the lists have a nice horizontal scrollbar if needed.
    AddHorizontalScrollToListBox lstFound
    AddHorizontalScrollToListBox lstToCopy
End Sub

' Copies over all of the files to the destination.
Private Sub cmdCopyFiles_Click()
    Dim idxFile As Integer
    Dim numCopied As Integer

    ' Check if the destination location is set.
    If txtDestination.Text = "" Then
        MsgBox "No destination location was set.", vbOKOnly + vbCritical, _
            "No Destination Set"
        Exit Sub
    End If
    
    ' Go through the To Copy list copying files.
    For idxFile = 0 To lstToCopy.ListCount - 1
        ' Copy the file.
        CopyFile lstToCopy.List(idxFile), txtSource.Text, txtDestination.Text, _
            chkPreserveStructure
        numCopied = numCopied + 1
    Next
    
    ' Finished message.
    MsgBox "Copied" & Str(numCopied) & " files to " & txtDestination.Text, _
        vbOKOnly + vbInformation, "Finished Copying Files"
End Sub

' Removes the selected items from the To Copy list.
Private Sub cmdRemove_Click()
    Dim i As Long
    Dim idxLast As Long
    
    ' Check if we only have one item selected.
    If lstToCopy.SelCount = 1 Then
        lstFound.AddItem lstToCopy.Text
        lstToCopy.RemoveItem lstToCopy.ListIndex
        Exit Sub
    End If
    
    ' Go through the list looking for the selected items.
    i = 0
    idxLast = lstToCopy.ListCount
    While i < idxLast
        If lstToCopy.Selected(i) Then
            lstFound.AddItem lstToCopy.List(i)
            lstToCopy.RemoveItem i
            
            i = i - 1
            idxLast = idxLast - 1
        End If
        
        i = i + 1
    Wend
    
    ' Make sure the ListBox can display its new contents.
    AddHorizontalScrollToListBox lstFound
    AddHorizontalScrollToListBox lstToCopy
End Sub

' Browse for a source location.
Private Sub cmdBrowseSource_Click()
    Dim strPath As String
    
    ' Get the path from the user.
    strPath = OpenDirectoryDialog(Me.hwnd, "Select the Source Directory")
    If strPath = "" Then
        Exit Sub
    End If
    
    ' Clear the lists.
    lstFound.Clear
    lstToCopy.Clear
    
    ' Set the source path.
    txtSource.Text = strPath
End Sub

' Browse for a destination location.
Private Sub cmdBrowseDestination_Click()
    Dim strPath As String
    
    ' Get the path from the user.
    strPath = OpenDirectoryDialog(Me.hwnd, "Select the Destination Directory")
    If strPath = "" Then
        Exit Sub
    End If
    
    txtDestination.Text = strPath
End Sub

' Move all the items from the To Copy list to the Found list.
Private Sub cmdRemoveAll_Click()
    Dim i As Integer
    
    ' Push all the To Copy items to the Found list.
    For i = 0 To lstToCopy.ListCount - 1
        lstFound.AddItem lstToCopy.List(i)
    Next
    
    ' Clear the To Copy list.
    lstToCopy.Clear
    
    ' Make sure the lists have a nice horizontal scrollbar if needed.
    AddHorizontalScrollToListBox lstFound
    AddHorizontalScrollToListBox lstToCopy
End Sub

' Search the source directory for files.
Private Sub cmdSearch_Click()
    ' Check if the source location is set.
    If txtSource.Text = "" Then
        MsgBox "No source location was set.", vbOKOnly + vbCritical, _
            "No Source Set"
        Exit Sub
    End If
    
    ' Clear the lists.
    lstFound.Clear
    lstToCopy.Clear
    
    ' Go through the source directory looking for its contents.
    GetDirectoryContents txtSource.Text, txtSource.Text
    
    ' Make sure the ListBox can display its new contents.
    AddHorizontalScrollToListBox lstFound
End Sub

' Add extensions to the list.
Private Sub mnuExtensionsAdd_Click()
    Dim strExtension As String
    
    ' Get the extension from the user.
    strExtension = InputBox("Enter a file extension (without the dot):", _
        "Add Extension")
    
    ' Add the item to the list if the user entered anything.
    If strExtension <> "" Then
       lstExtensions.AddItem strExtension
    End If
End Sub

' Remove the selected item on the extensions list.
Private Sub mnuExtensionsDelete_Click()
    ' Check if there are any items selected.
    If lstExtensions.SelCount = 0 Then
        MsgBox "No extensions in the list are selected.", vbOKOnly + vbCritical, _
            "Nothing Selected"
        Exit Sub
    End If
    
    ' Remove the selected item.
    lstExtensions.RemoveItem lstExtensions.ListIndex
End Sub

' Closes the application.
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Browse for a source location.
Private Sub mnuFileOpenSource_Click()
    cmdBrowseSource_Click
End Sub

' Browse for a destination location.
Private Sub mnuFileOpenDestination_Click()
    cmdBrowseDestination_Click
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "MP3 File Lister billymac@nbnet.nb.ca"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearchAll 
      Caption         =   "Search entire computer"
      Height          =   375
      Left            =   7020
      TabIndex        =   16
      Top             =   180
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   570
      Left            =   120
      TabIndex        =   14
      Top             =   4140
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1005
      SortKey         =   1
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MP3 Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox chkShowDirs 
      Caption         =   "Show Directories"
      Height          =   255
      Left            =   5220
      TabIndex        =   7
      Top             =   270
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.CheckBox chkRecurse 
      Caption         =   "Search Subdirectories"
      Height          =   255
      Left            =   3255
      TabIndex        =   6
      Top             =   270
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9480
      TabIndex        =   4
      Top             =   6390
      Width           =   9480
      Begin VB.CommandButton cmdCopyFiles 
         Caption         =   "Save List"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5310
         TabIndex        =   15
         Top             =   45
         Width           =   2145
      End
      Begin VB.CommandButton cmdClearList 
         Caption         =   "&Clear List"
         Height          =   360
         Left            =   135
         TabIndex        =   13
         Top             =   30
         Width           =   1590
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search Now"
         Height          =   360
         Left            =   1770
         TabIndex        =   12
         Top             =   30
         Width           =   1590
      End
      Begin VB.CommandButton cmdRemoveDuplicates 
         Caption         =   "Remove &Duplicates"
         Enabled         =   0   'False
         Height          =   360
         Left            =   3420
         TabIndex        =   10
         Top             =   30
         Width           =   1590
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   360
         Left            =   7950
         TabIndex        =   5
         Top             =   30
         Width           =   1530
      End
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   3675
      TabIndex        =   2
      Top             =   1065
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   135
      TabIndex        =   1
      Top             =   975
      Width           =   1290
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   2970
   End
   Begin VB.Label lblQuantFound 
      Caption         =   "Found:0"
      Height          =   255
      Left            =   150
      TabIndex        =   11
      Top             =   5025
      Width           =   4095
   End
   Begin VB.Label CurSearchPath 
      Caption         =   "c:\"
      Height          =   255
      Left            =   195
      TabIndex        =   9
      Top             =   675
      Width           =   9180
   End
   Begin VB.Label Label1 
      Caption         =   "MP3 Files Found"
      Height          =   255
      Left            =   160
      TabIndex        =   8
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblFilesFound 
      Height          =   210
      Left            =   180
      TabIndex        =   3
      Top             =   3510
      Width           =   4110
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Written By Bill MacIntyre
'Moncton NB Canada
'billymac@nbnet.nb.ca
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Please let me know if you use this program and what you think of it.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'I wrote this program as a tool I needed for my own use.
'I asked a friend that has a ton of MP3's to provide me with a list of what he has.
'The problem is that he had groups of MP3's in various dirs on multiple drives.
'He didn't even know what he actually had himself.
'I needed the list right away so I got him to make a list for each drive using DOS:
'dir *.mp3 /s > c:\filelisting.txt
'this works but was messy and not very good.
'I decided to throw a program together to provide a solution.
'This program is pretty basic.
'It's only purpose in life is to easily search a computer for all MP3 files and
'create a listing to be imported into a spreadsheet or copied to the clipboard to
'paste into an email or into a word processor to create a multi column listing.
'Yes, I realize there are probably a zillion such programs out there but this one
'does exactly what I want it to do and nothing more.

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Feel free to use parts of this code but please do not redistribute it of sell it as is.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim Searching As Boolean

Private Sub chkShowDirs_Click()
    Dir1.Visible = chkShowDirs.Value = vbChecked
    DoResize
End Sub

Private Sub cmdCopyFiles_Click()
    frmOutputoptions.Show , Me
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGenList_Click()
    SearchCurrentFolder
End Sub
Sub SearchCurrentFolder()
    Screen.MousePointer = vbHourglass
    Dim n As Integer

    Dim itmx
    For n = 0 To File1.ListCount - 1
        If Right$(UCase$(File1.List(n)), 4) = ".MP3" Then
            Set itmx = ListView1.ListItems.Add(, , Dir1.Path)
            itmx.SubItems(1) = Left$(File1.List(n), Len(File1.List(n)) - 4)
        End If
    Next

    cmdCopyFiles.Enabled = ListView1.ListItems.Count > 0
    cmdRemoveDuplicates.Enabled = ListView1.ListItems.Count > 0
    lblQuantFound.Caption = "Found: " & ListView1.ListItems.Count
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRemoveDuplicates_Click()
    RemoveDupes ListView1
    cmdCopyFiles.Enabled = ListView1.ListItems.Count > 0
    cmdRemoveDuplicates.Enabled = ListView1.ListItems.Count > 0
End Sub
Sub RemoveDupes(ByRef objList As Object)
    Screen.MousePointer = vbHourglass
    Dim n As Long
    Dim nn As Long
    Dim StartLetter As String
    On Error GoTo errs
    For n = 1 To objList.ListItems.Count - 1
        While UCase$(ListView1.ListItems(n).SubItems(1)) = UCase$(objList.ListItems(n + 1).SubItems(1))
            If n >= objList.ListItems.Count Then Exit For
            objList.ListItems.Remove (n + 1)
        Wend
    Next

errs:
    lblQuantFound.Caption = "Found: " & objList.ListItems.Count
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSearch_Click()
    Screen.MousePointer = vbHourglass
    Searching = Not Searching

    If Not Searching Then
        cmdSearch.Caption = "&Search Now"
    Else
        cmdSearch.Caption = "&Stop Search"
    End If

    If Not Searching Then Exit Sub

    ReDim sarray(0) As String

    If chkRecurse.Value = vbChecked Then
        SearchCurrentFolder
        Call DirWalk("*.mp3", Dir1.Path, sarray)
    Else
        SearchCurrentFolder
    End If

    ListView1.Refresh
    cmdSearch.Caption = "&Search Now"
    Searching = False
    Screen.MousePointer = vbDefault

End Sub


Private Sub cmdClearList_Click()
    ListView1.ListItems.Clear
    cmdCopyFiles.Enabled = False
End Sub

Private Sub cmdSearchAll_Click()

If MsgBox("are you sure you want to search all the drives on this computer for MP3 files?" & vbNewLine & "This may take a while", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
Dim OldChkVal As Boolean
Dim n As Long
OldChkVal = chkRecurse.Value
chkRecurse.Value = vbChecked
On Error Resume Next

cmdClearList_Click

For n = 1 To Drive1.ListCount - 1
    Drive1.ListIndex = n
    Dir1.Path = "\" 'Drive1.Drive
    cmdSearch_Click
    While Searching
    DoEvents
    Wend
Next

chkRecurse.Value = OldChkVal

End Sub

Private Sub Dir1_Change()
    
    File1.Path = Dir1.Path
    If chkShowDirs.Value <> 1 Then CurSearchPath.Caption = File1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo errs
    Dir1.Path = Drive1.Drive
    chkShowDirs.Value = vbChecked
    Exit Sub
errs:
    MsgBox Err.Description, vbExclamation + vbOKOnly, "Error"

End Sub

Private Sub Form_Load()
    CurSearchPath.Caption = ""
    File1.Pattern = "*.mp3"
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "File Path", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "File Name", ListView1.Width / 2
    ListView1.View = lvwReport
End Sub

Private Sub Form_Resize()
    DoResize
End Sub
Sub DoResize()
    On Error Resume Next
    If Dir1.Visible = True Then
        ListView1.Top = 4000
        Label1.Top = ListView1.Top - Label1.Height + 60
    Else
        ListView1.Top = 1240
        Label1.Top = 1030
    End If

    Dir1.Width = Me.Width - 400
    ListView1.Width = Me.Width - 400
    CurSearchPath.Width = Me.Width - 100
    ListView1.Height = Me.Height - ListView1.Top - Picture1.Height - 600
    cmdExit.Left = Me.Width - 250 - cmdExit.Width
    lblQuantFound.Top = ListView1.Top + ListView1.Height

    ListView1.ColumnHeaders.Item(1).Width = ListView1.Width \ 2
    ListView1.ColumnHeaders(2).Width = ListView1.Width - ListView1.ColumnHeaders.Item(1).Width - 100

End Sub
Sub DirWalk(ByVal Spattern As String, ByVal CurDir As String, SFound() As String)

' some of this sub was borrowed from someone.
' I am not sure who wrote it, but thanks to whoever
    
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    Dim i As Integer
    Dim sCurPath As String
    Dim sFile As String
    Dim ii As Integer
    Dim ifiles As Integer
    Dim ilen As Integer

    If Right$(CurDir, 1) <> "\" Then
        Dir1.Path = CurDir & "\"
    Else
        Dir1.Path = CurDir
    End If

    For i = 0 To Dir1.ListCount
        If Not Searching Then Exit Sub
        If Dir1.List(i) <> "" Then
            DoEvents
            Call DirWalk(Spattern, Dir1.List(i), SFound())
        Else
            If Right$(Dir1.Path, 1) = "\" Then
                sCurPath = Left$(Dir1.Path, Len(Dir1.Path) - 1)
            Else
                sCurPath = Dir1.Path
            End If
            File1.Path = sCurPath
            File1.Pattern = Spattern

            If File1.ListCount > 0 Then
                ' Matching files found in the dir
                For ii = 0 To File1.ListCount - 1
                    ReDim Preserve SFound(UBound(SFound) + 1)
                    SFound(UBound(SFound) - 1) = sCurPath & "\" & File1.List(ii)

                    Dim itmx
                    Set itmx = ListView1.ListItems.Add(, , sCurPath)
                    itmx.SubItems(1) = Left$(File1.List(ii), Len(File1.List(ii)) - 4)

                    lblQuantFound.Caption = "Found: " & ListView1.ListItems.Count

                Next ii
            End If

            ilen = Len(Dir1.Path)
            Do While Mid(Dir1.Path, ilen, 1) <> "\"
                ilen = ilen - 1
            Loop
            Dir1.Path = Mid(Dir1.Path, 1, ilen)
        End If
    Next i

    cmdCopyFiles.Enabled = ListView1.ListItems.Count > 0
    cmdRemoveDuplicates.Enabled = ListView1.ListItems.Count > 0

    Screen.MousePointer = vbDefault
End Sub



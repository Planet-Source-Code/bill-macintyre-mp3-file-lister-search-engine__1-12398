VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOutputoptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Output Format"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Listing"
      Height          =   300
      Left            =   3780
      TabIndex        =   16
      Top             =   3420
      Width           =   1485
   End
   Begin VB.Frame frmDelimit 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1740
      Visible         =   0   'False
      Width           =   3555
      Begin VB.OptionButton optDelimitType 
         Caption         =   "Tab"
         Height          =   195
         Index           =   1
         Left            =   1140
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optDelimitType 
         Caption         =   "Comma"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame frmIncludePath 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   1860
      TabIndex        =   8
      Top             =   1140
      Visible         =   0   'False
      Width           =   2955
      Begin VB.OptionButton optPathoption 
         Caption         =   "At beginning - f:\mp3\filename.mp3"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2865
      End
      Begin VB.OptionButton optPathoption 
         Caption         =   "At end - Filename.mp3 (f:\mp3)"
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2865
      End
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Delimited"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1740
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Include file path"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1635
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Include File Extension (.mp3)"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Value           =   1  'Checked
      Width           =   2385
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Include item Index ( 1. )"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Checked
      Width           =   4725
   End
   Begin VB.CommandButton cmdSaveToFile 
      Caption         =   "&Save to File"
      Height          =   300
      Left            =   2340
      TabIndex        =   3
      Top             =   3420
      Width           =   1350
   End
   Begin VB.CommandButton cmdCopytoClipBoard 
      Caption         =   "Copy to &Windows Clipboard"
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   3420
      Width           =   2100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   3780
      TabIndex        =   1
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Label lblSample 
      Caption         =   "Sample:"
      Height          =   675
      Left            =   120
      TabIndex        =   15
      Top             =   2580
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Sample:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Define Output Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   3285
   End
End
Attribute VB_Name = "frmOutputoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOptions_Click(Index As Integer)
    Select Case Index
    Case 0
    Case 1
    Case 2
        frmIncludePath.Visible = chkOptions(2).Value = vbChecked
    Case 3
      frmDelimit.Visible = chkOptions(3).Value = vbChecked
    End Select
    
    SetSample
    'lblSample.Caption = BuildOutputItem(1, "c:\mp3", "The Artist - The Group")
    
End Sub
Function BuildOutput()
    Screen.MousePointer = vbHourglass
    Dim n As Long
    Dim tmp As String

    For n = 1 To frmMain.ListView1.ListItems.Count
        tmp = tmp & BuildOutputItem(n, frmMain.ListView1.ListItems(n), frmMain.ListView1.ListItems(n).SubItems(1))
    Next
    BuildOutput = tmp
    Screen.MousePointer = vbDefault
End Function
Function SetSample()
    Screen.MousePointer = vbHourglass
    Dim n As Long
    Dim tmp As String

    For n = 1 To 3
        tmp = tmp & BuildOutputItem(n, frmMain.ListView1.ListItems(n), frmMain.ListView1.ListItems(n).SubItems(1))
    Next
    lblSample.Caption = tmp
    Screen.MousePointer = vbDefault
End Function

Function BuildOutputItem(itemindex As Long, strFilePath As String, strSong As String) As String
    Dim tmpItem As String
    Dim strDelimiter As String
    
    'determine delimiter
    strDelimiter = " - "
    If chkOptions(3).Value = vbChecked Then
        If optDelimitType(0).Value = True Then strDelimiter = ","
        If optDelimitType(1).Value = True Then strDelimiter = vbTab
    End If
    
    'add index
    If chkOptions(0).Value = vbChecked Then tmpItem = itemindex & "." & strDelimiter
    
    'add path part to beginning if selected
    If chkOptions(2).Value = vbChecked And optPathoption(0).Value = True Then
        tmpItem = tmpItem & strFilePath & strDelimiter
    End If
    
    'add song name
        tmpItem = tmpItem & strSong
    
    'add extension if selected
    If chkOptions(1).Value = vbChecked Then
        tmpItem = tmpItem & ".mp3"
    End If
    
    tmpItem = tmpItem & strDelimiter
    
    'add path part to end if selected
    If chkOptions(2).Value = vbChecked And optPathoption(1).Value = True Then
        tmpItem = tmpItem & strFilePath
    End If
    
    If Right$(tmpItem, Len(strDelimiter)) = strDelimiter Then tmpItem = Left$(tmpItem, Len(tmpItem) - Len(strDelimiter))
    
    'add a new line
    tmpItem = tmpItem & vbNewLine
    BuildOutputItem = tmpItem

End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCopytoClipBoard_Click()
  Clipboard.Clear
  Clipboard.SetText (BuildOutput)
End Sub

Private Sub cmdEdit_Click()
    frmlisting.strlisting = BuildOutput
    frmlisting.Show , frmMain
    Unload Me
End Sub

Private Sub cmdSaveToFile_Click()

    Screen.MousePointer = vbHourglass
    
    On Error GoTo errs
        
    CommonDialog1.DialogTitle = "Save MP3 listing as"
    Dim LastSavePath As String
    LastSavePath = GetSetting(App.EXEName, "Settings", "LastSavePath", App.Path)
    
    CommonDialog1.InitDir = LastSavePath
    CommonDialog1.FileName = "Mp3Listing.txt"
    
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files(*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
    Dim filenum As Integer
    Close
    filenum = FreeFile
    'see if file already exists
    If Dir(CommonDialog1.FileName, vbNormal) <> "" Then
    If MsgBox("The file " & CommonDialog1.FileName & " already exists." & vbNewLine & "Do you want to overwrite the file?", vbQuestion + vbDefaultButton2 + vbYesNo) <> vbYes Then
      GoTo errs
    End If
    
    End If
    
    Open CommonDialog1.FileName For Append As #filenum
    Close
    filenum = FreeFile
    Open CommonDialog1.FileName For Output As #filenum
    Dim n As Long
    Dim tmp As String
    For n = 1 To frmMain.ListView1.ListItems.Count
        Print #filenum, Replace(BuildOutputItem(n, frmMain.ListView1.ListItems(n), frmMain.ListView1.ListItems(n).SubItems(1)), vbNewLine, "")
    Next
    Close
    
   'remember the save path
   SaveSetting App.EXEName, "Settings", "LastSavePath", CommonDialog1.FileName
    End If
    
errs:
    Debug.Print Err.Description
    Close
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    SetSample
End Sub

Private Sub optDelimitType_Click(Index As Integer)
  SetSample

End Sub

Private Sub optPathoption_Click(Index As Integer)
 SetSample
 End Sub

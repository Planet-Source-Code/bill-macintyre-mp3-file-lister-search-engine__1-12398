VERSION 5.00
Begin VB.Form frmlisting 
   Caption         =   "Edit Listing"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4740
      TabIndex        =   1
      Top             =   4140
      Width           =   1485
   End
   Begin VB.TextBox txtListing 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   6075
   End
End
Attribute VB_Name = "frmlisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strlisting As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
txtListing.Text = strlisting
End Sub

Private Sub Form_Resize()
txtListing.Width = Me.Width - 500
txtListing.Height = Me.Height - 1200
cmdCancel.Left = Me.Width - cmdCancel.Width - 300
cmdCancel.Top = Me.Height - cmdCancel.Height - 400
End Sub

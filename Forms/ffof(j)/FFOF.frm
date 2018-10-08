VERSION 5.00
Begin VB.Form FFOF 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«‰ Œ«» ’Ê  "
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   Icon            =   "FFOF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2880
      Width           =   4455
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "DblClich to Clear"
      Top             =   120
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2920
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter :"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   6
      Top             =   195
      Width           =   420
   End
End
Attribute VB_Name = "FFOF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
Text1.Text = "*"
File1.Pattern = "*." & Text1.Text
File1.Refresh
Else
File1.Pattern = "*." & Text1.Text
File1.Refresh
Beep
End If
End Sub

Private Sub Command2_Click()
MP3F_N.Show
FFOF.Hide

MP3F_N.Enabled = True
MP3F_N.url_text.Text = Me.Text2.Text

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Dir1_Change()
On Error Resume Next

File1.Path = Dir1.Path
Text2.Text = Dir1.Path
End Sub
Private Sub Drive1_Change()
On Error Resume Next

Dir1.Path = Drive1.Drive
Text2.Text = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next

Text2.Text = File1.Path & "\" & File1.FileName

End Sub

Private Sub Form_Load()
On Error Resume Next

Text1 = "*"
Command1.Default = True
Text2.Text = File1.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

MP3F.Show
MP3F.Enabled = True
Unload Me
End Sub

Private Sub Label3_Click()
On Error Resume Next

MsgBox "Find Format Of File (FFOF)" + Chr$(10) + "By:Reza Mohiti" + Chr$(10) + "$Version 1.2$" + Chr$(10) + "#1387/04/18#", vbInformation, "FFOF"
End Sub
Private Sub Text1_DblClick()
On Error Resume Next

Text1.Text = ""
End Sub


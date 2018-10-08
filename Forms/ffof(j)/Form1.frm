VERSION 5.00
Begin VB.Form FFOF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Format Of File   (FFOF)"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2880
      Width           =   5415
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
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
      Caption         =   "R"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2920
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filter :"
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
Private Sub Dir1_Change()
File1.Path = Dir1.Path
Text2.Text = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Text2.Text = Drive1.Drive
End Sub
Private Sub Form_Load()
Text1 = "*"
Command1.Default = True
Text2.Text = File1.Path
End Sub
Private Sub Label3_Click()
MsgBox "Find Format Of File (FFOF)" + Chr$(10) + "By:Reza Mohiti" + Chr$(10) + "$Version 1.2$" + Chr$(10) + "#1387/04/18#", vbInformation, "FFOF"
End Sub
Private Sub Text1_DblClick()
Text1.Text = ""
End Sub


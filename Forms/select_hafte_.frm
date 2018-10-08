VERSION 5.00
Begin VB.Form select_hafte_F 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÇäÊÎÇÈ ÑæÒ åÇí åÝÊå"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404000&
   Icon            =   "select_hafte_.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox WATT 
      Height          =   1815
      Left            =   3720
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox ayam_hafte 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Height          =   2820
      ItemData        =   "select_hafte_.frx":030A
      Left            =   2400
      List            =   "select_hafte_.frx":030C
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox rozhaye_pahkhs 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Height          =   2820
      ItemData        =   "select_hafte_.frx":030E
      Left            =   120
      List            =   "select_hafte_.frx":0310
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Adobe Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Adobe Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Adobe Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Adobe Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "                  ËÈÊ                   "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   720
      TabIndex        =   6
      Top             =   3360
      Width           =   2640
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÑæÒåÇí ÎÔ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÇíÇã åÝÊå"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "select_hafte_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim what_ As String

Private Sub ayam_hafte_Click()
Command4.Enabled = True
Command3.Enabled = True

End Sub

Private Sub ayam_hafte_DblClick()
Call Command4_Click

End Sub

Private Sub Command1_Click()
rozhaye_pahkhs.Clear
ayam_hafte.Clear

ayam_hafte.AddItem ("ÔäÈå")
ayam_hafte.AddItem ("í˜ ÔäÈå")
ayam_hafte.AddItem ("Ïæ ÔäÈå")
ayam_hafte.AddItem ("Óå ÔäÈå")
ayam_hafte.AddItem ("åÇÑ ÔäÈå")
ayam_hafte.AddItem ("äÌ ÔäÈå")
ayam_hafte.AddItem ("ÌãÚå")
If rozhaye_pahkhs.ListCount = 0 Then
Command1.Enabled = False

Command2.Enabled = False
Else
Command1.Enabled = True

Command2.Enabled = True

End If
End Sub

Private Sub Command2_Click()

ayam_hafte.AddItem (rozhaye_pahkhs.Text)
rozhaye_pahkhs.RemoveItem (rozhaye_pahkhs.ListIndex)
rozhaye_pahkhs.Text = rozhaye_pahkhs.List(0)
If rozhaye_pahkhs.ListCount = 0 Then
Command1.Enabled = False

Command2.Enabled = False
Else
Command1.Enabled = True

Command2.Enabled = True

End If

End Sub

Private Sub Command3_Click()
ayam_hafte.Clear
rozhaye_pahkhs.Clear

rozhaye_pahkhs.AddItem ("ÔäÈå")
rozhaye_pahkhs.AddItem ("í˜ ÔäÈå")
rozhaye_pahkhs.AddItem ("Ïæ ÔäÈå")
rozhaye_pahkhs.AddItem ("Óå ÔäÈå")
rozhaye_pahkhs.AddItem ("åÇÑ ÔäÈå")
rozhaye_pahkhs.AddItem ("äÌ ÔäÈå")
rozhaye_pahkhs.AddItem ("ÌãÚå")
If ayam_hafte.ListCount = 0 Then
Command4.Enabled = False

Command3.Enabled = False
Else
Command3.Enabled = True

Command4.Enabled = True

End If

End Sub

Private Sub Command4_Click()
'On Error Resume Next

rozhaye_pahkhs.AddItem (ayam_hafte.Text)
ayam_hafte.RemoveItem (ayam_hafte.ListIndex)
ayam_hafte.Text = ayam_hafte.List(0)
If ayam_hafte.ListCount = 0 Then
Command4.Enabled = False

Command3.Enabled = False
Else
Command3.Enabled = True

Command4.Enabled = True

End If




End Sub

Private Sub Form_Load()
ayam_hafte.AddItem ("ÔäÈå")
ayam_hafte.AddItem ("í˜ ÔäÈå")
ayam_hafte.AddItem ("Ïæ ÔäÈå")
ayam_hafte.AddItem ("Óå ÔäÈå")
ayam_hafte.AddItem ("åÇÑ ÔäÈå")
ayam_hafte.AddItem ("äÌ ÔäÈå")
ayam_hafte.AddItem ("ÌãÚå")

End Sub

Private Sub Label2_Click()
If WATT.Text = "mehvar" Then
List_pakhsh_F.Label12.Caption = ""
For I = 1 To rozhaye_pahkhs.ListCount

List_pakhsh_F.Label12.Caption = List_pakhsh_F.Label12.Caption & rozhaye_pahkhs.List(I - 1) & " , "
Next I
Unload Me

ElseIf WATT.Text = "date" Then
List_pakhsh_F.Label6.Caption = ""
For I = 1 To rozhaye_pahkhs.ListCount

List_pakhsh_F.Label6.Caption = List_pakhsh_F.Label6.Caption & rozhaye_pahkhs.List(I - 1) & " , "
Next I
Unload Me
End If

End Sub

Private Sub rozhaye_pahkhs_Click()
Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub rozhaye_pahkhs_DblClick()
Call Command2_Click

End Sub

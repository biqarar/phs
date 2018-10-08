VERSION 5.00
Begin VB.Form WE 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5820
   ClientLeft      =   6450
   ClientTop       =   3780
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "»«“ê‘ "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”ŒÂ 2 :: 1392"
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
         Left            =   2160
         TabIndex        =   11
         Top             =   3000
         Width           =   1785
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»”„ «··Â «·—Õ„‰ «·—ÕÌ„"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·”·«„ ⁄·Ìò Ì« ›«ÿ„Â «·„⁄’Ê„Â"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   2505
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1390/07/17"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   2400
         TabIndex        =   8
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”«„«‰Â ÂÊ‘„‰œ ’Ê Ì"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   840
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   3525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—÷« „ÕÌÿÌ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   1800
         TabIndex        =   6
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":»—‰«„Â ‰ÊÌ”"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   3360
         TabIndex        =   5
         Top             =   3720
         Width           =   1020
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rm.Biqarar@Gmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   1680
         TabIndex        =   4
         Top             =   4680
         Width           =   2505
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":’‰œÊﬁ Å” Ì"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   3360
         TabIndex        =   3
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37148-16445"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   1320
         TabIndex        =   2
         Top             =   4200
         Width           =   1440
      End
   End
End
Attribute VB_Name = "WE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)

Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

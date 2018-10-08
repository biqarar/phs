VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form MP3F 
   Caption         =   "À»  «ÿ·«⁄«  ’Ê Ì"
   ClientHeight    =   8715
   ClientLeft      =   3015
   ClientTop       =   1920
   ClientWidth     =   15705
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MP3F.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   15705
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox list1 
      BackColor       =   &H80000002&
      Height          =   2475
      ItemData        =   "MP3F.frx":08CA
      Left            =   120
      List            =   "MP3F.frx":08CC
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      Picture         =   "MP3F.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "À»  ê—ÊÂ Ê  ⁄ÌÌ‰ ‘„«—Â ’Ê "
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã” ÃÊ"
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9240
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MP3F.frx":432B
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12640511
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   30
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«ÿ·«⁄«  ’Ê Ì"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "òœ ’Ê "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Goroh"
         Caption         =   "ê—ÊÂ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Name"
         Caption         =   "‰«„"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "OnOff"
         Caption         =   "›⁄«·0 €Ì— ›⁄«·"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "URL"
         Caption         =   "¬œ—”"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Tozih"
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Vol"
         Caption         =   "„Ì“«‰ ’œ«"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3195.213
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   4334.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4229.858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   945.071
         EndProperty
      EndProperty
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "URL::"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   14
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "_"
      DataField       =   "URL"
      DataSource      =   "MP3D"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4320
      TabIndex        =   13
      Top             =   6960
      Width           =   105
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP2 
      Height          =   675
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   3375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5953
      _cy             =   1191
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   " Ê÷ÌÕ« "
      Height          =   345
      Left            =   13560
      TabIndex        =   11
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "¬œ—”"
      Height          =   345
      Left            =   13680
      TabIndex        =   10
      Top             =   1680
      Width           =   390
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "‰«„"
      Height          =   345
      Left            =   13920
      TabIndex        =   9
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "ê—ÊÂ"
      Height          =   345
      Left            =   13800
      TabIndex        =   8
      Top             =   720
      Width           =   285
   End
   Begin VB.Menu mnuoption 
      Caption         =   " ‰ŸÌ„« "
      Begin VB.Menu mnusabt 
         Caption         =   "À»  ’Ê "
      End
      Begin VB.Menu mnudel 
         Caption         =   "Õ–› ’Ê "
      End
      Begin VB.Menu mnustrip 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuonoff 
         Caption         =   "›⁄«· -€Ì— ›⁄«·"
         Begin VB.Menu mnuon 
            Caption         =   "›⁄«·"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuoff 
            Caption         =   "€Ì— ›⁄«·"
            Shortcut        =   ^Q
         End
      End
      Begin VB.Menu DF 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MNOOPEN 
         Caption         =   "ÅŒ‘ ’Ê "
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu ssss 
      Caption         =   "ê—ÊÂ »‰œÌ"
      Begin VB.Menu dellgoroh 
         Caption         =   "Õ–› ê—ÊÂ"
      End
   End
End
Attribute VB_Name = "MP3F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If DataGrid1.AllowUpdate = False Then
DataGrid1.AllowUpdate = True
Else
DataGrid1.AllowUpdate = False
End If

End Sub

Private Sub Combo1_Click()
'On Error Resume Next
 Exit Sub

MP3D.Refresh
MP3D.RecordSource = "select * from mp3d" ' where kod like ('" & "" & "%')"
MP3D.Refresh




MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where kod like ('" & Mid(Combo1.Text, 1, 2) & "%')"
MP3D.Refresh

'MsgBox Mid(Combo1.Text, 3, 4)
    If Mid(Combo1.Text, 3, 4) = "List" Then
    Call ListCounter_Click
    
    Exit Sub
    End If
    

Dim x As Integer
x = Val(MP3D.Recordset.Fields("kod"))
For I = 1 To MP3D.Recordset.RecordCount
If Val(x) < Val(MP3D.Recordset.Fields("kod")) Then
x = Val(MP3D.Recordset.Fields("kod"))
End If
MP3D.Recordset.MoveNext
Next I

If x = 0 Then x = Mid(Combo1.Text, 1, 4)
Text2.Text = (x + 1)


If Check2.Value = 1 Then
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where goroh like ('%" + Combo1.Text + "%')"
MP3D.Refresh
End If



End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text3.Text = "" Or Combo1.Text = "«‰ Œ«» ò‰Ìœ" Then
MsgBox "«ÿ·«⁄«  —« »Â ’Ê—  ò«„· Ê«—œ ò‰Ìœ. œ—Ã ⁄‰Ê«‰ ê—ÊÂ° ‰«„ Ê ¬œ—” «·“«„Ì «” ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where kod like ('%" + Text2.Text + "%')"
MP3D.Refresh
If MP3D.Recordset.BOF = True Or MP3D.Recordset.EOF = True Then
GoTo 1
Else
MsgBox "»Â œ·Ì·  ò—«—Ì »Êœ‰ òœ ’Ê  ⁄„·Ì«  À»  ’Ê  „ Êﬁ› „Ì ‘Êœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
1:
If sg.Value = 1 Then
For I = 1 To List3.ListCount
MP3D.Refresh
MP3D.Recordset.AddNew
MP3D.Recordset.Fields("kod") = Text2.Text + I
MP3D.Recordset.Fields("goroh") = Combo1.Text
MP3D.Recordset.Fields("name") = Text3.Text & "(" & I & ")"
MP3D.Recordset.Fields("onoff") = "On"
MP3D.Recordset.Fields("systemonoff") = "On"

MP3D.Recordset.Fields("url") = Text1.Text & "\" & List3.List(I)
MP3D.Recordset.Fields("tozih") = Text5.Text
MP3D.Recordset.Fields("vol") = WMP2.settings.volume

MP3D.Recordset.AddNew
'MP3D.Recordset.Update
MP3D.Refresh
Next I
MsgBox "sdfsdfsdfsdf"

Else





MP3D.Refresh
MP3D.Recordset.AddNew
MP3D.Recordset.Fields("kod") = Text2.Text
MP3D.Recordset.Fields("goroh") = Combo1.Text
MP3D.Recordset.Fields("name") = Text3.Text
MP3D.Recordset.Fields("onoff") = "On"
MP3D.Recordset.Fields("systemonoff") = "On"

MP3D.Recordset.Fields("url") = Text1.Text
MP3D.Recordset.Fields("tozih") = Text5.Text
MP3D.Recordset.Fields("vol") = WMP2.settings.volume

MP3D.Recordset.AddNew
'MP3D.Recordset.Update
MP3D.Refresh

DataGrid1.Refresh

MsgBox "⁄„·Ì«  À»  «ÿ·«⁄«  ’Ê Ì »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ", vbInformation + vbOKOnly, "À»  ’Ê "
DataGrid1.Refresh

MP3D.Refresh
MP3D.Refresh
MP3D.RecordSource = "select * from MP3D where kod like ('%" & "" & "%') "
MP3D.Refresh


End If


End Sub

Private Sub Command10_Click()
Dim kod_gorohlist, kod_sot As String
kod_sot = Mid(List2.Text, 1, 4)

gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where xname like ('" & list1.Text & "')"
gorohlist.Refresh
'If gorohlist.Recordset.BOF = False Or gorohlist.Recordset.EOF = False Then

kod_gorohlist = gorohlist.Recordset.Fields("kodgoroh")
sotgoroh.Refresh
sotgoroh.RecordSource = "select * from sotgoroh where kodgoroh like ('" & kod_gorohlist & "') and kodsot like ('" & kod_sot & "')"
sotgoroh.Refresh
If sotgoroh.Recordset.BOF = True Or sotgoroh.Recordset.EOF = True Then
MsgBox "«Ì‰ ’Ê  œ— ê—ÊÂ ÊÃÊœ ‰œ«—œ", vbInformation + vbOKOnly, "ê—ÊÂ »‰œÌ ’Ê "
Exit Sub
Else
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ’Ê  «“ ê—ÊÂ Õ–› ‘Êœ", vbQuestion + vbYesNo, "ê—ÊÂ »‰œÌ ’Ê ") = vbYes Then
sotgoroh.Recordset.Delete
'Dim kod_gorohlist, kod_sot As String
'kod_sot = MP3D.Recordset.Fields("kod")
gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where xname like ('" & list1.Text & "')"
gorohlist.Refresh
'If gorohlist.Recordset.BOF = False Or gorohlist.Recordset.EOF = False Then

kod_gorohlist = gorohlist.Recordset.Fields("kodgoroh")

List2.Clear
sotgoroh.Refresh
sotgoroh.RecordSource = "select * from sotgoroh where kodgoroh like ('" & kod_gorohlist & "')" ' and kodsot like ('" & kod_sot & "')"
sotgoroh.Refresh
For I = 1 To sotgoroh.Recordset.RecordCount
mp3d2.Refresh
mp3d2.RecordSource = "select * from mp3d where kod like ('" & sotgoroh.Recordset.Fields("kodsot") & "')"
mp3d2.Refresh
If mp3d2.Recordset.BOF = True Or mp3d2.Recordset.EOF = True Then ' in fail qablan pak shode
sotgoroh.Recordset.MoveNext

Else
List2.AddItem (mp3d2.Recordset.Fields("kod") & " :: " & mp3d2.Recordset.Fields("name") & " :: " & mp3d2.Recordset.Fields("tozih"))
sotgoroh.Recordset.MoveNext

End If
Next I

End If
End If


End Sub

Private Sub Command2_Click()
FFOF.Show
MP3F.Enabled = False

End Sub

Private Sub Command3_Click()
On Error Resume Next

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ ’Ê  Õ–› ‘Êœ", vbQuestion + vbYesNo, "Õ–› ’Ê ") = vbYes Then
MP3D.Recordset.Delete
End If
End Sub



Private Sub Command4_Click()
 MP3D.Refresh
                    MP3D.RecordSource = "select * from MP3D where kod like ('" & "3" & "%') "
                    MP3D.Refresh
                    
                    'x = Int((Rnd(1000) * MP3D.Recordset.RecordCount) + 3000 + 1)
                    x = Int((Rnd(1000) * MP3D.Recordset.RecordCount) + 1)
                    For I = 1 To x - 1
                    MP3D.Recordset.MoveNext
                    Next I
                    MsgBox MP3D.Recordset.Fields("kod")
                'MP3D.Refresh
                'MP3D.RecordSource = "select * from MP3D where kod like ('%" & x & "%') "
                'MP3D.Refresh
'              Wmp1.URL = MP3D.Recordset.Fields("url")
            '
End Sub

Private Sub Command5_Click()
Dim kod_gorohlist, kod_sot As String
kod_sot = MP3D.Recordset.Fields("kod")
gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where xname like ('" & list1.Text & "')"
gorohlist.Refresh
'If gorohlist.Recordset.BOF = False Or gorohlist.Recordset.EOF = False Then

kod_gorohlist = gorohlist.Recordset.Fields("kodgoroh")

sotgoroh.Refresh
sotgoroh.RecordSource = "select * from sotgoroh where kodgoroh like ('" & kod_gorohlist & "') and kodsot like ('" & kod_sot & "')"
sotgoroh.Refresh
If sotgoroh.Recordset.BOF = False Or sotgoroh.Recordset.EOF = False Then
MsgBox "«Ì‰ ’Ê  œ— «Ì‰ ê—ÊÂ ÊÃÊœ œ«—œ", vbInformation + vbOKOnly, "ê—ÊÂ »‰œÌ ’œ«"
Exit Sub
Else
sotgoroh.Refresh
sotgoroh.Recordset.AddNew
sotgoroh.Recordset.Fields("kodgoroh") = kod_gorohlist
sotgoroh.Recordset.Fields("kodsot") = kod_sot
sotgoroh.Recordset.Update
sotgoroh.Refresh
List2.Clear
sotgoroh.Refresh
sotgoroh.RecordSource = "select * from sotgoroh where kodgoroh like ('" & kod_gorohlist & "')" ' and kodsot like ('" & kod_sot & "')"
sotgoroh.Refresh
For I = 1 To sotgoroh.Recordset.RecordCount
mp3d2.Refresh
mp3d2.RecordSource = "select * from mp3d where kod like ('" & sotgoroh.Recordset.Fields("kodsot") & "')"
mp3d2.Refresh
If mp3d2.Recordset.BOF = True Or mp3d2.Recordset.EOF = True Then ' in fail qablan pak shode
sotgoroh.Recordset.MoveNext

Else
List2.AddItem (mp3d2.Recordset.Fields("kod") & " :: " & mp3d2.Recordset.Fields("name") & " :: " & mp3d2.Recordset.Fields("tozih"))
sotgoroh.Recordset.MoveNext

End If
Next I
End If


End Sub

Private Sub Command6_Click()
Call Command1_Click


End Sub

Private Sub Command7_Click()
mp3d2.Refresh
mp3d2.RecordSource = "select * from mp3d"
mp3d2.Refresh
For I = 1000 To mp3d2.Recordset.RecordCount + 1000

mp3d2.Refresh
mp3d2.RecordSource = "select * from mp3d where kod like('%" & I & "%')"
mp3d2.Refresh
If mp3d2.Recordset.BOF = True Or mp3d2.Recordset.EOF = True Then
GoTo 1
End If
Next I
1:
Text2.Text = I
gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where xname like ('" & Combo1.Text & "')"
gorohlist.Refresh
If gorohlist.Recordset.BOF = True Or gorohlist.Recordset.EOF = True Then
'kasi peyda nashod
gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where kodgoroh like('%" & "" & "%')"
gorohlist.Refresh
For j = 100 To gorohlist.Recordset.RecordCount + 100

gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where kodgoroh like('" & j & "')"
gorohlist.Refresh
If gorohlist.Recordset.BOF = True Or gorohlist.Recordset.EOF = True Then
GoTo 2
End If
Next j
2:
gorohlist.Refresh
gorohlist.Recordset.AddNew
gorohlist.Recordset.Fields("kodgoroh") = j
gorohlist.Recordset.Fields("xname") = Combo1.Text
gorohlist.Recordset.Fields("xdate") = Taqvim.Tarikh.Caption

gorohlist.Recordset.Update
gorohlist.Refresh
gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where kodgoroh like('%" & "" & "%')"
gorohlist.Refresh
gorohlist.Refresh
Combo1.Clear
list1.Clear

For I = 1 To gorohlist.Recordset.RecordCount
Combo1.AddItem (gorohlist.Recordset.Fields("xname"))
list1.AddItem (gorohlist.Recordset.Fields("xname"))

gorohlist.Recordset.MoveNext
Next I

Else

End If


End Sub

Private Sub Command8_Click()
On Error Resume Next

MP3D.Recordset.MoveNext

End Sub

Private Sub Command9_Click()
On Error Resume Next

MP3D.Recordset.MovePrevious


End Sub


Private Sub DataGrid1_Click()
On Error Resume Next


MP3D.Recordset.Update

End Sub


Private Sub dellgoroh_Click()
Dim kod_gorohlist, kod_sot As String
'kod_sot = MP3D.Recordset.Fields("kod")
gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where xname like ('" & list1.Text & "')"
gorohlist.Refresh
'If gorohlist.Recordset.BOF = False Or gorohlist.Recordset.EOF = False Then

kod_gorohlist = gorohlist.Recordset.Fields("kodgoroh")
If gorohlist.Recordset.RecordCount = 1 Then
    '    If MsgBox("") = vbYes Then
        sotgoroh.Refresh
        sotgoroh.RecordSource = "select * from sotgoroh where kodgoroh like ('" & kod_gorohlist & "')" ' and kodsot like ('" & kod_sot & "')"
        sotgoroh.Refresh
            If sotgoroh.Recordset.BOF = True Or sotgoroh.Recordset.EOF = True Then
                    If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ê—ÊÂ —« «“ ·Ì”  Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› ê—ÊÂ") = vbYes Then
                    
                    gorohlist.Recordset.Delete
                        gorohlist.Refresh
                        gorohlist.RecordSource = "select * from gorohlist where kodgoroh like('%" & "" & "%')"
                        gorohlist.Refresh
                        gorohlist.Refresh
                        Combo1.Clear
                        list1.Clear
                        
                        For I = 1 To gorohlist.Recordset.RecordCount
                        Combo1.AddItem (gorohlist.Recordset.Fields("xname"))
                        list1.AddItem (gorohlist.Recordset.Fields("xname"))
                        
                        gorohlist.Recordset.MoveNext
                        Next I
                    Else
                    Exit Sub
                    End If
            Else
            MsgBox " ⁄œ«œÌ ’Ê  œ— «Ì‰ ê—ÊÂ Ì«›  ‘œ ·ÿ›« ﬁ»· «“ Õ–› ê—ÊÂ∫ ’Ê  Â« —« «“ ê—ÊÂ Å«ò ò‰Ìœ", vbExclamation + vbOKOnly, "Â‘œ«—"
            Exit Sub
            
            End If

End If
End Sub

 Private Sub Form_Load()
 Combo1.Clear
list1.Clear

 
gorohlist.Refresh
For I = 1 To gorohlist.Recordset.RecordCount
Combo1.AddItem (gorohlist.Recordset.Fields("xname"))
list1.AddItem (gorohlist.Recordset.Fields("xname"))
gorohlist.Recordset.MoveNext
Next I

 
 GoTo 1
 
Combo1.AddItem ("1000" & "-" & "ﬁ—¬‰")
Combo1.AddItem ("2000" & "-" & "«–«‰")
Combo1.AddItem ("3000" & "-" & "„‰«Ã« ")
Combo1.AddItem ("4000" & "-" & "œ⁄«")
Combo1.AddItem ("5000" & "-" & "„œ«ÕÌ")
Combo1.AddItem ("6000" & "-" & "ò·ÌÅ ’Ê Ì")

'Combo1.AddItem ("6000" & "-" & "„—ÀÌÂ")
'Combo1.AddItem ("7000" & "-" & "œò·„Â")
Combo1.AddItem ("7000" & "-" & "“‰ê")
Combo1.AddItem ("8000" & "-" & "”«Ì—")
Combo1.AddItem ("9000" & "-" & "¬ﬁ«")

Combo1.AddItem ("1-List")
Combo1.AddItem ("2-List")
Combo1.AddItem ("3-List")
Combo1.AddItem ("4-List")
Combo1.AddItem ("5-List")
Combo1.AddItem ("6-List")
Combo1.AddItem ("7-List")
Combo1.AddItem ("8-List")
Combo1.AddItem ("9-List")

1


End Sub


'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.StatusBar1.Panels(4).Text = Wmp2.settings.volume
'End Sub



Private Sub Form_Resize()
Exit Sub

On Error Resume Next

DataGrid1.Height = MP3F.Height - 4377
DataGrid1.Width = MP3F.Width - 330

End Sub


Private Sub Form_Unload(Cancel As Integer)
entekhab.Show

Unload Me
End Sub

Private Sub Label10_Change()
Exit Sub

If Label10.Caption = "On" Then
mnuon.Checked = True
mnuoff.Checked = False


Else
mnuon.Checked = False
mnuoff.Checked = True


End If



End Sub

Private Sub List1_Click()
Dim kod_gorohlist, kod_sot As String
'kod_sot = MP3D.Recordset.Fields("kod")
gorohlist.Refresh
gorohlist.RecordSource = "select * from gorohlist where xname like ('" & list1.Text & "')"
gorohlist.Refresh
'If gorohlist.Recordset.BOF = False Or gorohlist.Recordset.EOF = False Then

kod_gorohlist = gorohlist.Recordset.Fields("kodgoroh")

List2.Clear
sotgoroh.Refresh
sotgoroh.RecordSource = "select * from sotgoroh where kodgoroh like ('" & kod_gorohlist & "')" ' and kodsot like ('" & kod_sot & "')"
sotgoroh.Refresh
For I = 1 To sotgoroh.Recordset.RecordCount
mp3d2.Refresh
mp3d2.RecordSource = "select * from mp3d where kod like ('" & sotgoroh.Recordset.Fields("kodsot") & "')"
mp3d2.Refresh
If mp3d2.Recordset.BOF = True Or mp3d2.Recordset.EOF = True Then ' in fail qablan pak shode
sotgoroh.Recordset.MoveNext

Else
List2.AddItem (mp3d2.Recordset.Fields("kod") & " :: " & mp3d2.Recordset.Fields("name") & " :: " & mp3d2.Recordset.Fields("tozih"))
sotgoroh.Recordset.MoveNext

End If
Next I
End Sub

Private Sub List2_DblClick()
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where kod like ('" & Mid(List2.Text, 1, 4) & "%')"
MP3D.Refresh
End Sub

Private Sub ListCounter_Click()

MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where kod like ('" & "" & "%')"
MP3D.Refresh




MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where goroh like ('%" & Combo1.Text & "%')"
MP3D.Refresh

MP3D.Recordset.Sort = "kod"

If MP3D.Recordset.BOF = True Or MP3D.Recordset.EOF = True Then
Text2.Text = Combo1.Text + "140800"
Exit Sub
End If




MP3D.Recordset.MovePrevious

MP3D.Recordset.MoveLast

'MsgBox Mid(MP3D.Recordset.Fields("kod"), 7, 6)

Text2.Text = Combo1.Text & (Val(Mid(MP3D.Recordset.Fields("kod"), 7, 6)) + 1)


Beep


End Sub

Private Sub MNOOPEN_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ ’Ê  ÅŒ‘ ‘Êœ", vbQuestion + vbYesNo, "ÅŒ‘ ’Ê ") = vbYes Then
On Error GoTo 1
GoTo 2
1:
MsgBox "›«Ì· ÅÌœ« ‰‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
2
WMP2.URL = MP3D.Recordset.Fields("url")
WMP2.settings.volume = MP3D.Recordset.Fields("vol")


End If

End Sub

Private Sub mnudel_Click()
Call Command3_Click
End Sub

Private Sub mnuoff_Click()
MP3D.Recordset.Fields("onoff") = "Off"
MP3D.Recordset.Fields("systemonoff") = "Off"

MP3D.Recordset.Update

End Sub

Private Sub mnuon_Click()
MP3D.Recordset.Fields("onoff") = "On"
MP3D.Recordset.Fields("systemonoff") = "On"

MP3D.Recordset.Update

End Sub

Private Sub mnusabt_Click()
Call Command1_Click

End Sub

Private Sub Text1_Change()
If Check2.Value = 1 Then
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where url like ('%" + Text1.Text + "%')"
MP3D.Refresh
End If

End Sub

Private Sub Text2_Change()
If Check2.Value = 1 Then
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where id like ('" + Text2.Text + "')"
MP3D.Refresh
End If

End Sub

Private Sub Text3_Change()
If Check2.Value = 1 Then
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where name like ('%" + Text3.Text + "%')"
MP3D.Refresh
End If


End Sub

Private Sub Text4_Change()
If Check2.Value = 1 Then
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where sampel like ('%" + Text4.Text + "%')"
MP3D.Refresh
End If

End Sub

Private Sub Text5_Change()
If Check2.Value = 1 Then
MP3D.Refresh
MP3D.RecordSource = "select * from mp3d where tozih like ('%" + Text5.Text + "%')"
MP3D.Refresh
End If

End Sub

Private Sub WMP2_AudioLanguageChange(ByVal LangID As Long)
StatusBar1.Panels(4).Text = WMP2.settings.volume

End Sub

Private Sub WMP2_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
On Error GoTo 1
GoTo 2
1:
MsgBox "›«Ì· ÅÌœ« ‰‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
2
WMP2.URL = Text1.Text
End Sub


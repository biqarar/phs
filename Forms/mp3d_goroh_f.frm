VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Mp3d_goroh_F 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ã«»Ã«ÌÌ ’Ê  œ— ê—ÊÂ Â«"
   ClientHeight    =   6780
   ClientLeft      =   3360
   ClientTop       =   2685
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mp3d_goroh_f.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4920
      Top             =   5160
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Ã” ÃÊ"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3510
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Frame motor 
      BackColor       =   &H0000C0C0&
      Caption         =   "motor"
      Height          =   255
      Left            =   11280
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   3255
      Begin MSAdodcLib.Adodc N_oqate 
         Height          =   330
         Left            =   120
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_oqate"
         Caption         =   "N_oqate"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_user 
         Height          =   330
         Left            =   120
         Top             =   3600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_user"
         Caption         =   "N_user"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_shahr 
         Height          =   330
         Left            =   120
         Top             =   3240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_shahr"
         Caption         =   "N_shahr"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_setting 
         Height          =   330
         Left            =   120
         Top             =   2880
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_setting"
         Caption         =   "N_setting"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_goroh 
         Height          =   330
         Left            =   120
         Top             =   2520
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_goroh"
         Caption         =   "N_goroh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_mp3dgoroh 
         Height          =   330
         Left            =   120
         Top             =   2160
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_mp3dgoroh"
         Caption         =   "N_mp3dgoroh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_mp3d 
         Height          =   330
         Left            =   120
         Top             =   1800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_mp3d"
         Caption         =   "N_mp3d"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_history 
         Height          =   330
         Left            =   120
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_history"
         Caption         =   "N_history"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_hafte 
         Height          =   330
         Left            =   120
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_hafte"
         Caption         =   "N_hafte"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc N_date 
         Height          =   330
         Left            =   120
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_oqate"
         Caption         =   "N_date"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc10 
         Height          =   330
         Left            =   120
         Top             =   4320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_oqate"
         Caption         =   "000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc11 
         Height          =   330
         Left            =   120
         Top             =   3960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   32896
         ForeColor       =   16777215
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DataPHS.mdb;Mode=Share Deny None;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from N_oqate"
         Caption         =   "00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.ListBox goroh_list 
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4200
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.ListBox mp3_list 
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4200
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label soal 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "¬Ì« „ÿ„∆‰ Â” Ìœø"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7680
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label yes_l 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "»·Ì"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label no_l 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ŒÌ—"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label resetOnOff 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   " ‰ŸÌ„ „Ãœœ  „«„Ì ’Ê  Â«Ì «Ì‰ ê—ÊÂ »— ÅŒ‘ ‰‘œÂ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   5880
      Width           =   8175
   End
   Begin VB.Label pakhsh_nashode_label 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "ÅŒ‘ ‰‘œÂ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label pakhsh_shode_label 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "ÅŒ‘ ‘œÂ"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "›⁄«· / €Ì— ›⁄«·"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Õ–› ’Ê  «“ ê—ÊÂ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "«÷«›Â ò—œ‰ ’Ê  »Â ê—ÊÂ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "»Â —Ê“ —”«‰Ì"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "«÷«›Â ò—œ‰ ’Ê  ÃœÌœ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "Õ–› Ê «÷«›Â ò—œ‰ ê—ÊÂ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      DataField       =   "url"
      DataSource      =   "N_mp3d"
      BeginProperty Font 
         Name            =   "Adobe Arabic"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   14
      Top             =   6360
      Width           =   13140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”  ’Ê  Â«"
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
      Left            =   9840
      TabIndex        =   13
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”  ’Ê  Â«Ì „ÊÃÊœ œ— «Ì‰ ê—ÊÂ"
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
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   3990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”  ê—ÊÂ Â«"
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
      Left            =   5640
      TabIndex        =   11
      Top             =   480
      Width           =   1530
   End
End
Attribute VB_Name = "Mp3d_goroh_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timer_for_insert_mp3 As Integer

Private Sub Form_Load()
refresh_koli

End Sub
Function refresh_koli()

N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d "
N_mp3d.Refresh

List1.Clear

For I = 1 To N_mp3d.Recordset.RecordCount

List1.AddItem (N_mp3d.Recordset.Fields("id_mp3d") & " _ " & N_mp3d.Recordset.Fields("xname") & " _ " & N_mp3d.Recordset.Fields("tozih"))
N_mp3d.Recordset.MoveNext
Next I
List1.Text = List1.List(0)

N_goroh.Refresh
N_goroh.RecordSource = "select * from N_goroh" ' where xname like ('%" & "" & "%')"
N_goroh.Refresh
goroh_list.Clear

For I = 1 To N_goroh.Recordset.RecordCount
goroh_list.AddItem (N_goroh.Recordset.Fields("id_goroh") & " _   " & N_goroh.Recordset.Fields("xname") & "  ::  " & N_goroh.Recordset.Fields("tozih"))
N_goroh.Recordset.MoveNext
Next I
goroh_list.Text = goroh_list.List(0)

End Function



Private Sub goroh_list_Click()


Ref_resh_goroh_lis

End Sub
Function Ref_resh_goroh_lis()
mp3_list.Clear
a = Split(goroh_list.Text)
N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & a(0) & "') and onoff like ('on') "
N_mp3dgoroh.Refresh
pakhsh_nashode_label.Caption = " ÅŒ‘ ‰‘œÂ: " & N_mp3dgoroh.Recordset.RecordCount

N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & a(0) & "') and onoff like ('off') "
N_mp3dgoroh.Refresh
pakhsh_shode_label.Caption = " ÅŒ‘ ‘œÂ: " & N_mp3dgoroh.Recordset.RecordCount


N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & a(0) & "')"
N_mp3dgoroh.Refresh
For I = 1 To N_mp3dgoroh.Recordset.RecordCount
N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where id_mp3d like ('" & N_mp3dgoroh.Recordset.Fields("id_mp3d") & "')"
N_mp3d.Refresh
mp3_list.AddItem (N_mp3dgoroh.Recordset.Fields("id_mp3goroh") & " _ " & N_mp3d.Recordset.Fields("xname") & " :: " & N_mp3d.Recordset.Fields("tozih"))
N_mp3dgoroh.Recordset.MoveNext
Next I
mp3_list.Text = mp3_list.List(0)

End Function

Private Sub Label2_Click()

End Sub

Private Sub Label20_Click()
MP3F_N.Show

End Sub

Private Sub Label21_Click()
Add_goroh_f.Show

End Sub

Private Sub Label5_Click()
refresh_koli

End Sub

Private Sub Label6_Click()
id_m = Split(List1.Text, " _ ")
id_g = Split(goroh_list.Text, " _ ")

N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_mp3d like ('" & id_m(0) & "') and id_goroh like ('" & id_g(0) & "')"
N_mp3dgoroh.Refresh
If N_mp3dgoroh.Recordset.BOF = True Or N_mp3dgoroh.Recordset.EOF = True Then
N_mp3dgoroh.Refresh
N_mp3dgoroh.Recordset.AddNew
N_mp3dgoroh.Recordset.Fields("id_mp3d") = id_m(0)
N_mp3dgoroh.Recordset.Fields("id_goroh") = id_g(0)
N_mp3dgoroh.Recordset.Fields("onoff") = "on"
N_mp3dgoroh.Recordset.Fields("uonoff") = "on"

N_mp3dgoroh.Recordset.Update
N_mp3dgoroh.Refresh

a = Error_Label6("«÷«›Â ‘œ", "Green")
Ref_resh_goroh_lis

Else
a = Error_Label6("’Ê   ò—«—Ì «” ", "Red")
End If





End Sub
Function Mp3_ADD_successfully()

End Function

Private Sub Label7_Click()
On Error Resume Next

id_m = Split(mp3_list.Text, " _ ")

N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_mp3goroh like ('" & id_m(0) & "')" ' and id_goroh like ('" & id_g(0) & "')"
N_mp3dgoroh.Refresh
If N_mp3dgoroh.Recordset.BOF = False Or N_mp3dgoroh.Recordset.EOF = False Then N_mp3dgoroh.Recordset.Delete

Ref_resh_goroh_lis

End Sub

Private Sub Label8_Click()
On Error Resume Next

id_m = Split(mp3_list.Text, " _ ")

N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_mp3goroh like ('" & id_m(0) & "')" ' and id_goroh like ('" & id_g(0) & "')"
N_mp3dgoroh.Refresh
If N_mp3dgoroh.Recordset.BOF = False Or N_mp3dgoroh.Recordset.EOF = False Then
'If N_mp3dgoroh.Recordset.Fields("uonoff") = "on" Then
If Label8.Caption = "›⁄«·" Then
N_mp3dgoroh.Recordset.Fields("uonoff") = "off"
N_mp3dgoroh.Recordset.Update
Label8.BackColor = &H40C0&



Else
N_mp3dgoroh.Recordset.Fields("uonoff") = "on"
N_mp3dgoroh.Recordset.Update
Label8.BackColor = &HC000&

End If
End If
Call mp3_list_Click

End Sub

Private Sub List1_Click()

On Error Resume Next

a = Split(List1.Text, " _ ")
N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where id_mp3d like ('" & a(0) & "')"
N_mp3d.Refresh



End Sub

Private Sub mp3_list_Click()
On Error Resume Next

id_m = Split(mp3_list.Text, " _ ")

N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_mp3goroh like ('" & id_m(0) & "')" ' and id_goroh like ('" & id_g(0) & "')"
N_mp3dgoroh.Refresh
If N_mp3dgoroh.Recordset.BOF = False Or N_mp3dgoroh.Recordset.EOF = False Then
If N_mp3dgoroh.Recordset.Fields("uonoff") = "on" Then
Label8.Caption = "›⁄«·"
Label8.BackColor = &HC000&

Else
Label8.BackColor = &H40C0&


Label8.Caption = "€Ì— ›⁄«·"
End If
End If

End Sub

Private Sub no_l_Click()
soal.Visible = False
yes_l.Visible = False
no_l.Visible = False
resetOnOff.Visible = True

End Sub

Private Sub pakhsh_nashode_label_Click()
On Error Resume Next

mp3_list.Clear

a = Split(goroh_list.Text)
N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & a(0) & "') and onoff like ('on') "
N_mp3dgoroh.Refresh

For I = 1 To N_mp3dgoroh.Recordset.RecordCount
N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where id_mp3d like ('" & N_mp3dgoroh.Recordset.Fields("id_mp3d") & "')"
N_mp3d.Refresh
mp3_list.AddItem (N_mp3dgoroh.Recordset.Fields("id_mp3goroh") & " _ " & N_mp3d.Recordset.Fields("xname") & " :: " & N_mp3d.Recordset.Fields("tozih"))
N_mp3dgoroh.Recordset.MoveNext
Next I
mp3_list.Text = mp3_list.List(0)
End Sub

Private Sub pakhsh_shode_label_Click()
On Error Resume Next

mp3_list.Clear

a = Split(goroh_list.Text)
N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & a(0) & "') and onoff like ('off') "
N_mp3dgoroh.Refresh

For I = 1 To N_mp3dgoroh.Recordset.RecordCount
N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where id_mp3d like ('" & N_mp3dgoroh.Recordset.Fields("id_mp3d") & "')"
N_mp3d.Refresh
mp3_list.AddItem (N_mp3dgoroh.Recordset.Fields("id_mp3goroh") & " _ " & N_mp3d.Recordset.Fields("xname") & " :: " & N_mp3d.Recordset.Fields("tozih"))
N_mp3dgoroh.Recordset.MoveNext
Next I
mp3_list.Text = mp3_list.List(0)
End Sub

Private Sub resetOnOff_Click()
soal.Visible = True
yes_l.Visible = True
no_l.Visible = True
resetOnOff.Visible = False

End Sub

Private Sub Text1_Change()
On Error Resume Next

N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where xname like ('%" & Text1.Text & "%') or tozih like ('%" & Text1.Text & "%') or id_mp3d like ('%" & Text1.Text & "%') or url like ('%" & Text1.Text & "%')"
N_mp3d.Refresh
List1.Clear

For I = 1 To N_mp3d.Recordset.RecordCount

List1.AddItem (N_mp3d.Recordset.Fields("id_mp3d") & " _ " & N_mp3d.Recordset.Fields("xname") & " _ " & N_mp3d.Recordset.Fields("tozih"))
N_mp3d.Recordset.MoveNext
Next I
List1.Text = List1.List(0)
End Sub
Function Error_Label6(str_, color_)
Label6.Caption = str_
Timer_for_insert_mp3 = 20
If color_ = "Red" Then

Label6.BackColor = &H80&
ElseIf color_ = "Green" Then
Label6.BackColor = &HFF0000
End If


Timer1.Enabled = True

End Function
Private Sub Text1_Click()
If Text1.Text = "Ã” ÃÊ" Then Text1.Text = ""

End Sub

Private Sub Timer1_Timer()
Timer_for_insert_mp3 = Timer_for_insert_mp3 - 1
If Timer_for_insert_mp3 = 0 Then
Label6.BackColor = &H8000&

Label6.Caption = "«÷«›Â ò—œ‰ ’Ê  »Â ê—ÊÂ"
Timer1.Enabled = False

End If



End Sub

Private Sub yes_l_Click()
On Error Resume Next


a = Split(goroh_list.Text)
N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & a(0) & "') and onoff like ('off') "
N_mp3dgoroh.Refresh

For I = 1 To N_mp3dgoroh.Recordset.RecordCount
N_mp3dgoroh.Recordset.Fields("onoff") = "on"
N_mp3dgoroh.Recordset.Update
N_mp3dgoroh.Recordset.MoveNext

Next I
Ref_resh_goroh_lis





soal.Visible = False
yes_l.Visible = False
no_l.Visible = False
resetOnOff.Visible = True
End Sub

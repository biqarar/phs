VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Pakhsh_f 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”«„«‰Â ’Ê Ì"
   ClientHeight    =   7920
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pakhsh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox up_text 
      Height          =   855
      Left            =   3720
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox password_text 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "0"
      TabIndex        =   0
      Text            =   "ò·„Â ⁄»Ê—"
      ToolTipText     =   "ò·„Â ⁄»Ê—"
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   1320
      Width           =   6255
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1320
         Top             =   -120
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   720
         Top             =   -120
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   -120
      End
      Begin VB.Label wmp_zamen1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   4080
         TabIndex        =   39
         Top             =   240
         Width           =   660
      End
      Begin VB.Label wmp_zaman2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   2280
         TabIndex        =   38
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   360
         TabIndex        =   36
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "“„«‰ »«ﬁÌ„«‰œÂ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   5040
         TabIndex        =   35
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame motor 
      BackColor       =   &H0000C0C0&
      Caption         =   "motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   -120
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "”«⁄  Ê  «—ÌŒ ›⁄·Ì"
      ForeColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   120
      Width           =   6255
      Begin VB.Label date_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000000"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   21.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   765
         Left            =   3840
         TabIndex        =   26
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Time_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   21.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   765
         Left            =   840
         TabIndex        =   25
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14400
      Top             =   3840
   End
   Begin VB.Frame oqat_fream 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Height          =   1335
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   7695
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ì„Â ‘»"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   1440
         TabIndex        =   40
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«–«‰ ’»Õ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   6360
         TabIndex        =   31
         Top             =   480
         Width           =   765
      End
      Begin VB.Label sobh_l 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "12:25:31"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   5160
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ·Ê⁄ ¬› «»"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   6360
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
      Begin VB.Label toloe_l 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "12:25:31"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   5160
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.Label maqreb_l 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "12:25:31"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«–«‰ „€—»"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   1440
         TabIndex        =   22
         Top             =   480
         Width           =   870
      End
      Begin VB.Label nime_shab_l 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "12:25:31"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ì„Â ‘»"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   1440
         TabIndex        =   20
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label zohr_l 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "12:25:31"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   2760
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.Label qorob_l 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "12:25:31"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   2760
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«–«‰ ŸÂ—"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   3960
         TabIndex        =   17
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "€—Ê» ¬› «»"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   3960
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00404000&
      Caption         =   "·Ì”  «‰ Ÿ«— ÅŒ‘"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   6375
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   7695
      Begin VB.ListBox mehvar_list 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1780
         ItemData        =   "pakhsh.frx":030A
         Left            =   240
         List            =   "pakhsh.frx":030C
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   5775
      End
      Begin VB.ListBox date_list 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1780
         ItemData        =   "pakhsh.frx":030E
         Left            =   240
         List            =   "pakhsh.frx":0310
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   4440
         Width           =   5775
      End
      Begin VB.ListBox hafte_list 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1780
         ItemData        =   "pakhsh.frx":0312
         Left            =   240
         List            =   "pakhsh.frx":0314
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2520
         Width           =   5775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÅŒ‘ »« „ÕÊ—Ì  «Êﬁ«  ‘—⁄Ì"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1065
         Left            =   6120
         TabIndex        =   33
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÅŒ‘ »«  «—ÌŒ Œ«’"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   6120
         TabIndex        =   32
         Top             =   4560
         Width           =   1365
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÅŒ‘ Â«Ì Â› êÌ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   6240
         TabIndex        =   8
         Top             =   2640
         Width           =   1185
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÅŒ‘ ò‰‰œÂ ‘„«—Â 2"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   43
      Top             =   3840
      Width           =   1365
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÅŒ‘ ò‰‰œÂ ‘„«—Â 3"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   42
      Top             =   5520
      Width           =   1410
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÅŒ‘ ò‰‰œÂ ‘„«—Â 1"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   41
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label upd_l 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "»Â —Ê“ —”«‰Ì"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label tanzimat_l 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   " ‰ŸÌ„« "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Long_time 
      AutoSize        =   -1  'True
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   765
      Left            =   11520
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÅŒ‘ „Ê«—œ »«  ò—«— Â› êÌ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   4320
      TabIndex        =   14
      Top             =   3720
      Width           =   1980
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÅŒ‘ „Ê«—œ »«  «—ÌŒ Œ«’"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   4440
      TabIndex        =   13
      Top             =   5400
      Width           =   1920
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ÅŒ‘ «–«‰ ° ﬁ—¬‰ ° „‰«Ã«  "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4200
      TabIndex        =   12
      Top             =   2040
      Width           =   1920
   End
   Begin WMPLibCtl.WindowsMediaPlayer Wmp1 
      Height          =   1560
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   6255
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
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   11033
      _cy             =   2752
   End
   Begin WMPLibCtl.WindowsMediaPlayer Wmp2 
      Height          =   1560
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   6255
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
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   11033
      _cy             =   2752
   End
   Begin WMPLibCtl.WindowsMediaPlayer Wmp3 
      Height          =   1560
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   6255
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
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   11033
      _cy             =   2752
   End
   Begin VB.Menu mnuwe 
      Caption         =   "œ—»«—Â »—‰«„Â"
   End
End
Attribute VB_Name = "Pakhsh_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Zamane_Baqimande, Zamane_Baqimande1, Zamane_Baqimande2 As Integer
Dim WWMPPASLI As String

Dim H_1, M_1, S_1, H_2, M_2, S_2, H_3, M_3, S_3

Private Sub Form_Load()
TarikhShamsi

Set_oqat_labels
ref_resh_list_pahksh
Find_AFter_Befor
End Sub
Function Add_History_MP3()
On Error Resume Next

N_history.Refresh
N_history.Recordset.AddNew
N_history.Recordset.Fields("xtime") = Time_label.Caption
N_history.Recordset.Fields("xdate") = date_label.Caption
N_history.Recordset.Fields("id_mp3d") = N_mp3d.Recordset.Fields("id_mp3d")
N_history.Recordset.Fields("url") = N_mp3d.Recordset.Fields("url")
N_history.Recordset.Update
N_history.Refresh

End Function
Function TarikhShamsi(Optional date1 As String, Optional SmallDate1 As Boolean) As String

      '====================================================
      Dim d, p, w, mon, MM, Ym, u, v, rp, X, I, Ys, Ms, Dm, P1, D1, Ds, DateShamsi
      d = Array(20, 19, 20, 20, 21, 21, 22, 22, 22, 22, 21, 21)
      p = Array(11, 12, 10, 12, 11, 11, 10, 10, 10, 9, 10, 10)
      w = Array("Ìò‘‰»Â", "œÊ‘‰»Â", "”Â ‘‰»Â", "çÂ«—‘‰»Â", "Å‰Ã‘‰»Â", "Ã„⁄Â", "‘‰»Â")
      
      If SmallDate1 = True Then
            mon = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
      Else
            mon = Array("›—Ê—œÌ‰", "«—œÌ»Â‘ ", "Œ—œ«œ", " Ì—", "„—œ«œ", "‘Â—ÌÊ—", "„Â—", "¬»«‰", "¬–—", "œÌ", "»Â„‰", "«”›‰œ")
      End If
      
      If date1 = "" Then date1 = Date
      
      Dm = Day(date1) '»œ”  ¬Ê—œ‰ —Ê“
      MM = Month(date1) '»œ”  ¬Ê—œ‰ „«Â
      Ym = Year(date1) '»œ”  ¬Ê—œ‰ ”«·
      u = 0
      rp = 0
      If (Ym Mod 4) = 0 Then u = 1 ' ‘ŒÌ’ ò»Ì”Â »Êœ‰
      If ((Ym Mod 100) = 0 And (Ym Mod 400) <> 0) Then u = 0 ' ‘ŒÌ’ ò»Ì”Â ‰»Êœ‰
      Ys = Ym - 622 ' »œÌ· ”«· „Ì·«œÌ »Â ‘„”Ì
      X = Ys - 22
      X = X Mod 33
      If ((X Mod 4) = 0 And X <> 32) Then rp = 1
      I = Not (rp - 2) + Not (u - 2) * 2
      X = 0
      If (I = 0 And MM = 3) Then X = 1
      If I = 0 Then I = 3
      Ms = (9 + MM) Mod 13
      If Ms < 10 Then Ms = Ms + 1
      D1 = d(MM - 1)
      If (I = 1 And MM > 2) Then D1 = D1 - 1
      If (I = 2 And MM < 3) Then D1 = D1 - 1
      P1 = p(MM - 1)
      If (I = 1 And MM > 2) Then P1 = P1 + 1
      If (I = 2 And MM < 4) Then P1 = P1 + 1
      If (Dm > 0 And Dm <= D1) Then
             Ds = P1 + Dm + X - 1
          X = 1
      Else
          Ds = Dm - D1
          Ms = Ms + 1
          If Ms = 13 Then Ms = 1
          X = 2
      End If
      If ((MM = 3 And X = 2) Or MM > 3) Then Ys = Ys + 1
      If SmallDate1 = True Then
'     ??? ??? ?? ???? ???? ???????? ???????? ?? ??? ?? ?? ???? ????? ?? ?????
'            TarikhShamsi = Trim(Str(Ys)) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
           ' TarikhShamsi = Mid(Trim(str(Ys)), 3, 2) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(str(Ds))
           ' Tarikh.Caption = str(Ys) & "/" & (Ms) & "/" & str(Ds)
      Else
           ' TarikhShamsi = w(Weekday(Date) - 1) + " " + str(Ds) + " " + mon(Ms - 1) + " " + str(Ys)
           ' Tarikh.Caption = (Ys) & "/" & (Ms) & "/" & (Val(Ds))
            If Val(Ms) < 10 Then Ms = "0" & Ms
            
            If Val(Ds) < 10 Then Ds = "0" & Ds
             date_label.Caption = Ys & Ms & Ds
      End If

End Function

Function Set_oqat_labels()
On Error Resume Next

N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where kod =('" & "what_shahr" & "')"
N_setting.Refresh

N_shahr.Refresh
N_shahr.RecordSource = "select * from n_shahr where id_shahr like ('" & N_setting.Recordset.Fields("xvalue") & "')"

N_shahr.Refresh


oqat_fream.Caption = " «Êﬁ«  ‘—⁄Ì »Â «›ﬁ " & N_shahr.Recordset.Fields("shahr")

N_oqate.Refresh
N_oqate.RecordSource = "select * from n_oqate where ndate like ('" & date_label.Caption & "') and id_shahr like ('" & N_setting.Recordset.Fields("xvalue") & "')"
N_oqate.Refresh
If N_oqate.Recordset.BOF = True Or N_oqate.Recordset.EOF = True Then

sobh_l.Caption = "Œÿ« œ—  «—ÌŒ"
toloe_l.Caption = "Œÿ« œ—  «—ÌŒ"
zohr_l.Caption = "Œÿ« œ—  «—ÌŒ"
qorob_l.Caption = "Œÿ« œ—  «—ÌŒ"
maqreb_l.Caption = "Œÿ« œ—  «—ÌŒ"
nime_shab_l.Caption = "Œÿ« œ—  «—ÌŒ"



Else
sobh_l.Caption = N_oqate.Recordset.Fields("sobh")
toloe_l.Caption = N_oqate.Recordset.Fields("toloe")
zohr_l.Caption = N_oqate.Recordset.Fields("zohr")
qorob_l.Caption = N_oqate.Recordset.Fields("qorob")
maqreb_l.Caption = N_oqate.Recordset.Fields("maqreb")
nime_shab_l.Caption = N_oqate.Recordset.Fields("nimeshab")
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
If password_text.Text = "10012513" Then


Cancel = 0
End

Else
'MsgBox "ò·„Â ⁄»Ê— «‘ »«Â «” ", vbCritical + vbOKOnly, "Œÿ«"

Cancel = 1
End If

End Sub
Function RND_for_mp3(id_goroh_)
'Debug.Assert ""
On Error Resume Next

N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & id_goroh_ & "') and onoff like ('" & "on" & "') and uonoff like ('" & "on" & "')"
N_mp3dgoroh.Refresh
If N_mp3dgoroh.Recordset.RecordCount = 0 Then
N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & id_goroh_ & "') and uonoff like ('" & "on" & "')"
N_mp3dgoroh.Refresh
For I = 1 To N_mp3dgoroh.Recordset.RecordCount
N_mp3dgoroh.Recordset.Fields("onoff") = "on"
N_mp3dgoroh.Recordset.Update
N_mp3dgoroh.Recordset.MoveNext
Next I

End If
N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_goroh like ('" & id_goroh_ & "') and onoff like ('" & "on" & "') and uonoff like ('" & "on" & "')"
N_mp3dgoroh.Refresh

X = Int((Rnd() * Val(N_mp3dgoroh.Recordset.RecordCount)))
For I = 1 To X
N_mp3dgoroh.Recordset.MoveNext
Next I
N_mp3dgoroh.Recordset.Fields("onoff") = "off"
N_mp3dgoroh.Recordset.Update





End Function
Function ref_resh_list_pahksh()
On Error Resume Next
'mehvar_list.Clear
hafte_list.Clear
date_list.Clear

N_date.Refresh
N_date.RecordSource = "select * from n_date where xdate like ('" & date_label.Caption & "')"
N_date.Refresh
For I = 1 To N_date.Recordset.RecordCount
date_list.AddItem (N_date.Recordset.Fields("xtime") & " _ " & N_date.Recordset.Fields("id_mp3d") & " _ " & N_date.Recordset.Fields("xName") & "  ::  " & N_date.Recordset.Fields("zaman") & "  *  " & N_date.Recordset.Fields("player"))
N_date.Recordset.MoveNext
Next I

N_hafte.Refresh
N_hafte.RecordSource = "select * from n_hafte where week like ('%" & Weekday(Date) & "%')"
N_hafte.Refresh
For I = 1 To N_hafte.Recordset.RecordCount
'A = RND_for_mp3(N_hafte.Recordset.Fields("id_goroh"))

hafte_list.AddItem (N_hafte.Recordset.Fields("xtime") & " _ " & N_hafte.Recordset.Fields("id_goroh") & " _ " & N_hafte.Recordset.Fields("xName") & "  ::  " & N_hafte.Recordset.Fields("zaman") & "  *  " & N_hafte.Recordset.Fields("player"))
N_hafte.Recordset.MoveNext
Next I

End Function

Function Pakhsh_Oqate_sharee(Azan_)
On Error Resume Next



N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where kod like ('" & Azan_ & "')"
N_setting.Refresh
If N_setting.Recordset.BOF = True Or N_setting.Recordset.EOF = True Then Exit Function
yr = N_setting.Recordset.Fields("player")

a = RND_for_mp3(N_setting.Recordset.Fields("xname")) 'kod goroh
'N_mp3dgoroh.Recordset.Fields("onoff") = "off"
'N_mp3dgoroh.Recordset.Update

N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where id_mp3d like ('" & N_mp3dgoroh.Recordset.Fields("id_mp3d") & "')"
N_mp3d.Refresh

If yr = "1" Then
Wmp1.URL = N_mp3d.Recordset.Fields("url")
Wmp1.settings.volume = N_mp3d.Recordset.Fields("volome")
End If

If yr = "2" Then
Wmp2.URL = N_mp3d.Recordset.Fields("url")
Wmp2.settings.volume = N_mp3d.Recordset.Fields("volome")
End If

If yr = "3" Then
Wmp3.URL = N_mp3d.Recordset.Fields("url")
Wmp3.settings.volume = N_mp3d.Recordset.Fields("volome")
End If



Add_History_MP3





End Function
Function Find_AFter_Befor()
On Error Resume Next

'Dim a(0 To 3) As String
mehvar_list.Clear

N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where xvalue like ('%" & "::" & "%') and xvalue like ('%" & Weekday(Date) & "%')"
N_setting.Refresh

For I = 1 To N_setting.Recordset.RecordCount
a = Split(N_setting.Recordset.Fields("xvalue"), " :: ")
T = After_Befor(a(0), N_setting.Recordset.Fields("kod"), a(1), N_setting.Recordset.Fields("xname"), a(2), N_setting.Recordset.Fields("xtime"))

N_setting.Recordset.MoveNext



Next I

End Function
Function After_Befor(Af_Bf, A_B_what, minuts_, goroh_, zaman, x_name)
On Error Resume Next
H_1 = 0
H_2 = 0
H_3 = 0
M_1 = 0
M_2 = 0
M_3 = 0
S_1 = 0
S_2 = 0
S_3 = 0

If A_B_what = "sobh" Then time_asli_ = sobh_l.Caption
If A_B_what = "toloe" Then time_asli_ = toloe_l.Caption
If A_B_what = "zohr" Then time_asli_ = zohr_l.Caption
If A_B_what = "qorob" Then time_asli_ = qorob_l.Caption
If A_B_what = "maqreb" Then time_asli_ = maqreb_l.Caption
If A_B_what = "nimeshab" Then time_asli_ = nime_shab_l.Caption

a = Split(time_asli_, ":")
H_1 = a(0)
M_1 = a(1)
S_1 = a(2)
a = Split(minuts_, ":")
H_2 = a(0)
M_2 = a(1)
S_2 = a(2)

If Af_Bf = "B" Then
'kam konad
If Val(S_1) < Val(S_2) Then
M_1 = M_1 - 1
S_1 = S_1 + 59
End If

S_3 = S_1 - S_2

If Val(M_1) < Val(M_2) Then
H_1 = H_1 - 1
M_1 = M_1 + 59
End If

M_3 = M_1 - M_2

H_3 = H_1 - H_2

If H_3 < 10 Then H_3 = "0" & H_3

If M_3 < 10 Then M_3 = "0" & M_3

If S_3 < 10 Then S_3 = "0" & S_3

'MsgBox H_3 & ":" & m_3 & ":" & s_3
'End
End If
If Af_Bf = "A" Then
'ezafe konad


S_3 = Val(S_1) + Val(S_2)

If S_3 > 59 Then
M_1 = M_1 + 1
S_3 = S_3 - 59
End If

M_3 = Val(M_1) + Val(M_2)

If M_3 > 59 Then
H_1 = H_1 - 1
M_3 = M_3 - 59
End If

H_3 = Val(H_1) + Val(H_2)

If H_3 < 10 Then H_3 = "0" & H_3

If M_3 < 10 Then M_3 = "0" & M_3

If S_3 < 10 Then S_3 = "0" & S_3

'MsgBox H_3 & ":" & M_3 & ":" & S_3
'End
End If
mehvar_list.AddItem (H_3 & ":" & M_3 & ":" & S_3 & " _ " & goroh_ & " _ " & x_name & "  ::  " & zaman & "  *  " & N_setting.Recordset.Fields("player"))


End Function



Private Sub mnuwe_Click()
WE.Show

End Sub

Private Sub password_text_Change()
If password_text.Text = "140825132520" Then End
If password_text.Text = "10012513" Then
Wmp1.Enabled = True
Wmp2.Enabled = True
Wmp3.Enabled = True
upd_l.Visible = True
tanzimat_l.Visible = True


Else
Wmp1.Enabled = False
Wmp2.Enabled = False
Wmp3.Enabled = False
upd_l.Visible = False
tanzimat_l.Visible = False

End If

End Sub

Private Sub password_text_Click()
Me.password_text.Text = ""

End Sub

Private Sub tanzimat_l_Click()

List_pakhsh_F.Show

End Sub

Private Sub Time_label_Change()
On Error Resume Next
Dim Now_TIME As String
time__ = Split(Time_label.Caption, ":")
If time__(2) = "00" Then
password_text.Text = ""
up_text.Text = ""
up_text.Text = "10012513"
End If

Now_TIME = Me.Time_label.Caption
If Now_TIME = sobh_l.Caption Then a = Pakhsh_Oqate_sharee("azan_sobh")
If Now_TIME = zohr_l.Caption Then a = Pakhsh_Oqate_sharee("azan_zohr")
If Now_TIME = maqreb_l.Caption Then a = Pakhsh_Oqate_sharee("azan_maqreb")
If Now_TIME = qorob_l.Caption Then a = Pakhsh_Oqate_sharee("qorob")
If Now_TIME = toloe_l.Caption Then a = Pakhsh_Oqate_sharee("toloe")
If Now_TIME = nime_shab_l.Caption Then a = Pakhsh_Oqate_sharee("nimeshab")

For I = 1 To mehvar_list.ListCount
    a = Split(mehvar_list.List(I - 1), " _ ")
    If Time_label.Caption = a(0) Then
    ' pakhsh
    b = Split(mehvar_list.List(I - 1), "  *  ")
     If b(1) = 1 Then paksh = Play_Mehvar_list(a(1), Wmp1)
     If b(1) = 2 Then paksh = Play_Mehvar_list(a(1), Wmp2)
     If b(1) = 3 Then paksh = Play_Mehvar_list(a(1), Wmp3)
    a = Split(mehvar_list.List(I - 1), " :: ")
     T = Split(a(1), "  *  ")
    z = Enabeld_timer2_for_zaman(T(0), b(1))

    End If
Next I

For I = 1 To hafte_list.ListCount
    a = Split(hafte_list.List(I - 1), " _ ")
    If Time_label.Caption = a(0) Then
    ' pakhsh
    b = Split(hafte_list.List(I - 1), "  *  ")
     If b(1) = 1 Then paksh = Play_Mehvar_list(a(1), Wmp1)
     If b(1) = 2 Then paksh = Play_Mehvar_list(a(1), Wmp2)
     If b(1) = 3 Then paksh = Play_Mehvar_list(a(1), Wmp3)
     a = Split(hafte_list.List(I - 1), " :: ")
     T = Split(a(1), "  *  ")
    z = Enabeld_timer2_for_zaman(T(0), b(1))
    End If
Next I

For I = 1 To date_list.ListCount
    a = Split(date_list.List(I - 1), " _ ")
    If Time_label.Caption = a(0) Then
    ' pakhsh
    
    b = Split(date_list.List(I - 1), "  *  ")
    If b(1) = 1 Then paksh = Play_date_list(a(1), Wmp1)
    If b(1) = 2 Then paksh = Play_date_list(a(1), Wmp2)
    If b(1) = 3 Then paksh = Play_date_list(a(1), Wmp3)
    a = Split(date_list.List(I - 1), " :: ")
 T = Split(a(1), "  *  ")
    z = Enabeld_timer2_for_zaman(T(0), b(1))
    
    End If
Next I
End Sub
Function Play_date_list(id_mp3d_, WMPP)
On Error Resume Next


N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where id_mp3d like ('" & id_mp3d_ & "')"
N_mp3d.Refresh
WMPP.URL = N_mp3d.Recordset.Fields("url")
WMPP.settings.volume = N_mp3d.Recordset.Fields("volome")
Add_History_MP3
End Function
Function Play_Mehvar_list(id_goroh_, WWMP)
On Error Resume Next
'Dim WWMP1 As Object

'If WWMP = 1 Then WWMP1 = Wmp1
'If WWMP = 2 Then WWMP1 = Wmp2
'If WWMP = 3 Then WWMP1 = Wmp3
'If WWMP = "" Then WWMP1 = Wmp1

a = RND_for_mp3(id_goroh_) 'kod goroh
If N_mp3dgoroh.Recordset.BOF = True Or N_mp3dgoroh.Recordset.EOF = True Then Exit Function
'N_mp3dgoroh.Recordset.Fields("onoff") = "off"
'N_mp3dgoroh.Recordset.Update

N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where id_mp3d like ('" & N_mp3dgoroh.Recordset.Fields("id_mp3d") & "')"
N_mp3d.Refresh

WWMP.URL = N_mp3d.Recordset.Fields("url")
WWMP.settings.volume = N_mp3d.Recordset.Fields("volome")

Add_History_MP3



End Function
Private Sub Timer1_Timer()
Dim HH, MM, SS As String
HH = Hour(Now)
MM = Minute(Now)
SS = Second(Now)
If Val(HH) < 10 Then HH = "0" & HH
If Val(MM) < 10 Then MM = "0" & MM
If Val(SS) < 10 Then SS = "0" & SS


Time_label.Caption = HH & ":" & MM & ":" & SS
'Long_time.Caption = HH & ":" & MM

End Sub
Function Enabeld_timer2_for_zaman(zamane_, wwmpp)

On Error Resume Next

WWMPPASLI = wwmpp

'If WWMPPASLI = "wmp1" Then
If WWMPPASLI = 1 Then
Zamane_Baqimande = zamane_
Timer2.Enabled = True

End If

'If WWMPPASLI = "wmp2" Then
If WWMPPASLI = 2 Then

Zamane_Baqimande2 = zamane_
Timer4.Enabled = True
End If

'If WWMPPASLI = "wmp3" Then
If WWMPPASLI = 3 Then

Zamane_Baqimande1 = zamane_
Timer3.Enabled = True
End If

End Function
Private Sub Timer2_Timer()

Zamane_Baqimande = Zamane_Baqimande - 1
Label8.Caption = Zamane_Baqimande

If Zamane_Baqimande <= 0 Then
'WWMPT.Enabled = False
Zamane_Baqimande = 0
 Wmp1.Close

Timer2.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()
Zamane_Baqimande1 = Zamane_Baqimande1 - 1
wmp_zamen1.Caption = Zamane_Baqimande1

If Zamane_Baqimande1 <= 0 Then
'WWMPT.Enabled = False
Zamane_Baqimande1 = 0
 Wmp3.Close

Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Zamane_Baqimande2 = Zamane_Baqimande2 - 1
wmp_zaman2.Caption = Zamane_Baqimande2

If Zamane_Baqimande2 <= 0 Then
'WWMPT.Enabled = False
Zamane_Baqimande2 = 0
 Wmp2.Close

Timer4.Enabled = False
End If
End Sub

Private Sub up_text_Change()
If up_text.Text = "10012513" Then
TarikhShamsi

Set_oqat_labels
ref_resh_list_pahksh
Find_AFter_Befor
End If

End Sub

Private Sub upd_l_Click()
TarikhShamsi

Set_oqat_labels
ref_resh_list_pahksh
Find_AFter_Befor
End Sub

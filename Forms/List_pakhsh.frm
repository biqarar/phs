VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form List_pakhsh_F 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„œÌ—Ì  ·Ì”  «‰ Ÿ«— ÅŒ‘"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15300
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "List_pakhsh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   -840
      TabIndex        =   37
      Top             =   0
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
         RecordSource    =   "select * from N_date"
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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3480
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Caption         =   "»Œ‘ Â«Ì œÌê—"
      ForeColor       =   &H0000FFFF&
      Height          =   6255
      Left            =   12720
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   120
      Width           =   2415
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "»«“ ò—œ‰ ÅÊ‘Â"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "Œ—ÊÃ «“ »—‰«„Â"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "»«“ê‘  »Â ’›ÕÂ ÅŒ‘"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "«÷«›Â ò—œ‰ ‘Â— ÃœÌœ"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "œ— »«—Â »—‰«„Â"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "«‰ Œ«» ê—ÊÂ ’Ê  »—«Ì «–«‰"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "„‘«ÂœÂ „Ê«—œ ÅŒ‘ ‘œÂ"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   " €ÌÌ— ò·„Â ⁄»Ê—"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "«‰ Œ«» ‘Â—"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "Ã«»Ã«ÌÌ ’Ê  œ— ê—ÊÂ"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "Õ–› Ê «÷«›Â ò—œ‰ ê—ÊÂ"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "·Ì”  ’Ê  Â«Ì „ÊÃÊœ"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Height          =   2130
      ItemData        =   "List_pakhsh.frx":030A
      Left            =   120
      List            =   "List_pakhsh.frx":030C
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3720
      Width           =   6135
   End
   Begin VB.ListBox mehvar_list 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Height          =   2130
      ItemData        =   "List_pakhsh.frx":030E
      Left            =   6360
      List            =   "List_pakhsh.frx":0310
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "À»  “„«‰ ÅŒ‘"
      ForeColor       =   &H0000FFFF&
      Height          =   3495
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox Combo11 
         Height          =   465
         Left            =   3720
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Text            =   "(»Â „œ  ( À«‰ÌÂ "
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C000&
         Caption         =   "À» "
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2880
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   2760
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Text            =   "⁄‰Ê«‰"
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H80000018&
         Height          =   465
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H80000018&
         Height          =   465
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H80000018&
         Height          =   465
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H80000018&
         Height          =   465
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "»«"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5760
         TabIndex        =   63
         Top             =   3000
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "”«⁄ "
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   2160
         TabIndex        =   50
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   " «—ÌŒ ÅŒ‘"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   4200
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "⁄‰Ê«‰"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5400
         TabIndex        =   48
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "»« ê—ÊÂ ’Ê Ì"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "ÂÌç ê—ÊÂÌ Ê«—œ ‰‘œÂ «” "
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÂÌç —Ê“Ì  ‰ŸÌ„ ‰‘œÂ «” "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "«‰ Œ«» «Ì«„ Â› Â"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "«‰ Œ«» ’Ê "
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   2400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "ÂÌç ’Ê Ì «‰ Œ«» ‰‘œÂ «” "
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   2400
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "À»  ·Ì”  ÅŒ‘ »« „ÕÊ—Ì  «Êﬁ«  ‘—⁄Ì"
      ForeColor       =   &H0000FFFF&
      Height          =   3495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox Combo10 
         Height          =   465
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2760
         Width           =   1935
      End
      Begin VB.OptionButton AaA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404000&
         Caption         =   "»⁄œ"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton BbB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404000&
         Caption         =   "ﬁ»·"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5280
         TabIndex        =   15
         Top             =   960
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Text            =   "(»Â „œ  ( À«‰ÌÂ "
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H80000018&
         Height          =   465
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo7 
         BackColor       =   &H80000018&
         Height          =   465
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo8 
         BackColor       =   &H80000018&
         Height          =   465
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo9 
         Height          =   465
         Left            =   2280
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2880
         TabIndex        =   11
         Text            =   "⁄‰Ê«‰"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C000&
         Caption         =   "À» "
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "»«"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5640
         TabIndex        =   61
         Top             =   2880
         Width           =   105
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÂÌç —Ê“Ì  ‰ŸÌ„ ‰‘œÂ «” "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "œ— —Ê“ Â«Ì"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "»« ê—ÊÂ ’Ê Ì"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "œﬁÌﬁÂ"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   360
         TabIndex        =   41
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "ÅŒ‘"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5520
         TabIndex        =   40
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "ÂÌç ê—ÊÂÌ Ê«—œ ‰‘œÂ «” "
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2280
         Width           =   3255
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -120
      Top             =   3000
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ŒÌ—"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "»·Ì"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1200
      TabIndex        =   35
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "¬Ì« „ÿ„∆‰ Â” Ìœø"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2400
      TabIndex        =   59
      Top             =   6000
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "¬Ì« „ÿ„∆‰ Â” Ìœø"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   8640
      TabIndex        =   58
      Top             =   6000
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "»·Ì"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7440
      TabIndex        =   33
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ŒÌ—"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6360
      TabIndex        =   34
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Õ–›"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "«’·«Õ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Õ–›"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "«’·«Õ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "List_pakhsh_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timer_for_change_color_command_6, Timer_for_change_color_command_1 As Integer

Private Sub Combo1_Click()
a = Split(Combo1.Text, " _ ")
N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where id_setting like ('" & a(0) & "')"
N_setting.Refresh

If N_setting.Recordset.Fields("xname") = "N_hafte" Then
Label13.Visible = True
Label6.Visible = True
Label4.Visible = True
Label3.Visible = True

Label2.Visible = False
Label14.Visible = False
Label15.Visible = False
Combo5.Visible = False

Combo11.Text = Combo11.List(1)


ElseIf N_setting.Recordset.Fields("xname") = "N_date" Then
Label13.Visible = False
Label6.Visible = False
Label4.Visible = False
Label3.Visible = False

Label2.Visible = True
Label14.Visible = True
Label15.Visible = True
Combo5.Visible = True

Combo11.Text = Combo11.List(2)


End If



End Sub

Private Sub Command1_Click()

If Text1.Text = "⁄‰Ê«‰" Or Text1.Text = "" Then
a = Error_command1("⁄‰Ê«‰ —« Ê«—œ ò‰Ìœ")
Text1.SetFocus

Exit Sub
End If
If Len(Combo2.Text) < 2 Or Len(Combo3.Text) < 2 Or Len(Combo4.Text) < 2 Then
a = Error_command1("”«⁄  —« »Â ’Ê—  ’ÕÌÕ Ê«—œ ò‰Ìœ")
Combo2.SetFocus

Exit Sub
End If
If Text4.Text = "" Or Text4.Text = "(»Â „œ  ( À«‰ÌÂ " Then
a = Error_command1("“„«‰ ÅŒ‘ —« Ê«—œ ò‰Ìœ")
Text4.SetFocus

Exit Sub
End If
If Label6.Visible = True Then

If Label6.Caption = "ÂÌç —Ê“Ì  ‰ŸÌ„ ‰‘œÂ «” " Or Label6.Caption = "" Or Label6.Caption = " , " Then
a = Error_command1("«Ì«„ Â› Â —«  ‰ŸÌ„ òÌ‰œ")
Exit Sub
End If
ElseIf Combo5.Visible = True Then
If Combo5.Text = "" Or Len(Combo5.Text) <> 8 Then

a = Error_command1(" «—ÌŒ —«  ‰ŸÌ„ ò‰Ìœ")
Exit Sub
End If
End If

If Label6.Visible = True Then

If Label4.Caption = "ÂÌç ê—ÊÂÌ Ê«—œ ‰‘œÂ «” " Then
a = Error_command1("ê—ÊÂ ’Ê Ì —«  ‰ŸÌ„ ò‰Ìœ")
Exit Sub
End If
ElseIf Combo5.Visible = True Then
If Label15.Caption = "ÂÌç ’Ê Ì «‰ Œ«» ‰‘œÂ «” " Or Label15.Caption = "" Then

a = Error_command1("’Ê  —« «‰ Œ«» ò‰Ìœ")
Exit Sub
End If
End If

If Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  1" Then
wwww = "1"
ElseIf Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  2" Then
wwww = "2"
ElseIf Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  3" Then
wwww = "3"
End If

saat__ = (Combo2.Text & ":" & Combo3.Text & ":" & Combo4.Text)

If Label6.Visible = True Then 'haftegi
''''''''''''''

rooz = Split(Label6.Caption, " , ")
id_rooz = ""

For I = LBound(rooz) To UBound(rooz)
If rooz(I) = "‘‰»Â" Then id_rooz = id_rooz & 7 & "."
If rooz(I) = "Ìò ‘‰»Â" Then id_rooz = id_rooz & 1 & "."
If rooz(I) = "œÊ ‘‰»Â" Then id_rooz = id_rooz & 2 & "."
If rooz(I) = "”Â ‘‰»Â" Then id_rooz = id_rooz & 3 & "."
If rooz(I) = "çÂ«— ‘‰»Â" Then id_rooz = id_rooz & 4 & "."
If rooz(I) = "Å‰Ã ‘‰»Â" Then id_rooz = id_rooz & 5 & "."
If rooz(I) = "Ã„⁄Â" Then id_rooz = id_rooz & 6 & "."
Next I
a = Split(Label4.Caption, " _ ")
N_hafte.Refresh
N_hafte.RecordSource = "SELECT * from n_hafte where week like ('" & id_rooz & "') and  xtime like ('" & saat__ & "') and id_goroh like ('" & a(0) & "') and player like ('" & wwww & "')"
N_hafte.Refresh

If N_hafte.Recordset.BOF = False Or N_hafte.Recordset.EOF = False Then

a = Error_command1("«Ì‰ „Ê—œ ﬁ»·« «÷«›Â ‘œÂ «” ")
Exit Sub
End If

a = Split(Label4.Caption, " _ ")

N_hafte.Refresh
N_hafte.Recordset.AddNew
N_hafte.Recordset.Fields("xname") = Text1.Text
N_hafte.Recordset.Fields("xtime") = Combo2.Text & ":" & Combo3.Text & ":" & Combo4.Text
N_hafte.Recordset.Fields("week") = id_rooz
N_hafte.Recordset.Fields("id_goroh") = a(0)
N_hafte.Recordset.Fields("zaman") = Text4.Text

If Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  1" Then
N_hafte.Recordset.Fields("player") = "1"
ElseIf Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  2" Then
N_hafte.Recordset.Fields("player") = "2"
ElseIf Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  3" Then
N_hafte.Recordset.Fields("player") = "3"
End If

N_hafte.Recordset.Update
N_hafte.Refresh
Timer_for_change_color_command_1 = 20
Timer2.Enabled = True
Command1.BackColor = &HFF00&
Command1.Caption = "⁄„·Ì«  À»  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
Refresh_date_hafte_lis

'end sabt
''''''''''''''''''
ElseIf Combo5.Visible = True Then 'date khas
'''''''''''
a = Split(Label15.Caption, " _ ")


N_date.Refresh
N_date.RecordSource = "select * from n_date where xdate like ('" & Combo5.Text & "') and  xtime like ('" & saat__ & "') and id_mp3d like ('" & a(0) & "') and player like ('" & wwww & "')"
N_date.Refresh
If N_date.Recordset.BOF = False Or N_date.Recordset.EOF = False Then
a = Error_command1("«Ì‰ „Ê—œ ﬁ»·« «÷«›Â ‘œÂ «” ")
Exit Sub
End If


N_date.Refresh
N_date.Recordset.AddNew
N_date.Recordset.Fields("xname") = Text1.Text
N_date.Recordset.Fields("xtime") = Combo2.Text & ":" & Combo3.Text & ":" & Combo4.Text
N_date.Recordset.Fields("xdate") = Combo5.Text
N_date.Recordset.Fields("id_mp3d") = a(0)
N_date.Recordset.Fields("zaman") = Text4.Text

If Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  1" Then
N_date.Recordset.Fields("player") = "1"
ElseIf Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  2" Then
N_date.Recordset.Fields("player") = "2"
ElseIf Combo11.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  3" Then
N_date.Recordset.Fields("player") = "3"
End If

N_date.Recordset.Update
N_date.Refresh
Timer_for_change_color_command_1 = 20
Timer2.Enabled = True
Command1.BackColor = &HFF00&
Command1.Caption = "⁄„·Ì«  À»  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
Refresh_date_hafte_lis
'end sabt
''''''''''''
End If
End Sub

Private Sub Command6_Click()
' ErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrors
If Text2.Text = "⁄‰Ê«‰" Or Text2.Text = "" Then
a = Error_command6("⁄‰Ê«‰ —« Ê«—œ ò‰Ìœ")
Text2.SetFocus

Exit Sub
End If
If Len(Combo6.Text) < 2 Or Len(Combo7.Text) < 2 Or Len(Combo8.Text) < 2 Then
a = Error_command6("”«⁄  —« »Â ’Ê—  ’ÕÌÕ Ê«—œ ò‰Ìœ")
Combo7.SetFocus

Exit Sub
End If
If Combo9.Text = "" Then
a = Error_command6("„ÕÊ— ÅŒ‘ —« „‘Œ’ ò‰Ìœ")
Combo8.SetFocus

Exit Sub
End If

If Text3.Text = "" Or Text3.Text = "(»Â „œ  ( À«‰ÌÂ " Then
a = Error_command6("“„«‰ ÅŒ‘ —« Ê«—œ ò‰Ìœ")
Text3.SetFocus

Exit Sub
End If

If Label12.Caption = "ÂÌç —Ê“Ì  ‰ŸÌ„ ‰‘œÂ «” " Or Label12.Caption = "" Or Label12.Caption = " , " Then
a = Error_command6("«Ì«„ Â› Â —«  ‰ŸÌ„ òÌ‰œ")
Exit Sub
End If
If Label9.Caption = "ÂÌç ê—ÊÂÌ Ê«—œ ‰‘œÂ «” " Then
a = Error_command6("ê—ÊÂ ’Ê Ì —«  ‰ŸÌ„ ò‰Ìœ")
Exit Sub
End If



If Combo10.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  1" Then
wwww = 1
ElseIf Combo10.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  2" Then
wwww = 2
ElseIf Combo10.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  3" Then
wwww = 3
End If
' ErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrorsErrors
'sabt
time__ = Combo8.Text & ":" & Combo7.Text & ":" & Combo6.Text

'Dim rooz(0 To 6) As String


rooz = Split(Label12.Caption, " , ")
id_rooz = ""

For I = LBound(rooz) To UBound(rooz)
If rooz(I) = "‘‰»Â" Then id_rooz = id_rooz & 7 & "."
If rooz(I) = "Ìò ‘‰»Â" Then id_rooz = id_rooz & 1 & "."
If rooz(I) = "œÊ ‘‰»Â" Then id_rooz = id_rooz & 2 & "."
If rooz(I) = "”Â ‘‰»Â" Then id_rooz = id_rooz & 3 & "."
If rooz(I) = "çÂ«— ‘‰»Â" Then id_rooz = id_rooz & 4 & "."
If rooz(I) = "Å‰Ã ‘‰»Â" Then id_rooz = id_rooz & 5 & "."
If rooz(I) = "Ã„⁄Â" Then id_rooz = id_rooz & 6 & "."
Next I

id_goroh_ = Split(Label9.Caption, " _ ")
If Combo9.Text = "«–«‰ ’»Õ" Then mehvar__ = "sobh"
If Combo9.Text = "«–«‰ ŸÂ—" Then mehvar__ = "zohr"
If Combo9.Text = "«–«‰ „€—»" Then mehvar__ = "maqreb"
If Combo9.Text = "€—Ê» ¬› «»" Then mehvar__ = "qorob"
If Combo9.Text = "ÿ·Ê⁄ ¬› «»" Then mehvar__ = "toloe"
If Combo9.Text = "‰Ì„Â ‘»" Then mehvar__ = "nimeshab"

If AaA.Value = True Then ab = "A"
If BbB.Value = True Then ab = "B"
yyyyyy = ab & " :: " & time__ & " :: " & Text3.Text & " :: " & id_rooz

N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where xvalue like ('" & yyyyyy & "') and kod like ('" & mehvar__ & "') and player like ('" & wwww & "')"

N_setting.Refresh
'Debug.Assert ""

If N_setting.Recordset.BOF = False Or N_setting.Recordset.EOF = False Then
a = Error_command6("«Ì‰ „Ê—œ ﬁ»·« «÷«›Â ‘œÂ «” ")
Exit Sub
End If
N_setting.Refresh
N_setting.Recordset.AddNew
N_setting.Recordset.Fields("xname") = id_goroh_(0)
N_setting.Recordset.Fields("kod") = mehvar__
N_setting.Recordset.Fields("xtime") = Text2.Text
If Combo10.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  1" Then
N_setting.Recordset.Fields("player") = "1"
ElseIf Combo10.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  2" Then
N_setting.Recordset.Fields("player") = "2"
ElseIf Combo10.Text = "ÅŒ‘ ò‰‰œÂ ‘„«—Â  3" Then
N_setting.Recordset.Fields("player") = "3"
End If

N_setting.Recordset.Fields("xvalue") = ab & " :: " & time__ & " :: " & Text3.Text & " :: " & id_rooz

N_setting.Recordset.Update
N_setting.Refresh
Timer_for_change_color_command_6 = 20
Timer1.Enabled = True
Command6.BackColor = &HFF00&
Command6.Caption = "⁄„·Ì«  À»  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
Ref_resh_list_mehver_list1

'end sabt

End Sub
Function Error_command6(string_)
Timer_for_change_color_command_6 = 20
Timer1.Enabled = True
Command6.BackColor = &H80&
Command6.Caption = string_

End Function
Function Ref_resh_list_mehver_list1()
Dim TEXT_FOR_INSERT_LIST_1 As String
List1.Clear
N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where xvalue like ('%" & "::" & "%')"
N_setting.Refresh
For I = 1 To N_setting.Recordset.RecordCount
a = Split(N_setting.Recordset.Fields("xvalue"), " :: ")
TEXT_FOR_INSERT_LIST_1 = ""
TEXT_FOR_INSERT_LIST_1 = N_setting.Recordset.Fields("id_setting") & " _ "
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & "  ê—ÊÂ  " & N_setting.Recordset.Fields("xname") & " _ "
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & "   " & N_setting.Recordset.Fields("xtime") & "  "
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & a(1) & " œﬁÌﬁÂ "
If a(0) = "A" Then
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " »⁄œ «“ "
ElseIf a(0) = "B" Then
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " ﬁ»· «“ "
End If
Select Case N_setting.Recordset.Fields("kod")
Case "sobh"
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " «–«‰ ’»Õ "
Case "zohr"
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " «–«‰ ŸÂ— "
Case "maqreb"
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " «–«‰ „€—» "
Case "qorob"
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " €—Ê» ¬› «» "
Case "toloe"
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " ÿ·Ê⁄ ¬› «» "
Case "nimeshab"
TEXT_FOR_INSERT_LIST_1 = TEXT_FOR_INSERT_LIST_1 & " ‰Ì„Â ‘» "

End Select
List1.AddItem (TEXT_FOR_INSERT_LIST_1)
N_setting.Recordset.MoveNext
Next I

End Function
Function Error_command1(string_1)
Timer_for_change_color_command_1 = 20
Timer2.Enabled = True
Command1.BackColor = &H80&
Command1.Caption = string_1

End Function
'
'
'
Private Sub Form_Load()
Combo10.AddItem ("ÅŒ‘ ò‰‰œÂ ‘„«—Â  1")
Combo10.AddItem ("ÅŒ‘ ò‰‰œÂ ‘„«—Â  2")
Combo10.AddItem ("ÅŒ‘ ò‰‰œÂ ‘„«—Â  3")
Combo11.AddItem ("ÅŒ‘ ò‰‰œÂ ‘„«—Â  1")
Combo11.AddItem ("ÅŒ‘ ò‰‰œÂ ‘„«—Â  2")
Combo11.AddItem ("ÅŒ‘ ò‰‰œÂ ‘„«—Â  3")
Combo10.Text = Combo10.List(0)


N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where kod =('" & "Hafte_date" & "')"
N_setting.Refresh
For I = 1 To N_setting.Recordset.RecordCount
Combo1.AddItem (N_setting.Recordset.Fields("id_setting") & " _ " & N_setting.Recordset.Fields("xvalue"))
N_setting.Recordset.MoveNext
Next I
Combo1.Text = Combo1.List(0)

Combo9.AddItem ("«–«‰ ’»Õ")
Combo9.AddItem ("ÿ·Ê⁄ ¬› «»")
Combo9.AddItem ("«–«‰ ŸÂ—")
Combo9.AddItem ("€—Ê» ¬› «»")
Combo9.AddItem ("«–«‰ „€—»")
Combo9.AddItem ("‰Ì„Â ‘»")

For I = 0 To 23
If I < 10 Then
Combo2.AddItem ("0" & I)
Combo8.AddItem ("0" & I)
Else
Combo2.AddItem (I)
Combo8.AddItem (I)
End If
Next I
For I = 0 To 59
If I < 10 Then
Combo3.AddItem ("0" & I)
Combo7.AddItem ("0" & I)
Else
Combo3.AddItem (I)
Combo7.AddItem (I)
End If
Next I
For I = 0 To 59
If I < 10 Then
Combo4.AddItem ("0" & I)
Combo6.AddItem ("0" & I)
Else
Combo4.AddItem (I)
Combo6.AddItem (I)
End If
Next I
For I = Taqvim.KKK.Caption To Val(Taqvim.KKK.Caption) + 0
Combo5.AddItem (I)
Next I



Ref_resh_list_mehver_list1

Refresh_date_hafte_lis
End Sub

Private Sub Form_Unload(Cancel As Integer)
Pakhsh_f.up_text.Text = "10012513"
Pakhsh_f.Show

Unload Me
End Sub

Private Sub Label10_Click()
select_goroh_F.Show
select_goroh_F.WATT.Text = "mehvar"
End Sub
Function Refresh_date_hafte_lis()
mehvar_list.Clear

N_hafte.Refresh
N_hafte.RecordSource = "select * from N_hafte"
N_hafte.Refresh

For I = 1 To N_hafte.Recordset.RecordCount
mehvar_list.AddItem ("ÅŒ‘ Â› êÌ" & " _ " & N_hafte.Recordset.Fields("id_hafte") & " _ " & N_hafte.Recordset.Fields("xname") & " _ " & N_hafte.Recordset.Fields("xtime"))
N_hafte.Recordset.MoveNext
Next I

N_date.Refresh
N_date.RecordSource = "select * from n_date"
N_date.Refresh
For I = 1 To N_date.Recordset.RecordCount
mehvar_list.AddItem ("ÅŒ‘ »«  «—ÌŒ Œ«’" & " _ " & N_date.Recordset.Fields("id_date") & " _ " & N_date.Recordset.Fields("xname") & " _ " & N_date.Recordset.Fields("xtime"))
N_date.Recordset.MoveNext
Next I

End Function
Private Sub Label11_Click()
select_hafte_F.Show
select_hafte_F.WATT.Text = "mehvar"

End Sub

Private Sub Label13_Click()
select_hafte_F.Show
select_hafte_F.WATT.Text = "date"

End Sub

Private Sub Label14_Click()
select_mp3_f.Show

End Sub

Private Sub Label17_Click()
Label17.Visible = False
Label16.Visible = False
Label18.Visible = False
Label19.Visible = False
Label34.Visible = True
Label33.Visible = True
Label32.Visible = True

End Sub

Private Sub Label19_Click()
Label17.Visible = False
Label16.Visible = False
Label18.Visible = False
Label19.Visible = False
Label37.Visible = True
Label36.Visible = True
Label35.Visible = True
End Sub

Private Sub Label20_Click()
MP3F_N.Show

End Sub

Private Sub Label21_Click()
Add_goroh_f.Show

End Sub

Private Sub Label22_Click()
Mp3d_goroh_F.Show

End Sub

Private Sub Label23_Click()
select_shahr.Show

End Sub

Private Sub Label25_Click()
History_form.Show

End Sub

Private Sub Label26_Click()
select_goroh_for_azan_from.Show

End Sub

Private Sub Label27_Click()
WE.Show

End Sub

Private Sub Label28_Click()
add_shahr_F.Show

End Sub

Private Sub Label29_Click()
Pakhsh_f.Show
Pakhsh_f.up_text.Text = ""
Pakhsh_f.up_text.Text = "10012513"
Unload Me

End Sub

Private Sub Label3_Click()
select_goroh_F.Show
select_goroh_F.WATT.Text = "date"
End Sub

Private Sub Label30_Click()
End

End Sub

Private Sub Label31_Click()
Shell "explorer.exe " & App.Path, vbNormalFocus



End Sub

Private Sub Label32_Click()
Label17.Visible = True
Label16.Visible = True
Label18.Visible = False
Label19.Visible = False
Label34.Visible = False
Label33.Visible = False
Label32.Visible = False
End Sub

Private Sub Label33_Click()
On Error Resume Next

a = Split(mehvar_list.Text, " _ ")
If a(0) = "ÅŒ‘ Â› êÌ" Then

N_hafte.Refresh
N_hafte.RecordSource = "select * from n_hafte where id_hafte like ('" & a(1) & "')"
N_hafte.Refresh
If N_hafte.Recordset.BOF = False Or N_hafte.Recordset.EOF = False Then N_hafte.Recordset.Delete



ElseIf a(0) = "ÅŒ‘ »«  «—ÌŒ Œ«’" Then

N_date.Refresh
N_date.RecordSource = "select * from N_date where id_date like ('" & a(1) & "')"
N_date.Refresh
If N_date.Recordset.BOF = False Or N_date.Recordset.EOF = False Then N_date.Recordset.Delete

End If


Label17.Visible = True
Label16.Visible = True
Label18.Visible = False
Label19.Visible = False
Label34.Visible = False
Label33.Visible = False
Label32.Visible = False



Refresh_date_hafte_lis

End Sub


Private Sub Label36_Click()
On Error Resume Next
af = ""
af = Split(List1.Text, " _ ")

N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where id_setting like ('" & af(0) & "')"
N_setting.Refresh
If N_setting.Recordset.BOF = False Or N_setting.Recordset.EOF = False Then N_setting.Recordset.Delete




Label17.Visible = False
Label16.Visible = False
Label18.Visible = True
Label19.Visible = True
Label37.Visible = False
Label36.Visible = False
Label35.Visible = False
Ref_resh_list_mehver_list1

End Sub

Private Sub Label37_Click()
Label17.Visible = False
Label16.Visible = False
Label18.Visible = True
Label19.Visible = True
Label37.Visible = False
Label36.Visible = False
Label35.Visible = False
End Sub

Private Sub List1_Click()
Label16.Visible = False
Label17.Visible = False
Label18.Visible = True
Label19.Visible = True

End Sub

Private Sub mehvar_list_Click()
Label16.Visible = True
Label17.Visible = True
Label18.Visible = False
Label19.Visible = False

End Sub

Private Sub Text1_Click()
If Text1.Text = "⁄‰Ê«‰" Then Text1.Text = ""

End Sub

Private Sub Text2_Click()
If Text2.Text = "⁄‰Ê«‰" Then Text2.Text = ""

End Sub

Private Sub Text3_Click()
If Text3.Text = "(»Â „œ  ( À«‰ÌÂ " Then Text3.Text = ""

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Text4_Click()
If Text4.Text = "(»Â „œ  ( À«‰ÌÂ " Then Text4.Text = ""

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
Timer_for_change_color_command_6 = Timer_for_change_color_command_6 - 1
If Timer_for_change_color_command_6 = 0 Then
Command6.BackColor = &HC0C000
Command6.Caption = "À» "

Timer1.Enabled = False
End If


End Sub

Private Sub Timer2_Timer()
Timer_for_change_color_command_1 = Timer_for_change_color_command_1 - 1
If Timer_for_change_color_command_1 = 0 Then
Command1.BackColor = &HC0C000
Command1.Caption = "À» "

Timer2.Enabled = False
End If

End Sub

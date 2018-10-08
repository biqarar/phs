VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form MP3F_N 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À»  ’Ê "
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13095
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MP3F_N.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Update_co 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Picture         =   "MP3F_N.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "«÷«›Â ò—œ‰ ’Ê "
      Top             =   360
      Visible         =   0   'False
      Width           =   975
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
      Left            =   0
      TabIndex        =   20
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
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Picture         =   "MP3F_N.frx":405A
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Õ–› ’Ê  «“ ê—ÊÂ „‰ Œ»"
      Top             =   360
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "„Ì“«‰ ’œ«"
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   4575
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         Value           =   50
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "  0  "
         ForeColor       =   &H8000000B&
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   " 100 "
         ForeColor       =   &H8000000B&
         Height          =   345
         Left            =   4080
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.TextBox url_text 
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   5880
      TabIndex        =   2
      Top             =   1320
      Width           =   6255
   End
   Begin VB.TextBox xname_text 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   6255
   End
   Begin VB.TextBox tozih_text 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   6255
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   5880
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   6255
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
      Left            =   5400
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404000&
      Caption         =   "Ã” ÃÊ"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   1920
      Width           =   855
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
      Height          =   855
      Left            =   4800
      Picture         =   "MP3F_N.frx":861F
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "À»  ê—ÊÂ Ê  ⁄ÌÌ‰ ‘„«—Â ’Ê "
      Top             =   360
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MP3F_N.frx":C07C
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   4210688
      ForeColor       =   65535
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
         DataField       =   "ID_mp3d"
         Caption         =   "‘„«—Â ’Ê "
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
         DataField       =   "xname"
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
      BeginProperty Column02 
         DataField       =   "url"
         Caption         =   "¬—œ”"
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
         DataField       =   "volome"
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
      BeginProperty Column04 
         DataField       =   "tozih"
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
      BeginProperty Column05 
         DataField       =   "OnOff"
         Caption         =   "OnOff"
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
         DataField       =   "SOnOff"
         Caption         =   "SOnOff"
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
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2550.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3974.74
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin VB.Label laqv_l 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "·€Ê"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer Wmp1 
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   240
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
      _cy             =   1296
   End
   Begin VB.Label eror_l 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "«Ì‰ ’Ê  œ— ·Ì”  ÊÃÊœ œ«—œ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   12360
      TabIndex        =   15
      Top             =   360
      Width           =   285
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   12480
      TabIndex        =   14
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¬œ—”"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   12360
      TabIndex        =   13
      Top             =   1320
      Width           =   390
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Ê÷ÌÕ« "
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   12240
      TabIndex        =   12
      Top             =   1800
      Width           =   645
   End
   Begin VB.Menu mnuPlay 
      Caption         =   "ÅŒ‘ ’Ê "
      Begin VB.Menu mnuPlay_asli 
         Caption         =   "ÅŒ‘ ’Ê "
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu nmyedit 
      Caption         =   "«’·«Õ ’Ê  »« œÊ»«— ò·Ìò »— —ÊÌ ’Ê "
   End
End
Attribute VB_Name = "MP3F_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Secent_for_show_and_hide_label As Integer

Dim IdA() As String

Function SPL_(str)
IdA = Split(str, " _ ")


End Function
Private Sub Combo1_Click()
SPL_ (Combo1.Text)
If IdA(0) = 0 Then
Add_goroh_f.Show
Me.Enabled = False

End If
End Sub

Private Sub Command10_Click()
'On Error Resume Next

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ’Ê  —« Õ–› ò‰Ìœ", vbYesNo, "Õ–› ’Ê ") = vbYes Then
last_id_delete = N_mp3d.Recordset.Fields("id_mp3d")
N_mp3d.Recordset.Delete
N_mp3dgoroh.Refresh
N_mp3dgoroh.RecordSource = "select * from n_mp3dgoroh where id_mp3d = ('" & last_id_delete & "')"
N_mp3dgoroh.Refresh
For I = 1 To N_mp3dgoroh.Recordset.RecordCount
N_mp3dgoroh.Recordset.Delete
N_mp3dgoroh.Recordset.MoveNext
Next I

a = Show_error_or_insert("’Ê  »« „Ê›ﬁÌ  Õ–› ‘œ", &H80&, 20)



End If

End Sub

Private Sub Command2_Click()
FFOF.Show

End Sub

Private Sub Command7_Click()
If Combo1.Text = "" Or Combo1.Text = "0 _ «÷«›Â ò—œ‰ ê—ÊÂ ÃœÌœ" Or xname_text.Text = "" Or url_text.Text = "" Then Exit Sub

N_mp3d.Refresh
N_mp3d.RecordSource = "SELECT * FROM n_mp3d where url like ('" & url_text.Text & "%')"
N_mp3d.Refresh
If N_mp3d.Recordset.BOF = True Or N_mp3d.Recordset.EOF = True Then
N_mp3d.Refresh
N_mp3d.Recordset.AddNew
N_mp3d.Recordset.Fields("xname") = xname_text.Text
N_mp3d.Recordset.Fields("url") = url_text.Text
N_mp3d.Recordset.Fields("volome") = Slider1.Value
N_mp3d.Recordset.Fields("tozih") = tozih_text.Text

N_mp3d.Recordset.Update

last_id_inserted_N_mp3d = N_mp3d.Recordset.Fields("id_mp3d")

N_mp3d.Refresh


N_mp3dgoroh.Refresh
N_mp3dgoroh.Recordset.AddNew
N_mp3dgoroh.Recordset.Fields("id_mp3d") = last_id_inserted_N_mp3d
a = SPL_(Combo1.Text)
N_mp3dgoroh.Recordset.Fields("id_goroh") = IdA(0)
N_mp3dgoroh.Recordset.Update
N_mp3dgoroh.Refresh






a = Show_error_or_insert("’Ê  »« „Ê›ﬁÌ  «÷«›Â ‘œ", &HC000&, 20)






Else

a = Show_error_or_insert("«Ì‰ ’Ê  œ— ·Ì  ÊÃÊœ œ«—œ", &HFF&, 20)



End If

End Sub
Function Show_error_or_insert(str_, color_, secend_)
eror_l.Visible = True
eror_l.Caption = str_
eror_l.BackColor = color_
Secent_for_show_and_hide_label = secend_

Timer1.Enabled = True


End Function

Private Sub DataGrid1_DblClick()
On Error Resume Next
Check2.Value = 0
'.................

Update_co.Visible = True
 xname_text.Text = ""
url_text.Text = ""
Slider1.Value = 10
 tozih_text.Text = ""
 
 xname_text.Text = N_mp3d.Recordset.Fields("xname")
url_text.Text = N_mp3d.Recordset.Fields("url")
Slider1.Value = N_mp3d.Recordset.Fields("volome")
 tozih_text.Text = N_mp3d.Recordset.Fields("tozih")
DataGrid1.Enabled = False
Command7.Visible = False
Command10.Visible = False
laqv_l.Visible = True

Check2.Visible = False
'Command2.Visible = False


End Sub

Private Sub Form_Load()
N_goroh.Refresh
For I = 1 To N_goroh.Recordset.RecordCount
Combo1.AddItem (N_goroh.Recordset.Fields("id_goroh") & " _ " & N_goroh.Recordset.Fields("xname"))
N_goroh.Recordset.MoveNext
Next I

Combo1.AddItem ("0 _ «÷«›Â ò—œ‰ ê—ÊÂ ÃœÌœ")



End Sub

Private Sub laqv_l_Click()
DataGrid1.Enabled = True
Command7.Visible = True
Command10.Visible = True
Check2.Value = 1
Check2.Visible = True
Command2.Visible = True
laqv_l.Visible = False
Update_co.Visible = False
End Sub

Private Sub mnuPlay_asli_Click()
On Error Resume Next

Wmp1.URL = N_mp3d.Recordset.Fields("url")
Wmp1.settings.volume = N_mp3d.Recordset.Fields("volome")
End Sub

Private Sub nmyedit_Click()
Call DataGrid1_DblClick

End Sub

Private Sub Timer1_Timer()
Secent_for_show_and_hide_label = Secent_for_show_and_hide_label - 1
If Secent_for_show_and_hide_label = 0 Then
eror_l.Visible = False

Timer1.Enabled = False
End If

End Sub

Private Sub tozih_text_Change()
On Error Resume Next

If Check2.Value = 1 Then

N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where tozih like ('%" & tozih_text.Text & "%')"
N_mp3d.Refresh
End If
End Sub

Private Sub Update_co_Click()
Update_co.Visible = False

 N_mp3d.Recordset.Fields("xname") = xname_text.Text
N_mp3d.Recordset.Fields("url") = url_text.Text
N_mp3d.Recordset.Fields("volome") = Slider1.Value
N_mp3d.Recordset.Fields("tozih") = tozih_text.Text

N_mp3d.Recordset.Update

DataGrid1.Enabled = True
Command7.Visible = True
Command10.Visible = True
Check2.Value = 1
Check2.Visible = True
Command2.Visible = True
a = Show_error_or_insert(" €ÌÌ—«  «⁄„«· ‘œ", &HC000&, 20)


End Sub

Private Sub url_text_Change()
On Error Resume Next

If Check2.Value = 1 Then

N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where url like ('%" & url_text.Text & "%')"
N_mp3d.Refresh
End If
End Sub

Private Sub xname_text_Change()

On Error Resume Next

If Check2.Value = 1 Then

N_mp3d.Refresh
N_mp3d.RecordSource = "select * from n_mp3d where xname like ('%" & xname_text.Text & "%')"
N_mp3d.Refresh
End If

End Sub

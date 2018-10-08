VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Add_goroh_f 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ê—ÊÂ Â«"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Add_goroh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      Picture         =   "Add_goroh.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Õ–› ê—ÊÂ"
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   960
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
      Left            =   120
      Picture         =   "Add_goroh.frx":48CF
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "À»  ê—ÊÂ"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox NAME_GOROH_T 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.TextBox TOZIH_GOROH_T 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   6615
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
         Height          =   5055
         Left            =   1920
         TabIndex        =   13
         Top             =   2040
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
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         ItemData        =   "Add_goroh.frx":832C
         Left            =   120
         List            =   "Add_goroh.frx":832E
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ê—ÊÂ Â«Ì „ÊÃÊœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ŒÌ—"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "»·Ì"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ê—ÊÂ —« Õ–› ò‰Ìœø"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label ins_l 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ê—ÊÂ «÷«›Â ‘œ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label eror_l 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "«Ì‰ ê—ÊÂ œ— ·Ì”  ÊÃÊœ œ«—œ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Ê÷ÌÕ« "
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   6000
      TabIndex        =   10
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ê—ÊÂ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   6000
      TabIndex        =   9
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "Add_goroh_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Secent_for_show_and_hide_label As Integer

Private Sub Text3_Change()

End Sub

Private Sub Command10_Click()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Command10.Visible = False
Command7.Visible = False



End Sub

Private Sub Command7_Click()
If NAME_GOROH_T.Text = "" And TOZIH_GOROH_T.Text = "" Then Exit Sub

N_goroh.Refresh
N_goroh.RecordSource = "select * from n_goroh where xname like ('" & NAME_GOROH_T.Text & "')"
N_goroh.Refresh
If N_goroh.Recordset.BOF = True Or N_goroh.Recordset.EOF = True Then
N_goroh.Refresh
N_goroh.Recordset.AddNew
N_goroh.Recordset.Fields("xname") = NAME_GOROH_T.Text
N_goroh.Recordset.Fields("tozih") = TOZIH_GOROH_T.Text
N_goroh.Recordset.Update
N_goroh.Refresh
Refresh_goroh_list
ins_l.Visible = True
Secent_for_show_and_hide_label = 20

Timer1.Enabled = True
Else
eror_l.Visible = True
Secent_for_show_and_hide_label = 20

Timer1.Enabled = True

End If


End Sub
Function Refresh_goroh_list()
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
Private Sub Form_Load()
Refresh_goroh_list

End Sub

Private Sub Form_Unload(Cancel As Integer)
MP3F_N.Enabled = True
Unload Me

End Sub

Private Sub goroh_list_Click()
Command10.Visible = True

End Sub

Private Sub Label2_Click()
On Error Resume Next

N_goroh.Refresh
a = Split(goroh_list.Text, " _ ")
N_goroh.RecordSource = "select * from n_goroh where id_goroh like ('" & a(0) & "')"
N_goroh.Refresh

N_hafte.Refresh
N_hafte.RecordSource = "select * from n_hafte where id_goroh like ('" & a(0) & "')"
N_hafte.Refresh
If N_hafte.Recordset.BOF = False Or N_hafte.Recordset.EOF = False Then
T = Show_error_or_insert("«Ì‰ ê—ÊÂ œ— ·Ì”  ÅŒ‘ Â› êÌ œ— Õ«· «” ›«œÂ «” ")
Exit Sub
End If
N_setting.Refresh
N_setting.RecordSource = "select * from n_setting where xname like ('" & a(0) & "')"
N_setting.Refresh
If N_setting.Recordset.BOF = False Or N_setting.Recordset.EOF = False Then
T = Show_error_or_insert("«Ì‰ ê—ÊÂ œ— ·Ì”  ÅŒ‘ »« „ÕÊ—Ì  «Êﬁ«  ‘—⁄Ì «” ›«œÂ ‘œÂ «” ")
Exit Sub
End If

N_goroh.Recordset.Delete
Command10.Visible = True
Refresh_goroh_list
Command7.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
End Sub
Function Show_error_or_insert(str_)
Command10.Visible = True
'Refresh_goroh_list
Command7.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
eror_l.Visible = True
eror_l.Caption = str_
eror_l.BackColor = &HFF&
Secent_for_show_and_hide_label = 40

Timer1.Enabled = True


End Function
Private Sub Label3_Click()
Command10.Visible = True
Command7.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
End Sub

Private Sub Timer1_Timer()
Secent_for_show_and_hide_label = Secent_for_show_and_hide_label - 1
If Secent_for_show_and_hide_label = 0 Then
eror_l.Visible = False
ins_l.Visible = False

Timer1.Enabled = False
End If


End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form complaintsfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complaint Form"
   ClientHeight    =   11625
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9480
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11625
   ScaleWidth      =   9480
   Begin MSDataGridLib.DataGrid comp_grid 
      Bindings        =   "complaints.frx":0000
      Height          =   1215
      Left            =   0
      TabIndex        =   56
      Top             =   8880
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2143
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "comp_custph1"
         Caption         =   "comp_custph1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "comp_date"
         Caption         =   "comp_date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "comp_madeby"
         Caption         =   "Complaint Made by"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "comp_prdicno"
         Caption         =   "Inst. Card No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "comp_desc"
         Caption         =   "Complaint Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "comp_no"
         Caption         =   "Complaint No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "comp_status"
         Caption         =   "Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "comp_attdate"
         Caption         =   "Complaint Attd On"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "comp_attby"
         Caption         =   "Complaint Attd. By"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "comp_cause"
         Caption         =   "Complaint Cause"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "comp_action"
         Caption         =   "Action Taken"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   3390.236
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1904.882
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc comp_data 
      Height          =   330
      Left            =   7560
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\program files\custcareprj\custcaredb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\program files\custcareprj\custcaredb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "complaintstemp"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox comp_prodinfo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   5400
      TabIndex        =   55
      Top             =   4620
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Show_grid 
      Caption         =   "Show All &Complaints"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   47
      Top             =   8310
      Width           =   1905
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Details:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   30
      TabIndex        =   32
      Top             =   1620
      Width           =   7695
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Landmark:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   150
         TabIndex        =   51
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label lblcomp_landmark 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   2100
         TabIndex        =   50
         Top             =   1980
         Width           =   5205
      End
      Begin VB.Label lblemail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4290
         TabIndex        =   49
         Top             =   1590
         Width           =   675
      End
      Begin VB.Label lblcomp_custemail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   5340
         TabIndex        =   48
         Top             =   1590
         Width           =   1965
      End
      Begin VB.Label lblcomp_fax 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   5310
         TabIndex        =   46
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4290
         TabIndex        =   45
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblcomp_custph2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Left            =   5310
         TabIndex        =   44
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label lblcomp_custpin 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Left            =   2100
         TabIndex        =   43
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblcomp_custstate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Left            =   2130
         TabIndex        =   42
         Top             =   1410
         Width           =   2205
      End
      Begin VB.Label lblcomp_custcity 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Left            =   2130
         TabIndex        =   41
         Top             =   1140
         Width           =   2160
      End
      Begin VB.Label lblcomp_custadd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   2130
         TabIndex        =   40
         Top             =   660
         Width           =   3780
      End
      Begin VB.Label lblcomp_custname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2130
         TabIndex        =   39
         Top             =   330
         Width           =   2970
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone2:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4290
         TabIndex        =   38
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pincode:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   150
         TabIndex        =   37
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   150
         TabIndex        =   36
         Top             =   1410
         Width           =   675
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   150
         TabIndex        =   35
         Top             =   1140
         Width           =   555
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   150
         TabIndex        =   34
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Name:"
         DataSource      =   "comp_data"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   330
         Width           =   2475
      End
   End
   Begin VB.TextBox comp_no 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2880
      MaxLength       =   7
      TabIndex        =   5
      Top             =   6060
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   0
      TabIndex        =   30
      Top             =   780
      Width           =   4515
      Begin VB.ComboBox comp_custph1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2460
         TabIndex        =   0
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Phone No:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   270
         Width           =   2325
      End
   End
   Begin VB.Frame file_frame 
      BackColor       =   &H00E0E0E0&
      Height          =   840
      Left            =   0
      TabIndex        =   29
      Top             =   -60
      Width           =   4275
      Begin VB.CommandButton comp_delete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   855
         Picture         =   "complaints.frx":0018
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "DELETE"
         Top             =   120
         Width           =   820
      End
      Begin VB.CommandButton comp_undo 
         Caption         =   "&Undo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2535
         Picture         =   "complaints.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "UNDO"
         Top             =   120
         Width           =   820
      End
      Begin VB.CommandButton comp_query 
         Caption         =   "&Query"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1695
         Picture         =   "complaints.frx":080C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "QUERY"
         Top             =   120
         Width           =   820
      End
      Begin VB.CommandButton comp_help 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   3375
         Picture         =   "complaints.frx":090E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "HELP"
         Top             =   120
         Width           =   820
      End
      Begin VB.CommandButton comp_save 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   40
         Picture         =   "complaints.frx":0A10
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "SAVE"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   820
      End
   End
   Begin VB.Frame navigation_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   4320
      TabIndex        =   28
      Top             =   -60
      Width           =   3030
      Begin VB.CommandButton comp_bottom 
         Caption         =   "&Bottom"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   2280
         Picture         =   "complaints.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "BOTTOM"
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton comp_next 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1560
         Picture         =   "complaints.frx":14BC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton comp_top 
         Caption         =   "&Top"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   45
         Picture         =   "complaints.frx":18FE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "TOP"
         Top             =   120
         Width           =   675
      End
      Begin VB.CommandButton comp_previous 
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   735
         Picture         =   "complaints.frx":1D40
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "PREVIOUS"
         Top             =   120
         Width           =   820
      End
   End
   Begin VB.Frame complaint_frame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Complaint Details:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   0
      TabIndex        =   20
      Top             =   4380
      Width           =   9435
      Begin VB.ComboBox comp_prdicno1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Top             =   990
         Width           =   2520
      End
      Begin VB.TextBox comp_attdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox comp_madeby 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1039
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         MaxLength       =   25
         TabIndex        =   2
         Top             =   660
         Width           =   2500
      End
      Begin VB.ComboBox comp_status 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "complaints.frx":2182
         Left            =   2880
         List            =   "complaints.frx":218C
         TabIndex        =   6
         Top             =   2040
         Width           =   1470
      End
      Begin VB.TextBox comp_cause 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3060
         Width           =   6255
      End
      Begin VB.TextBox comp_date 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1039
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   1
         Top             =   330
         Width           =   1125
      End
      Begin VB.TextBox comp_desc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2865
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1350
         Width           =   4245
      End
      Begin VB.TextBox comp_attby 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2730
         Width           =   1725
      End
      Begin VB.TextBox comp_action 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         MaxLength       =   100
         TabIndex        =   10
         Top             =   3390
         Width           =   4245
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Installation Card No:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   54
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint Made By:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   53
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label lblcomp_no 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint No:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   52
         Top             =   1650
         Width           =   1530
      End
      Begin VB.Label lblcomp_action 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Action Taken:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   3390
         Width           =   1575
      End
      Begin VB.Label lbldate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   330
         Width           =   1515
      End
      Begin VB.Label lbldesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint Description:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   2475
      End
      Begin VB.Label lblattdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint Cause:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   3060
         Width           =   2025
      End
      Begin VB.Label lblcomp_attby 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint Attended. by:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   2730
         Width           =   2715
      End
      Begin VB.Label lblcomments 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Attended On:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label lblcomp_status 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   855
      End
   End
End
Attribute VB_Name = "complaintsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim compinsert, undoclick, ctr, recvalidated As Integer
Dim confirmadd, presentcontrol, curcomp_no, confirmdelete As String
Private Sub comp_action_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub comp_action_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_action Then
    complaints_wchange
End If
End Sub
Private Sub comp_attby_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub comp_attby_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_attby Then
    complaints_wchange
End If
End Sub
Private Sub comp_attdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 47 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
Private Sub comp_attdate_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_attdate Then
    complaints_wchange
End If
End Sub
Private Sub comp_attdate_LostFocus()
If comp_attdate = "" Then
    Exit Sub
End If
If IsDate(comp_attdate) Then
    comp_attdate.Text = CDate(comp_attdate)
Else
    MsgBox ("Please Enter A Valid Attendence Date")
    comp_attdate.SetFocus
End If
End Sub
Private Sub comp_cause_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub comp_cause_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_action Then
    complaints_wchange
End If
End Sub
Private Sub comp_custph1_Click()
mastershow
End Sub
Private Sub comp_custph1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
If KeyAscii = 8 Or KeyAscii = 32 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
Private Sub comp_custph1_LostFocus()
prodinstfrm.setcustcom
prodinstfrm.custrs.Filter = "cust_ph1='" & comp_custph1 & "'"
If prodinstfrm.custrs.EOF = True Or prodinstfrm.custrs.BOF = True Then
    compconfirmcustph1add
End If
End Sub

Private Sub comp_date_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 47 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
Private Sub comp_date_KeyUp(KeyCode As Integer, Shift As Integer)
detailshow
End Sub
Private Sub comp_date_LostFocus()
If IsDate(comp_date) Then
    comp_date.Text = CDate(comp_date)
Else
    MsgBox ("Please Enter A Valid Complaint Date")
    comp_date.SetFocus
End If
End Sub
Private Sub comp_desc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub comp_desc_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_desc Then
    complaints_wchange
End If
End Sub
Private Sub comp_grid_AfterDelete()
prodinstfrm.setcomptempcom
If prodinstfrm.comptemprs.RecordCount > 0 Then
    prodinstfrm.comptemprs.MoveNext
    If prodinstfrm.comptemprs.EOF = True Then
        prodinstfrm.comptemprs.MovePrevious
    End If
    comp_date.Text = prodinstfrm.comptemprs.Fields(1)
    showrecords1
    Else
    blankrecords
End If
End Sub
Private Sub comp_grid_BeforeDelete(Cancel As Integer)
prodinstfrm.custcaredb.Execute ("delete * from complaints where comp_custph1='" & comp_custph1 & "' and cstr(comp_date)='" & CStr(comp_date) & "'")
End Sub
Private Sub comp_grid_Click()
gridchange
End Sub
Private Sub comp_grid_KeyUp(KeyCode As Integer, Shift As Integer)
gridchange
End Sub
Private Sub comp_madeby_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub comp_madeby_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_madeby Then
    complaints_wchange
End If
End Sub
Private Sub comp_no_GotFocus()
If comp_no.Text = "" Then
    comp_no = "0" + CStr(Month(Date))
End If
End Sub
Private Sub comp_no_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
If KeyAscii = 8 Or KeyAscii = 32 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
Private Sub comp_no_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_no Then
    complaints_wchange
End If
End Sub
Private Sub comp_no_LostFocus()
If comp_no.Text = "" Then
    comp_no = "0" + CStr(Month(Date))
End If
End Sub
Private Sub comp_prdicno1_Click()
showprodinfo
End Sub
Private Sub comp_prdicno1_GotFocus()
showprodinfo
End Sub
Private Sub comp_prdicno1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
KeyAscii = 0
End Sub
Private Sub comp_prdicno1_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_prdicno1 Then
    complaints_wchange
End If
showprodinfo
End Sub
Private Sub comp_prdicno1_LostFocus()
comp_prodinfo.Visible = False
End Sub
Private Sub comp_prdicno11_GotFocus()
showprodinfo
End Sub

Private Sub comp_status_Click()
complaints_wchange
If comp_status = "Pending" Then
    comp_attdate.Enabled = False
    comp_attdate.BackColor = &H80000012
    comp_attby.Enabled = False
    comp_attby.BackColor = &H80000012
    comp_cause.Enabled = False
    comp_cause.BackColor = &H80000012
    comp_action.Enabled = False
    comp_action.BackColor = &H80000012
    Else
    comp_attdate.Enabled = True
    comp_attdate.BackColor = &HFFFFFF
    comp_attby.Enabled = True
    comp_attby.BackColor = &HFFFFFF
    comp_cause.Enabled = True
    comp_cause.BackColor = &HFFFFFF
    comp_action.Enabled = True
    comp_action.BackColor = &HFFFFFF
End If
End Sub
Private Sub Comp_status_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}", True
End If
KeyAscii = 0
End Sub
Private Sub comp_next_Click()
ctr = 1
If prodinstfrm.comprs.RecordCount > 0 Then
    comp_save = True
    If recvalidated = 0 Then
        Exit Sub
        Else
        recvalidated = 0
    End If
    prodinstfrm.setcomptempcom
    prodinstfrm.comptemprs.Filter = "comp_custph1='" & comp_custph1 & "' and comp_date='" & CDate(comp_date) & "'"
    prodinstfrm.setcomptempcom
    prodinstfrm.comptemprs.MovePrevious
    If prodinstfrm.comptemprs.BOF Then
        prodinstfrm.comptemprs.MoveNext
    End If
    comp_date.Text = prodinstfrm.comptemprs.Fields(1)
    showrecords1
End If
End Sub
Private Sub comp_custph1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 18 Or KeyCode = 80 Or KeyCode = 84 Or KeyCode = 66 Or KeyCode = 78 Or KeyCode = 13 Or KeyCode = 83 Then
    Exit Sub
End If
mastershow
End Sub
Private Sub comp_status_KeyUp(KeyCode As Integer, Shift As Integer)
If complaintsfrm.ActiveControl = comp_status Then
    complaints_wchange
End If
End Sub
Private Sub Form_Activate()
If newcomplaint = 1 Then
    comp_custph1.Text = curcustph1
    newcomplaint = 0
    comp_custph1.SetFocus
    mastershow
    Exit Sub
End If

If newcomplaint = 2 Then
    comp_custph1.Text = curcustph1
    comp_date.Text = curcustdate
    comp_custph1.SetFocus
    mastershow
    newcomplaint = 0
    Exit Sub
End If

ctr = 0
recvalidated = 0
compinsert = 0
comp_custph1.SetFocus
comp_undo.Enabled = False

'setting database connection
prodinstfrm.setcustcom
prodinstfrm.setprodtemp
prodinstfrm.setprodcom
prodinstfrm.setcomptempcom
prodinstfrm.setcompcom

'filling phone1 combo box
prodinstfrm.setcustcom
If prodinstfrm.custrs.RecordCount > 0 Then
    prodinstfrm.custrs.MoveFirst
    Do While Not prodinstfrm.custrs.EOF
        comp_custph1.AddItem (prodinstfrm.custrs.Fields(5))
        prodinstfrm.custrs.MoveNext
    Loop
    comp_custph1.Text = comp_custph1.List(0)
    prodinstfrm.custrs.Filter = "cust_ph1='" & comp_custph1 & "'"
    If prodinstfrm.custrs.BOF = False And prodinstfrm.custrs.EOF = False Then
        showrecords
    End If
End If

'filling product no combo box
fillprodicno

'if recordcount=0 then blank
prodinstfrm.setcompcom
If prodinstfrm.comprs.RecordCount = 0 Then
  compinsert = 1
  comp_date = Date
  comp_no = "1"
  blankrecords
  comp_delete.Enabled = False
  comp_query.Enabled = False
  comp_top.Enabled = False
  comp_bottom(0).Enabled = False
  comp_previous.Enabled = False
  comp_next.Enabled = False
  Else: mastershow
End If
End Sub
Public Sub blankrecords()
comp_madeby = ""
If comp_prdicno1.ListCount = 0 Then
    comp_prdicno1 = "0"
Else
    comp_prdicno1.Text = comp_prdicno1.List(0)
End If
If Month(Date) < 10 Then
    comp_no.Text = "0" + CStr(Month(Date))
    Else
    comp_no.Text = CStr(Month(Date))
End If
comp_desc = ""
comp_status.Text = comp_status.List(0)
comp_attdate.Text = ""
comp_attby.Text = ""
comp_cause.Text = ""
comp_action.Text = ""
comp_attdate.Enabled = False
comp_attdate.BackColor = &H80000012
comp_attby.Enabled = False
comp_attby.BackColor = &H80000012
comp_cause.Enabled = False
comp_cause.BackColor = &H80000012
comp_action.Enabled = False
comp_action.BackColor = &H80000012
End Sub
Private Sub comp_bottom_Click(Index As Integer)
ctr = 1
prodinstfrm.setcomptempcom
If prodinstfrm.comptemprs.RecordCount > 0 Then
  comp_save = True
  If recvalidated = 0 Then
    Exit Sub
  Else
    recvalidated = 0
  End If
  comp_custph1.SetFocus
  prodinstfrm.setcomptempcom
  prodinstfrm.comptemprs.MoveFirst
  comp_date.Text = prodinstfrm.comptemprs.Fields(1)
  showrecords1
End If
End Sub
Private Sub comp_delete_Click()
confirmdelete = MsgBox("Are You Sure You Want To Delete This Record", vbYesNo)
If confirmdelete = vbYes Then
    prodinstfrm.custcaredb.Execute ("delete * from complaints where comp_custph1 ='" & comp_custph1 & "' and cstr(comp_date)='" & CStr(comp_date) & "'")
    prodinstfrm.custcaredb.Execute ("delete * from complaintstemp")
    prodinstfrm.custcaredb.Execute ("insert into complaintstemp select * from complaints where comp_custph1='" & comp_custph1 & "'")
    prodinstfrm.setcompcom
    prodinstfrm.setcomptempcom

    If prodinstfrm.comptemprs.RecordCount > 0 Then
        prodinstfrm.comptemprs.MoveFirst
        comp_custph1.SetFocus
        comp_custph1 = prodinstfrm.comprs.Fields(0)
        comp_date = prodinstfrm.comptemprs.Fields(1)
        showrecords1
        comp_data.Refresh
    ElseIf prodinstfrm.comprs.RecordCount > 0 Then
        comp_custph1 = prodinstfrm.comprs.Fields(0)
        comp_date = prodinstfrm.comprs.Fields(1)
        comp_custph1.SetFocus
        mastershow
    Else
        compinsert = 1
        blankrecords
        comp_date = Date
        prodinstfrm.setcomptempcom
        comp_data.Refresh
        comp_delete.Enabled = False
        comp_query.Enabled = False
    End If
End If
End Sub
Private Sub comp_previous_Click()
ctr = 1
If prodinstfrm.comptemprs.RecordCount > 0 Then
  comp_save = True
  If recvalidated = 0 Then
    Exit Sub
  Else
    recvalidated = 0
  End If
  prodinstfrm.setcomptempcom
  prodinstfrm.comptemprs.Filter = "comp_custph1='" & comp_custph1 & " 'and comp_date='" & CDate(comp_date) & "'"
  prodinstfrm.setcomptempcom
  prodinstfrm.comptemprs.MoveNext
  If prodinstfrm.comptemprs.EOF Then
    prodinstfrm.comptemprs.MovePrevious
  End If
  
  comp_date.Text = prodinstfrm.comptemprs.Fields(1)
  showrecords1
  comp_custph1.SetFocus
End If
End Sub
Private Sub comp_query_Click()
compqueryfrm.Top = 1000
compqueryfrm.Left = 1000
compqueryfrm.Height = 3200
compqueryfrm.Width = 7500
compqueryfrm.Show
End Sub
Private Sub comp_save_Click()
  'validating data
  If comp_custph1.Text = "" Then
    MsgBox ("Phone No Cannot Be Empty")
    comp_custph1.SetFocus
    Exit Sub
  End If
  
  If comp_date.Text = "" Or Not IsDate(comp_date) Then
    MsgBox ("Please Enter Valid Complaint Date")
    comp_date.SetFocus
    Exit Sub
  End If
  
  If comp_desc.Text = "" Then
    MsgBox ("Description Cannot Be Empty")
    comp_desc.SetFocus
    Exit Sub
  End If
 
 If comp_madeby.Text = "" Then
    MsgBox ("Complaint Made By Cannot Be Empty")
    comp_madeby.SetFocus
    Exit Sub
  End If
 
 If comp_madeby.Text = "" Then
    MsgBox ("Complaint Made By Cannot Be Empty")
    comp_madeby.SetFocus
    Exit Sub
 End If
 
 If comp_no.Text = "" Then
    MsgBox ("Complaint No By Cannot Be Empty")
    comp_no.SetFocus
    Exit Sub
 End If
  
  If comp_status.Text = "" Then
    MsgBox ("Complaint Status Cannot Be Empty")
    comp_status.SetFocus
    Exit Sub
  End If
   
  If comp_status.Text = "Completed" Then
    If comp_attdate.Text = "" Or comp_attdate.Text = "0" Then
        MsgBox ("Complaint Attended On Date Cannot Be Empty")
        comp_attdate.SetFocus
        Exit Sub
    End If
    If CDate(comp_attdate.Text) < comp_date.Text Then
        MsgBox ("Complaint Attended On Date Should Be Greater Than Complaint Date.")
        comp_attdate.SetFocus
        Exit Sub
    End If
    If comp_attby.Text = "" Then
        MsgBox ("Complaint Attended By Cannot Be Empty")
        comp_attby.SetFocus
        Exit Sub
    End If
    If comp_cause.Text = "" Then
        MsgBox ("Complaint Cause Cannot Be Empty")
        comp_cause.SetFocus
        Exit Sub
    End If
    If comp_action.Text = "" Then
        MsgBox ("Action Taken Cannot Be Empty")
        comp_action.SetFocus
        Exit Sub
    End If
  End If
   
 'checking for duplicate complaint no
  prodinstfrm.setcompcom
  If prodinstfrm.comprs.RecordCount > 0 Then
    prodinstfrm.comprs.MoveFirst
    Do While Not prodinstfrm.comprs.EOF
        If (prodinstfrm.comprs.Fields(0) = Val(comp_custph1) And prodinstfrm.comprs.Fields(1) <> comp_date) Or (prodinstfrm.comprs.Fields(0) <> Val(comp_custph1)) Then
            If prodinstfrm.comprs.Fields(8) = Val(comp_no) Then
                MsgBox ("This Complaint No Already Exists")
                comp_no.SetFocus
                Exit Sub
            End If
        End If
        prodinstfrm.comprs.MoveNext
    Loop
  End If
  
  'transfer complaintstemp data to complaints table
  prodinstfrm.custcaredb.Execute ("delete * from complaints where comp_custph1='" & comp_custph1.Text & "'")
  prodinstfrm.custcaredb.Execute ("insert into complaints select * from complaintstemp where comp_custph1='" & comp_custph1.Text & "'")
  
  'transfer complaints data to complaintstemp table
  prodinstfrm.custcaredb.Execute ("delete * from complaintstemp")
  prodinstfrm.custcaredb.Execute ("insert into complaintstemp select * from complaints where comp_custph1='" & comp_custph1.Text & "'")
  
  compinsert = 0
  prodinstfrm.setcomptempcom
  comp_data.Refresh
  
  recvalidated = 1
  ctr = 1
  comp_delete.Enabled = True
  comp_query.Enabled = True
  comp_custph1.SetFocus
End Sub
Private Sub comp_top_Click()
ctr = 1
prodinstfrm.setcomptempcom
If prodinstfrm.comptemprs.RecordCount > 0 Then
  comp_save = True
  If recvalidated = 0 Then
    Exit Sub
  Else
    recvalidated = 0
  End If
  comp_custph1.SetFocus
  prodinstfrm.comptemprs.MoveLast
  comp_date.Text = prodinstfrm.comptemprs.Fields(1)
  showrecords1
End If
End Sub
Private Sub comp_undo_Click()
  undoclick = 1
  comp_undo.Enabled = False
  comp_delete.Enabled = True
  comp_query.Enabled = True
  comp_top.Enabled = True
  comp_bottom(0).Enabled = True
  comp_previous.Enabled = True
  comp_next.Enabled = True
  prodinstfrm.setcompcom
  If prodinstfrm.comprs.RecordCount > 0 Then
    prodinstfrm.comprs.Filter = "comp_custph1<>'" & comp_custph1.Text & "'"
    showrecords
    comp_custph1.Text = prodinstfrm.comprs.Fields(0)
    comp_custph1.SetFocus
    prodinstfrm.custcaredb.Execute ("Insert into complaintstemp select * from complaints where comp_custph1='" & comp_custph1.Text & "'")
    prodinstfrm.setcomptempcom
    prodinstfrm.comptemprs.MoveFirst
    comp_data.Refresh
    showrecords1
  End If
End Sub
Public Sub mastershow()
prodinstfrm.custrs.Filter = "cust_ph1='" & comp_custph1 & "'"

If prodinstfrm.custrs.EOF = False And prodinstfrm.custrs.BOF = False Then
    showrecords
    fillprodicno
Else
    lblcomp_custname.Caption = ""
    lblcomp_custadd.Caption = ""
    lblcomp_custcity.Caption = ""
    lblcomp_custstate.Caption = ""
    lblcomp_custpin.Caption = 0
    lblcomp_custph2.Caption = 0
    lblcomp_fax.Caption = 0
    lblcomp_custemail.Caption = ""
    lblcomp_landmark.Caption = ""
End If

prodinstfrm.custcaredb.Execute ("delete * from complaintstemp")
prodinstfrm.setcompcom

If prodinstfrm.comprs.RecordCount > 0 Then
    If newcomplaint = 2 Then
        prodinstfrm.comprs.Filter = "comp_custph1='" & curcustph1 & "' and comp_date='" & curcustdate & " '"
    Else
        prodinstfrm.comprs.Filter = "comp_custph1='" & comp_custph1 & "'"
    End If

    If prodinstfrm.comprs.EOF = False And prodinstfrm.comprs.BOF = False Then
        comp_undo.Enabled = False
        comp_delete.Enabled = True
        comp_query.Enabled = True
        comp_top.Enabled = True
        comp_bottom(0).Enabled = True
        comp_previous.Enabled = True
        comp_next.Enabled = True
        prodinstfrm.custcaredb.Execute ("Insert into complaintstemp select * from complaints where comp_custph1='" & comp_custph1 & "'")
        prodinstfrm.setcomptempcom
        comp_date = prodinstfrm.comptemprs.Fields(1)
        If prodinstfrm.comptemprs.RecordCount > 0 Then
            If newcomplaint = 2 Then
               prodinstfrm.comptemprs.Filter = "comp_custph1='" & curcustph1 & "' and comp_date='" & curcustdate & " '"
            Else
                prodinstfrm.comptemprs.MoveFirst
            End If
                showrecords1
        End If
        comp_data.Refresh
    Else: blankrecords
        comp_undo.Enabled = True
        comp_delete.Enabled = False
        comp_query.Enabled = False
        comp_top.Enabled = False
        comp_bottom(0).Enabled = False
        comp_previous.Enabled = False
        comp_next.Enabled = False
        comp_data.Refresh
        compinsert = 1
    End If
End If
End Sub
Public Sub complaints_wchange()
If Trim(comp_custph1.Text) <> "" And IsDate(comp_date) And comp_madeby.Text <> "" Then
    If compinsert = 1 Then
        prodinstfrm.custcaredb.Execute ("insert into complaintstemp values ('" & comp_custph1 & "','" & comp_date & "','" & comp_desc & "','" & comp_attdate & "','" & comp_attby & "','" & comp_cause & "','" & comp_action & "','" & comp_status & "','" & comp_no & "','" & comp_madeby & "','" & comp_prdicno1 & "')")
        compinsert = 0
    Else
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_desc='" & comp_desc.Text & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_attdate='" & comp_attdate.Text & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_attby='" & comp_attby.Text & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_cause='" & comp_cause.Text & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_action='" & comp_action.Text & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_status='" & Trim(comp_status.Text) & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_no='" & Trim(comp_no.Text) & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_madeby='" & Trim(comp_madeby.Text) & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
        prodinstfrm.custcaredb.Execute ("update complaintstemp set comp_prdicno='" & Trim(comp_prdicno1.Text) & "' where complaintstemp.comp_custph1='" & comp_custph1 & "' and cstr(complaintstemp.comp_date)='" & CStr(comp_date) & "'")
    End If
End If
End Sub
Public Sub showrecords1()
If prodinstfrm.comptemprs.RecordCount > 0 Then
    If prodinstfrm.comptemprs.Fields(9) <> "" Then
        comp_madeby = prodinstfrm.comptemprs.Fields(9)
    End If
    If prodinstfrm.comptemprs.Fields(10) <> "" Then
        comp_prdicno1 = prodinstfrm.comptemprs.Fields(10)
    End If
    If prodinstfrm.comptemprs.Fields(8) <> "" Then
        comp_no.Text = prodinstfrm.comptemprs.Fields(8)
    End If
    If prodinstfrm.comptemprs.Fields(2) <> "" Then
        comp_desc = prodinstfrm.comptemprs.Fields(2)
    End If
    If prodinstfrm.comptemprs.Fields(7) <> "" Then
        comp_status.Text = prodinstfrm.comptemprs.Fields(7)
    End If
    If prodinstfrm.comptemprs.Fields(3) <> "" Then
        comp_attdate.Text = prodinstfrm.comptemprs.Fields(3)
    End If
    If prodinstfrm.comptemprs.Fields(4) <> "" Then
        comp_attby.Text = prodinstfrm.comptemprs.Fields(4)
    End If
    If prodinstfrm.comptemprs.Fields(5) <> "" Then
        comp_cause.Text = prodinstfrm.comptemprs.Fields(5)
    End If
    If prodinstfrm.comptemprs.Fields(6) <> "" Then
        comp_action.Text = prodinstfrm.comptemprs.Fields(6)
    End If
    If Trim(comp_status) = "Completed" Then
        comp_attdate.Enabled = True
        comp_attdate.BackColor = &HFFFFFF
        comp_attby.Enabled = True
        comp_attby.BackColor = &HFFFFFF
        comp_cause.Enabled = True
        comp_cause.BackColor = &HFFFFFF
        comp_action.Enabled = True
        comp_action.BackColor = &HFFFFFF
        Else
        comp_attdate.Enabled = False
        comp_attdate.BackColor = &H80000012
        comp_attby.Enabled = False
        comp_attby.BackColor = &H80000012
        comp_cause.Enabled = False
        comp_cause.BackColor = &H80000012
        comp_action.Enabled = False
        comp_action.BackColor = &H80000012
    End If
End If
End Sub
Public Sub detailshow()
If comp_custph1 <> "" And IsDate(comp_date) Then
    prodinstfrm.setcomptempcom
    prodinstfrm.comptemprs.Filter = "comp_custph1='" & comp_custph1.Text & "' and comp_date='" & CDate(comp_date.Text) & "'"
    If prodinstfrm.comptemprs.EOF = False And prodinstfrm.comptemprs.BOF = False Then
        compinsert = 0
        comp_undo.Enabled = False
        showrecords1
        comp_undo.Enabled = False
        comp_delete.Enabled = True
        comp_query.Enabled = True
        comp_top.Enabled = True
        comp_bottom(0).Enabled = True
        comp_previous.Enabled = True
        comp_next.Enabled = True
        Else
        compinsert = 1
        comp_undo.Enabled = True
        comp_delete.Enabled = False
        comp_query.Enabled = False
        comp_top.Enabled = False
        comp_bottom(0).Enabled = False
        comp_previous.Enabled = False
        comp_next.Enabled = False
        blankrecords
    End If
End If
End Sub
Public Sub gridchange()
If comp_grid.Columns(0) <> "" Then
    prodinstfrm.setcomptempcom
    prodinstfrm.comptemprs.Filter = "comp_custph1='" & comp_custph1 & "'and comp_date='" & comp_grid.Columns(1) & "'"
    If prodinstfrm.comptemprs.EOF = False And prodinstfrm.comptemprs.BOF = False Then
        comp_date = comp_grid.Columns(1)
        showrecords1
    End If
End If
End Sub
Public Sub showrecords()
lblcomp_custname.Caption = prodinstfrm.custrs.Fields(0)
lblcomp_custadd.Caption = prodinstfrm.custrs.Fields(1)
lblcomp_custcity.Caption = prodinstfrm.custrs.Fields(2)
lblcomp_custstate.Caption = prodinstfrm.custrs.Fields(3)
lblcomp_custpin.Caption = prodinstfrm.custrs.Fields(4)
lblcomp_custph2.Caption = prodinstfrm.custrs.Fields(6)
lblcomp_fax.Caption = prodinstfrm.custrs.Fields(7)
lblcomp_custemail.Caption = prodinstfrm.custrs.Fields(8)
lblcomp_landmark.Caption = prodinstfrm.custrs.Fields(9)
End Sub

Private Sub Show_grid_Click()
If comp_grid.Visible = False Then
    comp_grid.Visible = True
    Show_grid.Caption = "&Hide Grid"
    
    Else
    comp_grid.Visible = False
    Show_grid.Caption = "Show &All Complaints"
End If
End Sub
Private Sub fillprodicno()
comp_prdicno1.Clear
prodinstfrm.setprodcom
If prodinstfrm.prodrs.RecordCount > 0 Then
    prodinstfrm.prodrs.MoveFirst
    Do While Not prodinstfrm.prodrs.EOF
        If comp_custph1 = prodinstfrm.prodrs.Fields(0) Then
            comp_prdicno1.AddItem (prodinstfrm.prodrs.Fields(5))
        End If
        prodinstfrm.prodrs.MoveNext
    Loop
    comp_prdicno1.Text = comp_prdicno1.List(0)
End If
End Sub
Public Sub showprodinfo()
If comp_prdicno1.ListCount > 0 Then
    comp_prodinfo.Visible = True
    comp_prodinfo.Clear
    prodinstfrm.setprodcom
    prodinstfrm.prodrs.MoveFirst
    Do While Not prodinstfrm.prodrs.EOF
    If prodinstfrm.prodrs.Fields(0) = comp_custph1.Text And prodinstfrm.prodrs.Fields(5) = comp_prdicno1.Text Then
        comp_prodinfo.AddItem ("PRODUCT NO:                  " + prodinstfrm.prodrs.Fields(1))
        comp_prodinfo.AddItem ("PRODUCT NAME:               " + prodinstfrm.prodrs.Fields(2))
        comp_prodinfo.AddItem ("PRODUCT INSTALLED ON:  " + CStr(prodinstfrm.prodrs.Fields(3)))
        comp_prodinfo.AddItem ("PRODUCT INSTALLED BY:  " + prodinstfrm.prodrs.Fields(4))
    End If
    prodinstfrm.prodrs.MoveNext
    Loop
End If
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form customerfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Form"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   9405
   Begin VB.CommandButton prod_delete 
      Cancel          =   -1  'True
      Caption         =   "Delete &Product"
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
      Left            =   6480
      TabIndex        =   45
      Top             =   5160
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid prod_grid 
      Bindings        =   "customer.frx":0000
      Height          =   1215
      Left            =   165
      TabIndex        =   44
      Top             =   7440
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   2143
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "prod_icno"
         Caption         =   "Installation Card No"
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
         DataField       =   "prod_no"
         Caption         =   "Product No"
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
         DataField       =   "prod_name"
         Caption         =   "Product Name"
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
         DataField       =   "prod_inston"
         Caption         =   "Installed On"
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
         DataField       =   "prod_instby"
         Caption         =   "Installed By"
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
         DataField       =   "prod_icno"
         Caption         =   "prod_icno"
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
            ColumnWidth     =   2280.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2310.236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc prod_data 
      Height          =   375
      Left            =   7440
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      RecordSource    =   "producttemp"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Details:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   120
      TabIndex        =   38
      Top             =   5010
      Width           =   6195
      Begin VB.TextBox cust_prodicno 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   10
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox cust_instby 
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
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1770
         Width           =   3165
      End
      Begin VB.TextBox cust_inston 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1410
         Width           =   1155
      End
      Begin VB.TextBox cust_prodname 
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
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1050
         Width           =   3075
      End
      Begin VB.TextBox cust_prodno 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblprod_icno 
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
         Height          =   285
         Left            =   390
         TabIndex        =   43
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblinstby 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Installed by:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   390
         TabIndex        =   42
         Top             =   1770
         Width           =   1455
      End
      Begin VB.Label lblprod_inston 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Installed On:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   390
         TabIndex        =   41
         Top             =   1410
         Width           =   1485
      End
      Begin VB.Label lblprod_name 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   40
         Top             =   1050
         Width           =   1635
      End
      Begin VB.Label lblprod_no 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product No:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   39
         Top             =   720
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   930
      Left            =   4260
      TabIndex        =   35
      Top             =   -90
      Width           =   3480
      Begin VB.CommandButton navigaterec 
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
         Index           =   3
         Left            =   2580
         Picture         =   "customer.frx":0018
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   210
         Width           =   820
      End
      Begin VB.CommandButton navigaterec 
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
         Index           =   2
         Left            =   1740
         Picture         =   "customer.frx":045A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   210
         Width           =   820
      End
      Begin VB.CommandButton navigaterec 
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
         Index           =   1
         Left            =   900
         Picture         =   "customer.frx":089C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   210
         Width           =   820
      End
      Begin VB.CommandButton navigaterec 
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
         Index           =   0
         Left            =   60
         Picture         =   "customer.frx":0CDE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   210
         Width           =   820
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Main Details:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   90
      TabIndex        =   25
      Top             =   930
      Width           =   7185
      Begin VB.TextBox cust_email 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3450
         Width           =   3075
      End
      Begin VB.TextBox cust_landmark 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2430
         Width           =   4395
      End
      Begin VB.TextBox cust_fax 
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
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   8
         Top             =   3090
         Width           =   1065
      End
      Begin VB.TextBox cust_state 
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
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1740
         Width           =   1065
      End
      Begin VB.TextBox cust_name 
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
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   1
         Top             =   720
         Width           =   2500
      End
      Begin VB.TextBox cust_add 
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
         Left            =   2280
         MaxLength       =   65
         TabIndex        =   2
         Top             =   1080
         Width           =   3285
      End
      Begin VB.TextBox cust_city 
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
         Left            =   2265
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1410
         Width           =   1065
      End
      Begin VB.TextBox cust_pin 
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2070
         Width           =   1065
      End
      Begin VB.TextBox cust_ph1 
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
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   0
         Top             =   360
         Width           =   1365
      End
      Begin VB.TextBox cust_ph2 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox cust_ph1find 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   3450
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   330
         TabIndex        =   36
         Top             =   2400
         Width           =   1185
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone1:"
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
         Left            =   390
         TabIndex        =   33
         Top             =   330
         Width           =   735
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
         Height          =   255
         Left            =   390
         TabIndex        =   32
         Top             =   1080
         Width           =   1035
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
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   1380
         Width           =   555
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
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1740
         Width           =   675
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
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2070
         Width           =   915
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   2730
         Width           =   975
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   3060
         Width           =   495
      End
   End
   Begin VB.CommandButton cust_delete 
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
      Picture         =   "customer.frx":1120
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "DELETE"
      Top             =   120
      Width           =   820
   End
   Begin VB.Frame file_frame 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   4245
      Begin VB.CommandButton cust_save 
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
         Picture         =   "customer.frx":12AA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "SAVE"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   820
      End
      Begin VB.CommandButton cust_help 
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
         Left            =   3345
         Picture         =   "customer.frx":1914
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "HELP"
         Top             =   120
         Width           =   820
      End
      Begin VB.CommandButton cust_query 
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
         Left            =   2505
         Picture         =   "customer.frx":1A16
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "QUERY"
         Top             =   120
         Width           =   820
      End
      Begin VB.CommandButton cust_undo 
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
         Left            =   1665
         Picture         =   "customer.frx":1B18
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "UNDO"
         Top             =   120
         Width           =   820
      End
   End
End
Attribute VB_Name = "customerfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim norec, ctr, recvalidated, custinsert, prodinsert As Long
Dim confirmdelete, maxprodicno, presentcontrol, prevcust_ph1, prevcust_name, newcust_ph1 As String
Dim custrsfilter As New ADODB.Recordset
Private Sub cust_add_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_city_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_email_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_instby_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_instby_KeyUp(KeyCode As Integer, Shift As Integer)
If customerfrm.ActiveControl = cust_instby Then
    customer_pchange
End If
End Sub
Private Sub cust_inston_KeyPress(KeyAscii As Integer)
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
Private Sub cust_inston_KeyUp(KeyCode As Integer, Shift As Integer)
If customerfrm.ActiveControl = cust_inston Then
    customer_pchange
End If
End Sub
Private Sub cust_inston_LostFocus()
If IsDate(cust_inston) Then
    cust_inston.Text = CDate(cust_inston)
Else
    MsgBox ("Please Enter A Valid Installation Date")
    cust_inston.SetFocus
End If
End Sub
Private Sub cust_landmark_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_ph1find_GotFocus()
prevcust_ph1 = cust_ph1
End Sub
Private Sub cust_ph1find_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 18 Or KeyCode = 81 Then
    Exit Sub
End If

prodinstfrm.custrs.Find "cust_ph1='" & cust_ph1find & "'"
If KeyCode = vbKeyEscape Then
  cust_ph1.Text = prevcust_ph1
  custShow
  Else
  cust_ph1.Text = cust_ph1find.Text
  custShow
End If

If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
  cust_ph1find.Visible = False
  cust_ph1.Visible = True
  cust_ph1.Enabled = True
  cust_ph1.SetFocus
  
'enabling controls
cust_name.Enabled = True
cust_add.Enabled = True
cust_city.Enabled = True
cust_state.Enabled = True
cust_pin.Enabled = True
cust_ph1.Enabled = True
cust_ph2.Enabled = True
cust_fax.Enabled = True
cust_email.Enabled = True
cust_landmark.Enabled = True
cust_prodno.Enabled = True
cust_prodname.Enabled = True
cust_inston.Enabled = True
cust_instby.Enabled = True
cust_prodicno.Enabled = True
cust_save.Enabled = True
cust_query.Enabled = True
cust_delete.Enabled = True
For i = 0 To 3
    navigaterec(i).Enabled = True
Next
End If
End Sub
Private Sub cust_fax_KeyPress(KeyAscii As Integer)
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
Private Sub cust_name_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_ph1_KeyPress(KeyAscii As Integer)
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
Private Sub cust_ph2_KeyPress(KeyAscii As Integer)
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
Private Sub cust_ph2_LostFocus()
If cust_ph2 = "" Then
    cust_ph2 = "0"
End If
End Sub
Private Sub cust_pin_KeyPress(KeyAscii As Integer)
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
Private Sub cust_prodicno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_prodicno_KeyUp(KeyCode As Integer, Shift As Integer)
detailshow
End Sub
Private Sub cust_prodicno_LostFocus()
'checking installation card no for duplicacy
  prodinstfrm.setprodcom
  If prodinstfrm.prodrs.RecordCount > 0 Then
     prodinstfrm.prodrs.MoveFirst
    Do While Not prodinstfrm.prodrs.EOF
        If prodinstfrm.prodrs.Fields(0) <> cust_ph1 Then
            If prodinstfrm.prodrs.Fields(5) = cust_prodicno Then
                MsgBox ("This Installation Card No Already Exists")
                cust_prodicno.SetFocus
                Exit Sub
            End If
        End If
        prodinstfrm.prodrs.MoveNext
    Loop
  End If
End Sub
Private Sub cust_prodname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_prodname_KeyUp(KeyCode As Integer, Shift As Integer)
If customerfrm.ActiveControl = cust_prodname Then
    customer_pchange
End If
End Sub
Private Sub cust_prodno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_prodno_KeyUp(KeyCode As Integer, Shift As Integer)
If customerfrm.ActiveControl = cust_prodno Then
    customer_pchange
End If
End Sub
Private Sub cust_state_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Private Sub cust_ph1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  SendKeys "{TAB}", True
  Exit Sub
End If
If KeyCode = 18 Or KeyCode = 83 Then
    Exit Sub
End If
custShow
End Sub
Private Sub cust_ph1find_Click()
prodinstfrm.custrs.Requery
prodinstfrm.custrs.Find "cust_ph1='" & cust_ph1find & "'"
If prodinstfrm.custrs.EOF = False And prodinstfrm.custrs.BOF = False Then
    showrecords
End If
End Sub
Private Sub Form_Activate()
prodinstfrm.setcustcom
prodinstfrm.setprodtemp
prodinstfrm.setprodcom
Set custrsfilter = prodinstfrm.custrs

i = 0
custinsert = 0
prodinsert = 0
ctr = 0
recvalidated = 0
norec = 0
maxprodicno = "01a"

cust_ph1.SetFocus

'moving to first record
If prodinstfrm.custrs.RecordCount > 0 Then
  cust_undo.Enabled = False
  findrecords
Else
  custinsert = 1
  cust_delete.Enabled = False
  cust_undo.Enabled = False
  cust_query.Enabled = False
  For i = 0 To 3
    navigaterec.Item(i).Enabled = False
  Next
  blankrecords
End If

If newcomplaint = 1 Then
    cust_ph1.Text = complaintsfrm.comp_custph1.Text
    curcustph1 = Trim(cust_ph1.Text)
    cust_ph1.SetFocus
    custShow
    Exit Sub
End If
End Sub
Public Sub showrecords()
  cust_name.Text = prodinstfrm.custrs.Fields(0)
  cust_add.Text = prodinstfrm.custrs.Fields(1)
  cust_city.Text = prodinstfrm.custrs.Fields(2)
  cust_state.Text = prodinstfrm.custrs.Fields(3)
  cust_pin.Text = prodinstfrm.custrs.Fields(4)
  cust_ph2.Text = prodinstfrm.custrs.Fields(6)
  cust_fax.Text = prodinstfrm.custrs.Fields(7)
  cust_email.Text = prodinstfrm.custrs.Fields(8)
  cust_landmark.Text = prodinstfrm.custrs.Fields(9)
  prevcust_ph1 = cust_ph1
End Sub
Public Sub blankrecords()
cust_name = ""
cust_add.Text = ""
cust_city.Text = ""
cust_state.Text = ""
cust_pin.Text = "0"
cust_ph2.Text = "0"
cust_fax.Text = "0"
cust_email.Text = ""
cust_landmark.Text = ""
cust_prodno.Text = ""
cust_prodname.Text = ""
cust_inston.Text = ""
cust_instby.Text = ""
cust_prodicno.Text = ""
End Sub
Private Sub cust_delete_Click()
confirmdelete = MsgBox("Are you Sure You Want To Delete This Record", vbYesNo)
If confirmdelete = vbYes Then
    prodinstfrm.custcaredb.Execute ("delete * from customer where cust_ph1 ='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("delete * from product where prod_custph1 ='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("delete * from producttemp")
    prodinstfrm.setcustcom
    If prodinstfrm.custrs.RecordCount > 0 Then
        prodinstfrm.custrs.MoveLast
        cust_ph1.SetFocus
        cust_ph1 = prodinstfrm.custrs.Fields(5)
        custShow
    Else
        norec = 1
        cust_ph1 = "0"
        cust_ph1.SetFocus
        cust_delete.Enabled = False
        cust_query.Enabled = False
        cust_undo.Enabled = False
        For i = 0 To 3
            navigaterec.Item(i).Enabled = False
       Next
       blankrecords
       mastershow
    End If
    Else
    cust_ph1.SetFocus
End If
End Sub
Private Sub cust_query_Click()
ctr = 1
cust_save = True
cust_ph1find.Clear
If recvalidated = 0 Then
  Exit Sub
  Else
  recvalidated = 0
End If

curcust_ph1 = cust_ph1.Text

prodinstfrm.custrs.Requery
If prodinstfrm.custrs.RecordCount = 0 Then
  MsgBox ("The Database Is Empty.No Records To Find")
  Exit Sub
End If

cust_ph1.Visible = False
cust_ph1find.Visible = True
cust_ph1find.SetFocus

'disabling controls
cust_name.Enabled = False
cust_add.Enabled = False
cust_city.Enabled = False
cust_state.Enabled = False
cust_pin.Enabled = False
cust_ph1.Enabled = False
cust_ph2.Enabled = False
cust_fax.Enabled = False
cust_email.Enabled = False
cust_landmark.Enabled = False
cust_prodno.Enabled = False
cust_prodname.Enabled = False
cust_inston.Enabled = False
cust_instby.Enabled = False
cust_prodicno.Enabled = False
cust_undo.Enabled = False
cust_save.Enabled = False
cust_query.Enabled = False
cust_delete.Enabled = False
For i = 0 To 3
    navigaterec.Item(i).Enabled = False
Next

'adding phone nos to list
prodinstfrm.custrs.MoveFirst
Do While Not prodinstfrm.custrs.EOF
  cust_ph1find.AddItem prodinstfrm.custrs.Fields(5)
  prodinstfrm.custrs.MoveNext
Loop
cust_ph1find.ListIndex = 0
End Sub
Private Sub cust_save_Click()
  'validating data
  If cust_ph2.Text = "" Then
    cust_ph2 = "0"
  End If
  If cust_fax.Text = "" Then
    cust_fax = "0"
  End If
  If cust_ph1.Text = "" Or cust_ph1 = "0" Then
    MsgBox ("Phone1 Cannot Be empty")
    cust_ph1.SetFocus
    Exit Sub
  End If
  If cust_name.Text = "" Then
    MsgBox ("customer Name Cannot Be empty")
    cust_name.SetFocus
    Exit Sub
  End If
  If cust_add.Text = "" Then
    MsgBox ("Address Cannot Be empty")
    cust_add.SetFocus
    Exit Sub
  End If
  If cust_state.Text = "" Then
    MsgBox ("State Cannot Be empty")
    cust_state.SetFocus
    Exit Sub
  End If
  If cust_city.Text = "" Then
    MsgBox ("City Name Cannot Be empty")
    cust_city.SetFocus
    Exit Sub
  End If
  If cust_pin.Text = "" Or cust_pin = "0" Then
    MsgBox ("Pincode Cannot Be empty")
    cust_pin.SetFocus
    Exit Sub
  End If
  If cust_prodno.Text <> "" Then
    If cust_prodname.Text = "" Then
        MsgBox ("Please Enter Something In Product Name")
        cust_prodname.SetFocus
        Exit Sub
    ElseIf Not IsDate(cust_inston) Then
        MsgBox ("Please Enter A Valid Installation date")
        cust_inston.SetFocus
        Exit Sub
    ElseIf cust_instby = "" Then
        MsgBox ("Please Enter Something In Installed By")
        cust_instby.SetFocus
        Exit Sub
    ElseIf cust_prodicno = "" Or cust_prodicno = "0" Then
        MsgBox ("Please Enter Something In Installation Card No")
        cust_prodicno.SetFocus
        Exit Sub
    End If
 End If
    
  'inserting and updating data to customer table
  If custinsert = 1 Then
    prodinstfrm.custcaredb.Execute ("insert into customer values ('" & cust_name & "','" & cust_add & "','" & cust_city & "','" & cust_state & "'," & cust_pin & ",'" & cust_ph1 & "'," & cust_ph2 & "," & cust_fax & ",'" & cust_email & "','" & cust_landmark & "')")
    custinsert = 0
  Else
    prodinstfrm.custcaredb.Execute ("update customer set cust_name='" & cust_name.Text & "' where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_add='" & cust_add.Text & "' where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_city='" & cust_city.Text & "' where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_state='" & cust_state.Text & "' where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_pin=" & cust_pin & " where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_ph2=" & cust_ph2 & " where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_fax=" & cust_fax & " where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_email='" & cust_email & "' where customer.cust_ph1='" & cust_ph1 & "'")
    prodinstfrm.custcaredb.Execute ("update customer set cust_landmark='" & cust_landmark & "' where customer.cust_ph1='" & cust_ph1 & "'")
 End If
  
 'transfer producttemp data to product table
  prodinstfrm.custcaredb.Execute ("delete * from product where prod_custph1='" & cust_ph1.Text & "'")
  prodinstfrm.custcaredb.Execute ("insert into product select * from producttemp where prod_custph1='" & cust_ph1.Text & "'")
    
 'transfer product data to producttemp table
  prodinstfrm.custcaredb.Execute ("delete * from producttemp")
  prodinstfrm.custcaredb.Execute ("insert into producttemp select * from product where prod_custph1='" & cust_ph1.Text & "'")
   
  prodinsert = 0
  prodinstfrm.prodtemprs.Requery
  prod_data.Refresh
  
  cust_ph1.SetFocus
  recvalidated = 1
  ctr = 1
  cust_delete.Enabled = True
  cust_query.Enabled = True
  For i = 0 To 3
    navigaterec.Item(i).Enabled = True
  Next
  prodinstfrm.custrs.Requery
  norec = 0
End Sub
Private Sub cust_undo_Click()
  cust_undo.Enabled = False
  If prodinstfrm.custrs.RecordCount > 0 Then
    prodinstfrm.custrs.MoveLast
    showrecords
    cust_ph1.Text = prodinstfrm.custrs.Fields(5)
  End If
  cust_ph1.SetFocus
  custShow
End Sub
Public Sub findrecords()
If prodinstfrm.prodrs.RecordCount > 0 Then
    prodinstfrm.prodrs.MoveFirst
    Do While Not prodinstfrm.prodrs.EOF
        If prodinstfrm.prodrs.Fields(5) > maxprodicno Then
            maxprodicno = prodinstfrm.prodrs.Fields(5)
        End If
        If Not prodinstfrm.prodrs.EOF Then
            newcust_ph1 = prodinstfrm.prodrs.Fields(0)
        End If
        prodinstfrm.prodrs.MoveNext
    Loop
 End If
  
  prodinstfrm.custrs.MoveFirst
  
  If newcust_ph1 <> "" Then
    Do While Not prodinstfrm.custrs.EOF
        If prodinstfrm.custrs.Fields(5) = newcust_ph1 Then
            Exit Do
        End If
        prodinstfrm.custrs.MoveNext
    Loop
  End If
  
    cust_ph1.Text = prodinstfrm.custrs.Fields(5)
    showrecords
    mastershow
End Sub
Public Sub custShow()
If cust_ph1.Text = "" Or cust_ph1.Text = "0" Then
  blankrecords
  cust_delete.Enabled = False
  cust_query.Enabled = False
  Exit Sub
End If
prodinstfrm.setcustcom
If prodinstfrm.custrs.RecordCount > 0 Then
    prodinstfrm.custrs.Find "cust_ph1='" & cust_ph1.Text & "'"
    If prodinstfrm.custrs.EOF = False And prodinstfrm.custrs.BOF = False Then
        custinsert = 0
        showrecords
        mastershow
        If cust_ph1find.Visible = False Then
            cust_delete.Enabled = True
            cust_query.Enabled = True
            cust_undo.Enabled = False
            For i = 0 To 3
                navigaterec(i).Enabled = True
            Next
        End If
    Else
        custinsert = 1
        blankrecords
        cust_delete.Enabled = False
        cust_query.Enabled = False
        cust_undo.Enabled = True
        For i = 0 To 3
            navigaterec(i).Enabled = False
        Next
        prodinstfrm.custcaredb.Execute ("delete * from producttemp")
        prod_data.Refresh
    End If
    Else
        custinsert = 1
        prodinstfrm.custcaredb.Execute ("delete * from producttemp")
        prod_data.Refresh
End If
End Sub

Private Sub navigaterec_Click(Index As Integer)
ctr = 1
If prodinstfrm.custrs.RecordCount > 0 Then
        cust_save = True
        If recvalidated = 0 Then
            Exit Sub
        Else
            recvalidated = 0
        End If
        cust_ph1.SetFocus
  
        If Index = 0 Then
            prodinstfrm.custrs.MoveFirst
        ElseIf Index = 1 Then
            prodinstfrm.custrs.Requery
            prodinstfrm.custrs.Find "cust_ph1='" & cust_ph1 & "'"
            prodinstfrm.custrs.MovePrevious
            If prodinstfrm.custrs.BOF Then
                 prodinstfrm.custrs.MoveNext
            End If
              
  ElseIf Index = 2 Then
        prodinstfrm.custrs.Requery
        prodinstfrm.custrs.Find "cust_ph1='" & cust_ph1 & "'"
        prodinstfrm.custrs.MoveNext
    If prodinstfrm.custrs.EOF Then
        prodinstfrm.custrs.MovePrevious
    End If
  ElseIf Index = 3 Then
    prodinstfrm.custrs.MoveLast
  End If
  End If
  'End If
  cust_ph1.Text = prodinstfrm.custrs.Fields(5)
  custShow
'End If
End Sub
Public Sub mastershow()
prodinstfrm.custcaredb.Execute ("delete * from producttemp")
prodinstfrm.setprodcom
prodinstfrm.prodrs.Filter = "prod_custph1='" & cust_ph1 & "'"
If prodinstfrm.prodrs.EOF = False And prodinstfrm.prodrs.BOF = False Then
    prodinsert = 0
    cust_undo.Enabled = False
    cust_delete.Enabled = True
    cust_query.Enabled = True
    For i = 0 To 3
        navigaterec(i).Enabled = True
    Next
   prodinstfrm.custcaredb.Execute ("Insert into producttemp select * from product where prod_custph1='" & cust_ph1 & "'")
   
   prodinstfrm.setprodtemp
   If prodinstfrm.prodtemprs.RecordCount > 0 Then
        prodinstfrm.prodtemprs.MoveFirst
        cust_prodicno = prodinstfrm.prodtemprs.Fields(5)
        showrecords1
        prod_data.Refresh
    End If
Else:
    prodinstfrm.custcaredb.Execute ("delete * from producttemp")
    prodinsert = 1
    cust_prodicno = ""
    cust_prodno = ""
    cust_prodname = ""
    cust_instby = ""
    cust_inston = ""
    If norec = 1 Then
        cust_undo.Enabled = True
        cust_delete.Enabled = False
        cust_query.Enabled = False
        For i = 0 To 3
            navigaterec(i).Enabled = False
        Next
    End If
    prod_data.Refresh
End If
End Sub
Public Sub showrecords1()
If prodinstfrm.prodtemprs.RecordCount > 0 Then
    cust_prodno = prodinstfrm.prodtemprs.Fields(1)
    cust_prodname = prodinstfrm.prodtemprs.Fields(2)
    If prodinstfrm.prodtemprs.Fields(3) <> "" Then
        cust_inston.Text = prodinstfrm.prodtemprs.Fields(3)
    End If
    cust_instby = prodinstfrm.prodtemprs.Fields(4)
End If
End Sub
Public Sub customer_pchange()
If Trim(cust_ph1.Text) <> "" And cust_prodno.Text <> "" And cust_prodname.Text <> "" And IsDate(cust_inston.Text) And cust_instby.Text <> "" Then
    If prodinsert = 1 Then
        prodinstfrm.custcaredb.Execute ("insert into producttemp values ('" & cust_ph1 & "','" & cust_prodno & "','" & cust_prodname & "','" & cust_inston & "','" & cust_instby & "','" & cust_prodicno & "')")
        prodinsert = 0
    Else
        prodinstfrm.custcaredb.Execute ("update producttemp set prod_no='" & cust_prodno.Text & "' where producttemp.prod_custph1='" & cust_ph1 & "' and producttemp.prod_icno='" & cust_prodicno & "'")
        prodinstfrm.custcaredb.Execute ("update producttemp set prod_name='" & cust_prodname.Text & "' where producttemp.prod_custph1='" & cust_ph1 & "' and producttemp.prod_icno='" & cust_prodicno & "'")
        prodinstfrm.custcaredb.Execute ("update producttemp set prod_inston='" & cust_inston.Text & "' where producttemp.prod_custph1='" & cust_ph1 & "' and producttemp.prod_icno='" & cust_prodicno & "'")
        prodinstfrm.custcaredb.Execute ("update producttemp set prod_instby='" & cust_instby.Text & "' where producttemp.prod_custph1='" & cust_ph1 & "' and producttemp.prod_icno='" & cust_prodicno & "'")
        prodinstfrm.prodtemprs.Requery
    End If
End If
End Sub
Public Sub detailshow()
If cust_ph1 <> "" And cust_prodicno <> "" Then
    prodinstfrm.setprodtemp
    prodinstfrm.prodtemprs.Filter = "prod_custph1='" & cust_ph1.Text & "' and prod_icno='" & cust_prodicno.Text & "'"
    If prodinstfrm.prodtemprs.EOF = False And prodinstfrm.prodtemprs.BOF = False Then
        prodinsert = 0
        showrecords1
        Else
        prodinsert = 1
        cust_prodno = ""
        cust_prodname = ""
        cust_inston = ""
        cust_instby = ""
    End If
End If
End Sub

Private Sub prod_delete_Click()
prodinstfrm.custcaredb.Execute ("delete * from producttemp where prod_custph1='" & cust_ph1 & "' and prod_no='" & cust_prodno & "'")
prodinstfrm.setprodtemp
If prodinstfrm.prodtemprs.RecordCount > 0 Then
    prodinstfrm.prodtemprs.MoveNext
    If prodinstfrm.prodtemprs.EOF Then
        prodinstfrm.prodtemprs.MovePrevious
    End If
    cust_prodicno.Text = prodinstfrm.prodtemprs.Fields(5)
    cust_prodno.Text = prodinstfrm.prodtemprs.Fields(1)
    prod_data.Refresh
    showrecords1
    Else
    cust_prodno.Text = ""
    cust_prodname.Text = ""
    cust_inston.Text = ""
    cust_instby.Text = ""
    cust_prodicno.Text = ""
    prod_data.Refresh
End If
End Sub
Private Sub prod_grid_Click()
gridchange
End Sub
Private Sub prod_grid_KeyUp(KeyCode As Integer, Shift As Integer)
gridchange
End Sub
Public Sub gridchange()
If prod_grid.ApproxCount > 0 Then
    If prod_grid.Columns(0) <> "" Then
        prodinstfrm.prodtemprs.Requery
        prodinstfrm.prodtemprs.Filter = "prod_custph1='" & cust_ph1 & "' and prod_icno='" & prod_grid.Columns(5) & "'"
        If prodinstfrm.prodtemprs.EOF = False And prodinstfrm.prodtemprs.BOF = False Then
            cust_prodicno = prod_grid.Columns(5)
            showrecords1
        End If
    End If
End If
End Sub


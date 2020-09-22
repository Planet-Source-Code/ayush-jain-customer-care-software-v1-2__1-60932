VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.tlb"
Begin VB.Form comp_pendingfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Pending Complaints Form"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   8295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport comp_pendingrpt 
      Left            =   5640
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ReportFileName  =   "c:\program files\custcareprj\COMP_PENDINGRPT.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1965
      Left            =   780
      TabIndex        =   3
      Top             =   420
      Width           =   3585
      Begin VB.TextBox to_compldate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox from_compldate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Top             =   300
         Width           =   1515
      End
      Begin VB.CommandButton compl_pendingshow 
         Caption         =   "&Show Complaints"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1080
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "To Date:"
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
         Left            =   360
         TabIndex        =   5
         Top             =   690
         Width           =   915
      End
      Begin VB.Label lblfrom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "From Date:"
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
         Left            =   360
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "comp_pendingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim custcaredb As New Connection
Dim compcom, comptempcom As New Command
Dim comprs, comptemprs As New Recordset

Private Sub Form_Activate()
'Set custcaredb = DBEngine.OpenDatabase("c:\program files\custcareprj\custcaredb.mdb")
'Set prodinstfrm.custrs = custcaredb.OpenRecordset("select * from customer ", dbOpenDynaset)
'Set comprs = custcaredb.OpenRecordset("select * from complaints ", dbOpenDynaset)
'Set comptemprs = custcaredb.OpenRecordset("select * from complaintstemp ", dbOpenDynaset)

'custcaredb.Execute ("delete * from complaintstemp")
'comprs.Requery
'If comprs.RecordCount > 0 Then
'    comprs.MoveFirst
'    Do While Not comprs.EOF
'        If comprs.Fields(7) = "Pending" Then
'            comptemprs.AddNew
'            comptemprs.Fields(0) = comprs.Fields(0)
'            comptemprs.Fields(1) = comprs.Fields(1)
'            comptemprs.Fields(2) = comprs.Fields(2)
'            comptemprs.Fields(9) = comprs.Fields(9)
'            comptemprs.Fields(10) = comprs.Fields(10)
'            comptemprs.Update
'        End If
'        comprs.MoveNext
'    Loop
'End If
comp_pendingrpt.DiscardSavedData = True
comp_pendingrpt.Action = 1
End Sub

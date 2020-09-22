VERSION 5.00
Begin VB.Form custcaresplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ProgressBar1 
      DragIcon        =   "custcaresplash.frx":0000
      DragMode        =   1  'Automatic
      Height          =   240
      Left            =   1050
      ScaleHeight     =   180
      ScaleWidth      =   1950
      TabIndex        =   7
      Top             =   1815
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3600
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2820
         Top             =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Loading"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label title 
         AutoSize        =   -1  'True
         Caption         =   "Customer Care System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Licensed To :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   4
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label companylbl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Oyster Bath Concepts Pvt Ltd"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   495
         TabIndex        =   3
         Top             =   1020
         Width           =   2325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dealers of : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1110
         TabIndex        =   2
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sanitaryware"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1020
         TabIndex        =   1
         Top             =   1500
         Width           =   915
      End
   End
End
Attribute VB_Name = "custcaresplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Timer1_Timer()
'ProgressBar1.Position = 1
'Do While Not ProgressBar1.Position = 100
'    ProgressBar1.Position = ProgressBar1.Position + 1
'    For i = 0 To 30000
'   Next
'Loop
Unload Me
Load prodinstfrm
prodinstfrm.Left = 0
prodinstfrm.Top = 0
prodinstfrm.Height = 0
prodinstfrm.Width = 0
prodinstfrm.Show

Load prodinst2frm
prodinst2frm.Left = 0
prodinst2frm.Top = 0
prodinst2frm.Height = 0
prodinst2frm.Width = 0
prodinst2frm.Show

Load custcaremdi
custcaremdi.Show
End Sub

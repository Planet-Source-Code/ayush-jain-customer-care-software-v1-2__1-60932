VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.tlb"
Begin VB.Form prodinst1frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Installation Form"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox no_days 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   555
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton Show 
         Caption         =   "&Show"
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Of Days:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport prodinstdlg 
      Left            =   3480
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "prodinst1frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub no_days_KeyPress(KeyAscii As Integer)
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
Private Sub Show_Click()
If no_days.Text = "" Then
    MsgBox ("Please Enter Something In 'No Of Days'")
    no_days.SetFocus
    Exit Sub
End If
prodinstdlg.DiscardSavedData = True
prodinstdlg.SelectionFormula = "Today - {product.prod_inston}+1>=" & no_days
prodinstdlg.Action = 1
End Sub

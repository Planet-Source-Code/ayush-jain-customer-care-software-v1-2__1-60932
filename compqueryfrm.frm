VERSION 5.00
Begin VB.Form compqueryfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Complaints Query Form"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   5805
   Begin VB.Frame Fieldframe 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   570
      TabIndex        =   3
      Top             =   330
      Width           =   6405
      Begin VB.ComboBox value 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   900
         Width           =   4395
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "&Goto This Record"
         Height          =   375
         Left            =   1470
         Picture         =   "compqueryfrm.frx":0000
         TabIndex        =   2
         Top             =   1470
         Width           =   1665
      End
      Begin VB.ComboBox field 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "compqueryfrm.frx":0442
         Left            =   1740
         List            =   "compqueryfrm.frx":0458
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblvalues 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Value:"
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
         Left            =   450
         TabIndex        =   5
         Top             =   870
         Width           =   825
      End
      Begin VB.Label lblfield 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search By:"
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
         Left            =   450
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "compqueryfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsearch_Click()
If prodinstfrm.custrs.RecordCount > 0 Then
    prodinstfrm.setcustcom
        If field.Text = "Phone1" Then
            prodinstfrm.custrs.Filter = "cust_ph1='" & value.Text & "'"
        End If
        If field.Text = "Phone2" Then
            prodinstfrm.custrs.Filter = "cust_ph2=" & value.Text
        End If
        If field.Text = "Address" Then
            prodinstfrm.custrs.Filter = "cust_add='" & value.Text & "'"
        End If
        If field.Text = "Email" Then
            prodinstfrm.custrs.Filter = "cust_email='" & value.Text & "'"
        End If
        If field.Text = "Fax" Then
            prodinstfrm.custrs.Filter = "cust_fax=" & value.Text
        End If
        If field.Text = "Complaint No" Then
            prodinstfrm.setcompcom
            prodinstfrm.comprs.Filter = "comp_no=" & value.Text
            If prodinstfrm.comprs.EOF = False And prodinstfrm.comprs.BOF = False Then
                curcustph1 = prodinstfrm.comprs.Fields(0)
                curcustdate = prodinstfrm.comprs.Fields(1)
                Unload compqueryfrm
                newcomplaint = 2
                Exit Sub
            Else
                MsgBox ("Record Not Found")
                value.SetFocus
            End If
        End If
        
        If prodinstfrm.custrs.BOF = False And prodinstfrm.custrs.EOF = False Then
            curcustph1 = prodinstfrm.custrs.Fields(5)
            Unload compqueryfrm
            Else
            MsgBox ("Record Not Found")
            value.SetFocus
        End If
        newcomplaint = 1
End If
End Sub
Private Sub field_Click()
fillvalues
End Sub
Private Sub field_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
KeyAscii = 0
End Sub
Private Sub field_KeyUp(KeyCode As Integer, Shift As Integer)
fillvalues
End Sub
Private Sub Form_Activate()
prodinstfrm.setcustcom
prodinstfrm.setcompcom
field.Text = field.List(0)
If prodinstfrm.custrs.RecordCount > 0 Then
    prodinstfrm.custrs.MoveFirst
    Do While Not prodinstfrm.custrs.EOF
        value.AddItem (prodinstfrm.custrs.Fields(5))
        prodinstfrm.custrs.MoveNext
    Loop
End If
value.Text = value.List(0)
End Sub
Private Sub value_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}", True
End If
End Sub
Public Sub fillvalues()
If prodinstfrm.custrs.RecordCount = 0 Then
    Exit Sub
    Else
    value.Clear
End If

prodinstfrm.custrs.MoveFirst
Do While Not prodinstfrm.custrs.EOF
    If field.Text = "Phone1" Then
        value.AddItem prodinstfrm.custrs.Fields(5)
    End If
    If field.Text = "Phone2" Then
        value.AddItem prodinstfrm.custrs.Fields(6)
    End If
    If field.Text = "Address" Then
        value.AddItem prodinstfrm.custrs.Fields(1)
    End If
    If field.Text = "Email" Then
        value.AddItem prodinstfrm.custrs.Fields(8)
    End If
    If field.Text = "Fax" Then
        value.AddItem prodinstfrm.custrs.Fields(7)
    End If
    prodinstfrm.custrs.MoveNext
Loop
If field.Text = "Complaint No" Then
    If prodinstfrm.comprs.RecordCount > 0 Then
        prodinstfrm.comprs.MoveFirst
        Do While Not prodinstfrm.comprs.EOF
            value.AddItem prodinstfrm.comprs.Fields(8)
            prodinstfrm.comprs.MoveNext
        Loop
    End If
End If
value.Text = value.List(0)
End Sub

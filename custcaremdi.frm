VERSION 5.00
Begin VB.MDIForm custcaremdi 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Customer Care System For Oysters"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "custcaremdi.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "custcaremdi.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu customer 
      Caption         =   "Customer"
   End
   Begin VB.Menu complaints 
      Caption         =   "Complaints"
   End
   Begin VB.Menu reports 
      Caption         =   "Reports"
      Begin VB.Menu comp_pending 
         Caption         =   "Pending Complaints"
      End
      Begin VB.Menu prodinst 
         Caption         =   "Product Installation Report"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "custcaremdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim custcaredb As New ADODB.Connection
Dim prodcom As New Command
Dim prodrs As New Recordset
Private Sub comp_pending_Click()
Unload complaintsfrm
Unload customerfrm
Unload comp_pendingfrm
Unload prodinstfrm
Unload prodinst1frm
Unload prodinst2frm
Load comp_pendingfrm
comp_pendingfrm.Show
comp_pendingfrm.Left = 0
comp_pendingfrm.Top = 0
comp_pendingfrm.Height = 0
comp_pendingfrm.Width = 0
End Sub
Private Sub complaints_Click()
complaintsfrm.Left = 0
complaintsfrm.Top = 0
complaintsfrm.Width = 9800
complaintsfrm.Height = 12000
complaintsfrm.Show
Unload customerfrm
Unload prodinstfrm
Unload prodinst1frm
Unload prodinst2frm
Unload comp_pendingfrm
End Sub
Private Sub customer_Click()
customerfrm.Left = 0
customerfrm.Top = 0
customerfrm.Width = 9800
customerfrm.Height = 9000
customerfrm.Show
Unload complaintsfrm
Unload prodinstfrm
Unload prodinst1frm
Unload prodinst2frm
End Sub
Private Sub exit_Click()
End
End Sub
'Private Sub retrievedata_Click()
'DBEngine.CompactDatabase "a:\custcaredb.mdb", "c:/program files/custcareprj/app.path/custcaredb,dbLangGeneral"
'End Sub
'Private Sub save_Click()
'DBEngine.CompactDatabase "c:/program files/custcareprj/app.path/custcaredb.mdb", "a:\custcaredb"
'End Sub
Private Sub three_Click()
Unload complaintsfrm
Unload customerfrm
Unload comp_pendingfrm
Unload prodinstfrm
Unload prodinst1frm
Unload prodinst2frm
Unload prodinst3frm

Load prodinstfrm
prodinst1frm.Left = 0
prodinst1frm.Top = 0
prodinst1frm.Height = 0
prodinst1frm.Width = 0
prodinst1frm.Show
End Sub
Private Sub two_Click()
Unload complaintsfrm
Unload customerfrm
Unload comp_pendingfrm
Unload prodinstfrm
Unload prodinst1frm
Unload prodinst2frm
Unload prodinst3frm

Load prodinst3frm
prodinst3frm.Left = 0
prodinst3frm.Top = 0
prodinst3frm.Height = 0
prodinst3frm.Width = 0
prodinst3frm.Show
End Sub

Private Sub prodinst_Click()
Unload complaintsfrm
Unload customerfrm
Unload comp_pendingfrm
Unload prodinstfrm
Unload prodinst1frm
Unload prodinst2frm
Load prodinst1frm
prodinst1frm.Left = 3000
prodinst1frm.Top = 2000
prodinst1frm.Height = 2500
prodinst1frm.Width = 4300
prodinst1frm.Show
End Sub

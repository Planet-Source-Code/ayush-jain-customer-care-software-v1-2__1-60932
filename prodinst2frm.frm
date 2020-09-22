VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.tlb"
Begin VB.Form prodinst2frm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport prodinstdlg 
      Left            =   960
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "prodinst2frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
If prodinstfrm.prodrs.RecordCount > 0 Then
    Do While Not prodinstfrm.prodrs.EOF
        If Date - prodrs.Fields(3) = 732 Then
           Exit Do
        End If
        prodrs.MoveNext
    Loop
    Else
    Exit Sub
End If
If prodinstfrm.prodrs.EOF Then
    Exit Sub
End If
prodinstdlg.DiscardSavedData = True
prodinstdlg.SelectionFormula = "Today - {product.prod_inston}+1=730"
prodinstdlg.Action = 1
End Sub


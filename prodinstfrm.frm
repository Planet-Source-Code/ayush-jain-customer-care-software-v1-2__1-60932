VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.tlb"
Begin VB.Form prodinstfrm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Product Installation Form"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport prodinstdlg 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ReportFileName  =   "c:\program files\custcareprj\prodinstrpt.rpt"
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "prodinstfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public custcaredb As New ADODB.Connection
Public prodcom As New Command
Public custcom As New Command
Public prodtempcom As New Command
Public compcom As New Command
Public comptempcom As New Command
Public prodrs As New Recordset
Public custrs As New Recordset
Public comprs As New Recordset
Public comptemprs As New Recordset
Public prodtemprs As New Recordset
Private Sub Form_Activate()
'custcaredb.Open "Driver={Microsoft Access Driver (*.mdb)};" & "Dbq=\\oysterbath\c\custcare\custcaredb.mdb;" & "Uid=;" & "Pwd="
custcaredb.Open "Driver={Microsoft Access Driver (*.mdb)};" & "Dbq=c:\program files\custcareprj\custcaredb.mdb;" & "Uid=;" & "Pwd="
setprodcom
If prodrs.RecordCount > 0 Then
    Do While Not prodrs.EOF
        If Date - prodrs.Fields(3) = 92 Then
           Exit Do
        End If
        prodrs.MoveNext
    Loop
    Else
    Exit Sub
End If
If prodrs.EOF Then
    Exit Sub
End If
prodinstdlg.DiscardSavedData = True
prodinstdlg.SelectionFormula = "Today - {product.prod_inston}+1=90"
prodinstdlg.Action = 1
End Sub
Public Sub setcustcom()
With custcom
    .ActiveConnection = custcaredb
    .CommandText = "customer"
    .CommandType = adCmdTable
End With

Set custrs = custcom.Execute
With custrs
    .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .Open
End With
End Sub
Public Sub setprodcom()
With prodcom
   .ActiveConnection = custcaredb
   .CommandText = "product"
   .CommandType = adCmdTable
End With

Set prodrs = prodcom.Execute

With prodrs
    .Close
    .CursorLocation = adUseClient
    .Open
End With
End Sub
Public Sub setprodtemp()
With prodtempcom
        .ActiveConnection = custcaredb
        .CommandText = "producttemp"
        .CommandType = adCmdTable
End With

Set prodtemprs = prodtempcom.Execute

With prodtemprs
    .Close
    .CursorLocation = adUseClient
    .Open
End With
End Sub
Public Sub setcompcom()
With compcom
   .ActiveConnection = custcaredb
   .CommandText = "complaints"
   .CommandType = adCmdTable
End With
Set comprs = compcom.Execute
comprs.Close
comprs.CursorLocation = adUseClient
comprs.Open
End Sub
Public Sub setcomptempcom()
With comptempcom
   .ActiveConnection = custcaredb
   .CommandText = "complaintstemp"
   .CommandType = adCmdTable
End With
Set comptemprs = comptempcom.Execute
comptemprs.Close
comptemprs.CursorLocation = adUseClient
comptemprs.Open
End Sub

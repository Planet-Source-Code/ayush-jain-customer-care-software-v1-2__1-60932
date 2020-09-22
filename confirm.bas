Attribute VB_Name = "Module1"
Public Sub compconfirmcustph1add()
confirm = MsgBox("This Phone No Deos Not Exist In The Books.Do You Want To Add It To The List", vbYesNo)
If confirm = vbYes Then
    curcustph1 = complaintsfrm.comp_custph1.Text
    newcomplaint = 1
    Load customerfrm
    customerfrm.Left = 0
    customerfrm.Top = 0
    customerfrm.Width = 9800
    customerfrm.Height = 9000
    customerfrm.Show
    Exit Sub
Else
    complaintsfrm.comp_custph1.SetFocus
End If
End Sub

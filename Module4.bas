Attribute VB_Name = "Module4"
Dim custcaredb As Database
Dim custrs As Recordset
Dim comprs As Recordset
Public Sub globset()
Set custcaredb = DBEngine.OpenDatabase(App.Path & "\custcaredb.mdb")
Set custrs = custcaredb.OpenRecordset("select * from customer ", dbOpenDynaset)
Set complaintsrs = custcaredb.OpenRecordset("select * from complaints ", dbOpenDynaset)
End Sub


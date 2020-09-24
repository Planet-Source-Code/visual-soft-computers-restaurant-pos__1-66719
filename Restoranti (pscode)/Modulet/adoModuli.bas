Attribute VB_Name = "ModDB"
Public ac As New ADODB.Connection
Public ar As New ADODB.Recordset
Public CurrentForm As Form
Public strConek, pword, CurrentUser As String
Public rc, ctr, passFlag, liCtr, dbFlag, menuFlag, saveFlag As Integer
Public Function dblidhja()
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
strConek = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\dbaza.mdb;Persist " & _
"Security Info=False;Jet OLEDB:Database Password=cc03bn01"
End Function
Public Function dbshitja()
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
strConek = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\dbshitja.mdb;Persist " & _
"Security Info=False;Jet OLEDB:Database Password=cc03bn01"
End Function



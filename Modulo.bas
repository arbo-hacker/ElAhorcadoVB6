Attribute VB_Name = "Modulo"
Public DB As Database
Public RS As Recordset
Public Sql As String

Public Sub AbrirBD()
Set DB = OpenDatabase(App.Path & "\ahorcado.mdb")
End Sub

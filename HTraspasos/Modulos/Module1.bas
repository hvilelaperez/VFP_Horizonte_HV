Attribute VB_Name = "Module1"
Public Con As New cConexion

Sub CargarRegistros(Rs As Recordset, tabla As String)
    Rs.CursorLocation = adUseClient
    'para trabajar en linea debe ser dinamica y con bloqueo optimista
    Rs.Open "Select * from " + tabla + " order by 1 ", Con.Cnx, adOpenDynamic, adLockOptimistic
End Sub

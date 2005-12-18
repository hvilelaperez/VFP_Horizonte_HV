VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim WithEvents cnConexion As Connection
Attribute cnConexion.VB_VarHelpID = -1

Dim strSQLServer As String
Dim strSQLUser As String
Dim strSQLUserPassword As String
Dim strSQLDataBase As String
Dim strStringConection As String
Dim strMsgError As String
Dim bolConexionEstado As Boolean
Public Property Get SQLServer() As String
    SQLServer = strSQLServer
End Property
Public Property Let SQLServer(ByVal strServer As String)
    strSQLServer = strServer
End Property
Public Property Get SQLUser() As String
    SQLUser = strSQLUser
End Property
Public Property Let SQLUser(ByVal strUser As String)
    strSQLUser = strUser
End Property
Public Property Get SQLUserPassword() As String
    SQLUserPassword = strSQLUserPassword
End Property
Public Property Let SQLUserPassword(ByVal strPassword As String)
    strSQLUserPassword = strPassword
End Property
Public Property Get SQLDataBase() As String
    SQLDataBase = strSQLDataBase
End Property
Public Property Let SQLDataBase(ByVal strDataBase As String)
    strSQLDataBase = strDataBase
End Property
Public Property Get MsgError() As String
    MsgError = strMsgError
End Property
Public Property Let MsgError(ByVal strMensaje As String)
    strMsgError = strMensaje
End Property
Public Property Get ConexionActiva() As Boolean
    'ConexionActiva = bolConexionEstado
    ConexionActiva = (cnConexion.State = 1)
End Property
Public Property Let ConexionActiva(ByVal bolActivo As Boolean)
    bolConexionEstado = bolActivo
End Property
Public Property Get Conexion() As Connection
    Set Conexion = cnConexion
End Property
Public Property Let Conexion(Cnx As Connection)
    Set cnConexion = Cnx
End Property
Public Sub Conectar()
On Error GoTo Err:
    strStringConection = "Provider=sqloledb;Data Source=" & strSQLServer & ";Initial Catalog=" & strSQLDataBase & ";User Id=" & strSQLUser & ";Password=" & strSQLUserPassword
    Set cnConexion = New Connection
    If cnConexion.State = 1 Then cnConexion.Close
    cnConexion.ConnectionString = strStringConection
    cnConexion.CursorLocation = adUseClient
    cnConexion.CommandTimeout = 1000
    cnConexion.Open
    bolConexionEstado = (cnConexion.State = 1)
    Exit Sub
Err:
    bolConexionEstado = (Err.Number = 0)
    strMsgError = Err.Description
End Sub
Public Sub Desconectar()
On Local Error Resume Next
    If Not cnConexion Is Nothing Then
        cnConexion.RollbackTrans
        If cnConexion.State = 1 Then cnConexion.Close
    End If
    bolConexionEstado = False
End Sub
Private Sub Class_Initialize()
    strSQLServer = ""
    strSQLUser = ""
    strSQLUserPassword = ""
    strSQLDataBase = ""
    strStringConection = ""
End Sub
Private Sub Class_Terminate()
'    If cnConexion.State = 1 Then cnConexion.Close
'    Set cnConexion = Nothing
End Sub

Public Sub cnConexion_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)

End Sub


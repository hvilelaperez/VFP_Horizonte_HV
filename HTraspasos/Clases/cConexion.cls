VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Cnx As ADODB.Connection
Public vbdev As Variant
Dim er As Error
Public ERCM As String


Public Function ConexionSQL(Usuario As String, _
                        Password As String, _
                        Servidor As String, _
                        BaseDatos As String) As Boolean
    On Error GoTo merror
    Set Cnx = New ADODB.Connection
    Cnx.Provider = "SQLOLEDB"
        
    Cnx.Provider = "MSDataShape.1"
    Cnx.ConnectionString = "Data Provider=SQLOLEDB ;User ID=" & Usuario & _
                           ";Password=" & LimpiarCadena(Password) & _
                          ";Data Source=" & Servidor & _
                         ";Initial Catalog=" & BaseDatos
    'Cnx.ConnectionTimeout = 30
    Cnx.CursorLocation = adUseClient
    'Cnx.Provider = "SQLOLEDB"
    'Cnx.ConnectionString = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CATASTROP;Data Source= " & Servidor & " "
    Cnx.ConnectionTimeout = 120
    Cnx.CommandTimeout = 120
    Cnx.Open
    ConexionSQL = True
    Exit Function
merror:
    ConexionSQL = False
    Select Case Err.Number
    Case -2147467259
                  MsgBox "No se encontro el Servidor Especificado  o no cuenta con el Servico SQL", 0 + 64 + 0, "Mensaje"
    Case Else
                  MsgBox Err.Description, 0 + 64 + 0, "Mensaje"
    End Select
End Function

Public Sub DesConectar()
    Cnx.Close
End Sub

Public Sub Abrir()
If Cnx.State = 0 Then Cnx.Open
End Sub

'TCursor indica que tipo de cursor es
'   0=adOpenForwardOnly
'   1=adOpenKeySet
'   2=adOpenDynamic
'   3=adOpenStatic
Public Sub AbrirTabla(StoreProcedure As String, Rsx As ADODB.Recordset, Optional TCursor As CursorTypeEnum, Optional TBloqueo As LockTypeEnum)
    On Error GoTo merror
    Set Rsx = New ADODB.Recordset
    Rsx.CursorLocation = adUseClient
    Rsx.Open StoreProcedure, Cnx, TCursor, TBloqueo
    Exit Sub
merror:
    MsgBox "'" & Err.Description & "'"
End Sub

Public Sub EjecutarSQL(StoreProcedure As String, Valores As String)
Dim cmx As Command
Set cmx = New Command
cmx.ActiveConnection = Cnx
cmx.CommandText = "" & StoreProcedure & " " & Valores & ""
cmx.Execute
End Sub

Public Function FEjecutarSQL(cmd As ADODB.Command) As Boolean
On Error GoTo TRATA_ERROR
    Dim RsEjecutarSQL As ADODB.Recordset
    Set RsEjecutarSQL = New ADODB.Recordset
    Set RsEjecutarSQL = cmd.Execute()
    Set cmd.ActiveConnection = Nothing
    If Trim(UCase(RsEjecutarSQL.Fields(0))) <> "EXITO" Then
        FEjecutarSQL = False
        Exit Function
    End If
    FEjecutarSQL = True
    RsEjecutarSQL.Close
    Set cmd.ActiveConnection = Nothing
    Set RsEjecutarSQL = Nothing
    Exit Function
TRATA_ERROR:
    If RsEjecutarSQL.State = 1 Then 'verifica si está abierto el recordset
        RsEjecutarSQL.Close
        Set RsEjecutarSQL = Nothing
    End If
    MsgBox Err.Description, 0 + 64 + 0, "Mensaje"
    FEjecutarSQL = False
End Function

Public Function DataOfSystem(Optional ByRef sDate As String, _
    Optional ByRef sHour As String, Optional ByRef sAnio As String, _
    Optional ByRef sTrim As String)

Dim cmd  As ADODB.Command
Dim data As ADODB.Recordset
Dim Fechas As String
Set cmd = New ADODB.Command
Set data = New ADODB.Recordset

cmd.ActiveConnection = Cnx
cmd.CommandText = "Select 'sDate'=Convert(Char(10),getdate(),103), " _
                & " 'sHour'=SubString(Convert(Char(22),Convert(datetime,getdate(),103)),12,11), " _
                & " 'sAnio'=DateName(year,getdate()), " _
                & " 'sTrim'=DateName(quarter,getdate()) "

Set data = cmd.Execute()
Fechas = Format(data.Fields(0), "dd/mm/yyyy")
sDate = Fechas
sHour = data.Fields(1)
sAnio = data.Fields(2)
sTrim = data.Fields(3)

gDate = Fechas
gHour = data.Fields(1)
gAnio = data.Fields(2)
gTrim = data.Fields(3)

data.Close
End Function

Public Function GetFEjecutarSQL(cmd As Command) As Boolean
On Error GoTo merror
Dim rs As New ADODB.Recordset
Set rs = cmd.Execute
vbdev = ""
If Not (rs.BOF And rs.EOF) Then
    vbdev = rs(0).Value
    GetFEjecutarSQL = True
Else
    GetFEjecutarSQL = False
End If
Set cmd.ActiveConnection = Nothing
Set rs = Nothing
For Each er In Cnx.Errors
    MsgBox er.Description, 0 + 64 + 0, "Mensaje"
    GetFEjecutarSQL = False
Next
Exit Function
merror:
Select Case Err.Number
    Case -2147217900:
        MsgBox "No Puede eliminar este registro porque se encuentra relacionado (" & Err.Description & ")", 0 + 64 + 0, "Mensaje"
    Case -2147217911:
        MsgBox Err.Description & "( Consulte con el Administrador de la Base de Datos)", 0 + 64 + 0, "Mensaje"
    Case Else
        MsgBox Err.Description, 0 + 64 + 0, "Mensaje"
End Select
End Function

Public Function GetFEjecutarSQLCM(cmd As Command) As Boolean
On Error GoTo merror
    Dim rs As New ADODB.Recordset
    Dim Arch As String
    Set rs = cmd.Execute
    vbdev = ""
    ERCM = ""
    If Not (rs.BOF And rs.EOF) Then
       vbdev = rs(0).Value
       GetFEjecutarSQLCM = True
    Else
       GetFEjecutarSQLCM = False
    End If
    Set cmd.ActiveConnection = Nothing
    Set rs = Nothing
    Exit Function
merror:
      GetFEjecutarSQLCM = False
      ERCM = Err.Description
'         Arch = Dir("C:\ECM.txt")
'         If Trim(Arch) = "" Then
'             Open "C:\ECM.txt" For Output As #1
'             Print #1, Trim(cmd.Parameters(0)), Trim(cmd.Parameters(1)), ERR.Description
'             Close #1
'           Else
'              Open "C:\ECM.txt" For Input As #1
'              Print #1, Trim(cmd.Parameters(0)), Trim(cmd.Parameters(1)), ERR.Description
'              Close #1
'           End If
End Function

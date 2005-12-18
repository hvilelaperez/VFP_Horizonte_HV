Attribute VB_Name = "Mod_IntegraFlow"
Option Explicit

Public strCadenaConexion As String
Public cnSQLData As New Connection
Public cnSQLData2 As New Connection
Public rsSqlDatos As New Recordset

Public cnConexion As New Connection
Public cnServer As New Connection
Public rsDatos As New Recordset
Public rsPc As New Recordset
Public rsSQL As New Recordset
Public csql As String


Global strRutaDP As String
Global strDirDP As String

Public strPassword As String
Public strPermisos As String

Public strUsuarioTablaDatos As String
Public strUsuarioTablaLote As String

Public strDirectorioImagenes As String
Public StrServidor As String
Public StrServidorSQL As String
Public strData As String
Public strDataSQL As String
Public strCliente As String
Public strDriveServer As String
Public strDriveCliente As String
Public strCadena As String
Public chTipo As String
Public bolSeleccionar As Boolean
Public strUsuarioActivo As String
Public strEjecutar As String
Public intRespuesta As Integer
Public strBaseUsuarioAvon As String

Public StrRutaAF As String


Public StrTipo As String
Public StrTipoApli As String

Public strBaseDatos2 As String
Public strServidor2 As String
Public strUsuario2 As String
Public strPassword2 As String

Public RutaDestino As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long
         
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Public Const MAX_COMPUTERNAME_LENGTH = 255
         


Public Function LeerIni_Integra(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
    'Los parámetros son:
    'lpFileName:    La Aplicación (fichero INI)
    'lpAppName:     La sección que suele estar entrre corchetes
    'lpKeyName:     Clave
    'vDefault:      Valor opcional que devolverá
    '               si no se encuentra la clave.
    '
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String

    'Si no se especifica el valor por defecto,
    'asignar incialmente una cadena vacía
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If

    sRetVal = String$(255, 0)

    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        LeerIni_Integra = lpString
    Else
        LeerIni_Integra = Left(sRetVal, LTmp)
    End If
End Function
Public Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    'Guarda los datos de configuración
    'Los parámetros son los mismos que en LeerIni
    'Siendo lpString el valor a guardar
    '
    Dim LTmp As Long

    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub
Public Function Padln(Cadena As String, Cantidad As Integer, Caracter As String) As String
Dim intTemp As Integer
Dim strTemp As String
    If (Len(Trim(Cadena)) >= Cantidad) Then
        Padln = Cadena
    Else
        strTemp = ""
        For intTemp = 1 To Cantidad - Len(Trim(Cadena))
            strTemp = strTemp & Caracter
        Next
        Padln = strTemp & Cadena
    End If
End Function
Public Function EncryptText(strText As String, ByVal strPwd As String)
Dim i As Integer, c As Integer
Dim strBuff As String

    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

Public Function DecryptText(strText As String, ByVal strPwd As String)
Dim i As Integer, c As Integer
Dim strBuff As String
    
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function
Public Sub Setear()
    StrServidor = LeerIni(App.Path & "\flow.ini", "Seteo", "SERVIDOR")
    strCliente = LeerIni(App.Path & "\flow.ini", "Seteo", "CLIENTE")
    strDriveCliente = LeerIni(App.Path & "\flow.ini", "Seteo", "RUTACLIENTE")
    strDriveServer = LeerIni(App.Path & "\flow.ini", "Seteo", "RUTASERVER")
    strData = LeerIni(App.Path & "\flow.ini", "Seteo", "DATA")
    strDataSQL = LeerIni(App.Path & "\flow.ini", "Seteo", "basedatossql")
    StrServidorSQL = LeerIni(App.Path & "\flow.ini", "Seteo", "serversql")
    StrRutaAF = LeerIni(App.Path & "\flow.ini", "Seteo", "Rutaaf")
    StrTipoApli = LeerIni(App.Path & "\flow.ini", "Seteo", "Tipo")
    strDirectorioLST = LeerIni(App.Path & "\flow.ini", "Seteo", "directoriolst")
    strPlantillaLF = LeerIni(App.Path & "\flow.ini", "Seteo", "plantillalf")
    StrBarra_Chek = LeerIni(App.Path & "\flow.ini", "Seteo", "Barra")
End Sub
Public Sub EstadoMenu(bolestado As Boolean)
    Flow_Frmpanel.mnulote.Enabled = bolestado
    Flow_Frmpanel.mnuprocesos.Enabled = bolestado
    Flow_Frmpanel.mnuusuarios.Enabled = bolestado
End Sub
Public Function ComputerName() As String
    'Devuelve el nombre del equipo actual
    Dim sComputerName As String
    Dim ComputerNameLength As Long
    
    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
     ComputerName = Mid(sComputerName, 1, ComputerNameLength)
End Function

Sub CrearDirectorio(RUTA As String)
On Error GoTo errores
 
 If Not ExisteDirectorio(RUTA) Then
  MkDir RUTA
 End If
 
 
 Exit Sub
errores:
Select Case Err.Number
      Case 75:
        Call MsgBox("No tiene Permisos para Crear Carpetas", vbCritical, "Creación de Directorio")
      Case Else
        Call MsgBox(Err.Description, vbCritical, "Creación de Directorio " & RUTA)
End Select

End Sub
Function ExisteDirectorio(RUTA As String) As Boolean
On Error Resume Next
Dim tempdir As String
ExisteDirectorio = False
If Trim(RUTA) = "" Then ExisteDirectorio = False: Exit Function
tempdir = Dir(RUTA, vbDirectory)
ExisteDirectorio = IIf(Err = 0 And tempdir <> "", True, False)
End Function

Public Sub Archivos_Liquidacion(Xcarpeta As String, Subdirectorio As String, Intbatch As Double)
Dim sArch As String, sDir As String
Dim IntArchPro As Integer
Dim intNumero As String
Dim Directorio  As String
On Error Resume Next
sDir = Xcarpeta & "\" & Subdirectorio

sArch = Dir(sDir & "\*.tif")

IntArchPro = 0
MkDir Xcarpeta & "\" & strPeriodo & "-" & strAño & "-" & Format(Intbatch, "0000000")
Directorio = Xcarpeta & "\" & strPeriodo & "-" & strAño & "-" & Format(Intbatch, "0000000")
Do While sArch <> ""
   IntArchPro = IntArchPro + 1
   intNumero = Format(IntArchPro, "0000")
   If IntArchPro = 1 Then
   
   Else
      sArch = Dir
   End If
  
  If sArch <> "" Then
    FileCopy Xcarpeta & "\" & Subdirectorio & "\" & sArch, Directorio & "\" & strPeriodo & strAño & Format(Intbatch, "0000000") & intNumero & ".tif"
    
    Insertar_Imagen_Scaneada Format(Intbatch, "0000000"), strPeriodo & strAño & Format(Intbatch, "0000000") & intNumero & ".tif"
    
    Kill Xcarpeta & "\" & Subdirectorio & "\" & sArch
  End If
Loop
'lblcantnopro.Caption = "Cantidad de Archivos: " & IntArchPro
If Dir(Xcarpeta & "\" & Subdirectorio) <> "" Then
Else
Kill Xcarpeta & "\" & Subdirectorio & "\*.*"
RmDir UCase(Xcarpeta & "\" & Subdirectorio)
'xx = CurDir
'ChDir ".."
End If
End Sub

Public Sub Archivos_No_Procesados(Xcarpeta As String)
Dim sArch As String, sDir As String
Dim IntArchPro As Integer
sDir = Xcarpeta
sArch = Dir(sDir & "\*.tif")
'IntArchPro = 0
Do While sArch <> ""
 '   lstnoproc.AddItem sArch
    sArch = Dir
 '   IntArchPro = IntArchPro + 1
Loop
'lblcantnopro.Caption = "Cantidad de Archivos: " & IntArchPro
End Sub

Public Sub Archivos_Liquidacion_DroPoup(Xcarpeta As String, Subdirectorio As String, Intbatch As Double)
Dim sArch As String, sDir As String
Dim IntArchPro As Integer
Dim intNumero As String
Dim Directorio  As String
On Error Resume Next
sDir = Xcarpeta & "\" & Subdirectorio

sArch = Dir(sDir & "\*.tif")

IntArchPro = 0
MkDir Xcarpeta & "\" & strPeriodo & "-" & strAño & "-" & Format(Intbatch, "0000000") & "SP"
Directorio = Xcarpeta & "\" & strPeriodo & "-" & strAño & "-" & Format(Intbatch, "0000000") & "SP"
Do While sArch <> ""
   IntArchPro = IntArchPro + 1
   intNumero = Format(IntArchPro, "0000")
   If IntArchPro = 1 Then
   
   Else
      sArch = Dir
   End If
  
  If sArch <> "" Then
    FileCopy Xcarpeta & "\" & Subdirectorio & "\" & sArch, Directorio & "\" & strPeriodo & strAño & Format(Intbatch, "0000000") & intNumero & ".tif"
    
    ''''''''Copiar los datos de imagenes a la base de datos'''''''''''''
    Insertar_Imagen_Scaneada Format(Intbatch, "0000000"), strPeriodo & strAño & Format(Intbatch, "0000000") & intNumero & ".tif"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Kill Xcarpeta & "\" & Subdirectorio & "\" & sArch
  End If
Loop
'lblcantnopro.Caption = "Cantidad de Archivos: " & IntArchPro
If Dir(Xcarpeta & "\" & Subdirectorio) <> "" Then
Else
Kill Xcarpeta & "\" & Subdirectorio & "\*.*"
RmDir UCase(Xcarpeta & "\" & Subdirectorio)
'xx = CurDir
'ChDir ".."
End If
End Sub

Public Sub Insertar_Imagen_Scaneada(strLote As String, strImagen As String)
Dim bolExiste As Boolean
Dim strCadena As String
Dim rsExiste As New Recordset
On Error Resume Next
    strCadena = "exec sp_InsertarImagenes '" & strLote & "','" & strImagen & "'"
    cnServer.Execute strCadena
End Sub

Function Select_Imagen_Min(strLote As String) As String
Dim bolExiste As Boolean
Dim strCadena As String
Dim rsExiste As New Recordset
Dim TClientes As New Recordset
On Error Resume Next
    strCadena = "select * from tb_imagenes where batch= '" & strLote & "' order by id_imagen "
    TClientes.Open strCadena, cnServer, adOpenStatic
    'Set Select_Imagen_Min = .OpenRecordset(Sql)
    If TClientes.RecordCount > 0 Then
    Select_Imagen_Min = Mid(TClientes.Fields("imagen"), Len(TClientes.Fields("imagen")) - 7, 4)
    Else
    Select_Imagen_Min = "0"
    End If
End Function

Function Select_Imagen_Max(strLote As String) As String
Dim bolExiste As Boolean
Dim strCadena As String
Dim rsExiste As New Recordset
Dim TClientes As New Recordset
On Error Resume Next
    strCadena = "select * from tb_imagenes where batch= '" & strLote & "' order by id_imagen desc "
    TClientes.Open strCadena, cnServer, adOpenStatic
    'Set Select_Imagen_Min = .OpenRecordset(Sql)
    If TClientes.RecordCount > 0 Then
    Select_Imagen_Max = Mid(TClientes.Fields("imagen"), Len(TClientes.Fields("imagen")) - 7, 4)
    Else
    Select_Imagen_Max = "0"
    End If
End Function

Function Select_Imagen_Existe(strLote As String, strImagen As String) As String
Dim bolExiste As Boolean
Dim strCadena As String
Dim rsExiste As New Recordset
Dim TClientes As New Recordset
On Error Resume Next
    strCadena = "select * from tb_imagenes where batch= '" & strLote & "' and substring(imagen,18,4)='" & strImagen & "'"
    TClientes.Open strCadena, cnServer, adOpenStatic
    'Set Select_Imagen_Min = .OpenRecordset(Sql)
    If TClientes.RecordCount > 0 Then
    Select_Imagen_Existe = TClientes.Fields("imagen")
    Else
    Select_Imagen_Existe = ""
    End If
End Function

Public Sub Borrar_Imagenes_Scan(strLote As String)
Dim cmd As Command
    Set cmd = New Command
    cmd.ActiveConnection = cnServer
    cmd.CommandText = "DELETE  FROM TB_IMAGENES WHERE BATCH ='" & Format(strLote, "0000000") & "'"
    cmd.Execute
End Sub

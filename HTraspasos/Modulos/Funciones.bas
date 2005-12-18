Attribute VB_Name = "Funciones"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long

Const BIF_RETURNONLYFSDIRS = &H1

Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
                
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

    
Function BrowseFolder(szDialogTitle As String) As String
Dim X As Long, bi As BrowseInfo, dwIList As Long
Dim szPath As String, wPos As Integer

  With bi
    .hOwner = hWndAccessApp
    .lpszTitle = szDialogTitle
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  dwIList = SHBrowseForFolder(bi)
  szPath = Space$(512)
  X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
  If X Then
    wPos = InStr(szPath, Chr(0))
    BrowseFolder = Left$(szPath, wPos - 1)
  Else
    BrowseFolder = ""
  End If

End Function


Function Comparar(valor1 As Variant, valor2 As Variant, Optional indtexto As Boolean) As Boolean
'Si indicador de texto es 1 significa q la comparacion por texto
If indtexto = True Then
    If valor1 = valor2 Then
    Comparar = True
    Else
    Comparar = False
    End If
Else
'Comparacion Numerica
    If val(valor1) = val(valor2) Then
    Comparar = True
    Else
    Comparar = False
    End If
End If
End Function
Public Sub TabSig(i As Integer)
Select Case i
    Case 13
'        SendKeys "{TAB}"
        'Simular que pulsamos la tecla
        keybd_event VK_TAB, 0, 0, 0
        'Simular que soltamos la tecla
                                                                                         keybd_event VK_TAB, 0, KEYEVENTF_KEYUP, 0

End Select

End Sub
Function Modulo11(numruc As String) As Boolean
Dim arruc(10) As Integer
Dim Peso As Integer
Dim AcumRuc As Integer
Dim DigCheq As Integer
Peso = 2
AcumRuc = 0
If Len(numruc) <> 11 Then Modulo11 = False: Exit Function
For i = 10 To 1 Step -1
 arruc(11 - i) = Mid(numruc, i, 1)
Next i
For i = 1 To 10
 AcumRuc = AcumRuc + arruc(i) * Peso
 Peso = Peso + 1
 If Peso > 7 Then
   Peso = 2
 End If
Next i
DigCheq = (11 - (AcumRuc Mod 11)) Mod 10
If numruc = Left(numruc, 10) & DigCheq Then
 Modulo11 = True
Else
 Modulo11 = False
End If
End Function
Public Function fgDatosql(sql As String) As Variant
On Error GoTo TRATA_ERROR
Dim wrs As ADODB.Recordset
Dim cmd As ADODB.Command
Set wrs = New ADODB.Recordset
Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = Con.Cnx
    .CommandText = sql
End With
Set wrs = cmd.Execute
Set cmd.ActiveConnection = Nothing
If Not (wrs.EOF And wrs.BOF) Then
    fgDatosql = Trim(IIf(IsNull(wrs.Fields(0)), "", wrs.Fields(0)))
Else
    fgDatosql = ""
End If
wrs.Close
Exit Function
TRATA_ERROR:
If wrs.State = 1 Then
    wrs.Close
End If
MsgBox Err.Description
End Function
Public Sub Marcado(Caja As Control)
Caja.SelStart = 0
Caja.SelLength = Len(Caja)
End Sub
Public Function DatosServidor()
    Con.DataOfSystem
End Function
Function Conectar() As Boolean
Set Con = New cConexion
If Not Con.ConexionSQL(gUsuario, gPassword, gServidor, gBaseDatos) Then
Conectar = False
Else
Conectar = True
End If
End Function
Public Function FLlenagrilla(grilla As DataGrid, sql As String) As Boolean
On Error GoTo TRATA_ERROR
Dim wrs As ADODB.Recordset
Set wrs = New ADODB.Recordset

With wrs
     .ActiveConnection = Con.Cnx
     .CursorType = adOpenStatic
     .CursorLocation = adUseClient
     .LockType = adLockBatchOptimistic
     .Open sql
End With
If Not (wrs.EOF And wrs.BOF) Then
    Set grilla.DataSource = wrs
    FLlenagrilla = True
Else
    Set grilla.DataSource = wrs
    FLlenagrilla = False
End If
Set wrs.ActiveConnection = Nothing
Exit Function

TRATA_ERROR:
Set wrs.ActiveConnection = Nothing
MsgBox Err.Description
End Function
Public Function LimpiarCadena(expresion As String) As Variant
Dim aux As String
Dim i As Integer
aux = ""
For i = 1 To Len(expresion)
 If Not (Asc(Mid(expresion, i, 1)) >= Asc("A") And Asc(Mid(expresion, i, 1)) <= Asc("Z")) Then
    If Not (Asc(Mid(expresion, i, 1)) >= Asc("a") And Asc(Mid(expresion, i, 1)) <= Asc("z")) Then
        If Not (Asc(Mid(expresion, i, 1)) >= Asc("0") And Asc(Mid(expresion, i, 1)) <= Asc("9")) Then
            If Mid(expresion, i, 1) <> " " Then
                Exit For
            End If
        End If
     End If
End If
aux = aux + Mid(expresion, i, 1)
Next i
LimpiarCadena = aux
End Function

Public Function AccesoBotones(Impr As String, Agr As String, Modi As String, Eli As String, Conf As String) As Boolean
If ValorBarra = "" Then
    ValorBarra = "00000"
    Exit Function
Else
    If Impr = "1" Then   'imprime
        Acc = "1"
    ElseIf Agr = "1" Then  'agrega
        Acc = "2"
    ElseIf Modi = "1" Then 'modifica
        Acc = "3"
    ElseIf Eli = "1" Then  'elimina
        Acc = "4"
    ElseIf Conf = "1" Then  'Confidencial
        Acc = "5"
    End If
    If Valor_Acc Then
        AccesoBotones = True
    Else
        Call Msje
        AccesoBotones = False
    End If
End If
End Function
Public Sub Msje()
MsgBox "Ud. No Tiene Permiso para esta Operacion", 0 + 64 + 0, UCase(gNomUsu)
End Sub
Public Function Valor_Acc() As Boolean
Select Case Acc
    Case "1":   'IMPRIMIR
        If Mid(ValorBarra, 1, 1) = "1" Then
           Valor_Acc = True
        Else
           Valor_Acc = False
        End If
    Case "2":   'AGREGAR
        If Mid(ValorBarra, 2, 1) = "1" Then
           Valor_Acc = True
        Else
           Valor_Acc = False
        End If
    Case "3":   'MODIFICAR
        If Mid(ValorBarra, 3, 1) = "1" Then
           Valor_Acc = True
        Else
           Valor_Acc = False
        End If
    Case "4":   'ELIMINAR
        If Mid(ValorBarra, 4, 1) = "1" Then
           Valor_Acc = True
        Else
           Valor_Acc = False
        End If
    Case "5":   'CONFIDENCIAL
        If Mid(ValorBarra, 5, 1) = "1" Then
           Valor_Acc = True
        Else
           Valor_Acc = False
        End If
End Select
End Function
Public Function FRetornamaximosql(sql, formato As String) As String
On Error GoTo TRATA_ERROR
Dim wrs As ADODB.Recordset
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
Set wrs = New ADODB.Recordset
With cmd
    .ActiveConnection = Con.Cnx
    .CommandText = sql
End With
Set wrs = cmd.Execute
Set cmd.ActiveConnection = Nothing
If IsNull(wrs.Fields(0)) Then
    FRetornamaximosql = Format("1", formato)
Else
    FRetornamaximosql = Format(CLng(wrs.Fields(0)) + 1, formato)
End If
wrs.Close
Exit Function
TRATA_ERROR:
If wrs.State = 1 Then
   wrs.Close
End If
Set cmd.ActiveConnection = Nothing
MsgBox Err.Description
End Function

Function FechaTransm(Fecha As String) As String
'Funcion que convierte el formato dd/mm/aaaa a ddmmaaaa
'If CDate(fecha) Is Date Then
   'FechaTransm = Right(Trim(fecha), 4) & Mid(Trim(fecha), 4, 2) & Left(Trim(fecha), 2)
   FechaTransm = Left(Trim(Fecha), 2) & Mid(Trim(Fecha), 4, 2) & Right(Trim(Fecha), 4)
   
'End If
End Function


Function EnvioForFecha(Fecha As String) As String
'Funcion que convierte el formato dd/mm/aaaa a aaaammdd
'If CDate(fecha) Is Date Then
   EnvioForFecha = Right(Trim(Fecha), 4) & Mid(Trim(Fecha), 4, 2) & Left(Trim(Fecha), 2)
'End If
End Function
Function ReceForFecha(Fecha As String) As String
'Funcion que convierte el formato aaaammdd a dd/mm/aaaa
   ReceForFecha = Right(Trim(Fecha), 2) & "/" & Mid(Trim(Fecha), 5, 2) & "/" & Left(Trim(Fecha), 4)
End Function
Function EnvioMesDev(Fecha As String) As String
'Funcion que convierte el formato mmaaaa a aaaamm
   EnvioMesDev = Right(Trim(Fecha), 4) & Left(Trim(Fecha), 2)
End Function
Function RecepMesDev(Fecha As String) As String
'Funcion que convierte el formato aaaamm a mmaaaa
   RecepMesDev = Right(Trim(Fecha), 2) & "/" & Left(Trim(Fecha), 4)
End Function

Function FCuentaDec(Caja As Control, dec As Integer) As Boolean
    If InStr(Caja, ".") > 0 Then
        If dec = Len(Mid(Caja, InStr(Caja, ".") + 1, dec)) Then
        FCuentaDec = True
        Else
        FCuentaDec = False
        End If
    Else
    FCuentaDec = False
    End If
End Function
Function ObtenerHora() As String
Dim Hora As String
    Hora = fgDatosql("select convert(datetime,getdate(),114)")
    ObtenerHora = Format(Hora, "HH:mm:ss")
End Function
Public Sub CargarComboSql(Cbo As ComboBox, sql As String)
On Error GoTo TRATA_ERROR
Dim wrs As ADODB.Recordset
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
Set wrs = New ADODB.Recordset
With cmd
      .ActiveConnection = Con.Cnx
      .CommandText = sql
End With
Set wrs = cmd.Execute
Set cmd.ActiveConnection = Nothing
Do Until wrs.EOF
   DoEvents
   Cbo.AddItem wrs.Fields(0) & Space(50) & wrs.Fields(1)
   wrs.MoveNext
Loop
wrs.Close
Exit Sub

TRATA_ERROR:
If wrs.State = 1 Then
   wrs.Close
End If
   MsgBox Err.Description

End Sub
Public Sub CargarListaSql(Cbo As ListBox, sql As String)
On Error GoTo TRATA_ERROR
Dim wrs As ADODB.Recordset
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
Set wrs = New ADODB.Recordset
With cmd
      .ActiveConnection = Con.Cnx
      .CommandText = sql
End With
Set wrs = cmd.Execute
Set cmd.ActiveConnection = Nothing
Do Until wrs.EOF
   DoEvents
   Cbo.AddItem wrs.Fields(0) & Space(50) & wrs.Fields(1)
   wrs.MoveNext
Loop
wrs.Close
Exit Sub

TRATA_ERROR:
If wrs.State = 1 Then
   wrs.Close
End If
   MsgBox Err.Description

End Sub
Public Function CodIde(Letra As Integer) As Boolean
On Error GoTo Lerror
Dim rsLetra As ADODB.Recordset
Dim num As Integer
Set rsLetra = New ADODB.Recordset
rsLetra.Open "select CODIDE from tidepla where codafp = '" & vbGAFP & "' and codide= '" & UCase(Chr(Letra)) & "'", Con.Cnx, adOpenForwardOnly
If Not (rsLetra.EOF And rsLetra.BOF) Then
    CodIde = True
Else
    CodIde = False
End If
Set rsLetra = Nothing
Exit Function
Lerror:
MsgBox Err.Description
Set rsLetra = Nothing
End Function

'************************************************************************************
'* Procedimiento que llena un combo con un determinado conjunto de
'* resultados.
'* cbo : nombre del combo que se va cargar
'* sql : selección que se desea hacer, debe incluir 1 campo, primero
'*       el que se desea mostrar en el combo y segundo el que se guarda
'*       a la derecha del combo
'* Utilizada para cargar solo descripciones (carga un solo campo)
'*************************************************************************************
Public Sub ccombosql(Cbo As ComboBox, sql As String)
On Error GoTo TRATA_ERROR
Dim wrs As ADODB.Recordset
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
Set wrs = New ADODB.Recordset
With cmd
    .ActiveConnection = Con.Cnx
    .CommandText = sql
End With
Set wrs = cmd.Execute
Set cmd.ActiveConnection = Nothing
Do Until wrs.EOF
    DoEvents
    Cbo.AddItem wrs.Fields(0)
    wrs.MoveNext
Loop
wrs.Close
Exit Sub
TRATA_ERROR:
If wrs.State = 1 Then
    wrs.Close
End If
    MsgBox Err.Description
End Sub
'********************************************************************************************
'* Posiciona a un combo en una determinada fila, según el dato que se le envie.
'* Se debe tener en cuenta que el combo debe estar cargado de la sgte manera:
'* codigo  - descripción
'* combo  : el nombre del combo que quiere posicionar en una determinada fila
'* dato   : es el código que se encuentra en el lado izquierdo de cada fila
'* lencod : longitud del código
'********************************************************************************************
Public Sub IzqBuscarEnCombo(combo As Control, ByRef Dato As String, lencod)
Dim xfilacombo As String
Dim iCont As Integer
Dim nroItems As Integer
nroItems = combo.ListCount
For iCont = 0 To nroItems - 1
    DoEvents
    xfilacombo = combo.list(iCont)
    If Trim(Left(xfilacombo, lencod)) = Dato Then
        combo.ListIndex = iCont
        Exit Sub
    End If
Next
If nroItems = iCont Then
    combo.ListIndex = -1
End If
End Sub
Function CompararTexto(texto1 As String, texto2 As String, Keyascci As Integer, pos As Byte) As Boolean
'texto2 = texto2 & LimpiarCadena(Trim((Chr(Keyascci))))
If Mid(texto1, pos, 1) = ((Chr(Keyascci))) Then
  CompararTexto = True
 Else
  CompararTexto = False
End If
End Function
Function Rutaimag(RutaStand As String, lotafp As String, nombimag As String) As String
'Rutaimag = RutaStand & "\" & lotafp & "\" & Trim(nombimag) & ".TIF"
 Rutaimag = RutaStand & "\" & Format(lotafp, "0000000#") & "\" & Trim(nombimag) & ".TIF"
End Function
Public Sub pgBuscarEnCombo(combo As Control, ByRef Dato As String, lencod)
Dim xfilacombo As String
Dim iCont As Integer
Dim nroItems As Integer
nroItems = combo.ListCount
For iCont = 0 To nroItems - 1
    DoEvents
    xfilacombo = combo.list(iCont)
    If Trim(Right(Trim(xfilacombo), lencod)) = Dato Then
        combo.ListIndex = iCont
        Exit Sub
    End If
Next
If nroItems = iCont Then
    combo.ListIndex = -1
End If
End Sub
Function ExisteArchivo(Ruta As String) As Boolean
On Error Resume Next
Dim tempdir As String
ExisteArchivo = False
If Trim(Ruta) = "" Then ExisteArchivo = False: Exit Function
tempdir = Dir(Ruta)
ExisteArchivo = IIf(Err = 0 And tempdir <> "", True, False)
End Function

Function ExisteDirectorio(Ruta As String) As Boolean
On Error Resume Next
Dim tempdir As String
ExisteDirectorio = False
If Trim(Ruta) = "" Then ExisteDirectorio = False: Exit Function
tempdir = Dir(Ruta, vbDirectory)
ExisteDirectorio = IIf(Err = 0 And tempdir <> "", True, False)
End Function


Public Sub HDTextBox(Estado As Boolean)
Dim n As Integer
For n = 0 To Screen.ActiveForm.Controls.Count - 1
    If TypeOf Screen.ActiveForm.Controls(n) Is TextBox Then
        Screen.ActiveForm.Controls(n).Enabled = Estado
    End If
Next n
End Sub
'****************************************************************
'* Se utliza para limpiar textbox de un formulario activo
'****************************************************************
Public Sub LimpiarTextBox()
Dim n As Integer
For n = 0 To Screen.ActiveForm.Controls.Count - 1
    If TypeOf Screen.ActiveForm.Controls(n) Is TextBox Then
        Screen.ActiveForm.Controls(n).Text = ""
    End If
Next n
End Sub

Function ValCusp2(texto As String) As Boolean
ValCusp2 = True
If Len(Trim(texto)) < 11 Then ValCusp2 = False: Exit Function

If IsNumeric(Left(Trim(texto), 6)) Then
    ValCusp2 = True
 Else
    ValCusp2 = False
    Exit Function
End If

For i = 7 To 11
  If Not ((Asc(Mid(Trim(texto), i, 1)) >= 65 And Asc(Mid(Trim(texto), i, 1)) <= 90) Or Asc(Mid(Trim(texto), i, 1)) = 165) Then
     ValCusp2 = False
     Exit Function
  End If
Next i

If IsNumeric(Mid(Trim(texto), Len(Trim(texto)), 1)) Then
    ValCusp2 = True
 Else
    ValCusp2 = False
    Exit Function
End If

End Function

Sub GrabarBitacora(tabla As String, campo As String, valnue As String, _
    valant As String, codafp As String, lotafp As String, numfol As String, _
    numpre As String, Numpart As String, Secafi As String, _
    Fecdig As Date, horainim As String, horafinm As String, IdTran As String, Codusu As String)
    
    Set comando = New Command
    With comando
       .ActiveConnection = Con.Cnx
       .CommandText = "sp_Grabarbitacora ?,?,?,?,?, ?,?,?,?,?, ?,?,?,?,? ,?,?,?,? "
       .Parameters("@fecmod") = gDate
       .Parameters("@codmod") = gNomUsu
       .Parameters("@tabla") = tabla
       .Parameters("@codafp") = codafp
       .Parameters("@lotafp") = lotafp
       .Parameters("@numfol") = numfol
       .Parameters("@numpre") = numpre
       .Parameters("@numpart") = Numpart
       .Parameters("@secafi") = Secafi
       .Parameters("@campo") = campo
       .Parameters("@valant") = valant
       .Parameters("@valnue") = valnue
       .Parameters("@estacion") = nHostName
       .Parameters("@horinim") = horainim
       .Parameters("@horfiNm") = horafinm
       .Parameters("@codusu") = Codusu 'gNomUsu
       .Parameters("@usured") = LimpiarCadena(nGetUser)
       .Parameters("@fecdig") = Fecdig
       .Parameters("@idtran") = IdTran
       
       .Execute
    End With

End Sub


Function nGetUser() As String
Dim usern$
Dim Nlen%
Dim X%
nGetUser = ""
usern$ = String$(255, 0)
Nlen% = Len(usern$)
X% = GetUserName(usern$, Len(usern$))
If X% <> 0 Then nGetUser = Left$(usern$, Len(usern$) - 1)
End Function

Function nHostName() As String
Dim HostNam$
Dim NlenH%
Dim d%
nHostName = ""
HostNam$ = String$(255, 0)
NlenH% = Len(HostNam$)
d% = GetComputerName(HostNam$, Len(HostNam$))
If d% <> 0 Then nHostName = Trim(Left$(HostNam$, Len(HostNam$) - 1))
nHostName = LimpiarCadena(nHostName)
End Function


'Public Sub IzqBuscarEnCombo(combo As Control, ByRef dato As String, lencod)
'Dim xfilacombo As String
'Dim iCont As Integer
'Dim nroItems As Integer
'nroItems = combo.ListCount
'For iCont = 0 To nroItems - 1
'    DoEvents
'    xfilacombo = combo.List(iCont)
'    If Trim(Left(xfilacombo, lencod)) = dato Then
'        combo.ListIndex = iCont
'        Exit Sub
'    End If
'Next
'If nroItems = iCont Then
'    combo.ListIndex = -1
'End If
'End Sub


Public Sub SiempreVisible(Formulario As Form, Estado As Boolean)
    
    If Estado Then
       SetWindowPos Formulario.hWnd, HWND_TOPMOST, Formulario.Left / 15, _
                    Formulario.Top / 15, Formulario.Width / 15, _
                    Formulario.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
       SetWindowPos Formulario.hWnd, HWND_NOTOPMOST, Formulario.Left / 15, _
                    Formulario.Top / 15, Formulario.Width / 15, _
                    Formulario.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If

End Sub

Function LeerIni(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String

    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If

    sRetVal = String$(255, 0)

    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        LeerIni = lpString
    Else
        LeerIni = Left(sRetVal, LTmp)
    End If
End Function
Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    Dim LTmp As Long
    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub
Function CartetaAnterior(Ruta As String) As String
Dim Arreglo() As String
Arreglo = Split(Ruta, "\")
NElem = UBound(Arreglo)
If NElem > 0 Then
    CartetaAnterior = Arreglo(NElem - 1)
Else
    CartetaAnterior = ""
End If
End Function
Sub CambiarEstadoLote(Lote As String, Estado As String)
On Error GoTo errores
Dim cmdAct As Command
Set cmdAct = New Command

With cmdAct
    .ActiveConnection = Con.Cnx
    .CommandText = " Update _scan set Estado = '" & Estado & "' where LoteFisico = '" & Lote & "'"
    .Execute
End With

Set cmdAct = Nothing

Exit Sub
errores:
MsgBox Err.Description

End Sub
Sub ActualizaIndDig(Lote As String)
'Actualiza indicadores de digitacion
Dim sqltexto As String
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command

With cmd

sqltexto = " Update Solicitud " & _
    " set IndFecSus = case when inddoc = 'C' and FecRegPro is null then '1' else '0' end, " & _
    " IndFecDec = case when inddoc = 'T' and FecFirmSol is null then '1' else '0' end, " & _
    " IndFecVer = case when inddoc = 'T' and FecRegPro is null then '1' else '0' end, " & _
    " IndSinAfpO = case when inddoc = 'T' and (CodAFPOri is null or CodAFPOri = '') then '1' else '0' end, " & _
    " IndTipFon = case when inddoc = 'T' and ( isnull(TipoFonDes,'') = '') then '1' else '0' end, " & _
    " IndCusBlan = case when inddoc = 'T' and (CUSPP is null or CUSPP = '') then '1' else '0' end " & _
    " where lote = '" & Lote & "'"
    
.ActiveConnection = Con.Cnx
.CommandText = sqltexto
.Execute

End With


Set cmd = Nothing

End Sub
Function CopiarFile(rutaOrigen As String, rutaDestino As String) As Boolean
On Error GoTo errores
CopiarFile = True
FileCopy rutaOrigen, rutaDestino
  
Exit Function
errores:
CopiarFile = False
MsgBox Err.Description
End Function

Function CenterForm(frmForm As Form)
frmForm.Left = (MDIPrincipal.ScaleWidth - frmForm.Width) / 2
frmForm.Top = ((MDIPrincipal.ScaleHeight - frmForm.Height) / 2)
End Function
Sub VerificarRevision(Lote As String)
'Verifica que todos las solicitudes del lote esten revisadas para dar por verificado el lote
If fgDatosql("Select count(numsol) from solicitud where lote = '" & Lote & "' and isnull(indVerD2,0) = 0 and secu = 0 ") > 0 Then
    Call CambiarEstadoLote(Lote, "6")
Else
    Call CambiarEstadoLote(Lote, "7")
End If
End Sub
Function CargarImagenes(Numsol As String, lotefisico As String) As Recordset
On Error Resume Next
Dim RsImag As New Recordset

Dim strsql As String

strsql = " Select imagen, c.Ruta, LoteFisico, fecsal,posicion " & _
         " from _Imagenes b join _Volumenes c on c.idvolumen = b.idvolumen " & _
         " join CtrlDoc d on d.Lote = b.Lotefisico " & _
         " where barcode = 0 and b.Lotefisico = '" & lotefisico & "' " & _
         " and idimagen = '01" & Numsol & "' " & _
         " order by  convert(int,posicion)  "
         
         'and  b.posicion = " & Trim(TDBGrid1.Columns(12)) & "

    If RsImag.State = 1 Then RsImag.Close
    RsImag.Open strsql, Con.Cnx, adOpenStatic, adLockReadOnly
    
    Set CargarImagenes = RsImag
End Function

Function CrearDirectorio(strRuta As String) As Boolean
On Local Error GoTo Err:
Dim bolOk As Boolean
    bolOk = True
    MkDir strRuta
    CrearDirectorio = bolOk
    Exit Function
Err:
    MsgBox Err.Description, vbCritical, "Crear directorio volumen"
    bolOk = False
    CrearDirectorio = bolOk
End Function

Function FrmtCampo(Dato As String, Tipo As String, Tamano As Integer)
On Error GoTo errores
Select Case Tipo
Case "A"
    FrmtCampo = Left(Trim(Dato) & Replica(" ", Tamano), Tamano)
Case "N"
    FrmtCampo = Right(Replica("0", Tamano) & Trim(Dato), Tamano)
End Select
Exit Function
errores:
MsgBox Err.Description & " campo " & Dato, vbCritical, "Error en dato"
End Function

Function Replica(caracter As String, veces As Integer) As String
Replica = ""
For i = 1 To veces
    Replica = Replica & caracter
Next

End Function

Function ValidarDigitacion(Cajatexto As TextBox, KeyCode As Integer, CadenaVal As String) As Integer

KeyCode = Asc(Chr(KeyCode))

If Not ValidarCaracteres(KeyCode, CadenaVal) Then
    ValidarDigitacion = 0
Else
    ValidarDigitacion = KeyCode
End If

End Function
Function ValidarCaracteres(key As Integer, CaracteresValidos As String) As Boolean
Dim CaracterDig As String
If key = vbKeyBack Then
 ValidarCaracteres = True
 Exit Function
End If
 
If key = 13 Then
 ValidarCaracteres = True
 Exit Function
End If
 
 
CaracterDig = Chr(key)
If InStr(CaracteresValidos, CaracterDig) > 0 Then
    ValidarCaracteres = True
Else
    ValidarCaracteres = False
End If

End Function


Public Function Validar_Email(ByVal Email As String) As Boolean

Dim strTmp As String
Dim n As Long
Dim sEXT As String

mensajeError = ""
Validar_Email = True

sEXT = Email

Do While InStr(1, sEXT, ".") <> 0
   sEXT = Right(sEXT, Len(sEXT) - InStr(1, sEXT, "."))
Loop

If Email = "" Then
   Validar_Email = False
   mensajeError = mensajeError & "No se indicó ninguna dirección de " & _
                  "mail para verificar!" & vbNewLine
ElseIf InStr(1, Email, "@") = 0 Then
   Validar_Email = False
   mensajeError = mensajeError & "La dirección de email con contiene el signo @" & vbNewLine
ElseIf InStr(1, Email, "@") = 1 Then
   Validar_Email = False
   mensajeError = mensajeError & "El @ No puede estar al principio" & vbNewLine
ElseIf InStr(1, Email, "@") = Len(Email) Then
   Validar_Email = False
   mensajeError = mensajeError & "El @ no puede estar al final de la dirección" & vbNewLine
ElseIf EXTisOK(sEXT) = False Then
   Validar_Email = False
   mensajeError = mensajeError & "LA dirección no tiene un dominio válido, "
   mensajeError = mensajeError & "por ejemplo : "
   mensajeError = mensajeError & ".com, .net, .gov, .org, .edu, .biz, .tv etc.. " & vbNewLine
ElseIf Len(Email) < 6 Then
   Validar_Email = False
   mensajeError = mensajeError & "La dirección no puede ser menor a 6 caracteres." & vbNewLine
End If
strTmp = Email
Do While InStr(1, strTmp, "@") <> 0
   n = 1
   strTmp = Right(strTmp, Len(strTmp) - InStr(1, strTmp, "@"))
Loop
If n > 1 Then
   Validar_Email = False
   mensajeError = mensajeError & "Solo puede haber un @ en la dirección de email" & vbNewLine
End If

    Dim pos As Integer

    pos = InStr(1, Email, "@")

    If Mid(Email, pos + 1, 1) = "." Then
        Validar_Email = False
        mensajeError = mensajeError & "El punto no puede estar seguido del @" & vbNewLine
    End If

    If mensajeError <> "" Then
        MsgBox mensajeError, vbCritical
    End If

End Function


Public Function EXTisOK(ByVal sEXT As String) As Boolean
Dim EXT As String, X As Long
EXTisOK = False
If Left(sEXT, 1) <> "." Then sEXT = "." & sEXT
    sEXT = UCase(sEXT) 'just to avoid errors
    EXT = EXT & ".COM.EDU.GOV.NET.BIZ.ORG.TV"
    EXT = EXT & ".AF.AL.DZ.As.AD.AO.AI.AQ.AG.AP.AR.AM.AW.AU.AT.AZ.BS.BH.BD.BB.BY"
    EXT = EXT & ".BE.BZ.BJ.BM.BT.BO.BA.BW.BV.BR.IO.BN.BG.BF.MM.BI.KH.CM.CA.CV.KY"
    EXT = EXT & ".CF.TD.CL.CN.CX.CC.CO.KM.CG.CD.CK.CR.CI.HR.CU.CY.CZ.DK.DJ.DM.DO"
    EXT = EXT & ".TP.EC.EG.SV.GQ.ER.EE.ET.FK.FO.FJ.FI.CS.SU.FR.FX.GF.PF.TF.GA.GM.GE.DE"
    EXT = EXT & ".GH.GI.GB.GR.GL.GD.GP.GU.GT.GN.GW.GY.HT.HM.HN.HK.HU.IS.IN.ID.IR.IQ"
    EXT = EXT & ".IE.IL.IT.JM.JP.JO.KZ.KE.KI.KW.KG.LA.LV.LB.LS.LR.LY.LI.LT.LU.MO.MK.MG"
    EXT = EXT & ".MW.MY.MV.ML.MT.MH.MQ.MR.MU.YT.MX.FM.MD.MC.MN.MS.MA.MZ.NA"
    EXT = EXT & ".NR.NP.NL.AN.NT.NC.NZ.NI.NE.NG.NU.NF.KP.MP.NO.OM.PK.PW.PA.PG.PY"
    EXT = EXT & ".PE.PH.PN.PL.PT.PR.QA.RE.RO.RU.RW.GS.SH.KN.LC.PM.ST.VC.SM.SA.SN.SC"
    EXT = EXT & ".SL.SG.SK.SI.SB.SO.ZA.KR.ES.LK.SD.SR.SJ.SZ.SE.CH.SY.TJ.TW.TZ.TH.TG.TK"
    EXT = EXT & ".TO.TT.TN.TR.TM.TC.TV.UG.UA.AE.UK.US.UY.UM.UZ.VU.VA.VE.VN.VG.VI"
    EXT = EXT & ".WF.WS.EH.YE.YU.ZR.ZM.ZW"
    EXT = UCase(EXT) 'just to avoid errors
    If InStr(1, EXT, sEXT, 0) <> 0 Then
        EXTisOK = True
    End If
End Function

Sub EliminarLoteCompleto(Lote As String)
Dim CmdEliminar As New ADODB.Command
With CmdEliminar
    .ActiveConnection = Con.Cnx
    .CommandText = "Eliminar_Lote_Completo ?,?"
    .Parameters("@Lote") = Lote
    .Parameters("@usuario") = gNomUsu
    .Execute , , adExecuteNoRecords
End With
End Sub

Sub EliminarLoteDigitacion(Lote As String)

Dim CmdEliminar As New ADODB.Command
With CmdEliminar
    .ActiveConnection = Con.Cnx
    .CommandText = "Eliminar_Lote_Digitacion ?,?"
    .Parameters("@Lote") = Lote
    .Parameters("@usuario") = gNomUsu
    .Execute , , adExecuteNoRecords
End With

End Sub
Function ExisteLote(strLoteFisico As String) As Boolean
Dim rsTempo As New Recordset
Dim bolOk As Boolean
    bolOk = False
    If rsTempo.State = 1 Then rsTempo.Close
    Set rsTempo = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Scan,[lote],'where idvolumen=" & lngIdVolumenActual & " and lotefisico = " & strLoteFisico & "',''")
    If rsTempo.RecordCount > 0 Then
        bolOk = True
    End If
    If rsTempo.State = 1 Then rsTempo.Close
    Set rsTempo = Nothing
     
   
    ExisteLote = bolOk
End Function

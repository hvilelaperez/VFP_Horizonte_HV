Attribute VB_Name = "Mod_IntegraICR"
'Public cnServer As New Connection
Public strZonaActual As String
Public strCampañaActual As String
Public intProductos As Integer
Public intCantidad_Minimo As Integer

'Public strDataSQL As String
'Public StrServidorSQL As String
Public strDirectorioLST As String
Public strPlantillaLF As String

Global StrDupli As String

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

Public Sub MostrarAvance(picAvance As Control, ByVal intValor As Integer, intMax As Integer, ByVal bolVisible As Boolean)
    Dim strPintarValor As String
    Dim intLeft As Integer
    Dim intTop As Integer
    Dim intRight As Integer
    Dim intBottom As Integer
    Dim intLineWidth As Integer

    Const DGREYCOLOUR& = &H808080
    Const LGREYCOLOUR& = &HC0C0C0
    Const WHITECOLOUR& = &HFFFFFF

    Const COPYPEN = 13
    Const XORPEN = 7

    If intValor < 0 Then
        intValor = 0
    Else
        If intValor > intMax Then
            intValor = intMax
        End If
    End If

    intLineWidth = Screen.TwipsPerPixelX

    If picAvance.BorderStyle <> 0 Then
        picAvance.BorderStyle = 0
    End If

    If bolVisible = False Then
        If picAvance.Visible <> False Then
            picAvance.Visible = False
        End If
    Else
        If picAvance.Visible <> True Then
            picAvance.Visible = True
        End If
    End If

    picAvance.Tag = intValor
    strPintarValor = CStr(intValor) & "%"
    intLeft = intLineWidth
    intTop = intLineWidth
    intRight = picAvance.ScaleWidth - intLineWidth
    intBottom = picAvance.ScaleHeight - intLineWidth

    picAvance.DrawMode = COPYPEN
    picAvance.Line (intLeft, intTop)-(intRight, intBottom), picAvance.BackColor, BF
    
    picAvance.CurrentX = (picAvance.ScaleWidth - picAvance.TextWidth(strPintarValor)) / 2
    picAvance.CurrentY = (picAvance.ScaleHeight - picAvance.TextHeight(strPintarValor)) / 2
    picAvance.Print strPintarValor
    
    If intValor > 0 Then
        picAvance.DrawMode = XORPEN
        picAvance.Line (intLeft, intTop)-((intRight / intMax) * intValor, intBottom), picAvance.FillColor, BF
        picAvance.Line (intLeft, intTop)-((intRight / intMax) * intValor, intBottom), picAvance.BackColor, BF
    End If
    
    picAvance.DrawMode = COPYPEN
    picAvance.Line (intRight, intLineWidth)-(intRight, intBottom), WHITECOLOUR&, BF
    picAvance.Line (intLineWidth, intBottom)-(intRight, intBottom), WHITECOLOUR&, BF
    picAvance.Line (0, 0)-(intRight, 0), DGREYCOLOUR&, BF
    picAvance.Line (0, 0)-(0, intBottom), DGREYCOLOUR&, BF
    picAvance.Line (intLeft, intTop)-(intRight - intLineWidth, intBottom - intLineWidth), LGREYCOLOUR, B
End Sub
'Function Padln(strTexto As String, intLongitud As Integer, strCaracter As String)
'Dim strCadena As String
'Dim intIndex As Integer
'    strCadena = Trim(strTexto)
'    For intIndex = 1 To (intLongitud - Len(Trim(strTexto)))
'        strCadena = strCaracter & strCadena
'    Next
'    Padln = strCadena
'End Function
Public Sub Centrar(frmPadre As Form, frmHijo As Form)
    frmHijo.Left = (frmPadre.Width / 2) - (frmHijo.Width / 2)
    frmHijo.Top = (frmPadre.Height / 2) - (frmHijo.Height / 2)
End Sub
Function ExisteAsesora(strAsesora As String, strZona As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & strZona & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    ExisteAsesora = bolExiste
End Function
Function DigitoOK(ByVal StrCodigo As String) As Boolean
Dim bolOk As Boolean
Dim strParte1 As String
Dim strparte2 As String
Dim intIndice As Integer
Dim intFactor As Integer
    StrCodigo = Format(RTrim(LTrim(StrCodigo)), "00000")
    strParte1 = Mid(StrCodigo, 1, 4)
    strparte2 = Mid(StrCodigo, 5, 1)
    bolOk = False
    intFactor = 0
    intFactor = intFactor + Mid(strParte1, 1, 1) * 9
    intFactor = intFactor + Mid(strParte1, 2, 1) * 7
    intFactor = intFactor + Mid(strParte1, 3, 1) * 3
    intFactor = intFactor + Mid(strParte1, 4, 1) * 1
    If Val(strparte2) = Right(intFactor, 1) Then
        bolOk = True
    Else
        bolOk = False
    End If
    DigitoOK = bolOk
End Function
Function ExisteCL(ByVal strCL As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_cl from cl where substring(id_cl,1,4)='" & Mid(Format(strCL, "00000"), 1, 4) & "'"
    If Len(Trim(strCL)) = 5 Then
        strCL = Mid(strCL, 1, 4)
    End If
    rsExiste.Source = "select id_cl from cl where substring(id_cl,1,4)='" & strCL & "'"
    
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    ExisteCL = bolExiste
End Function
Public Sub Setear_Icr()
    strDataSQL = LeerIni(App.Path & "\dyclass.ini", "Seteo", "basedatossql")
    StrServidorSQL = LeerIni(App.Path & "\dyclass.ini", "Seteo", "serversql")
    strDirectorioLST = LeerIni(App.Path & "\dyclass.ini", "Seteo", "directoriolst")
    strPlantillaLF = LeerIni(App.Path & "\dyclass.ini", "Seteo", "plantillalf")
End Sub

'Public Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
'    'Guarda los datos de configuración
'    'Los parámetros son los mismos que en LeerIni
'    'Siendo lpString el valor a guardar
'    '
'    Dim LTmp As Long
'
'    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
'End Sub
Function ExisteZona(strZona As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    rsExiste.Source = "select id_zona from zonas where id_zona='" & strZona & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    ExisteZona = bolExiste
End Function

Function ExisteBarra(strBarra As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select nrolIQUI from Tb_Formularios where NROLIQUI='" & "HO" & Trim$(Mid(strBarra, 1, 9)) & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    ExisteBarra = bolExiste
End Function
Function ExisteNroliq(strNroliq As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select nroLiqui from Tb_Formularios where NROLiqui='" & Trim$(strNroliq) & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    ExisteNroliq = bolExiste
End Function
Function ExisteBarra_Liqui(strBarra As String, strNroliq As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select nrolqp,nroliqui from Tb_Formularios where NROLQP='" & Trim$(Mid(strBarra, 2, 9)) & "' and nroliqui='" & strNroliq & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    ExisteBarra_Liqui = bolExiste
End Function

Function Existe_Liqui_Pag(StrPag As String, strNroliq As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select NROLIQPAG,nroliqui from Tb_Formularios where NROLIQPAG='" & "HO" & strNroliq & Trim$(StrPag) & "' and nroliqui='" & "HO" & strNroliq & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    Existe_Liqui_Pag = bolExiste
End Function



Function ExisteBarra_Liqui_pag(strBarra As String, strNroliq As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select nrolqp,nroliqui,nroliqpag from Tb_Formularios where NROLIQPAG='" & strNroliq & Trim$(Mid(strBarra, 11, 3)) & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    ExisteBarra_Liqui_pag = bolExiste
End Function

Function Liqui_pag_Procesado(strBarra As String, strNroliq As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select nrolqp,nroliqui,nroliqpag from Tb_Formularios where NROLIQPAG='" & strNroliq & Trim$(Mid(strBarra, 11, 3)) & "' and procesado ='0'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    Liqui_pag_Procesado = bolExiste
End Function

Function Liqui_pagina_Procesado(strpagina As String, strNroliq As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select nroliqui,nroliqpag from Tb_Formularios where NROLIQPAG='" & strNroliq & strpagina & "' and procesado ='0'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    Liqui_pagina_Procesado = bolExiste
End Function

Function LiquidacionFormulario(Strliqui As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select nrolqp,nroliqui,nroliqpag from Tb_Formularios where NROLIQPAG='" & strNroliq & Trim$(Strliqui) & "' and procesado ='0'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    LiquidacionFormulario = bolExiste
End Function
Function LiquidacionProcesada(Strliqui As String, StrPag As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select * from Tb_liquidacion where ds_nroliq='" & Strliqui & "' and ds_pag ='" & StrPag & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    LiquidacionProcesada = bolExiste
End Function
Function CupssProcesada(Strliqui As String, StrPag As String, StrCupss As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
    rsExiste.Source = "Select * from Tb_FORMULARIOS where nroliqUI='" & Strliqui & "' and CAMPO2 ='" & StrPag & "' and (" & StrCupss & " IS NULL OR " & StrCupss & "='')"
    'Select * from Tb_FORMULARIOS where nroliqUI='IN10060912' and CAMPO2 ='PAG.:001' and (CUSPP3 is null OR CUSPP3='')
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    CupssProcesada = bolExiste
End Function

Function Verificar_Codigo(StrCodigo As String) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
bolExiste = False
Select Case Format(StrCodigo, "00")
    Case ""
    bolExiste = True
    Case "01"
    bolExiste = True
    Case "02"
    bolExiste = True
    Case "03"
    bolExiste = True
    Case "04"
    bolExiste = True
    Case "05"
    bolExiste = True
    Case "09"
    bolExiste = True
    Case "10"
    bolExiste = True
    Case "11"
    bolExiste = True
    Case "12"
    bolExiste = True
    Case "20"
    bolExiste = True
    Case "50"
    bolExiste = True
    Case "60"
    bolExiste = True
    Case "70"
    bolExiste = True
    Case "80"
    bolExiste = True
    Case "90"
    bolExiste = True
'    Case Else
'        If IsNumeric((Format(StrCodigo, "00"))) Then
'        bolExiste = True
'        Else
'        bolExiste = False
'        End If
    
End Select
    Verificar_Codigo = bolExiste
End Function

Function Verificar_Fecha(strFecha As String) As Boolean
Dim StrfechaAux As String
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
bolExiste = False

If Len(strFecha) = 6 Then
    If IsNumeric((strFecha)) Then
       bolExiste = True
       If Val(Mid(strFecha, 5, 2)) <= 31 Then
          If Mid(strFecha, 3, 2) = "02" Then
            If Val(Mid(strFecha, 1, 2)) <= 28 Then
            Else
            bolExiste = False
            End If
          Else
            If Val(Mid(strFecha, 3, 2)) <= 12 Then
                
                If Val(Mid(strFecha, 1, 2)) <= 99 And Val(Mid(strFecha, 1, 2)) >= 50 Then
                StrfechaAux = "19" & Mid(strFecha, 1, 2)
                Else
                StrfechaAux = "20" & Mid(strFecha, 1, 2)
                End If
                
                If IsDate(Mid(strFecha, 5, 2) & "/" & Mid(strFecha, 3, 2) & "/" & StrfechaAux) Then
                bolExiste = True
                Else
                bolExiste = False
                End If
            Else
            bolExiste = False
            End If
          
          End If
       Else
       bolExiste = False
       End If
    Else
    bolExiste = False
    End If
    
Else
   If strFecha = "" Then
   bolExiste = True
   Else
   bolExiste = False
   End If
    
End If

    Verificar_Fecha = bolExiste
End Function


Public Function Mostrar_Fecha_Ref(Xcodigo As String, xinicio As String, xfin As String, xmonto As String, xPlani As String) As String

Select Case Format(Xcodigo, "00")
  Case "":
   If Len(xinicio) > 0 Then
   
       If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
         If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
          
          If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
              Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
          Else
              Mostrar_Fecha = ""
          End If
         Else
          Mostrar_Fecha = ""
         End If
       Else
          If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
          
            If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
               Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
            Else
               Mostrar_Fecha = ""
            End If
          
          Else
          Mostrar_Fecha = ""
          End If
       End If
       
   End If
   If Len(xfin) > 0 Then
    'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
    
       If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
         If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
             If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
             Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
             Else
             Mostrar_Fecha = ""
             End If
         Else
          Mostrar_Fecha = ""
         End If
       Else
          If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
            If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
            Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
            Else
            Mostrar_Fecha = ""
            End If
          Else
          Mostrar_Fecha = ""
          End If
       End If
    
   End If
    
    If Len(xinicio) > 0 And Len(xfin) > 0 Then
    'Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
    
       If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
         If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
            If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
             Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
             Else
             Mostrar_Fecha = ""
            End If
         Else
          Mostrar_Fecha = ""
         End If
       Else
          If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
            If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
             Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
            Else
             Mostrar_Fecha = ""
            End If
          Else
          Mostrar_Fecha = ""
          End If
       End If
    
    End If
    
    If Len(xinicio) = 0 And Len(xfin) = 0 Then
       If Len(xmonto) > 0 Then
       Mostrar_Fecha = "" '"RIA"
       End If
     End If
    
     If Len(xPlani) > 0 Then
       Mostrar_Fecha = "" '"RIA"
     End If

  Case "01":
        
        If Len(xinicio) > 0 Then
         'Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
                  If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
                     If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                     Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
                     Else
                     Mostrar_Fecha = ""
                     End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
                    If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                    Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                    Else
                    Mostrar_Fecha = ""
                    End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
        End If
        
        If Len(xfin) > 0 Then
         'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
'                If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
'                  If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
'                   Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
'                  Else
'                   Mostrar_Fecha = ""
'                  End If
'                Else
'                   If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
'                   Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
'                   Else
'                   Mostrar_Fecha = ""
'                   End If
'                End If
                 Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
         Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
          'Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
                  If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
                      If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                      Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
                      Else
                      Mostrar_Fecha = ""
                      End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
                      If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                      Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                      Else
                      Mostrar_Fecha = ""
                      End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
          
        End If
        
        If Len(xinicio) = 4 Then
        'Mostrar_Fecha = "01" & "/" & Mid(xinicio, 1, 2) & "/" & "20" & Mid(xinicio, 3, 2)
           xdia = "01" 'DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
           If Val(Mid(xinicio, 3, 2)) > 70 And Val(Mid(xinicio, 3, 2)) <= 99 Then
             If IsDate(xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)) Then
                If Val(Mid(xinicio, 1, 2)) > 0 And Val(Mid(xinicio, 1, 2)) <= 12 Then
                Mostrar_Fecha = xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)
                Else
                Mostrar_Fecha = ""
                End If
             Else
             Mostrar_Fecha = ""
             End If
           Else
             xdia = "01" 'DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
             If Val(Mid(xinicio, 3, 2)) > 70 And Val(Mid(xinicio, 3, 2)) <= 99 Then
               If IsDate(xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)) Then
                If Val(Mid(xinicio, 1, 2)) > 0 And Val(Mid(xinicio, 1, 2)) <= 12 Then
                Mostrar_Fecha = xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)
                Else
                Mostrar_Fecha = ""
                End If
               Else
                Mostrar_Fecha = ""
               End If
             Else
               If IsDate(xdia & "/" & Mid(xinicio, 1, 2) & "/" & "20" & Mid(xinicio, 3, 2)) Then
                    If Val(Mid(xinicio, 1, 2)) > 0 And Val(Mid(xinicio, 1, 2)) <= 12 Then
                    Mostrar_Fecha = xdia & "/" & Mid(xinicio, 1, 2) & "/" & "20" & Mid(xinicio, 3, 2)
                    Else
                    Mostrar_Fecha = ""

                    End If
               Else
                Mostrar_Fecha = ""
               End If
             End If
           End If
        End If
        
        If Len(xfin) = 4 Then
        Mostrar_Fecha = ""
        End If
        
  Case "02":
        
        If Len(xfin) > 0 Then
          'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
               If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
                  If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                       If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                           If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                           Mostrar_Fecha = ""
                           Else
                           Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
                           End If
                            
                       Else
                       Mostrar_Fecha = ""
                       End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                       
                       If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                           If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                           Mostrar_Fecha = ""
                           Else
                           Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
                           End If
                       Else
                       Mostrar_Fecha = ""
                       End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
        End If
        
        If Len(xinicio) > 0 Then
        Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
         Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
         'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
              If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
                  If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                         If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                             If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                             Mostrar_Fecha = ""
                             Else
                             Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
                             End If
                         Else
                         Mostrar_Fecha = ""
                         End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                         If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                             If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                             Mostrar_Fecha = ""
                             Else
                             Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
                             End If
                         Else
                         Mostrar_Fecha = ""
                         End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
        End If
        
        If Len(xinicio) = 4 And Len(xfin) = 4 Then
         'xdia = DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
         'Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 4)
         
           xdia = DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
           If Val(Mid(xfin, 3, 2)) > 70 And Val(Mid(xfin, 3, 2)) <= 99 Then
             If IsDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                 If Val(Mid(xfin, 1, 2)) > 0 And Val(Mid(xfin, 1, 2)) <= 12 Then
                     
                     If Format(Now, "dd/mm/yyyy") < CDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                     Mostrar_Fecha = ""
                     Else
                     Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)
                     End If
                 Else
                 Mostrar_Fecha = ""
                 End If
             Else
             Mostrar_Fecha = ""
             End If
           Else
             xdia = DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
             If Val(Mid(xfin, 3, 2)) > 70 And Val(Mid(xfin, 3, 2)) <= 99 Then
               If IsDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                    
                    If Val(Mid(xfin, 1, 2)) > 0 And Val(Mid(xfin, 1, 2)) <= 12 Then
                    
                        If Format(Now, "dd/mm/yyyy") < CDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                        Mostrar_Fecha = ""
                        Else
                        Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)
                        End If
                    
                    
                    Else
                    Mostrar_Fecha = ""
                    End If
               Else
                Mostrar_Fecha = ""
               End If
             Else
               If IsDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 2)) Then
                    If Val(Mid(xfin, 1, 2)) > 0 And Val(Mid(xfin, 1, 2)) <= 12 Then
                        
                        If Format(Now, "dd/mm/yyyy") < CDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 2)) Then
                        Mostrar_Fecha = ""
                        Else
                        Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 2)
                        End If
                    Else
                    Mostrar_Fecha = ""
                    End If
               Else
                Mostrar_Fecha = ""
               End If
             End If
           End If
         
        End If

Case "04":
        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
         Mostrar_Fecha = ""
        End If
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Fecha = ""
        End If
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        End If
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Fecha = ""
        End If

Case "05":
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        End If
        
        Case Xcodigo = "18"
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Fecha = ""
        End If

Case "20":
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        End If

Case "21":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        'End If
Case "22":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""

Case "RIA":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""

Case "99":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        
        If Len(xmonto) > 0 Then
        Mostrar_Fecha = "" '"RIA"
        End If
        If Len(xmonto) = 0 Then
        Mostrar_Fecha = ""
        End If
 Case Else:
        Mostrar_Fecha = ""
End Select
Mostrar_Fecha_Ref = Mostrar_Fecha
End Function

Public Function Mostrar_Comentario(Xcodigo As String, xinicio As String, xfin As String, xmonto As String, xPlani As String) As String
StrDupli = "0"
Select Case Format(Xcodigo, "00")
  Case "":
'        If Len(xinicio) > 0 Then
'         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
'
'        If Len(xfin) > 0 Then
'         Mostrar_Comentario = "CESE DE RELACION LABORAL"
'        End If
'
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'         StrDupli = 1
'        End If
'
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'          If Len(xmonto) > 0 Then
'          Mostrar_Comentario = "ACTUALIZACION DE RIA"
'          End If
'        End If
'
'        If Len(xPlani) > 0 Then
'         Mostrar_Comentario = "OTROS MOTIVOS-ED"
'        End If
         
  Case "01":
  Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        If Len(xinicio) > 0 Then
'         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
'        If Len(xfin) > 0 Then
'         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
'
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        StrDupli = 1
'        End If
'        If Len(xinicio) = 4 Then
'        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
  Case "02":
  Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
'         Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        End If
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'        Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        End If
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        End If
'        If Len(xfin) = 4 Then
'        Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        End If
Case "03":
 Mostrar_Comentario = "INICIO DE SUBSIDIO"
'        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
'         Mostrar_Comentario = "INICIO DE SUBSIDIO"
'        End If
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'         Mostrar_Comentario = "INICIO DE SUBSIDIO"
'        End If

Case "04":
Mostrar_Comentario = "INICIO DE LICENCIA"
'        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
'         Mostrar_Comentario = "INICIO DE LICENCIA"
'        End If
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'         Mostrar_Comentario = "INICIO DE LICENCIA"
'        End If


Case "05":
Mostrar_Comentario = "INICIO DE PERIODO VACACIONAL"
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "INICIO DE PERIODO VACACIONAL"
'        Else
'        Mostrar_Comentario = "INICIO DE PERIODO VACACIONAL"
'        End If
Case "09":
Mostrar_Comentario = "RECEPCION EXTEMPORANEA DE CONTRATO"
Case "10":
Mostrar_Comentario = "AFILIADO NUNCA LABORO EN LA EMPRESA"
Case "11":
Mostrar_Comentario = "APORTE CANCELADO A AFP"
Case "12":
Mostrar_Comentario = "APORTE CANCELADO A OTRA AFP"
Case "20":
Mostrar_Comentario = "DEUDA REAL"


'Case Xcodigo = "18"
'
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'        Mostrar_Comentario = "AFILIADO NUNCA LABORO EN LA EMPRESA"
'        End If
'
'Case "20":
'
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "POSTERGAR DEVENGUE Lima y provincias"
'        Else
'        Mostrar_Comentario = "POSTERGAR DEVENGUE Lima y provincias"
'        End If
'
'Case "21":
'
'        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "APORTE CANCELADO A AFP Horizonte"
'        'End If
'Case "22":
'
'        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "APORTE CANCELADO A OTRA AFP"
'
'Case "RIA":
'
'        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "ACTUALIZACION DE RIA"
'
'Case "99":
'
'        If Len(xmonto) > 0 Then
'        Mostrar_Comentario = "ACTUALIZACION DE RIA"
'        End If
'        If Len(xmonto) = 0 Then
'        Mostrar_Comentario = "EMPLEADOR CONFIRMA DEUDA"
'        End If
Case Else:
        'Mostrar_Comentario = "OTROS MOTIVOS-ED"
        Mostrar_Comentario = ""
End Select
End Function

Public Function Mostrar_Rechazo(Xcodigo As String, xinicio As String, xfin As String, xmonto As String, xPlani As String) As String
StrDupli = "0"
Mostrar_Rechazo = ""
Select Case Format(Xcodigo, "00")
  Case "":
    
  If Xcodigo <> "" And xinicio <> "" And xfin <> "" And xPlani <> "" And Val(xmonto) <> 0 Then
      Mostrar_Rechazo = "CODIGO EN BLANCO"
  End If
'        If Len(xinicio) > 0 Then
'         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
'
'        If Len(xfin) > 0 Then
'         Mostrar_Comentario = "CESE DE RELACION LABORAL"
'        End If
'
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'         StrDupli = 1
'        End If
'
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'          If Len(xmonto) > 0 Then
'          Mostrar_Comentario = "ACTUALIZACION DE RIA"
'          End If
'        End If
'
'        If Len(xPlani) > 0 Then
'         Mostrar_Comentario = "OTROS MOTIVOS-ED"
'        End If
         
  Case "01":
  'Mostrar_Comentario = "INICIO DE RELACION LABORAL"
        If Len(xinicio) > 0 Then
         Mostrar_Rechazo = ""
        Else
         Mostrar_Rechazo = "FECHA INICIAL INVALIDA"
        End If
'        If Len(xfin) > 0 Then
'         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
'
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        StrDupli = 1
'        End If
'        If Len(xinicio) = 4 Then
'        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
'        End If
  Case "02":
  'Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
        If Len(xfin) > 0 Then
         Mostrar_Rechazo = ""
        Else
         Mostrar_Rechazo = "FECHA FINAL INVALIDA"
        End If
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'        Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        End If
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        End If
'        If Len(xfin) = 4 Then
'        Mostrar_Comentario = "TERMINO DE RELACION LABORAL"
'        End If
Case "03":
'     Mostrar_Comentario = "INICIO DE SUBSIDIO"
'        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
'         Mostrar_Comentario = "INICIO DE SUBSIDIO"
'        End If
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'         Mostrar_Comentario = "INICIO DE SUBSIDIO"
'        End If

Case "04":
'     Mostrar_Comentario = "INICIO DE LICENCIA"
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
         Mostrar_Rechazo = ""
         Else
         Mostrar_Rechazo = "FECHA INVALIDA"
        End If
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'         Mostrar_Comentario = "INICIO DE LICENCIA"
'        End If


Case "05":
'Mostrar_Comentario = "INICIO DE PERIODO VACACIONAL"
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "INICIO DE PERIODO VACACIONAL"
'        Else
'        Mostrar_Comentario = "INICIO DE PERIODO VACACIONAL"
'        End If
Case "09":
'Mostrar_Comentario = "RECEPCION EXTEMPORANEA DE CONTRATO"
         'If Len(xinicio) > 0 And Len(xfin) > 0 Then
         If Len(xinicio) > 0 Then
           Mostrar_Rechazo = ""
         Else
           Mostrar_Rechazo = "FECHA INICIO INVALIDA"
         End If
Case "10":
'Mostrar_Comentario = "AFILIADO NUNCA LABORO EN LA EMPRESA"
Case "11":
Mostrar_Rechazo = ""
'         If Len(xPlani) > 0 Then
'           Mostrar_Rechazo = ""
'         Else
'           Mostrar_Rechazo = "INCOSISTENCIA"
'         End If
'Mostrar_Comentario = "APORTE CANCELADO A AFP"
Case "12":
           Mostrar_Rechazo = ""
'Mostrar_Comentario = "APORTE CANCELADO A OTRA AFP"
'         If Len(xPlani) > 0 Then
'           Mostrar_Rechazo = ""
'         Else
'           Mostrar_Rechazo = "INCOSISTENCIA"
'         End If
Case "20":
'Mostrar_Comentario = "DEUDA REAL"
         If Len(xmonto) > 0 Then
           Mostrar_Rechazo = ""
         Else
           Mostrar_Rechazo = "INCOSISTENCIA"
         End If


'Case Xcodigo = "18"
'
'        If Len(xinicio) = 0 And Len(xfin) = 0 Then
'        Mostrar_Comentario = "AFILIADO NUNCA LABORO EN LA EMPRESA"
'        End If
'
'Case "20":
'
'        If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "POSTERGAR DEVENGUE Lima y provincias"
'        Else
'        Mostrar_Comentario = "POSTERGAR DEVENGUE Lima y provincias"
'        End If
'
'Case "21":
'
'        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "APORTE CANCELADO A AFP Horizonte"
'        'End If
'Case "22":
'
'        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "APORTE CANCELADO A OTRA AFP"
'
'Case "RIA":
'
'        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "ACTUALIZACION DE RIA"
'
'Case "99":
'
'        If Len(xmonto) > 0 Then
'        Mostrar_Comentario = "ACTUALIZACION DE RIA"
'        End If
'        If Len(xmonto) = 0 Then
'        Mostrar_Comentario = "EMPLEADOR CONFIRMA DEUDA"
'        End If
'Case Else:
        'Mostrar_Comentario = "OTROS MOTIVOS-ED"
'        Mostrar_Comentario = ""
End Select
End Function

Public Function Codigo_Nulo(xinicio As String, xfin As String, xmonto As String) As String
Dim StrCodigo As String
    StrCodigo = ""
         If Len(xinicio) > 0 Then
         StrCodigo = "01"
         End If
        
         If Len(xfin) > 0 Then
         StrCodigo = "02"
         End If
        
         If Len(xinicio) > 0 And Len(xfin) > 0 Then
         StrCodigo = "01"
         End If
         
         If Len(xinicio) = 0 And Len(xfin) = 0 Then
          If Len(xmonto) > 0 Then
            'StrCodigo = "99"
          End If
         End If
         
 Codigo_Nulo = StrCodigo
 End Function

Function Verificar_Pag(StrPag As String) As Boolean
Dim bolExiste As Boolean
Dim StrNropag As String
Dim rsExiste As New Recordset
  
  If Len(StrPag) = 5 Then
     If IsNumeric(StrPag) Then
     bolExiste = True
     Else
     bolExiste = False
     End If
  Else
  bolExiste = False
  End If
    
'    rsExiste.CursorType = adOpenDynamic
'    rsExiste.LockType = adLockBatchOptimistic
'    rsExiste.ActiveConnection = cnServer
'    'rsExiste.Source = "select id_asesora from asesoras where id_asesora='" & strAsesora & "' and id_zona='" & StrZona & "'"
'    rsExiste.Source = "Select nroLiqui from Tb_Formularios where NROLiqui='" & Trim$(strNroliq) & "'"
'    rsExiste.Open
'    bolExiste = False
'    If rsExiste.RecordCount > 0 Then bolExiste = True
'    If rsExiste.State = 1 Then rsExiste.Close
'    Set rsExiste = Nothing

    Verificar_Pag = bolExiste
End Function

Function Verificar_No_Existe(StrCodigo As String) As String
Dim bolExiste As String
Dim rsExiste As New Recordset
bolExiste = ""
Select Case Format(StrCodigo, "00")
    Case ""
    bolExiste = ""
    Case "01"
    bolExiste = "01"
    Case "02"
    bolExiste = "02"
    Case "03"
    bolExiste = "03"
    Case "05"
    bolExiste = "05"
    Case "04"
    bolExiste = "04"
    Case "09"
    bolExiste = "09"
    Case "10"
    bolExiste = "10"
    Case "11"
    bolExiste = "11"
    Case "12"
    bolExiste = "12"
    Case "20"
    bolExiste = "20"
    Case "50"
    bolExiste = "50"
    Case "60"
    bolExiste = "60"
    Case "70"
    bolExiste = "70"
    Case "80"
    bolExiste = "80"
    Case "90"
    bolExiste = "90"
    Case Else
     If IsNumeric(Val(StrCodigo)) Then
     bolExiste = Format(StrCodigo, "00")
     Else
     bolExiste = ""
     End If
    
End Select
    Verificar_No_Existe = bolExiste
End Function
Public Function Formatear_Fecha(xinicio As String)
If Len(xinicio) > 0 Then
   If Len(xinicio) = "6" Then
                If Val(Mid(xinicio, 1, 2)) > 70 And Val(Mid(xinicio, 1, 2)) <= 99 Then
                  If IsDate(Mid(xinicio, 5, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 1, 2)) Then
                   
                   If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                       'Formatear_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
                       Formatear_Fecha = "19" & Mid(xinicio, 1, 2) & Mid(xinicio, 3, 2) & Mid(xinicio, 5, 2)
                       
                   Else
                       Formatear_Fecha = ""
                   End If
                  Else
                   Formatear_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xinicio, 5, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 1, 2)) Then
                   
                     If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                        'Formatear_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                        Formatear_Fecha = "20" & Mid(xinicio, 1, 2) & Mid(xinicio, 3, 2) & Mid(xinicio, 5, 2)
                        
                     Else
                        Formatear_Fecha = ""
                     End If
                   
                   Else
                   Formatear_Fecha = ""
                   End If
                End If
    Else
       Formatear_Fecha = xinicio
    End If
 End If
End Function

Public Function Verificar_Firma(strdni As String, StrDatos As String)
Verificar_Firma = False
    If Len(Trim$(strdni)) > 0 Then
    Verificar_Firma = True
    Else
    Verificar_Firma = False
    End If
End Function

Public Function Verificar_Periodo(strPeriodo As String) As String
Dim VAR As String
Dim StrAno As String
Verificar_Periodo = ""
    
Select Case UCase(Mid(strPeriodo, 1, 3))
 Case "ENE"
 VAR = "01"
 Case "FEB"
 VAR = "02"
 Case "MAR"
 VAR = "03"
 Case "ABR"
 VAR = "04"
 Case "MAY"
 VAR = "05"
 Case "JUN"
 VAR = "06"
 Case "JUL"
 VAR = "07"
 Case "AGO"
 VAR = "08"
 Case "SEP"
 VAR = "09"
 Case "OCT"
 VAR = "10"
 Case "NOV"
 VAR = "11"
 Case "DIC"
 VAR = "12"
End Select
    StrAno = Mid(strPeriodo, InStr(strPeriodo, "/") + 1, Len(strPeriodo))
    
    
    Verificar_Periodo = StrAno & VAR
End Function

Public Function Validar_Fecha(strFecha As String, strFecha1 As String) As String
  Dim Validar_Fecha1 As String
  Dim Validar_Fecha2 As String
  
  Validar_Fecha = ""
  Validar_Fecha1 = ""
  Validar_Fecha2 = ""
  If Len(strFecha) > 0 Then
    If Len(strFecha) = 8 Then
    
                 If Val(Mid(strFecha, 1, 4)) > 1970 And Val(Mid(strFecha, 1, 4)) <= 1999 Then
                    If IsDate(Mid(strFecha, 7, 2) & "/" & Mid(strFecha, 5, 2) & "/" & Mid(strFecha, 1, 4)) Then
                      If Val(Mid(strFecha, 5, 2)) > 0 And Val(Mid(strFecha, 5, 2)) <= 12 Then
                          'Formatear_Fecha = Mid(STRFECHA, 1, 2) & "/" & Mid(STRFECHA, 3, 2) & "/" & "19" & Mid(STRFECHA, 5, 2)
                          Validar_Fecha1 = ""
                      Else
                          Validar_Fecha1 = "FECHA INVALIDA"
                      End If
                    Else
                         Validar_Fecha1 = "FECHA INVALIDA"
                    End If
                 Else
                     If IsDate(Mid(strFecha, 7, 2) & "/" & Mid(strFecha, 5, 2) & "/" & Mid(strFecha, 1, 4)) Then
                     
                       If Val(Mid(strFecha, 5, 2)) > 0 And Val(Mid(strFecha, 5, 2)) <= 12 Then
                          'Formatear_Fecha = Mid(STRFECHA, 1, 2) & "/" & Mid(STRFECHA, 3, 2) & "/" & "20" & Mid(STRFECHA, 5, 2)
                          Validar_Fecha1 = ""
                          
                       Else
                          Validar_Fecha1 = "FECHA INVALIDA"
                       End If
                     
                     Else
                     Validar_Fecha1 = "FECHA INVALIDA"
                     End If
                 End If
    Else
    Validar_Fecha1 = "FECHA INVALIDA"
    End If
  
  End If
  
  
  
  If Len(strFecha1) > 0 Then
    If Len(strFecha1) = 8 Then
    
                 If Val(Mid(strFecha1, 1, 4)) > 1970 And Val(Mid(strFecha1, 1, 4)) <= 1999 Then
                    If IsDate(Mid(strFecha1, 7, 2) & "/" & Mid(strFecha1, 5, 2) & "/" & Mid(strFecha1, 1, 4)) Then
                      If Val(Mid(strFecha1, 5, 2)) > 0 And Val(Mid(strFecha1, 5, 2)) <= 12 Then
                          'Formatear_Fecha = Mid(STRFECHA, 1, 2) & "/" & Mid(STRFECHA, 3, 2) & "/" & "19" & Mid(STRFECHA, 5, 2)
                          Validar_Fecha2 = ""
                      Else
                          Validar_Fecha2 = "FECHA INVALIDA"
                      End If
                    Else
                         Validar_Fecha2 = "FECHA INVALIDA"
                    End If
                 Else
                     If IsDate(Mid(strFecha1, 7, 2) & "/" & Mid(strFecha1, 5, 2) & "/" & Mid(strFecha1, 1, 4)) Then
                     
                       If Val(Mid(strFecha1, 5, 2)) > 0 And Val(Mid(strFecha1, 5, 2)) <= 12 Then
                          'Formatear_Fecha = Mid(STRFECHA, 1, 2) & "/" & Mid(STRFECHA, 3, 2) & "/" & "20" & Mid(STRFECHA, 5, 2)
                          Validar_Fecha2 = ""
                          
                       Else
                          Validar_Fecha2 = "FECHA INVALIDA"
                       End If
                     
                     Else
                     Validar_Fecha2 = "FECHA INVALIDA"
                     End If
                 End If
    Else
    Validar_Fecha2 = "FECHA INVALIDA"
    End If
  
  End If
  
  
  If Len(Validar_Fecha1) <> 0 Or Len(Validar_Fecha2) <> 0 Then
       Validar_Fecha = "FECHA INVALIDA"
  Else
       Validar_Fecha = ""
  End If
End Function

Public Function Validar_Observacion(StrCodigo) As String
Dim bolExiste As String
Select Case Format(StrCodigo, "00")
    Case "50"
    bolExiste = "JUBILADO"
    Case "60"
    bolExiste = "PENSIONISTA"
    Case "70"
    bolExiste = "PAGO EN LA ONP"
    Case "80"
    bolExiste = "DESAFILIADO"
    Case "90"
    bolExiste = "OTROS"
End Select
Validar_Observacion = bolExiste
End Function

Public Sub Grabar_TemporalIcr(Xperiodo As String, Xlote As String, Xfecha As String, _
                               Xnroliqui As String, Xbarra)
                              
Dim Dbo As Database
Dim Sql As String
Dim Rst As Recordset
Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")
'Sql = "SELECT * FROM tmpprpt"
'Set Rst = Dbo.OpenRecordset(Sql, dbOpenDynaset)
'Rst.AddNew
'Rst.Fields("periodo") = (Xperiodo)
'Rst.Fields("lote") = Xlote
'Rst.Fields("fecha") = (Xfecha)
'Rst.Fields("nroliq") = Xnroliqui
'Rst.Fields("barra") = Xbarra
'Rst.Update

Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")
Sql = "INSERT  into  tmpprpt  (periodo,lote,fecha,nroliq,barra) values ('" & Trim$(Xperiodo) & "','" & Trim$(Xlote) & "','" & Trim$(Xfecha) & "','" & Trim$(Xnroliqui) & "','" & Trim$(Xbarra) & "')"

Dbo.Execute Sql
End Sub

Public Sub Grabar_Temporal_Imagenes(Xperiodo As String, Xlote As String, Xfecha As String, _
                               Ximagenes As Integer, Ximagenest As Integer, Xanno As String)
                              
Dim Dbo As Database
Dim Sql As String
Dim Rst As Recordset
Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")
'Sql = "SELECT * FROM tmpprpt"
'Set Rst = Dbo.OpenRecordset(Sql, dbOpenDynaset)
'Rst.AddNew
'Rst.Fields("periodo") = (Xperiodo)
'Rst.Fields("lote") = Xlote
'Rst.Fields("fecha") = (Xfecha)
'Rst.Fields("imag") = Ximagenes
'Rst.Fields("anno") = Xanno
'Rst.Update

Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")
Sql = "INSERT  into  tmpprpt  (periodo,lote,fecha,imag,imagt,anno) values ('" & Trim$(Xperiodo) & "','" & Trim$(Xlote) & "','" & Trim$(Xfecha) & "'," & Trim$(Ximagenes) & "," & Trim$(Ximagenest) & ",'" & Trim$(Xanno) & "')"

Dbo.Execute Sql

End Sub

Public Sub Grabar_Temporal_Envios(Xorden As String, Xlote As String, Xfecha As String)
                              
Dim Dbo As Database
Dim Sql As String
Dim Rst As Recordset
Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")

Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")
Sql = "INSERT  into  tmpprpt  (orden,lote,fecha) values ('" & Trim$(Xorden) & "','" & Trim$(Xlote) & "','" & Trim$(Xfecha) & "')"

Dbo.Execute Sql
End Sub



Public Sub Existe_Temporal_icr()
     If Dir(App.Path & "\rpticr.dbf") = "" Then
     Else
     Kill App.Path & "\rpticr.dbf"
     End If
End Sub

'Public Sub Create_TemporalDBF()
'Dim NewDB As Database
'Dim NewTable As New TableDef
'Dim F01 As New Field, F02 As New Field, F03 As New Field
'Dim F04 As New Field, F05 As New Field
'
'On Error GoTo Error_Handling_Create_TemDBF
'Set NewDB = OpenDatabase(App.Path, 0, 0, "DBASE IV;")
'NewTable.Name = "Rtempo"
'F01.Name = "periodo": F01.Type = dbText: F01.Size = 15: NewTable.Fields.Append F01
'F02.Name = "lote": F02.Type = dbText: F02.Size = 15: NewTable.Fields.Append F02
'F03.Name = "fecha": F03.Type = dbDate:  NewTable.Fields.Append F03
'F04.Name = "nroliq": F04.Type = dbText: F04.Size = 15: NewTable.Fields.Append F04
'F05.Name = "barra": F05.Type = dbText: F05.Size = 20: NewTable.Fields.Append F05
'NewDB.TableDefs.Append NewTable
'NewDB.Close
'Exit Sub
'Error_Handling_Create_TemDBF:
'If Err.Number = 3010 Then
'  Resume Next
'End If
'End Sub

Public Sub delete_TemporalIcr()
                              
Dim Dbo As Database
Dim Sql As String
Dim Rst As Recordset
Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")
Sql = "delete * FROM tmpprpt"
Dbo.Execute Sql
End Sub

Public Sub Grabar_Rechazos(Xfecha As String, Xnroliqui As String, Xlote As String, xruc As String, xrason As String, xcomen As String, xcuspp As String)
                              
Dim Dbo As Database
Dim Sql As String
Dim Rst As Recordset


Set Dbo = OpenDatabase(App.Path, False, False, "DBASE IV")
Sql = "INSERT  into  tmpprpt  (fecha,nroliq,lote,ruc,razon,comen,cuspp) values ('" & Trim$(Xfecha) & "','" & Trim$(Xnroliqui) & "','" & Trim$(Xlote) & "','" & Trim$(xruc) & "','" & Trim$(xrason) & "','" & Trim$(xcomen) & "','" & Trim$(xcuspp) & "')"

Dbo.Execute Sql
End Sub


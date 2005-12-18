Attribute VB_Name = "basScan"
Const MAX_PATH = 255

Public lngMiSession As Long
Public PlantillaZDF As String
Public strLoteCaptura As String

Public bolModoForm As Boolean
Public clsDataScan As clsDataSQL

Public strFechaSeleccionada As String
Public strLoteActual As String
Public bolVolumenSeleccionado As Boolean
Public lngIdVolumenActual As Long
Public lngLoteActual As Long
Public Const cExtensionImagen = "TIF"
Public intMenu As Byte
Public Enum eBIF
    BIF_RETURNONLYFSDIRS = &H1            'Sólo directorios del sistema
    BIF_DONTGOBELOWDOMAIN = &H2           'No incluir carpetas de red
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_BROWSEFORCOMPUTER = &H1000        'Buscar PCs
    BIF_BROWSEFORPRINTER = &H2000         'Buscar impresoras
End Enum

Private Type BrowseInfo
    hwndOwner               As Long
    pIDLRoot                As Long             'Especifica dónde se empezará a mostrar
    pszDisplayName          As Long
    lpszTitle               As Long
    ulFlags                 As Long
    lpfnCallback            As Long
    lParam                  As Long
    iImage                  As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        (lpbi As BrowseInfo) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
        (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" _
        (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        (ByVal pidList As Long, ByVal lpBuffer As String) As Long



Const ATTR_DIRECTORY = 16

Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFind As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 64
    cAlternate As String * 14
End Type
Type Datos_Directorio
    Path As String
    Tamaño As Long
    NumDir As Long
    NumFic As Long
End Type


Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function FindWindow Lib "User32" Alias _
  "FindWindowA" (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long

  Private Declare Function PostMessage Lib "User32" Alias _
  "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
  ByVal wParam As Long, lParam As Any) As Long


Const WM_CLOSE = &H10
  

Public Function BrowseForFolder(ByVal hwndOwner As Long, ByVal sPrompt As String, Optional ByVal vFlags As eBIF) As String
    '
    Dim INull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    Dim lFlags As Long

    If Not IsMissing(vFlags) Then
        lFlags = CInt(vFlags)
    End If

    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = lFlags Or BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        INull = InStr(sPath, vbNullChar)
        If INull Then
            sPath = Left$(sPath, INull - 1)
        End If
    Else
        'Se ha pulsado en cancelar
        sPath = ""
    End If

    BrowseForFolder = sPath
End Function
Public Function AñadirBarra(Path As String) As String
     If Mid(Path, Len(Path), 1) <> "\" Then
          AñadirBarra = Path & "\"
     Else
          AñadirBarra = Path
     End If
End Function

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
Public Function LoteActivo() As Boolean
Dim bolEstadoLote As Boolean
Dim rsSql As Recordset
    bolEstadoLote = False
    Set rsSql = New Recordset
    Set rsSql = clsDataScan.Conexion.Execute("EXECUTE sp_EstadoScan " & lngIdVolumenActual & "," & lngLoteActual)
    If rsSql.RecordCount > 0 Then
        bolEstadoLote = rsSql!scan
    End If
    If rsSql.State = 1 Then rsSql.Close
    Set rsSql = Nothing
    LoteActivo = bolEstadoLote
End Function
'Function LeerIni(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
'    Dim lpString As String
'    Dim LTmp As Long
'    Dim sRetVal As String
'
'    If IsMissing(vDefault) Then
'        lpString = ""
'    Else
'        lpString = vDefault
'    End If
'
'    sRetVal = String$(255, 0)
'
'    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
'    If LTmp = 0 Then
'        LeerIni = lpString
'    Else
'        LeerIni = Left(sRetVal, LTmp)
'    End If
'End Function
Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    Dim LTmp As Long
    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub

Function ExisteDirVolumen(strDirVolumen As String) As Boolean
On Local Error GoTo Err:
Dim bolOk As Boolean
    bolOk = True
    If Dir(strDirVolumen, vbDirectory) = "" Then bolOk = False
    ExisteDirVolumen = bolOk
    Exit Function
Err:
    bolOk = False
    ExisteDirVolumen = bolOk
End Function




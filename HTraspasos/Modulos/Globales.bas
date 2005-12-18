Attribute VB_Name = "Globales"
Public gUsuario As String
Public gPassword As String
Public gServidor As String
Public gBaseDatos As String
Public sDate As String
Public sHour As String
Public sAnio As String
Public sTrim As String
Public gDate As String
Public gHour As String
Public gAnio As String
Public gTrim As String
Public gNomUsu As String
Public UsuarioA As String
Public PasswordA As String
Public gTipuso As String
Public Acc As String      'accesos del usuario que entra al sistema
Public ValorBarra As String

'* Se declara una variable de tipo cConexion
'Public Con As Sctro.cConexion
'Public Con As cConexion
Public comando As ADODB.Command
'* Definir una Matriz linea de tipo variant para retornar el resultado de la Búsqueta
'* Se utiliza con el buscador
Public ModResult()

Public vbGAFP As String 'Varible global de Afp
Public vbGLte As String
Public vbGNumfol As String
Public vbGNumpre As String
Public vbGNumpart As String
Public vbGTipLey As String
Public vbGFecsal As String
Public vbGTippla As String
Public vbGFecPag As String
Public vbGCodBco As String
Public vbGTippag As String
Public vbGMesd As String
Public vbGDsAfp As String
Public vbGDsBco As String
Public vbGDsTpag As String
Public vbGDsTpla As String
Public vbGRuta As String
Public vbGNimag As String
Public vbrefrescar As Boolean 'Si actualiza la informacion
Public vbCaracteresNumero As String


Public vbGRutaStand As String 'Ruta Standar de las imagenes

Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const HWND_TOPMOST = -1 ' permanente
Public Const HWND_NOTOPMOST = -2 ' no permanente
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'Global m_Report As CRAXDDRT.Report
Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer

Public Const VK_TAB = &H9
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)






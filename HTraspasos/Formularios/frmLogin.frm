VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2415
   ClientLeft      =   4380
   ClientTop       =   4275
   ClientWidth     =   4350
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1426.862
   ScaleMode       =   0  'User
   ScaleWidth      =   4084.415
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   2640
      TabIndex        =   1
      Top             =   255
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   1260
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   855
      Width           =   1275
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre de usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   375
      Width           =   1875
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1140
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsU As New Recordset        'Usuarios

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "2326", "si-006", "TraspasosPiloto"
'    Con.ConexionSQL "sa", "", "192.168.20.57", "BDTraspasosDig2"
'    DatosServidor
    
On Error GoTo ERRORES
Ruta = "" & App.Path & "\Arccon.cnx"
Open Ruta For Input Shared As #1
Cont = 0
While Not EOF(1)
   Conta = Conta + 1
      Select Case Conta
              Case 1
              Line Input #1, gUsuario
              Case 2
              Line Input #1, gPassword
              Case 3
              Line Input #1, gServidor
              Case 4
              Line Input #1, gBaseDatos
           End Select
        Wend
Close #1
'gUsuario = "sa"
If Trim(gPassword) <> "" Then
   gPassword = Trim(Desencrip(Desencrip(gPassword, 8), 8))
End If

DoEvents
    If Not Conectar Then Exit Sub
    DatosServidor
    anioCalculo = gAnio
    DatosServidor

    Call Con.AbrirTabla("select * from usuarios", RsU, adOpenStatic, adLockReadOnly)
    'Conectar
    'CargarRegistros RsU, "Usuarios"
    cmdOK.Enabled = False
    
Exit Sub
ERRORES:
Select Case Err.Number
  Case 53:
     MsgBox "No se configuró la conexión, Consulte con su Administrador de  Base de Datos", vbExclamation, "Mensaje"
    End
End Select

End Sub

Private Sub txtUserName_Change()
    cmdOK.Enabled = Len(txtUserName) > 0 And Len(txtPassword) > 0
End Sub

Private Sub txtPassword_Change()
    cmdOK.Enabled = Len(txtUserName) > 0 And Len(txtPassword) > 0
End Sub

Private Sub cmdOK_Click()
    With RsU
        .MoveFirst
        .Find "Login='" + txtUserName + "'"
        If .EOF Then
            MsgBox "Usuario No registrado", , "Sistema de Traspasos"
            txtUserName = ""
            txtPassword = ""
            txtUserName.SetFocus
        Else
            If Trim(txtPassword) = Trim(!clave) Then
                gNomUsu = txtUserName
                'DesConectar
                Unload Me
                MDIPrincipal.Show
            Else
                MsgBox "La Clave es Incorrecta", , "Sistema de Traspasos"
                txtPassword = ""
                txtPassword.SetFocus
            End If
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Function Encrip(P_dato, P_leng) As String
Ret_Val = ""
Xdato = P_dato + Space(8 - Len(P_dato))
For i = 1 To Len(Xdato)
    Xcar = Mid(Xdato, i, 1)
    Xpos = Chr(Asc(Xcar) + i)
    If Xpos = "'" Then
       Xpos = Chr(Asc(Xcar) + i + 3)
    End If
    Ret_Val = Ret_Val + Xpos
Next
Encrip = Ret_Val
End Function

Function Desencrip(P_dato, P_leng) As String
Ret_Val = ""
Xdato = P_dato + Space(8 - Len(P_dato))
For i = 1 To Len(Xdato)
    Xcar = Mid(Xdato, i, 1)
    Xpos = Chr(Asc(Xcar) - i)
    If Xpos = "#" Then
       Xpos = Chr(Asc(Xcar) - i - 3)
    End If
    Ret_Val = Ret_Val + Xpos
Next
Desencrip = Ret_Val
End Function


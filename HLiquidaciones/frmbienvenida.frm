VERSION 5.00
Begin VB.Form Flow_frmbienvenida 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Liquidacion Previa Horizonte"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5490
   ControlBox      =   0   'False
   DrawMode        =   3  'Not Merge Pen
   DrawStyle       =   5  'Transparent
   Icon            =   "frmbienvenida.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cmbUsuarios 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmbienvenida.frx":030A
         Left            =   2520
         List            =   "frmbienvenida.frx":030C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contraseña:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Flow_frmbienvenida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strPass As String
Private Sub cmbUsuarios_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
    Unload Flow_Frmpanel
End Sub
Private Sub cmdOk_Click()
Dim strPass As String
Dim strEstacion As String

    If txtPassword.Text <> "" Then
        cmbUsuarios.Text = UCase(cmbUsuarios.Text)
        strPass = ""
        strPass = EncryptText(UCase(txtPassword.Text), "DMM")
        If DevolverPassword(cmbUsuarios.Text) = True Then
            If strPass = strPassword Then
            'If strPass <> strPassword Then
                strEstacion = ComputerName
                strUsuarioActivo = cmbUsuarios.Text
                strEjecutar = "Update tb_Usuarios set ds_estado = 'A' , ds_estacion = '" & strEstacion & "' where ds_usuario = '" & strUsuarioActivo & "'"
                cnServer.Execute strEjecutar
                MenuUsuario (strPermisos)
                Unload Me
            Else
                MsgBox "Password incorrecto", vbCritical + vbOKOnly, "Password"
            End If
        Else
            MsgBox "Usuario no registrado", vbCritical + vbOKOnly, "Password"
        End If
    Else
        MsgBox "Ingrese el password", vbCritical + vbOKOnly, "Password"
    End If
End Sub
Private Sub Form_Load()
    Me.Left = (Flow_frmbienvenida.Width / 2) + (Me.Width / 2)
    Me.Top = (Flow_frmbienvenida.Height / 2) + (Me.Height / 2)
    
     rsDatos.CursorType = adOpenDynamic
     rsDatos.LockType = adLockBatchOptimistic
     rsDatos.ActiveConnection = cnServer

     If rsDatos.State = 1 Then rsDatos.Close
     rsDatos.Source = "Select * from TB_USUARIOS order by id_usuario"
     rsDatos.Open
     Cargar_Usuarios
     If rsDatos.State = 1 Then rsDatos.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err:
    If rsDatos.State = 1 Then rsDatos.Close
    Exit Sub
Err:
End Sub
Sub Cargar_Usuarios()
    cmbUsuarios.Clear
    If rsDatos.RecordCount > 0 Then
        rsDatos.MoveFirst
        Do While Not rsDatos.EOF
            cmbUsuarios.AddItem rsDatos!ds_Usuario
            rsDatos.MoveNext
        Loop
    End If
End Sub
Sub MenuUsuario(strAcceso As String)
On Local Error GoTo Err:
    If Trim(strAcceso) = "A" Then
       Flow_Frmpanel.mnulote.Visible = True
       Flow_Frmpanel.mnuasiglote.Visible = True
       'mdiMain.mnuSeleccionar.Visible = False
       '''MENU SELECCIONAR
       Flow_Frmpanel.mnuseleccionarlote.Visible = True
       '''
       Flow_Frmpanel.mnuusuarios.Visible = True
       'Frmpanel.mnuVentanas.Visible = True
       Flow_Frmpanel.mnusalir.Visible = True
       Flow_Frmpanel.mnuprocesos.Visible = True
       
        rsSqlDatos.ActiveConnection = cnServer
    Else
       Flow_Frmpanel.mnuarchivo.Visible = False
       Flow_Frmpanel.mnulote.Visible = True
       Flow_Frmpanel.mnuprocesos.Visible = False
       Flow_Frmpanel.mnuasiglote.Visible = False
       'Frmpanel.mnuVentanas.Visible = True
       Flow_Frmpanel.mnusalir.Visible = True
       
       Flow_Frmpanel.Toolbar1.Buttons(4).Enabled = False
       Flow_Frmpanel.Toolbar1.Buttons(2).Enabled = False
    End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, App.EXEName
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
        cmdOk_Click
    End If
End Sub

Private Function DevolverPassword(strUser As String) As Boolean
    DevolverPassword = False
    If rsDatos.State = 1 Then rsDatos.Close
    rsDatos.Source = "select ds_password,ds_permisos, ds_data,ds_lote from tb_usuarios where ds_usuario = '" & cmbUsuarios.Text & "'"
    rsDatos.Open
    If rsDatos.RecordCount > 0 Then
        strPassword = Trim("" & rsDatos!ds_password)
        strPermisos = Trim("" & rsDatos!ds_permisos)
        strUsuarioTablaDatos = Trim("" & rsDatos!ds_data)
        strUsuarioTablaLote = Trim("" & rsDatos!ds_lote)
        DevolverPassword = True
    End If
    If rsDatos.State = 1 Then rsDatos.Close
End Function


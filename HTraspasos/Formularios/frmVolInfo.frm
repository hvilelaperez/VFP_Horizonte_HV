VERSION 5.00
Begin VB.Form frmVolInfo 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información del Volumen"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox dirArchivos 
      Height          =   570
      Left            =   3960
      TabIndex        =   17
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   330
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdDirectorio 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Seleccione el directorio ..."
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   7050
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   7050
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7050
      Width           =   1215
   End
   Begin VB.TextBox txtFechaCierre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3720
      Width           =   5175
   End
   Begin VB.TextBox txtFechaInicio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3090
      Width           =   5175
   End
   Begin VB.TextBox txtEspacio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4365
      Width           =   5175
   End
   Begin VB.TextBox txtCantidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2430
      Width           =   5175
   End
   Begin VB.TextBox txtDirectorio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   5175
   End
   Begin VB.Label lblEspacio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño en disco"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   4125
      Width           =   1230
   End
   Begin VB.Label lblFechaFin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de cierre"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   3510
      Width           =   1140
   End
   Begin VB.Label lblFechaInicio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de inicio"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label lblImagenes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de Imagenes"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   2220
      Width           =   1590
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   750
      Width           =   690
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmVolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVolumen As Recordset
Dim RutaVol As String
Private Sub cmdAceptar_Click()
Dim intID As Long
    If txtNombre.Text <> "" Then
        Dim strDirVolumen As String
        'strDirVolumen = Mid(drvDrive.Drive, 1, 2) & "\" & txtNombre.Text
        strDirVolumen = RutaVol & "\" & txtNombre.Text
        
        strDirVolumen = UCase(strDirVolumen)
        If DirOK = True Then
            If Dir(strDirVolumen, vbDirectory) = "" Then
                If vbYes = MsgBox("Desea crear el directorio de volumen " & strDirVolumen, vbQuestion + vbYesNo, "Nuevo volumen") Then
                    If CrearDirectorio(strDirVolumen) = False Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
        'GuardarIni App.Path & "\digitacion.ini", "seteos", "DirectorioScan", "" & "" & UCase(Mid(drvDrive.Drive, 1, 2))
        txtNombre.Text = UCase(txtNombre.Text)
        Set rsVolumen = New Recordset
        If rsVolumen.State = 1 Then rsVolumen.Close
        Set rsVolumen = clsDataScan.Conexion.Execute("EXECUTE sp_VolumenNuevo [" & UCase(txtNombre.Text) & "], [" & strDirVolumen & "]")
        If rsVolumen.RecordCount > 0 Then
            frmVolumen.lstvolumen.AddItem txtNombre.Text
            frmVolumen.lstIDVolumen.AddItem rsVolumen!NuevoID
            frmVolumen.lstEstado.AddItem False
        End If
        If rsVolumen.State = 1 Then rsVolumen.Close
        Unload Me
    Else
        MsgBox "Complete el nombre del volumen", vbCritical, "Nuevo volumen"
    End If
End Sub
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
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdDirectorio_Click()
Dim strRuta As String
    strRuta = ""
    strRuta = BrowseForFolder(hWnd, "Selecciona el directorio donde se grabarán las imagenes del Volumen")
    If Len(strRuta) > 0 Then txtDirectorio.Text = UCase(strRuta)
End Sub
Private Sub drvDrive_Change()
    DirOK
End Sub
Function DirOK() As Boolean
On Error GoTo errores
Dim bolOk As Boolean
    bolOk = True
    dirArchivos.Path = drvDrive.Drive
    DirOK = bolOk
    Exit Function
errores:
    bolOk = False
    MsgBox "Drive no habilitado", vbCritical, "Drive"
    DirOK = bolOk
End Function
Private Sub Form_Load()
    If bolModoForm = True Then
        Caption = "Nuevo volumen"
        'drvDrive.Drive = LeerIni(App.Path & "\digitacion.ini", "seteos", "DirectorioScan")
        RutaVol = LeerIni(App.Path & "\digitacion.ini", "seteos", "DirectorioScan")
        lblRuta.Caption = "Seleccione la unidad donde se creará el volumen"
    Else
        Caption = "Información del volumen"
        lblRuta.Caption = "Directorio del volumen"
        If clsDataScan.ConexionActiva = True Then
            Set rsVolumen = New Recordset
            If rsVolumen.State = 1 Then rsVolumen.Close
            clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTotalImagenes " & lngIdVolumenActual)
            Set rsVolumen = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Volumenes,[*],'where idvolumen=" & lngIdVolumenActual & "',''")
            If rsVolumen.RecordCount > 0 Then
                txtNombre.Text = " " & "" & UCase(rsVolumen!nombre)
                txtDirectorio.Text = " " & "" & UCase(rsVolumen!Ruta)
                txtCantidad.Text = " " & "" & rsVolumen!total_imagenes
                txtEspacio.Text = " " & "" & rsVolumen!Tamaño
                txtFechaInicio.Text = " " & "" & rsVolumen!fecha_hora_inicio
                txtFechaCierre.Text = " " & "" & rsVolumen!fecha_hora_fin
                txtEspacio.Text = " " & "" & rsVolumen!Tamaño
            End If
            If rsVolumen.State = 1 Then rsVolumen.Close
        End If
    End If
    PrenderBotones bolModoForm
End Sub

Sub PrenderBotones(bolEstado As Boolean)
    cmdAceptar.Visible = bolEstado
    cmdCancelar.Visible = bolEstado
    cmdCerrar.Visible = Not bolEstado
    cmdAceptar.TabStop = bolEstado
    txtDirectorio.Visible = Not bolEstado
    'drvDrive.Visible = bolEstado
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rsVolumen = Nothing
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtEspacio_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtFechaCierre_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtFechaInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If bolModoForm = False Then KeyAscii = 0
End Sub
Private Sub txtDirectorio_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

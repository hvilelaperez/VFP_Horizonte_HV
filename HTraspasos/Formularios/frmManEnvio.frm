VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmManEnvio 
   Caption         =   "Mantenimiento de Envios"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
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
   ScaleHeight     =   4650
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DtgEnvio 
      Height          =   1935
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Envio"
         Caption         =   "Envio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FecSal"
         Caption         =   "Fec. Salida"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Envio"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtFecsal 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtenvio 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Envio"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Número de Envio"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmManEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private RsEnvio As ADODB.Recordset
Private strAccion As String


Private Sub cmdCancelar_Click()
RefrescarPantalla
End Sub

Private Sub cmdEliminar_Click()
Call EliminarEnvio(txtenvio.Text)
End Sub

Private Sub CmdGrabar_Click()
Grabar
End Sub

Private Sub cmdModificar_Click()
strAccion = "M"
Modificar
End Sub

Private Sub cmdNuevo_Click()
strAccion = "N"
Nuevo
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DtgEnvio_Click()
RefrescarPantalla
End Sub
Sub Nuevo()
HabilitarBotones False, False, True, True
txtenvio.Enabled = True
txtFecsal.Enabled = True
txtenvio = ""
txtFecsal = ""
txtenvio.SetFocus
End Sub
Sub Modificar()
HabilitarBotones False, False, True, True
txtenvio.Enabled = False
txtFecsal.Enabled = True
txtFecsal.SetFocus

End Sub
Sub Grabar()
On Error GoTo errores
    Dim cmdEnvio As Command
    Set cmdEnvio = New ADODB.Command
    
    With cmdEnvio
        .ActiveConnection = Con.Cnx
        
        If strAccion = "N" Then
            .CommandText = "pGrabarEnvio ?,?"
        Else
            .CommandText = "pModificarEnvio ?,?"
        End If
        .Parameters("@Envio") = Trim(txtenvio)
        .Parameters("@Fecsal") = Trim(txtFecsal)
       ' .Execute
        
        If Con.GetFEjecutarSQL(cmdEnvio) Then
               If strAccion = "N" Then
                    If Con.vbdev = "Exito" Then
                        CargarDatos
                        RefrescarPantalla
                        MsgBox "Envio Guardado Correctamente", 64, Caption
                    Else
                        MsgBox "Envio ya existe", vbCritical, Caption
                    End If
               Else
                    If Con.vbdev = "Exito" Then
                        CargarDatos
                        RefrescarPantalla
                        MsgBox "Envio Actualizado Correctamente", 64, Caption
                    Else
                        MsgBox "Envio no existe", vbCritical, Caption
                    End If
               End If
        
        End If
        
    End With
Exit Sub
errores:
MsgBox Err.Description, vbCritical, Caption
End Sub

Private Sub Form_Load()
Set RsEnvio = New ADODB.Recordset
CargarDatos
RefrescarPantalla
End Sub

Private Sub txtenvio_KeyPress(KeyAscii As Integer)
FEntDec txtenvio, KeyAscii, 7, 0
End Sub

Private Sub txtEnvio_LostFocus()
txtenvio.Text = Format(txtenvio.Text, "000000#")
End Sub

Private Sub txtFecsal_KeyPress(KeyAscii As Integer)
FVFecha txtFecsal, KeyAscii
End Sub

Private Sub txtFecsal_Validate(Cancel As Boolean)
ValidaFecha txtFecsal
End Sub
Sub RefrescarPantalla()
With RsEnvio
    If .RecordCount > 0 Then
        txtenvio = .Fields("Envio")
        txtFecsal = .Fields("Fecsal")
        HabilitarBotones True, True, False, False
    Else
        txtenvio = ""
        txtFecsal = ""
        HabilitarBotones True, False, False, False
    End If
End With
txtenvio.Enabled = False
txtFecsal.Enabled = False
End Sub
Sub CargarDatos()
Dim sqltexto As String
sqltexto = "select * from Envio order by envio desc"
With RsEnvio
    If .State = 1 Then .Close
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
    Set DtgEnvio.DataSource = RsEnvio
    
End With




End Sub
Sub HabilitarBotones(Nuevo As Boolean, Modificar As Boolean, Grabar As Boolean, Cancelar As Boolean)
cmdNuevo.Enabled = Nuevo
cmdModificar.Enabled = Modificar
cmdGrabar.Enabled = Grabar
cmdCancelar.Enabled = Cancelar
End Sub
Sub EliminarEnvio(Envio As String)

'Verificar q no tenga lotes asociado

Dim cmdEliminar As New ADODB.Command

If fgDatosql("select top 1 envio from ctrldoc where envio = '" & Envio & "'") = "" Then
    
    With cmdEliminar
        .ActiveConnection = Con.Cnx
        .CommandText = "Delete Ctrldoc where Envio = '" & Envio & "' "
        .Execute , , adExecuteNoRecords
    End With
    
    MsgBox "Envio Eliminado ", 64, Caption
    

Else
    MsgBox "El Envio tiene lotes asociados", vbCritical, Caption
End If

End Sub

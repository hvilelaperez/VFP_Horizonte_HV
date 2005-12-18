VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCdoc 
   Caption         =   "Ingreso Control Documentario"
   ClientHeight    =   7890
   ClientLeft      =   3045
   ClientTop       =   1530
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   8685
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agregar  Agencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   8295
      Begin VB.TextBox TxtCantImag 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DtCAgen 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Agencias"
      End
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   7080
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LblImag 
         AutoSize        =   -1  'True
         Caption         =   "Imagenes:"
         Height          =   195
         Left            =   6120
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Top             =   360
         Width           =   315
      End
      Begin VB.Label LblCantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Doc:"
         Height          =   195
         Left            =   5040
         TabIndex        =   17
         Top             =   360
         Width           =   765
      End
      Begin VB.Label LblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   5280
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DtGAgencia 
      Height          =   5535
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9763
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Lote"
         Caption         =   "Lote"
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
         DataField       =   "Agencia"
         Caption         =   "Agencia"
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
      BeginProperty Column02 
         DataField       =   "CantSol"
         Caption         =   "Cantidad"
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
      BeginProperty Column03 
         DataField       =   "CantImag"
         Caption         =   "Imagenes"
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
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtFecha 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtEnvio 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label LblTotImag 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Imag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7560
      TabIndex        =   21
      Top             =   3120
      Width           =   420
   End
   Begin VB.Label LblTotDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label LblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total Docs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   13
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   600
   End
   Begin VB.Label LblEnvio 
      AutoSize        =   -1  'True
      Caption         =   "Envio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "FrmCdoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public RsAFP As New Recordset        'AFP
Public RsAgenAFP As New Recordset    'Agencias AFP
Public RsCtrlD As New Recordset      'CtrlDoc
Dim Activado As Boolean
Dim xcodafp As String
Dim xcodigo As String
Dim xstatus As String
Dim xcodusu As String

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "DG-63\LASERFICHE", "BDTraspasosDig2"
'    DatosServidor
'***
'    Call Con.AbrirTabla("select * from tafp", RsAFP, adOpenStatic, adLockReadOnly)
    Call Con.AbrirTabla("select * from tagenafp order by DesAgen", RsAgenAFP, adOpenStatic, adLockReadOnly)
    Call Con.AbrirTabla("select * from CtrlDoc", RsCtrlD, adOpenStatic, adLockBatchOptimistic)
    Nuevo
    'llena datos al Datacombo
'    Set DtCAfp.RowSource = RsAFP
'    DtCAfp.ListField = "Descri"
    Set DtCAgen.RowSource = RsAgenAFP
    DtCAgen.ListField = "DesAgen"
    xcodusu = gNomUsu
    LblTotDoc = "0"
    LblTotImag = "0"
    xcodafp = "IN"
    xcodigo = ""
    xstatus = "1"
End Sub

'Private Sub DtCAfp_Click(Area As Integer)
'    If DtCAfp = "" Then Exit Sub
'    With RsAFP  'coneccion a tabla de afp
'        .MoveFirst
'        .Find "[descri]='" + DtCAfp + "'"
'        xcodafp = !codafp
'    End With
'End Sub

Private Sub DtCAgen_Click(Area As Integer)
    If DtCAgen = "" Then Exit Sub
    With RsAgenAFP  'coneccion a agencias afp
        .MoveFirst
        .Find "[DesAgen]='" + DtCAgen + "'"
        xcodigo = !Codigo
    End With
    'CmdAnadir.Enabled = True
End Sub

Private Sub DtGAgencia_SelChange(Cancel As Integer)
    Activado = True
End Sub

Sub Nuevo()
    With RsCtrlD
    'captura el ultimo numero de envio
    If .BOF And .EOF Then   'pregunta si tabla esta vacia
        TxtEnvio = 1
        TxtEnvio = Format(TxtEnvio, "0000#")
    Else
        .MoveLast           'desplaza hasta la ultima registro y suma 1
        TxtEnvio = Val(!Envio) + 1
        TxtEnvio = Format(TxtEnvio, "0000#")
    End If
    TxtFecha = Date
    DtCAfp = ""
    txtLote = ""
    DtCAgen = ""
    TxtCantidad = ""
    TxtCantImag = ""
    LblTotDoc = "0"
    LblTotImag = "0"
    
    CmdBuscar.Enabled = True
    CmdAgregar.Enabled = True
    CmdGrabar.Enabled = False
    CmdSalir.Enabled = True
    End With
End Sub

Private Sub CmdBuscar_Click()
    Dim strsql As String
    TxtEnvio = Format(TxtEnvio, "0000#")
    LblTotDoc = "0"
    'strSQL = "Select * From CtrlDoc Where " + Trim(Envio) + " like '" + Trim(TxtEnvio) & "%' Order By Codafp, Envio, Agencia"
    strsql = "Select * From CtrlDoc Where Envio = '" + Trim(TxtEnvio) & "' Order By Codafp, Lote, Agencia"
    If RsCtrlD.State = 1 Then RsCtrlD.Close
    RsCtrlD.Open strsql, Con.Cnx, adOpenStatic, adLockOptimistic
    
    Set DtGAgencia.DataSource = RsCtrlD
    DtGAgencia.Refresh
    TxtFecha = RsCtrlD.Fields("fecsal")
    'Formato_Grid
    If DtGAgencia.ApproxCount <> 0 Then
        RsCtrlD.MoveFirst
        For i = 1 To DtGAgencia.ApproxCount
            DtGAgencia.Col = 2
            LblTotDoc = Val(LblTotDoc) + Val(DtGAgencia.Text)
            DtGAgencia.Col = 3
            LblTotImag = Val(LblTotImag) + Val(DtGAgencia.Text)
            RsCtrlD.MoveNext
        Next i
'        Dim Comienzo, Final, numReg
'        Comienzo = DtGAgencia.Columns(1).CellText(DtGAgencia.FirstRow)
'        Final = DtGAgencia.Columns(1).CellText(DtGAgencia.ApproxCount)
'        numClientes = DtGAgencia.ApproxCount
'        'CmdVisualizar.Enabled = True
'        'txtCaptura.Text = "Se Filtraron los Clientes desde " & Comienzo & " a " & Final & " con un total de " & numClientes & " Clientes."
    Else
        MsgBox " No hay registros de ese envio "
    End If
    txtLote.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    If txtLote = "" Or Val(txtLote) = 0 Then
        MsgBox "Debe ingresar Lote ", vbCritical, "Error en Ingreso de Lote"
        txtLote.SetFocus
        Exit Sub
    End If
    If TxtCantidad = "" Or Val(TxtCantidad) = 0 Then
        MsgBox "Debe ingresar Cantidad ", vbCritical, "Error en Ingreso de Cantidad"
        TxtCantidad.SetFocus
        Exit Sub
    End If
    If TxtCantImag = "" Or Val(TxtCantImag) = 0 Then
        MsgBox "Debe ingresar cantidad de Imagenes ", vbCritical, "Error en Ingreso de cantidad de Imagenes"
        TxtCantImag.SetFocus
        Exit Sub
    End If
'    If xcodafp = "" Then
'        MsgBox "Debe ingresar AFP ", vbCritical, "Error en Ingreso de AFP"
'        DtCAfp.SetFocus
'        Exit Sub
'    End If
    If xcodigo = "" Then
        MsgBox "Debe ingresar Agencia ", vbCritical, "Error en Ingreso de Agencia"
        DtCAgen.SetFocus
        Exit Sub
    End If
    If DtCAgen = "" Then
        MsgBox "Debe ingresar Agencia ", vbCritical, "Error en Ingreso de Agencia"
        DtCAgen.SetFocus
        Exit Sub
    End If

    With RsCtrlD
        .AddNew
        .Fields("CodAfp") = xcodafp
        .Fields("FecSal") = CDate(TxtFecha)
        .Fields("Lote") = Val(txtLote)
        .Fields("Envio") = Format(TxtEnvio, "0000#")
        .Fields("Agencia") = Trim(xcodigo)
        .Fields("CantSol") = Val(TxtCantidad)
        .Fields("CantImag") = Val(TxtCantImag)
        .Fields("Status") = xstatus
        .Fields("CodUsu") = xcodusu
        .Fields("Fecdig") = gDate
        .Fields("Horini") = ObtenerHora
        .Fields("Horfin") = ObtenerHora
        '.Update
    End With
    LblTotDoc = Val(LblTotDoc) + Val(TxtCantidad)
    LblTotImag = Val(LblTotImag) + Val(TxtCantImag)
    txtLote = ""
    DtCAgen = ""
    TxtCantidad = ""
    TxtCantImag = ""
    CmdGrabar.Enabled = True
    txtLote.SetFocus
End Sub

Private Sub CmdGrabar_Click()
    RsCtrlD.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub CmdEliminar_Click()
    If Activado Then
        On Error GoTo Salir
        With RsCtrlD
            xstatus = .Fields("Status")
            If xstatus = "1" Then
                DtGAgencia.Col = 0
                DtGAgencia.Col = 2
                LblTotDoc = Val(LblTotDoc) - Val(DtGAgencia.Text)
                DtGAgencia.Col = 3
                LblTotImag = Val(LblTotImag) - Val(DtGAgencia.Text)
                .Delete
            Else
                MsgBox "No se puede eliminar registro con Datos", vbCritical, "Error en Eliminación "
            End If
        End With
    Else
        MsgBox "Selecciona la celda a eliminar", vbCritical, "Error en Eliminación "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

'Private Sub TxtEnvio_Change()
'    TxtEnvio = Format(TxtEnvio, "0000#")
'End Sub

Private Sub TxtEnvio_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtFecha.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
End Sub

Private Sub TxtFecha_LostFocus()
    ValidaFecha TxtFecha
End Sub

Private Sub TxtLote_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            DtCAgen.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtCantImag.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCantImag_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdAgregar.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Formato_Grid()
With DtGAgencia
    Set .DataSource = RsCtrlD
    .Refresh
    .Caption = "Cantidades por Agencia"
    .AllowUpdate = False
    .Columns(0).Width = 1300
    .Columns(1).Width = 1300
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
    .Splits(0).AllowRowSizing = False
    .Splits(0).SizeMode = dbgExact
End With
End Sub

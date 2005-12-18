VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmCdoc 
   Caption         =   "Ingreso Control Documentario"
   ClientHeight    =   9570
   ClientLeft      =   -225
   ClientTop       =   1035
   ClientWidth     =   15240
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkVis2 
      Caption         =   "Visualizar 2da Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   37
      Top             =   360
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   34
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   3360
      TabIndex        =   23
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtEnvio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtFecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LblEnvio 
         AutoSize        =   -1  'True
         Caption         =   "Envio:"
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
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   585
      End
      Begin VB.Label LblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         Left            =   1800
         TabIndex        =   26
         Top             =   360
         Width           =   645
      End
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   6615
      Left            =   9000
      TabIndex        =   22
      Top             =   1800
      Width           =   6015
      _Version        =   131074
      _ExtentX        =   10610
      _ExtentY        =   11668
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   411411200
      OcrZoneVisibility=   -3292
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3292
      sReserved2      =   -3292
      lReserved1      =   38
      lReserved2      =   38
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Frame Frame3 
      Caption         =   "Envio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdActEnvio 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   35
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Envios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstEnvio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lotes sin Envio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7095
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   2175
      Begin VB.CommandButton cmdActulizarLSE 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox lstSinEnvio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5160
         ItemData        =   "FrmCdoc.frx":0000
         Left            =   240
         List            =   "FrmCdoc.frx":0002
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   14
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   6
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   5
      Top             =   8640
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DtGAgencia 
      Height          =   3975
      Left            =   3360
      TabIndex        =   7
      Top             =   4440
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin Traspasos27.UsCBusqueda2 UsCBusqueda21 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   3735
      _ExtentX        =   6165
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agregar  Agencias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   3360
      TabIndex        =   10
      Top             =   1680
      Width           =   5535
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Documento"
         Height          =   735
         Left            =   1080
         TabIndex        =   29
         Top             =   1800
         Width           =   2895
         Begin VB.OptionButton optTraspaso 
            Caption         =   "Traspaso"
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
            Left            =   1560
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optContrato 
            Caption         =   "Afiliación"
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
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtCantImag 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   3120
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Actualizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   1080
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label LblImag 
         AutoSize        =   -1  'True
         Caption         =   "Imagenes"
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
         Left            =   2160
         TabIndex        =   15
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label LblLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
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
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label LblCantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Doc"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label LblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
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
         Left            =   240
         TabIndex        =   11
         Top             =   900
         Width           =   660
      End
   End
   Begin VB.Label LblTotImag 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10680
      TabIndex        =   17
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Imag"
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
      Left            =   10080
      TabIndex        =   16
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label LblTotDoc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label LblTotal 
      Caption         =   "Total Docs:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   8
      Top             =   240
      Width           =   645
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
Dim CodAgen As String  ' codigo de agencia
Dim RsImag As Recordset

Private Sub cmdActEnvio_Click()
ActualizarListaEnvios
End Sub

Private Sub cmdActulizarLSE_Click()
ListarLotesSinEnvio
End Sub

Private Sub cmdAsignar_Click()
AsignarLote
End Sub

Private Sub cmdNuevo_Click()
frmManEnvio.Show 1
End Sub

Private Sub cmdQuitar_Click()
If txtEnvio.Text <> "" And Me.TxtLote.Text <> "" Then
    QuitarLoteEnvio
End If
End Sub

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "DG-63\LASERFICHE", "BDTraspasosDig2"
'    DatosServidor
'***
'    Call Con.AbrirTabla("select * from tafp", RsAFP, adOpenStatic, adLockReadOnly)
    Call Con.AbrirTabla("select busqueda = DesAgen + space(30) +  '/'+ Codigo, *  from tagenafp order by DesAgen", RsAgenAFP, adOpenStatic, adLockReadOnly)
    Call Con.AbrirTabla("select * from CtrlDoc", RsCtrlD, adOpenStatic, adLockBatchOptimistic)
   ' Nuevo
    'llena datos al Datacombo
'    Set DtCAfp.RowSource = RsAFP
'    DtCAfp.ListField = "Descri"
    'Set DtCAgen.RowSource = RsAgenAFP
    'DtCAgen.ListField = "DesAgen"
    xcodusu = gNomUsu
    LblTotDoc = "0"
    LblTotImag = "0"
    xcodafp = "IN"
    CodAgen = ""
    xstatus = "1"
    
    Set UsCBusqueda21.cConexion = Con.Cnx
    
    'ucUbigeo.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    UsCBusqueda21.Campocondicion = "busqueda"
    UsCBusqueda21.Busqinicial = False
    UsCBusqueda21.nroCaracteres = 3
    
    Set UsCBusqueda21.RecordBusqueda = RsAgenAFP
    UsCBusqueda21.BusqRecord = True
    
    Set RsImag = New ADODB.Recordset
    ActualizarListaEnvios
    ListarLotesSinEnvio
    
End Sub

'Private Sub DtCAfp_Click(Area As Integer)
'    If DtCAfp = "" Then Exit Sub
'    With RsAFP  'coneccion a tabla de afp
'        .MoveFirst
'        .Find "[descri]='" + DtCAfp + "'"
'        xcodafp = !codafp
'    End With
'End Sub

'Private Sub DtCAgen_Click(Area As Integer)
'    If DtCAgen = "" Then Exit Sub
'    With RsAgenAFP  'coneccion a agencias afp
'        .MoveFirst
'        .Find "[DesAgen]='" + DtCAgen + "'"
'        xcodigo = !Codigo
'    End With
'    'CmdAnadir.Enabled = True
'End Sub

Private Sub DtGAgencia_SelChange(Cancel As Integer)
    Activado = True
    RefrescarDatos
    CargarImagenes
End Sub

Sub RefrescarDatos()

With RsCtrlD
    If .RecordCount > 0 Then
        TxtLote = .Fields("Lote")
        UsCBusqueda21.cTexto = fgDatosql("select desagen from tagenafp where Codigo = '" & .Fields("Agencia") & "'") & Space(50) & "/" & .Fields("Agencia")
        CodAgen = .Fields("Agencia")
        txtCantidad = .Fields("CantSol")
        TxtCantImag = .Fields("CantImag")
        optContrato.Value = IIf(.Fields("indTipDoc"), True, False)
        optTraspaso.Value = Not optContrato.Value
    End If
End With

End Sub


Sub Nuevo()
    With RsCtrlD
    'captura el ultimo numero de envio
    If .BOF And .EOF Then   'pregunta si tabla esta vacia
        txtEnvio = 1
        txtEnvio = Format(txtEnvio, "000000#")
    Else
        .MoveLast           'desplaza hasta la ultima registro y suma 1
        txtEnvio = val(!Envio) + 1
        txtEnvio = Format(txtEnvio, "000000#")
    End If
    TxtFecha = Date
    TxtLote = ""
    txtNumsol = ""
    TxtCantImag = ""
    LblTotDoc = "0"
    LblTotImag = "0"
    
    cmdBuscar.Enabled = True
    CmdAgregar.Enabled = True
    cmdGrabar.Enabled = False
    CmdSalir.Enabled = True
    End With
End Sub

Private Sub cmdBuscar_Click()
    Dim strsql As String
    txtEnvio = Format(txtEnvio, "000000#")
    LblTotDoc = "0"
    'strSQL = "Select * From CtrlDoc Where " + Trim(Envio) + " like '" + Trim(TxtEnvio) & "%' Order By Codafp, Envio, Agencia"
    strsql = "Select * From CtrlDoc Where Envio = '" + Trim(txtEnvio) & "' Order By Codafp, Lote, Agencia"
    If RsCtrlD.State = 1 Then RsCtrlD.Close
    RsCtrlD.Open strsql, Con.Cnx, adOpenStatic, adLockOptimistic
    
    Set DtGAgencia.DataSource = RsCtrlD
    
    DtGAgencia.Refresh
    
    'Formato_Grid
    If DtGAgencia.ApproxCount <> 0 Then
        TxtFecha = RsCtrlD.Fields("fecsal")
        RsCtrlD.MoveFirst
        For i = 1 To DtGAgencia.ApproxCount
            DtGAgencia.Col = 2
            LblTotDoc = val(LblTotDoc) + val(DtGAgencia.Text)
            DtGAgencia.Col = 3
            LblTotImag = val(LblTotImag) + val(DtGAgencia.Text)
            RsCtrlD.MoveNext
        Next i
'        Dim Comienzo, Final, numReg
'        Comienzo = DtGAgencia.Columns(1).CellText(DtGAgencia.FirstRow)
'        Final = DtGAgencia.Columns(1).CellText(DtGAgencia.ApproxCount)
'        numClientes = DtGAgencia.ApproxCount
'        'CmdVisualizar.Enabled = True
'        'txtCaptura.Text = "Se Filtraron los Clientes desde " & Comienzo & " a " & Final & " con un total de " & numClientes & " Clientes."
    Else
        MsgBox " No hay registros de ese envio ", 64, Caption
    End If
    TxtLote.SetFocus
    CodAgen = ""
End Sub

Private Sub CmdAgregar_Click()
    If TxtLote = "" Or val(TxtLote) = 0 Then
        MsgBox "Debe ingresar Lote ", vbCritical, "Error en Ingreso de Lote"
        TxtLote.SetFocus
        Exit Sub
    End If
    If txtCantidad = "" Or val(txtCantidad) = 0 Then
        MsgBox "Debe ingresar Cantidad ", vbCritical, "Error en Ingreso de Cantidad"
        txtCantidad.SetFocus
        Exit Sub
    End If
    If TxtCantImag = "" Or val(TxtCantImag) = 0 Then
        MsgBox "Debe ingresar cantidad de Imagenes ", vbCritical, "Error en Ingreso de cantidad de Imagenes"
        TxtCantImag.SetFocus
        Exit Sub
    End If
    If CodAgen = "" Then
        MsgBox "Debe ingresar Agencia ", vbCritical, "Error en Ingreso de Agencia"
        UsCBusqueda21.SetFocus
        Exit Sub
    End If
    
    If Not optContrato.Value And Not optTraspaso.Value Then
       MsgBox "Debe indicar que tipo de documentos contiene el Lote (Contratos o Traspasos) ", vbCritical, "Tipo de Documento"
       Exit Sub
    End If

'    With RsCtrlD
'        .AddNew
'        .Fields("CodAfp") = xcodafp
'        .Fields("FecSal") = CDate(TxtFecha)
'        .Fields("Lote") = Val(TxtLote)
'        .Fields("Envio") = Format(TxtEnvio, "0000#")
'        .Fields("Agencia") = Trim(CodAgen)
'        .Fields("CantSol") = Val(TxtCantidad)
'        .Fields("CantImag") = Val(TxtCantImag)
'        .Fields("Status") = xstatus
'        .Fields("CodUsu") = xcodusu
'        .Fields("Fecdig") = gDate
'        .Fields("Horini") = ObtenerHora
'        .Fields("Horfin") = ObtenerHora
'        '.Update
'    End With

    ActualizaCtrlDoc
    Call cmdBuscar_Click
    CodAgen = ""
    UsCBusqueda21.cTexto = ""
    
    LblTotDoc = val(LblTotDoc) + val(txtCantidad)
    LblTotImag = val(LblTotImag) + val(TxtCantImag)
    TxtLote = ""
    'DtCAgen = ""
    txtCantidad = ""
    TxtCantImag = ""
    cmdGrabar.Enabled = True
    TxtLote.SetFocus
    Call ListarLotesSinEnvio
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
                LblTotDoc = val(LblTotDoc) - val(DtGAgencia.Text)
                DtGAgencia.Col = 3
                LblTotImag = val(LblTotImag) - val(DtGAgencia.Text)
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

Private Sub lstEnvio_DblClick()
txtEnvio.Text = Trim(Left(lstEnvio.Text, 30))
TxtFecha.Text = Trim(Right(lstEnvio.Text, 30))
End Sub

Private Sub lstSinEnvio_Click()
'AsignarLote
End Sub

'Private Sub TxtEnvio_Change()
'    TxtEnvio = Format(TxtEnvio, "0000#")
'End Sub

Private Sub txtenvio_KeyPress(KeyAscii As Integer)
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
            UsCBusqueda21.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
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

Private Sub UsCBusqueda21_Pulsacion(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Activado = True
CodAgen = ""
If KeyAscii = 13 Then
  'If LLenarDatosUbigeo Then TabSig KeyAscii
   With UsCBusqueda21.RecordBusqueda
        .MoveFirst
        .Find "Codigo = '" & Trim(Mid(Trim(UsCBusqueda21.cTexto), InStr(UsCBusqueda21.cTexto, "/") + 1, 10)) & "'"
        If Not .EOF Then
            CodAgen = .Fields("Codigo")
            TabSig KeyAscii
        Else
            CodAgen = ""
            MsgBox "El Dato ingresado no existe", 64, "Mensaje"
        End If
  
  End With
    
End If

End Sub

Private Sub UsCBusqueda21_Validate(Cancel As Boolean)
If CodAgen = "" Then
   With UsCBusqueda21.RecordBusqueda
        .MoveFirst
        .Find "Codigo = '" & Trim(Mid(Trim(UsCBusqueda21.cTexto), InStr(UsCBusqueda21.cTexto, "/") + 1, 10)) & "'"
        If Not .EOF Then
            CodAgen = .Fields("Codigo")
            Cancel = False
        Else
            CodAgen = ""
            MsgBox "El Dato ingresado no existe", 64, "Mensaje"
        End If
  
  End With
End If
End Sub

Sub ActualizarListaEnvios()
lstEnvio.Clear
CargarListaSql lstEnvio, "Select Envio,fecsal from Envio order by Envio desc"

End Sub
Sub ListarLotesSinEnvio()
lstSinEnvio.Clear
CargarListaSql lstSinEnvio, "select lotefisico, vacio='' from _scan where Lotefisico not in ( Select lote from ctrldoc) and Estado = 1 "
End Sub
Sub AsignarLote()
TxtLote = Left(lstSinEnvio.list(lstSinEnvio.ListIndex), 10)
TxtLote.SetFocus

'VisualizarImagen
Call CargarImagenes


End Sub
Sub CargarImagenes()

On Error Resume Next
Dim strsql As String
'Lectura de 1 Imagen
strsql = " Select imagen, c.Ruta, LoteFisico, posicion " & _
         " from _Imagenes b join _Volumenes c on c.idvolumen = b.idvolumen " & _
         " where b.Lotefisico = '" & Trim(TxtLote) & "' order by imagen"
                 'idImagen is null
    If RsImag.State = 1 Then RsImag.Close
    RsImag.Open strsql, Con.Cnx, adOpenStatic, adLockReadOnly
    
    RsImag.MoveFirst
    
    If chkVis2.Value = 1 Then RsImag.MoveNext
      
    ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
    ImgEdit1.Display
    If Err.Number > 0 Then ImgEdit1.ClearDisplay
    
    ImgEdit1.FitTo 0
    
    
End Sub

Sub ActualizaCtrlDoc()
Dim cmdCtrlDoc As ADODB.Command
Set cmdCtrlDoc = New ADODB.Command
With cmdCtrlDoc
    .ActiveConnection = Con.Cnx
    .CommandText = "pActCtrlDoc ?,?,?,?,?,?,?,?,?,?,?,?,? "
    .Parameters("@CodAfp") = xcodafp
    .Parameters("@Fecsal") = CDate(TxtFecha)
    .Parameters("@Lote") = Trim(TxtLote)
    .Parameters("@Envio") = Format(txtEnvio, "000000#")
    .Parameters("@Agencia") = Trim(CodAgen)
    .Parameters("@CantSol") = val(txtCantidad)
    .Parameters("@CantImag") = val(TxtCantImag)
    .Parameters("@Status") = xstatus
    .Parameters("@CodUsu") = xcodusu
    .Parameters("@Fecdig") = gDate
    .Parameters("@Horini") = ObtenerHora
    .Parameters("@Horfin") = ObtenerHora
    .Parameters("@indTipDoc") = IIf(optContrato.Value, 1, 0)
    .Execute
End With

End Sub
Sub QuitarLoteEnvio()
Dim cmdQuitar As ADODB.Command
Set cmdQuitar = New ADODB.Command
With cmdQuitar
    .ActiveConnection = Con.Cnx
    .CommandText = "Delete CtrlDoc where Envio = '" & Trim(txtEnvio.Text) & "' and Lote = '" & Trim(TxtLote.Text) & "'"
    .Execute
    
    .ActiveConnection = Con.Cnx
    .CommandText = "Update _Scan set Estado = 1 where Lotefisico = '" & Trim(TxtLote.Text) & "'"
    .Execute
    
    
    
End With

Set cmdQuitar = Nothing
Call ListarLotesSinEnvio
Call cmdBuscar_Click

End Sub

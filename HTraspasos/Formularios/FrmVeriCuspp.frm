VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmVeriCuspp 
   Caption         =   "Verifica CUSPP"
   ClientHeight    =   10500
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   14550
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
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
      Height          =   1335
      Left            =   5760
      TabIndex        =   13
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton CmdCons 
         Caption         =   "&Consulta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   220
         Width           =   1095
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   280
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptSoloAfi 
         Caption         =   "Solo Afiliaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton OptSoloTra 
         Caption         =   "Solo Traspasos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   4575
         Begin VB.CheckBox chkDifFecNacDni 
            Caption         =   "Dif. DNI o Fec.Nac."
            Height          =   255
            Left            =   2520
            TabIndex        =   21
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkDifNomApe 
            Caption         =   "Dif. Nombres o Apell."
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Modo de Consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtEnvio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   4200
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optEnvio 
         Caption         =   "Por Envio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptLote 
         Caption         =   "Por Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Por Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1480
         MaxLength       =   8
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtFecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   1480
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DtGVeriCuspp 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   21
      BeginProperty Column00 
         DataField       =   "lote"
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
         DataField       =   "inddoc"
         Caption         =   "Tipo"
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
         DataField       =   "numsol"
         Caption         =   "N Solicitud"
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
         DataField       =   "secu"
         Caption         =   "sec"
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
      BeginProperty Column04 
         DataField       =   "cuspp"
         Caption         =   "CUSPP"
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
      BeginProperty Column05 
         DataField       =   "CUSSBS"
         Caption         =   "CUSPP SBS"
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
      BeginProperty Column06 
         DataField       =   "mcusp"
         Caption         =   "Digitado"
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
      BeginProperty Column07 
         DataField       =   "xcusp"
         Caption         =   "Algoritmo"
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
      BeginProperty Column08 
         DataField       =   "indcuserr"
         Caption         =   "Cus_Err"
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
      BeginProperty Column09 
         DataField       =   "apepat"
         Caption         =   "Ap Paterno"
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
      BeginProperty Column10 
         DataField       =   "apemat"
         Caption         =   "Ap Materno"
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
      BeginProperty Column11 
         DataField       =   "nombre1"
         Caption         =   "Nombre 1"
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
      BeginProperty Column12 
         DataField       =   "nombre2"
         Caption         =   "Nombre 2"
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
      BeginProperty Column13 
         DataField       =   "ApepatCus"
         Caption         =   "Apellidos y Nombres"
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
      BeginProperty Column14 
         DataField       =   "Numdoc"
         Caption         =   "NumDoc"
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
      BeginProperty Column15 
         DataField       =   "NumdocCus"
         Caption         =   "Numdoc Cus"
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
      BeginProperty Column16 
         DataField       =   "FecNac"
         Caption         =   "Fec.Nac"
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
      BeginProperty Column17 
         DataField       =   "Fecnaccus"
         Caption         =   "Fec. Nac. Cus"
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
      BeginProperty Column18 
         DataField       =   "idimagen"
         Caption         =   "IDimagen"
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
      BeginProperty Column19 
         DataField       =   "codusu"
         Caption         =   "Codusu"
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
      BeginProperty Column20 
         DataField       =   "UsuRepCus"
         Caption         =   "Usu. Mod."
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
            Alignment       =   2
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column19 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column20 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
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
      Left            =   13320
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   12000
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdVerImag 
      Caption         =   "Ver &Imagen"
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
      Left            =   10800
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   4695
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   14295
      _Version        =   131074
      _ExtentX        =   25215
      _ExtentY        =   8281
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   377961216
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   45
      lReserved2      =   45
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Label LblTotReg 
      Alignment       =   2  'Center
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
      Left            =   12120
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Registros"
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
      Left            =   12120
      TabIndex        =   6
      Top             =   720
      Width           =   1305
   End
End
Attribute VB_Name = "FrmVeriCuspp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVeriCuspp As New Recordset
Dim RsImag As New Recordset
Public xfrmrep As String
Public xnumsol As String
Public xidimagen As String
Public xlote As String
Dim xlote2 As String
Dim xsecu As String
Dim xcuspp As String
Dim xapepat As String
Dim xapemat As String
Dim xnombre1 As String
Dim xtipimag As String
Dim xrutavol As String
Dim ximagen As String
Dim xnumpag As Integer
Dim B1 As String
Dim Activado As Boolean

Private Sub DtGVeriCuspp_HeadClick(ByVal ColIndex As Integer)
    RsVeriCuspp.Sort = DtGVeriCuspp.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "2326", "si-006", "BDTraspasosdig2"
'    Con.ConexionSQL "sa", "", "192.168.20.57", "BDTraspasosdig2"
'    DatosServidor
'***
    cmdGrabar.Enabled = False
    CmdVerImag.Enabled = False
    xfrmrep = "C"
    TxtFecha = gDate
    B1 = ""
End Sub

Private Sub CmdCons_Click()
    Dim Condicion As String
    
    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = "delete tmpErrCuspp "
        .Execute
        
        
        If Me.OptFecha Then
           Condicion = " a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) order by a.lote "
        Else
           If optEnvio.Value Then
            Condicion = " b.Envio = '" & txtEnvio & "' order by a.lote "
           Else
            Condicion = " a.lote = '" & TxtLote & "' order by a.lote "
           End If
        End If
        
        
        .CommandText = "insert into tmpErrCuspp (fecsal,lote,inddoc,numsol,secu,apepat,apemat,nombre1,nombre2,cuspp,idimagen,posicion,codusu) " & _
        "SELECT a.fecsal,a.lote,a.inddoc,numsol,secu,apepat,apemat,nombre1,nombre2,cuspp,idimagen,posicion,a.codusu " & _
        "from solicitud a join ctrldoc b on a.lote = b.lote " & _
        "where " & Condicion
        
        
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2 = ApePat, " & _
        "apemat2 = ApeMat " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,'VDA DE ',''), " & _
        "apemat2=REPLACE(apemat2,'VDA DE ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,'VDA. DE ',''), " & _
        "apemat2=REPLACE(apemat2,'VDA. DE ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,'DE LA ',''), " & _
        "apemat2=REPLACE(apemat2,'DE LA ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,'DE LOS ',''), " & _
        "apemat2=REPLACE(apemat2,'DE LOS ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,'DEL ',''), " & _
        "apemat2=REPLACE(apemat2,'DEL ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,'DE ',''), " & _
        "apemat2=REPLACE(apemat2,'DE ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,'LA ',''), " & _
        "apemat2=REPLACE(apemat2,'LA ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,' Y ',''), " & _
        "apemat2=REPLACE(apemat2,' Y ','') " & _
        "From tmpErrCuspp "
        .Execute
        .CommandText = "update tmpErrCuspp " & _
        "set apepat2=REPLACE(apepat2,' ',''), " & _
        "apemat2=REPLACE(apemat2,' ','') " & _
        "From tmpErrCuspp "
        .Execute
    End With

    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""
    B1_5 = ""
    B1_6 = ""

'    SQLstr = "SELECT s.fecsal,s.lote,s.inddoc,s.numsol,s.secu,s.cuspp,mcusp=substring(t.cuspp,7,5),xcusp=substring(t.nombre1,1,1)+substring(t.apepat2,1,1)+substring(t.apemat2,1,1)+(case when len(t.apepat2)>3 then substring(t.apepat2,4,1) else substring(t.apepat2,len(t.apepat2),1) end)+(case when len(t.apemat2)>3 then substring(t.apemat2,4,1) else substring(t.apemat2,len(t.apemat2),1) end),s.apepat,s.apemat,s.nombre1,s.nombre2,s.idimagen,s.codusu " & _
'                "FROM tmpErrCuspp t, solicitud s " & _
'                "where s.lote = t.lote and s.numsol = t.numsol and s.secu = t.secu " & _
'                " and ( substring(t.cuspp,7,1)<> substring(t.nombre1,1,1) or substring(t.cuspp,8,1)<> substring(t.apepat2,1,1) or substring(t.cuspp,9,1)<> substring(t.apemat2,1,1) " & _
'                " or substring(t.cuspp,10,1)<> (case when len(t.apepat2)>3 then substring(t.apepat2,4,1) else substring(t.apepat2,len(t.apepat2),1) end) " & _
'                " or substring(t.cuspp,11,1)<> (case when len(t.apemat2)>3 then substring(t.apemat2,4,1) else substring(t.apemat2,len(t.apemat2),1) end) ) and t.cuspp<>'' ORDER BY s.lote,s.numsol "

    B1_1 = "SELECT s.fecsal,s.lote,s.inddoc,s.numsol,cussbs, s.secu,s.cuspp,mcusp=substring(t.cuspp,7,5),xcusp=substring(t.nombre1,1,1)+substring(t.apepat2,1,1)+substring(t.apemat2,1,1)+(case when len(t.apepat2)>3 then substring(t.apepat2,4,1) else substring(t.apepat2,len(t.apepat2),1) end)+(case when len(t.apemat2)>3 then substring(t.apemat2,4,1) else substring(t.apemat2,len(t.apemat2),1) end),s.apepat,s.apemat,s.nombre1,s.nombre2,s.numdoc, s.fecnac , s.apepatcus,s.apematcus,s.nombre1cus,s.nombre2cus,s.numdoccus, s.fecnaccus, s.idimagen,s.codusu,s.indcuserr, UsuRepCus "
    B1_2 = "FROM solicitud s join tmpErrCuspp t on s.lote = t.lote and s.numsol = t.numsol and s.secu = t.secu join ctrldoc c on s.lote = c.lote and s.fecsal = c.fecsal "
'*****
    If OptTodos.Value Then
        B1_3 = "WHERE s.indDoc is not NULL and s.secu=0 "
    End If
    If OptSoloAfi.Value Then
        B1_3 = "WHERE s.indDoc = 'C' and s.secu=0 "
    End If
    If OptSoloTra.Value Then
        B1_3 = "WHERE s.indDoc = 'T' and s.secu=0 "
    End If
    
    
'*****
    If OptFecha.Value Then
        If TxtFecha <> "  /  /   " And TxtFecha <> "" Then
            B1_4 = " and s.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
        Else
            MsgBox "Ingresar Fecha a buscar", vbCritical, "Error en la busqueda "
            TxtFecha.SetFocus
            Exit Sub
        End If
    Else
        If OptLote Then
            If val(TxtLote) > 0 Then
                B1_4 = " and s.Lote = '" & TxtLote & "' "
            Else
                MsgBox "Ingresar Lote a buscar", vbCritical, "Error en la busqueda "
                TxtLote.SetFocus
                Exit Sub
            End If
        Else
            If val(txtEnvio) > 0 Then
                B1_4 = " and c.Envio = '" & txtEnvio & "' "
            Else
                MsgBox "Ingresar Lote a buscar", vbCritical, "Error en la busqueda "
                txtEnvio.SetFocus
                Exit Sub
            End If
        End If
    End If
'*****

    If chkDifNomApe.Value = 1 And chkDifFecNacDni.Value = 1 Then
    B1_5 = " and ( s.cuspp <> isnull(cussbs,'') or substring(t.cuspp,7,1)<> substring(t.nombre1,1,1) or substring(t.cuspp,8,1)<> substring(t.apepat2,1,1) or substring(t.cuspp,9,1)<> substring(t.apemat2,1,1) " & _
                " or substring(t.cuspp,10,1)<> (case when len(t.apepat2)>3 then substring(t.apepat2,4,1) else substring(t.apepat2,len(t.apepat2),1) end) " & _
                " or substring(t.cuspp,11,1)<> (case when len(t.apemat2)>3 then substring(t.apemat2,4,1) else substring(t.apemat2,len(t.apemat2),1) end) ) and t.cuspp<>'' "
    Else
    
        If OptSoloTra.Value Then
        
            If chkDifNomApe.Value = 1 Then
                B1_5 = " and (ltrim(rtrim(s.Apepat)) + ' ' + ltrim(rtrim(s.Apemat)) + ' ' + ltrim(rtrim(s.Nombre1)) + ' ' + ltrim(rtrim(s.Nombre2)) <> isnull(s.ApepatCus,'')) "
            End If
            
            If chkDifFecNacDni.Value = 1 Then
                B1_5 = " and (s.Numdoc <> isnull(s.NumdocCus,'') or convert(datetime,isnull(s.Fecnac,''),103) <> convert(datetime,isnull(s.FecNacCus,''),103) ) "
            End If
            
            If chkDifNomApe.Value = 1 And chkDifFecNacDni.Value = 1 Then
                B1_5 = " and ( ltrim(rtrim(s.Apepat)) + ' ' + ltrim(rtrim(s.Apepat)) + ' ' + ltrim(rtrim(s.Nombre1)) + ' ' + ltrim(rtrim(s.Nombre2)) <> isnull(s.ApepatCus,'') or s.Numdoc <> isnull(s.NumdocCus,'') or  convert(datetime,isnull(s.Fecnac,''),103) <> convert(datetime,isnull(s.FecNacCus,''),103) ) "
            End If
        End If
    
    End If
        
'*****
    B1_6 = " ORDER BY s.lote,s.numsol,s.secu "

    B1 = B1_1 + B1_2 + B1_3 + B1_4 + B1_5 + B1_6
     
    LlenaControles
End Sub

Private Sub LlenaControles()
    LblTotReg = 0
    
    If RsVeriCuspp.State = 1 Then RsVeriCuspp.Close
    RsVeriCuspp.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    With RsVeriCuspp
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            Set DtGVeriCuspp.DataSource = Nothing
            DtGVeriCuspp.Refresh
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Verificacion"
            Exit Sub
        Else
            Set DtGVeriCuspp.DataSource = RsVeriCuspp
            DtGVeriCuspp.Refresh
            cmdGrabar.Enabled = True
            LblTotReg.Caption = RsVeriCuspp.RecordCount
            CargaImag
        End If
    End With
End Sub

Private Sub DtGVeriCuspp_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "C"
    If Activado Then
        On Error GoTo Salir
        With RsVeriCuspp
            xlote = .Fields("lote")
            xnumsol = .Fields("numsol")
            xidimagen = .Fields("idimagen")
            FrmMuestraImg.Show
        End With
    Else
        MsgBox "Selecciona la solicitud a buscar", vbCritical, "Error en la busqueda "
    End If
Salir:
   Exit Sub
End Sub

Private Sub CmdGrabar_Click()
    If Activado Then
        RsVeriCuspp.Fields("UsuRepCus") = gNomUsu
        RsVeriCuspp.UpdateBatch
        MsgBox "Registro Actualizado", 64, "Revisión de CUSPP"
    Else
        MsgBox "Seleccione la Solicitud a Modificar", vbCritical, "Error en la Modificacion "
    End If
End Sub

Private Sub txtEnvio_LostFocus()
 txtEnvio.Text = Format(txtEnvio.Text, "000000#")
End Sub

Private Sub TxtFecha_GotFocus()
    Marcado TxtFecha
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
    If KeyAscii = 13 Then CmdCons.SetFocus
End Sub

Private Sub TxtFecha_LostFocus()
    ValidaFecha TxtFecha
End Sub

Private Sub TxtLote_GotFocus()
    Marcado TxtLote
End Sub

Private Sub TxtLote_LostFocus()
    TxtLote = Format(TxtLote, "0000000#")
End Sub

Private Sub TxtLote_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdCons.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CargaImag()
    With RsVeriCuspp
        xlote = Format(!Lote, "0000000#")
        xlote2 = Format(!Lote, "0000000#")
        xnumsol = Format(!Numsol, "000000#")
        xtipimag = "S"
        xidimagen = "01" + xnumsol

    End With
    xnumpag = 1
      
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient
    
    SQLImag = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and TipImag = '" & xtipimag & "' and substring(idimagen,3,9)= '" & xnumsol & "' order by posicion "
    
    RsImag.Open SQLImag, Con.Cnx, adOpenStatic, adLockReadOnly
    
    With RsImag          'coneccion a _Imagenes
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            .MoveFirst
            MuestraImg
        End If
    End With
               
End Sub

Private Sub MuestraImg()
    
    With RsImag  'coneccion a _Imagenes
        .Find "[posicion]='" & xnumpag & "'"
'        .Find "[idimagen]='" & xidimagen & "'"
        
        If .EOF Then .MovePrevious
        ximagen = IIf(IsNull(!Imagen), "", !Imagen)
        xrutavol = .Fields("ruta")
        xrutavol = Replace(xrutavol, "\\", "\")
        xidvolumen = IIf(IsNull(!idvolumen), "", !idvolumen)
    End With
      
    ImgEdit1.Image = xrutavol & "\" & xlote2 & "\" & ximagen & ".TIF"
    ImgEdit1.Enabled = True

    ImgEdit1.FitTo 1
'    ImgEdit1.Zoom = 50
    ImgEdit1.Display
    
'    ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 3)
'    ImgEdit1.ScrollPositionX = 0
   
    ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    ImgEdit1.Refresh
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
' alto
    ImgEdit1.Height = FrmVeriCuspp.Height - 7800
' ancho
    ImgEdit1.Width = FrmVeriCuspp.Width - 300
' alto
'    DtGHV.Height = 1200 - ImgEdit1.Height
    ImgEdit1.Height = FrmVeriCuspp.Height - 6500
End Sub

Sub MovImg(KeyCode As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF12 Then
        If Not (RsImag.EOF And RsImag.BOF) Then
            RsImag.MoveNext
            If RsImag.EOF Then RsImag.MoveLast
            'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
            ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If

    If KeyCode = vbKeyF11 Then
        If Not (RsImag.EOF And RsImag.BOF) Then
            RsImag.MovePrevious
            If RsImag.BOF Then RsImag.MoveFirst
            'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
            ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If
End Sub

Private Sub ImgEdit1_DblClick()
    On Error Resume Next
    ImgEdit1.FitTo 1
End Sub

Private Sub ImgEdit1_KeyDown(KeyCode As Integer, Shift As Integer)
    MovImg KeyCode
End Sub

Private Sub ImgEdit1_MarkEnd(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MarkType As Integer, ByVal GroupName As String)
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

Private Sub ImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

'********************
Private Sub ImgEdit1_Click()
    On Error Resume Next
End Sub

Private Sub UpDown1_DownClick()
    On Error Resume Next
    If Not (RsImag.EOF And RsImag.BOF) Then
        RsImag.MoveNext
        If RsImag.EOF Then RsImag.MoveLast
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
        ImgEdit1.Image = Rutaimag(xrutavol, RsImag!lotefisico, RsImag!Imagen)
        ImgEdit1.Display
        ImgEdit1.Refresh
    End If
End Sub

Private Sub UpDown1_UpClick()
    On Error Resume Next
    If Not (RsImag.EOF And RsImag.BOF) Then
        RsImag.MovePrevious
        If RsImag.BOF Then RsImag.MoveFirst
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
        ImgEdit1.Image = Rutaimag(xrutavol, RsImag!lotefisico, RsImag!Imagen)
        ImgEdit1.Display
        ImgEdit1.Refresh
    End If
End Sub



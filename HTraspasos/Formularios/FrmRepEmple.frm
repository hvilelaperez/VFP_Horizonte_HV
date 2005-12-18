VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "Imgedit.ocx"
Begin VB.Form FrmRepEmple 
   Caption         =   "Reportes de Empleadores"
   ClientHeight    =   10815
   ClientLeft      =   1275
   ClientTop       =   1950
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   15030
   Begin VB.CommandButton cmdPrefijos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prefijos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame frmPrefijos 
      Height          =   3615
      Left            =   360
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ListBox lstPrefDir 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   360
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
      Begin VB.ListBox LstPrefLoc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   2880
         TabIndex        =   31
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3960
         TabIndex        =   30
         Top             =   3120
         Width           =   1215
      End
   End
   Begin VB.Frame FraRep 
      Caption         =   "Criterios "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1455
      Left            =   5520
      TabIndex        =   18
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton CmdCons 
         Caption         =   "&Consulta "
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
         Left            =   3840
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptErrUbigeo 
         Caption         =   "Error Ubigeo"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton OptRucB 
         Caption         =   "RUC = blanco"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton OptTodosF 
         Caption         =   "Todos "
         Height          =   255
         Left            =   3840
         TabIndex        =   23
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptVarEmple 
         Caption         =   "Varios Empleadores"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton OptDirB 
         Caption         =   "Direcc = blanco"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   300
         Width           =   1575
      End
      Begin VB.OptionButton OptErrTel 
         Caption         =   "Error Telefono"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   620
         Width           =   1455
      End
      Begin VB.OptionButton OptUbigeoB 
         Caption         =   "Ubigeo = b"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   620
         Width           =   1215
      End
   End
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
      Height          =   1455
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos los Registros"
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
         TabIndex        =   17
         Top             =   280
         Value           =   -1  'True
         Width           =   2055
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
         TabIndex        =   16
         Top             =   620
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
         TabIndex        =   15
         Top             =   960
         Width           =   1815
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
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2895
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
         Left            =   120
         TabIndex        =   28
         Top             =   1040
         Width           =   1215
      End
      Begin VB.TextBox txtEnvio 
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
         Height          =   360
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   27
         Top             =   960
         Width           =   1215
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
         TabIndex        =   13
         Top             =   645
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
         TabIndex        =   12
         Top             =   285
         Value           =   -1  'True
         Width           =   1335
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
         Height          =   360
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtFecha 
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
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdVerImag 
      Caption         =   "&Imagen"
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
      Left            =   12360
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdUbigeo 
      Caption         =   "Ubigeo"
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
      Left            =   12360
      TabIndex        =   6
      Top             =   1080
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
      Height          =   615
      Left            =   10920
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DtGEmple 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   6588
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
      ColumnCount     =   20
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
         DataField       =   "Numsol"
         Caption         =   "Solicitud"
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
         DataField       =   "Secuencial"
         Caption         =   "Sec"
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
         DataField       =   "NumRuc"
         Caption         =   "RUC"
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
         DataField       =   "RazSoc"
         Caption         =   "Raz. Soc."
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
         DataField       =   "PrefDirEmp"
         Caption         =   "Tipo Via"
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
         DataField       =   "Direccion"
         Caption         =   "Direccion"
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
         DataField       =   "Numero"
         Caption         =   "Numero"
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
         DataField       =   "PrefLocEmp"
         Caption         =   "Tipo Urba"
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
         DataField       =   "Urb"
         Caption         =   "Urbanizacion"
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
         DataField       =   "Distrito"
         Caption         =   "Distrito"
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
         DataField       =   "Provincia"
         Caption         =   "Provincia"
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
         DataField       =   "Departamento"
         Caption         =   "Departamento"
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
         DataField       =   "Ubigeo"
         Caption         =   "Ubigeo"
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
         DataField       =   "Telefono"
         Caption         =   "Telefono"
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
         DataField       =   "RiaAfil"
         Caption         =   "RIA Afil"
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
         DataField       =   "IDimagen"
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
      BeginProperty Column18 
         DataField       =   "codusu"
         Caption         =   "Cod Usu"
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
         DataField       =   "usuRepEmp"
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
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3165.166
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column18 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column19 
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
      Height          =   375
      Left            =   13680
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton CmdRepo 
      Caption         =   "&Reporte Vista Prev"
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
      Left            =   13680
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   5175
      Left            =   120
      TabIndex        =   26
      Top             =   5520
      Width           =   14775
      _Version        =   131074
      _ExtentX        =   26061
      _ExtentY        =   9128
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   350712320
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   55
      lReserved2      =   55
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
      Left            =   10920
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Registros"
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
      Left            =   10920
      TabIndex        =   9
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "FrmRepEmple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsConsEmple As New Recordset        'TempConsulta
Public RsRepoEmple As New Recordset        'TempConsulta
Dim RsImag As New Recordset
Public xnumsol As String
Public xidimagen As String
Public xlote As String
Dim xlote2 As String
Public xfrmrep As String
Dim xtipimag1 As String
Dim xtipimag2 As String
Dim xrutavol As String
Dim ximagen As String
Dim xnumpag As Integer
'Public xsecuencial As Integer
Dim xprimer As String
Dim SQLstr As String
Dim B1 As String
Dim Activado As Boolean

Private Sub cmdCerrar_Click()
frmPrefijos.Visible = False
End Sub

Private Sub cmdPrefijos_Click()
frmPrefijos.Visible = True
End Sub

Private Sub DtGEmple_HeadClick(ByVal ColIndex As Integer)
    RsConsEmple.Sort = DtGEmple.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "2326", "si-006", "BDTraspasosdig2"
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    DatosServidor
'***
    CmdGrabar.Enabled = False
    CmdVerImag.Enabled = False
    xfrmrep = "E"
    TxtFecha = gDate
    B1 = ""
    CargarListaSql lstPrefDir, "select des = rtrim(Coddir) +' - ' +AbreDir, Coddir from prefdir"
    CargarListaSql LstPrefLoc, "select des = rtrim(CodLoc) +' - ' +AbreLoc, CodLoc from prefloc"
End Sub

Private Sub CmdCons_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""
    B1_5 = ""
    B1_6 = ""
    
'    If OptTodos.Value Then
'        SQLstr = "SELECT a.lote,b.numsol,a.secu,b.secuencial,a.cantemple,b.numruc,b.direccion,b.numero,b.urb,b.distrito,b.provincia,b.departamento,b.ubigeo,b.telefono,b.riaafil,a.idimagen,b.codusu " & _
'        "FROM empleadores b, solicitud a " & _
'        "WHERE b.numsol = a.numsol and b.secuencial = a.secu and b.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) " & _
'        "ORDER BY a.lote,b.numsol,b.secuencial"
'    End If

    If OptTodos.Value Then
        B1_1 = "SELECT a.fecsal,a.lote,a.secu,b.secuencial,b.numsol,b.inddoc,a.cantemple,b.numruc,b.RazSoc, b.direccion,b.numero,b.urb,b.distrito,b.provincia,b.departamento,b.ubigeo,b.telefono,a.idimagen,b.codusu,a.posicion, prefdiremp, preflocemp, b.usuRepEmp "
        B1_2 = "FROM SOLICITUD a join EMPLEADORES b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.secuencial join CtrlDoc c on a.Lote = c.Lote and a.fecsal = c.fecsal "
        B1_3 = "WHERE b.indDoc is not NULL "
    End If
    If OptSoloAfi.Value Then
        B1_1 = "SELECT a.fecsal,a.lote,a.secu,b.secuencial,b.numsol,b.inddoc,b.numruc,b.direccion,b.RazSoc,b.numero,b.urb,b.distrito,b.provincia,b.departamento,b.ubigeo,b.telefono,b.riaafil,a.idimagen,b.codusu,a.posicion, prefdiremp, preflocemp, b.usuRepEmp "
        B1_2 = "FROM SOLICITUD a join EMPLEADORES b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.secuencial join CtrlDoc c on a.Lote = c.Lote and a.fecsal = c.fecsal "
        B1_3 = "WHERE b.indDoc = 'C' "
    End If
    If OptSoloTra.Value Then
        B1_1 = "SELECT a.fecsal,a.lote,a.secu,b.secuencial,b.numsol,b.inddoc,b.numruc,b.direccion,b.RazSoc,b.numero,b.urb,b.distrito,b.provincia,b.departamento,b.ubigeo,b.telefono,b.riaafil,a.idimagen,b.codusu,a.posicion,prefdiremp, preflocemp, b.usuRepEmp "
        B1_2 = "FROM SOLICITUD a join EMPLEADORES b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.secuencial join CtrlDoc c on a.Lote = c.Lote and a.fecsal = c.fecsal  "
        B1_3 = "WHERE b.indDoc = 'T' "
    End If
'*****
    If OptFecha.Value Then
        If Trim(TxtFecha) <> "" Or Trim(TxtFecha) <> "//" Then
            B1_4 = " and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
        Else
            MsgBox "Debe Ingresar Fecha a buscar", vbCritical, "Búsqueda"
            TxtFecha.SetFocus
            Exit Sub
        End If
    Else
        If OptLote Then
            If Trim(TxtLote) <> "" Then
                B1_4 = " and a.Lote = '" & TxtLote & "' "
            Else
                MsgBox "Debe Ingresar Lote a buscar", vbCritical, "Búsqueda"
                TxtLote.SetFocus
                Exit Sub
            End If
        Else
            If Trim(txtEnvio) <> "" Then
                B1_4 = " and c.Envio = '" & txtEnvio & "' "
            Else
                MsgBox "Debe Ingresar Envio a buscar", vbCritical, "Búsqueda"
                txtEnvio.SetFocus
                Exit Sub
            End If
            B1_4 = " and c.Envio = '" & txtEnvio & "' "
            
        End If
    End If
'*****
    If OptTodosF.Value Then
        B1_5 = ""
    End If
    If OptDirB.Value Then
        B1_5 = " and b.direccion='' "
    End If
    If OptUbigeoB.Value Then
        B1_5 = " and b.ubigeo = '' "
    End If
    If OptErrUbigeo.Value Then
        B1_5 = " and ( b.ubigeo is null or ( b.ubigeo like '__00__' ) or ( b.ubigeo like '____00' ) or b.distrito is null or b.provincia is null or a.departamento is null )  "
    End If
    If OptRucB.Value Then
        B1_5 = " and ( b.numruc='' or b.numruc is null ) "
    End If
    If OptErrTel.Value Then
        B1_5 = " and ((len(b.telefono)<7 and b.telefono<>'') or b.telefono like '%-%' or b.telefono='' or b.telefono like '%11111%' or b.telefono is null ) "
    End If
    If OptVarEmple.Value Then
        B1_1 = "SELECT distinct a.fecsal,a.lote,a.secu,b.secuencial,b.numsol,b.inddoc,a.cantemple,b.numruc,b.direccion,b.numero,b.urb,b.distrito,b.provincia,b.departamento,b.ubigeo,b.telefono,a.idimagen,b.codusu, a.posicion "
        B1_5 = " and (a.cantemple>1 or a.cantemple=0) "
    End If
       
'    If OptVarEmple.Value Then
'        SQLstr = "SELECT distinct a.lote,b.numsol,b.secuencial,b.numruc,b.direccion,b.numero,b.urb,b.distrito,b.provincia,b.departamento,b.ubigeo,b.telefono,b.riaafil,a.idimagen,empleadores.codusu FROM empleadores b, solicitud a WHERE b.numsol = a.numsol and b.secuencial = a.secu and a.cantemple>1 and b.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
'    End If
   
'*****
    B1_6 = " order by a.lote,b.numsol,b.secuencial "

    B1 = B1_1 + B1_2 + B1_3 + B1_4 + B1_5 + B1_6
    
    Call CargarEmple

End Sub

Sub CargarEmple()
    LblTotReg = 0
    If RsConsEmple.State = 1 Then RsConsEmple.Close
    
    RsConsEmple.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
       
    With RsConsEmple
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Empleadores"
            DtGEmple.Refresh
            Exit Sub
        Else
            Set DtGEmple.DataSource = RsConsEmple
            DtGEmple.Refresh
            CmdGrabar.Enabled = True
            LblTotReg.Caption = RsConsEmple.RecordCount
            CargaImag
        End If
    End With
End Sub

Private Sub CmdRepo_Click()
    If RsRepoEmple.State = 1 Then RsRepoEmple.Close
    If OptTodos.Value Then
        SQLstr = "SELECT solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If
    If OptRia.Value Then
        SQLstr = "SELECT solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and empleadores.riaafil<=1 and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If
    If OptRucB.Value Then
        SQLstr = "SELECT solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and (numruc='' or numruc is null) and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If
    If OptDirB.Value Then
        SQLstr = "SELECT solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and empleadores.direccion='' and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If
    If OptErrTel.Value Then
        SQLstr = "SELECT solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and len(empleadores.telefono)<7 and empleadores.telefono<>'' and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If
    If OptVarEmple.Value Then
        SQLstr = "SELECT distinct solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and cantemple>1 and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If
    If OptUbigeoB.Value Then
        SQLstr = "SELECT solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and empleadores.ubigeo = '' and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If
    If OptErrUbigeo.Value Then
        SQLstr = "SELECT solicitud.lote,empleadores.numsol,secuencial,numruc,empleadores.direccion,empleadores.numero,empleadores.urb,empleadores.distrito,empleadores.provincia,empleadores.departamento,empleadores.ubigeo,empleadores.telefono,empleadores.riaafil,solicitud.idimagen,empleadores.codusu FROM empleadores, solicitud, ubigeo WHERE empleadores.numsol = solicitud.numsol and empleadores.secuencial = solicitud.secu and empleadores.ubigeo = ubigeo.codgeo and (empleadores.distrito <> ubigeo.distrito or empleadores.provincia <> ubigeo.provincia or empleadores.departamento <> ubigeo.departamento or ( empleadores.ubigeo like '__00__' ) or ( empleadores.ubigeo like '____00' )) and empleadores.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote,empleadores.numsol,secuencial"
    End If

    RsRepoEmple.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtREmple.DataSource = RsRepoEmple
    DtREmple.Show
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

Private Sub CmdGrabar_Click()
    RsConsEmple.Fields("usuRepEmp") = gNomUsu
    RsConsEmple.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub CmdSalir_Click()
    xfrmrep = ""
    Unload Me
End Sub

Private Sub CmdUbigeo_Click()
    FrmConsUbi.Show
End Sub

Private Sub DtGEmple_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "E"
    If Activado Then
        On Error GoTo Salir
        With RsConsEmple
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

Private Sub CargaImag()
    With RsConsEmple
        xlote = Format(!Lote, "0000000#")
        xlote2 = Format(!Lote, "0000000#")
        xnumsol = Format(!Numsol, "000000#")
        xtipimag1 = "S"
        xtipimag2 = "E"
        xidimagen = "01" + xnumsol

    End With
    xnumpag = 1
      
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient
    
    SQLImag = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and ( TipImag = '" & xtipimag1 & "' or TipImag = '" & xtipimag2 & "' ) and substring(idimagen,3,9)= '" & xnumsol & "' and posicion = " & RsConsEmple.Fields("posicion") & " "
    
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
   
    ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 9) * 2)
    ImgEdit1.ScrollPositionX = 0
   
'    ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    
    ImgEdit1.Refresh
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
' alto
    ImgEdit1.Height = FrmRepEmple.Height - 7800
' ancho
    ImgEdit1.Width = FrmRepEmple.Width - 300
' alto
'    DtGHV.Height = 1200 - ImgEdit1.Height
    ImgEdit1.Height = FrmRepEmple.Height - 6250
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




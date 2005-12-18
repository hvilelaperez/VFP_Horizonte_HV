VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmRepHV 
   Caption         =   "Revision Datos Hoja de Vida"
   ClientHeight    =   10545
   ClientLeft      =   1140
   ClientTop       =   1275
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   14760
   Begin VB.CommandButton cmdVerAyuda 
      Caption         =   "Prefijos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   42
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame frmPrefijos 
      Height          =   3855
      Left            =   9240
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   5415
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
         TabIndex        =   41
         Top             =   480
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
         TabIndex        =   40
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   39
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Urbanización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdCargos 
      Caption         =   "Cargos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   28
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CmdCons 
      Caption         =   "&Consulta "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   11640
      TabIndex        =   27
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton CmdPromo 
      Caption         =   "Promotor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   25
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton CmdVerImag 
      Caption         =   "&Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   5
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton CmdSucur 
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton CmdNacion 
      Caption         =   "Nacionalidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
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
      Left            =   11280
      TabIndex        =   22
      Top             =   120
      Width           =   3375
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
         Left            =   240
         TabIndex        =   37
         Top             =   1030
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
         Height          =   330
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   36
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
         Left            =   240
         TabIndex        =   24
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
         Left            =   240
         TabIndex        =   23
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
         Height          =   330
         Left            =   1680
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
         Height          =   330
         Left            =   1680
         TabIndex        =   0
         Top             =   240
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
      Height          =   1095
      Left            =   11280
      TabIndex        =   21
      Top             =   1560
      Width           =   1575
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   280
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptSoloAfi 
         Caption         =   "Afiliaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton OptSoloTra 
         Caption         =   "Traspasos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   800
         Width           =   1335
      End
   End
   Begin VB.Frame FraRep 
      Caption         =   "Criterios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   12960
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
      Begin VB.OptionButton OptmailErr 
         Caption         =   "Email Error"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton OptEmail 
         Caption         =   "Email marca"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton OptUbiSuc 
         Caption         =   "Ubigeo Suc=b"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton OptUbiCorre 
         Caption         =   "Ubigeo Corre=b"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton OptPromERR 
         Caption         =   "Promotor Error"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton OptRentaB 
         Caption         =   "Renta = 0"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptCInterB 
         Caption         =   "Cod Interno = b"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton OptEstCivB 
         Caption         =   "Est Civil = b"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptFVisitaB 
         Caption         =   "Fec Visita = b"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptTodosF 
         Caption         =   "Todos"
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
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
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
      Left            =   11280
      TabIndex        =   4
      Top             =   5400
      Width           =   975
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
      Height          =   375
      Left            =   12480
      TabIndex        =   2
      Top             =   6480
      Width           =   975
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
      Left            =   13560
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   6735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   11055
      _Version        =   131074
      _ExtentX        =   19500
      _ExtentY        =   11880
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   377962752
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
   Begin MSDataGridLib.DataGrid DtGHV 
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   6960
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   6165
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
      ColumnCount     =   39
      BeginProperty Column00 
         DataField       =   "Lote"
         Caption         =   "Lote"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "sechv"
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
         DataField       =   "nomcorres"
         Caption         =   "Nombre Corresp"
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
         DataField       =   "codcargo"
         Caption         =   "C Cargo"
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
         DataField       =   "celular"
         Caption         =   "Celular"
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
         DataField       =   "fecnaci"
         Caption         =   "Fec Nacim"
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
         DataField       =   "ubigeonacren"
         Caption         =   "Ubi Reniec"
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
         DataField       =   "ubigeonacine"
         Caption         =   "Ubi INEI"
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
         DataField       =   "nacion"
         Caption         =   "C Nacion"
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
         DataField       =   "sexo"
         Caption         =   "Sexo"
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
         DataField       =   "codestciv"
         Caption         =   "Est Civil"
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
         DataField       =   "fecsus"
         Caption         =   "Fec Visita"
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
         DataField       =   "codinpro"
         Caption         =   "Cod Inter"
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
         DataField       =   "IndCIntErr"
         Caption         =   "Error C. Int"
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
         DataField       =   "renta"
         Caption         =   "Renta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "IndRam"
         Caption         =   "Sin Renta"
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
         DataField       =   "indenvmail"
         Caption         =   "Env mail"
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
         DataField       =   "indmail"
         Caption         =   "Casa/Trab"
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
         DataField       =   "emaildoc"
         Caption         =   "email Casa"
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
      BeginProperty Column21 
         DataField       =   "emailtra"
         Caption         =   "email Trabajo"
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
      BeginProperty Column22 
         DataField       =   "PrefDirCor"
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
      BeginProperty Column23 
         DataField       =   "desdomcor"
         Caption         =   "Direcc Correspondencia"
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
      BeginProperty Column24 
         DataField       =   "nrodomcor"
         Caption         =   "Nro Corresp"
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
      BeginProperty Column25 
         DataField       =   "PrefLocCor"
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
      BeginProperty Column26 
         DataField       =   "urbadomcor"
         Caption         =   "Urb Correspondencia"
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
      BeginProperty Column27 
         DataField       =   "codgeocor"
         Caption         =   "Ubi Corresp"
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
      BeginProperty Column28 
         DataField       =   "codsuc"
         Caption         =   "Cod Suc"
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
      BeginProperty Column29 
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
      BeginProperty Column30 
         DataField       =   "direemp"
         Caption         =   "Sede Trabajo"
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
      BeginProperty Column31 
         DataField       =   "nroemp"
         Caption         =   "Nro Trabajo"
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
      BeginProperty Column32 
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
      BeginProperty Column33 
         DataField       =   "urbaemp"
         Caption         =   "Urb Trabajo"
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
      BeginProperty Column34 
         DataField       =   "codgeo"
         Caption         =   "Ubi Trab"
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
      BeginProperty Column35 
         DataField       =   "teleofi"
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
      BeginProperty Column36 
         DataField       =   "anexofi"
         Caption         =   "Anexo"
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
      BeginProperty Column37 
         DataField       =   "refdirec"
         Caption         =   "Referencia"
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
      BeginProperty Column38 
         DataField       =   "codusu"
         Caption         =   "CodUsu"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column14 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column16 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column18 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column19 
            Alignment       =   2
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column25 
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column27 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column28 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column29 
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column32 
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column34 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column35 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column36 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column38 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Descrip Cargo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   200
      Left            =   11400
      TabIndex        =   33
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label LblCargo 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   11400
      TabIndex        =   32
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Label LblCodcargo 
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
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11400
      TabIndex        =   31
      Top             =   4200
      Width           =   735
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
      Left            =   11520
      TabIndex        =   11
      Top             =   2940
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
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   11520
      TabIndex        =   10
      Top             =   2700
      Width           =   1260
   End
End
Attribute VB_Name = "FrmRepHV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsHV As New Recordset
Dim RsImag As New Recordset
Dim RsCargo As New Recordset
Public xnumsol As String
Public xidimagen As String
Public xlote As String
Dim xlote2 As String
Public xfrmrep As String
Dim xtipimag As String
Dim xrutavol As String
Dim ximagen As String
Dim xnumpag As Integer
'Public xsecuencial As Integer
Dim xcodcargo As String
Dim Activado As Boolean
Dim B1 As String

Private Sub cmdCerrar_Click()
frmPrefijos.Visible = False
End Sub

Private Sub DtGHV_HeadClick(ByVal ColIndex As Integer)
    RsHV.Sort = DtGHV.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
    cmdGrabar.Enabled = False
    CmdVerImag.Enabled = False
    xfrmrep = "V"
    TxtFecha = gDate
    B1 = ""
    LblCargo = ""
    CargarListaSql lstPrefDir, "select des = rtrim(Coddir) +' - ' +AbreDir, Coddir from prefdir"
    CargarListaSql LstPrefLoc, "select des = rtrim(CodLoc) +' - ' +AbreLoc, CodLoc from prefloc"
End Sub

Private Sub CmdCons_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""

    If OptTodos.Value Then
        B1_1 = "SELECT a.fecsal,b.lote,a.numsol,a.inddoc,a.sechv,a.nomcorres,a.codcargo,a.celular,a.fecnaci,a.ubigeonacren,a.ubigeonacine,a.nacion,a.sexo,b.indenvmail,b.indmail,a.emaildoc,a.emailtra,a.desdomcor,a.nrodomcor,a.urbadomcor,a.codgeocor,a.codsuc,a.direemp,a.nroemp,a.urbaemp,a.codgeo,a.teleofi,a.anexofi,a.refdirec,a.codestciv,a.fecsus,a.codinpro,a.renta,b.codusu,indRam,IndCIntErr,PrefDirCor, PrefLocCor, PrefDirEmp, PrefLocEmp " & _
        "FROM HOJAVIDA a join SOLICITUD b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.sechv=b.secu join Ctrldoc c on b.lote = c.lote and b.fecsal = c.fecsal   " & _
        "WHERE a.indDoc is not NULL "
    End If
    If OptSoloAfi.Value Then
        B1_1 = "SELECT a.fecsal,b.lote,a.numsol,a.inddoc,a.sechv,a.nomcorres,a.codcargo,a.celular,a.fecnaci,a.ubigeonacren,a.ubigeonacine,a.nacion,a.sexo,b.indenvmail,b.indmail,a.emaildoc,a.emailtra,a.desdomcor,a.nrodomcor,a.urbadomcor,a.codgeocor,a.codsuc,a.direemp,a.nroemp,a.urbaemp,a.codgeo,a.teleofi,a.anexofi,a.refdirec,a.codestciv,a.fecsus,a.codinpro,a.renta,b.codusu, indRam,IndCIntErr,,PrefDirCor, PrefLocCor, PrefDirEmp, PrefLocEmp " & _
        "FROM HOJAVIDA a join SOLICITUD b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.sechv=b.secu join Ctrldoc c on b.lote = c.lote and b.fecsal = c.fecsal " & _
        "WHERE a.indDoc = 'C' "
    End If
    If OptSoloTra.Value Then
        B1_1 = "SELECT a.fecsal,b.lote,a.numsol,a.inddoc,a.sechv,a.nomcorres,a.codcargo,a.celular,a.fecnaci,a.ubigeonacren,a.ubigeonacine,a.nacion,a.sexo,b.indenvmail,b.indmail,a.emaildoc,a.emailtra,a.desdomcor,a.nrodomcor,a.urbadomcor,a.codgeocor,a.codsuc,a.direemp,a.nroemp,a.urbaemp,a.codgeo,a.teleofi,a.anexofi,a.refdirec,a.codestciv,a.fecsus,a.codinpro,a.renta,b.codusu,indRam,IndCIntErr,PrefDirCor, PrefLocCor, PrefDirEmp, PrefLocEmp " & _
        "FROM HOJAVIDA a join SOLICITUD b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.sechv=b.secu join Ctrldoc c on b.lote = c.lote and b.fecsal = c.fecsal " & _
        "WHERE a.indDoc = 'T' "
    End If
'*****
    If OptFecha.Value Then
        B1_2 = " and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
    Else
        If OptLote Then
            B1_2 = " and b.Lote = '" & TxtLote & "' "
        Else
            B1_2 = " and c.envio= '" & txtEnvio & "' "
        End If
    End If
'*****
    If OptTodosF.Value Then
        B1_3 = ""
    End If
    If OptEstCivB.Value Then
        B1_3 = " and a.codestciv = '' "
    End If
    If OptFVisitaB.Value Then
        B1_3 = " and (a.fecsus is NULL or a.fecsus > CONVERT(datetime,'" & TxtFecha & "',103) ) "
    End If
    If OptCInterB.Value Then
        B1_3 = " and a.codinpro = '' "
    End If
    If OptRentaB.Value Then
        B1_3 = " and a.renta <= 100 "
    End If
    If OptPromERR.Value Then
        B1_3 = " and a.codinpro not in (select codproad from tpromotor b) "
    End If
    If OptUbiCorre.Value Then
        B1_3 = " and a.codgeocor = '' "
    End If
    If OptUbiSuc.Value Then
        B1_3 = " and a.codgeo = '' "
    End If
    If OptEmail.Value Then
        B1_3 = " and (b.indenvmail is not null or b.indmail is not null) "
    End If
    If OptmailErr.Value Then
        B1_3 = " and ((b.indenvmail is not null and a.emaildoc='' and a.emailtra='') or (b.indmail is not null and a.emaildoc='' and a.emailtra='') or (b.indmail is null and (a.emaildoc<>'' or a.emailtra<>'')) or (b.indmail=1 and a.emaildoc='') or (b.indmail=0 and a.emailtra='')) "
    End If

'*****
    B1_4 = " order by a.numsol,a.sechv "

    B1 = B1_1 + B1_2 + B1_3 + B1_4
    
    Call CargarHV

End Sub

Sub CargarHV()
    If RsHV.State = 1 Then RsHV.Close
    
    RsHV.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    If RsHV.EOF And RsHV.BOF Then
        Set DtGHV.DataSource = Nothing
        DtGHV.Refresh
        cmdGrabar.Enabled = False
'        CmdElimina.Enabled = False
        MsgBox "No se encontraron registros", vbInformation, "Revision de Datos de Hoja de Vida "
    Else
        Set DtGHV.DataSource = RsHV
        DtGHV.Refresh
        cmdGrabar.Enabled = True
'        CmdElimina.Enabled = True
        LblTotReg.Caption = RsHV.RecordCount
        CargaImag
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

Private Sub CmdGrabar_Click()
On Error GoTo errores
    RsHV.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
    Exit Sub
errores:
    If Err.Number = -2147217864 Then
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
    Else
    MsgBox Err.Description, vbCritical, Caption
    End If
    
End Sub

Private Sub CmdSalir_Click()
    xfrmrep = ""
    Unload Me
End Sub

Private Sub CmdUbigeo_Click()
    FrmConsUbi.Show
End Sub

Private Sub CmdPromo_Click()
    FrmConsPromo.Show
End Sub

Private Sub CmdCargos_Click()
    FrmCargos.Show
End Sub

Private Sub CmdNacion_Click()
    FrmNacion.Show
End Sub

Private Sub DtGHV_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "D"
    If Activado Then
        On Error GoTo Salir
        With RsHV
            xlote = .Fields("lote")
            xnumsol = .Fields("numsol")
            xidimagen = "01" + xnumsol
            FrmMuestraImg.xrutavol = xrutavol
            FrmMuestraImg.Show
        End With
    Else
        MsgBox "Seleccione la solicitud a buscar", vbCritical, "Error en la busqueda "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CargaImag()
    With RsHV
        xlote = Format(!Lote, "0000000#")
        xlote2 = Format(!Lote, "0000000#")
        xnumsol = Format(!Numsol, "000000#")
        xtipimag = "V"
        xidimagen = "01" + xnumsol
        xcodcargo = IIf(IsNull(!codcargo), "", !codcargo)
    End With
    xnumpag = 1
    
    LblCodcargo = xcodcargo
    
    If RsCargo.State = 1 Then RsCargo.Close
    RsCargo.CursorLocation = adUseClient
    
    SQLCargo = "SELECT * FROM Cargo where codcargo = '" & xcodcargo & "' order by codcargo "
    
    RsCargo.Open SQLCargo, Con.Cnx, adOpenStatic, adLockReadOnly
    
    
    If xcodcargo = "" Then
        LblCargo = " *** "
    Else
        With RsCargo
'            .Find "[codcargo] = '" & xcodcargo & "' "
            If .EOF Then
                LblCargo = " XXX "
            Else
                LblCargo = .Fields("descargo")
            End If
        End With
    End If
      
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
    
    If OptTodosF.Value Or OptUbiCorre.Value Or OptEmail.Value Or OptmailErr.Value Then
        ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    Else
        ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 3)
        ImgEdit1.ScrollPositionX = 0
    End If
    ImgEdit1.Refresh
    
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'' alto
'    ImgEdit1.Height = FrmRepHV.Height - 7800
'' ancho
'    ImgEdit1.Width = FrmRepHV.Width - 300
'' alto
''    DtGHV.Height = 1200 - ImgEdit1.Height
'    ImgEdit1.Height = FrmRepHV.Height - 6500
'End Sub

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


Private Sub cmdVerAyuda_Click()
frmPrefijos.Visible = True
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmRepSolic 
   Caption         =   "Reportes de Solicitudes"
   ClientHeight    =   10590
   ClientLeft      =   1005
   ClientTop       =   990
   ClientWidth     =   15075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   15075
   Begin VB.Frame frmPrefijos 
      Height          =   3975
      Left            =   480
      TabIndex        =   36
      Top             =   2280
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3960
         TabIndex        =   40
         Top             =   3480
         Width           =   1215
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
         TabIndex        =   38
         Top             =   600
         Width           =   2295
      End
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
         TabIndex        =   37
         Top             =   600
         Width           =   2295
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
         Left            =   3000
         TabIndex        =   42
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
         Left            =   480
         TabIndex        =   41
         Top             =   240
         Width           =   1575
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
      Left            =   3000
      TabIndex        =   26
      Top             =   120
      Width           =   2295
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
         TabIndex        =   29
         Top             =   960
         Width           =   1815
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
         TabIndex        =   28
         Top             =   620
         Width           =   1815
      End
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
         TabIndex        =   27
         Top             =   280
         Value           =   -1  'True
         Width           =   2055
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
      TabIndex        =   23
      Top             =   120
      Width           =   2775
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
         TabIndex        =   33
         Top             =   960
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
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   32
         Top             =   960
         Width           =   1095
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1095
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
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   1095
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
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
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
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdPromo 
      Caption         =   "&Promotor"
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
      Left            =   12600
      TabIndex        =   6
      Top             =   600
      Width           =   1095
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
      Left            =   12600
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdUbigeo 
      Caption         =   "&Ubigeo"
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
      Left            =   12600
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
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
      Left            =   13800
      TabIndex        =   3
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
      Height          =   375
      Left            =   13800
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DtGSolic 
      Height          =   3735
      Left            =   120
      TabIndex        =   19
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
      ColumnCount     =   59
      BeginProperty Column00 
         DataField       =   "Fecsal"
         Caption         =   "Fec Salida"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "secu"
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
      BeginProperty Column06 
         DataField       =   "Apepat"
         Caption         =   "Ape Pat"
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
         DataField       =   "Apemat"
         Caption         =   "Ape Mat"
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
         DataField       =   "Fecnac"
         Caption         =   "Fecha Nac"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "ubigeonac"
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
      BeginProperty Column13 
         DataField       =   "ubigeoren"
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
      BeginProperty Column14 
         DataField       =   "tipodoc"
         Caption         =   "T Doc"
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
         DataField       =   "numdoc"
         Caption         =   "Num Doc"
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
         DataField       =   "nac"
         Caption         =   "Nacional"
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
      BeginProperty Column18 
         DataField       =   "PrefDirAfi"
         Caption         =   "Tip.Via"
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
      BeginProperty Column20 
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
      BeginProperty Column21 
         DataField       =   "PrefLocAfi"
         Caption         =   "Tip Urba"
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
      BeginProperty Column23 
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
      BeginProperty Column24 
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
      BeginProperty Column25 
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
      BeginProperty Column26 
         DataField       =   "ubigeo"
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
      BeginProperty Column27 
         DataField       =   "telefono"
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
      BeginProperty Column28 
         DataField       =   "IndDirCorr"
         Caption         =   "Ind Dir. Corres."
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
         DataField       =   "empleadores"
         Caption         =   "Emplea"
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
         DataField       =   "cantemple"
         Caption         =   "Cant Emp"
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
         DataField       =   "tiptrabajador"
         Caption         =   "Tp Tra"
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
         DataField       =   "IPSS"
         Caption         =   "IPSS"
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
         DataField       =   "proc19990"
         Caption         =   "19990"
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
         DataField       =   "proc20530"
         Caption         =   "20530"
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
         DataField       =   "primempleo"
         Caption         =   "1 emp"
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
         DataField       =   "tipafil"
         Caption         =   "Tp Afil"
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
         DataField       =   "codpromo"
         Caption         =   "Cod Promo"
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
         DataField       =   "IndCSBSErr"
         Caption         =   "Ind. Error SBS"
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
      BeginProperty Column39 
         DataField       =   "bolpago"
         Caption         =   "Bol Pago"
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
      BeginProperty Column40 
         DataField       =   "pagoultapo"
         Caption         =   "Pag Ult Apo"
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
      BeginProperty Column41 
         DataField       =   "Copiasim"
         Caption         =   "Copia Sim"
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
      BeginProperty Column42 
         DataField       =   "otrodoc"
         Caption         =   "Otro Doc"
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
      BeginProperty Column43 
         DataField       =   "descotrodoc"
         Caption         =   "Desc Otro Doc"
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
      BeginProperty Column44 
         DataField       =   "fecregpro"
         Caption         =   "Fec Veri Tra"
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
      BeginProperty Column45 
         DataField       =   "codafpori"
         Caption         =   "Afp Ori"
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
      BeginProperty Column46 
         DataField       =   "tipofonori"
         Caption         =   "Fd Ori"
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
      BeginProperty Column47 
         DataField       =   "tipofondes"
         Caption         =   "Fd Des"
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
      BeginProperty Column48 
         DataField       =   "Fecfirmsol"
         Caption         =   "Fec Firma"
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
      BeginProperty Column49 
         DataField       =   "e_mail"
         Caption         =   "e-mail"
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
      BeginProperty Column50 
         DataField       =   "indenvmail"
         Caption         =   "eecc"
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
      BeginProperty Column51 
         DataField       =   "indmail"
         Caption         =   "Tip Email"
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
      BeginProperty Column52 
         DataField       =   "status"
         Caption         =   "Status"
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
      BeginProperty Column53 
         DataField       =   "idimagen"
         Caption         =   "IDImagen"
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
      BeginProperty Column54 
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
      BeginProperty Column55 
         DataField       =   "error"
         Caption         =   "Cod Errores"
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
      BeginProperty Column56 
         DataField       =   "IndCodGeoDoc"
         Caption         =   "Ind. Geo Doc Afil."
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
      BeginProperty Column57 
         DataField       =   "IndCodGeoEmp"
         Caption         =   "Ind. Geo Doc Emp"
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
      BeginProperty Column58 
         DataField       =   "UsuRepSol"
         Caption         =   "Usu Mod"
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column14 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column16 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column27 
         EndProperty
         BeginProperty Column28 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column31 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column33 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column34 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column35 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column36 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column38 
            Alignment       =   2
         EndProperty
         BeginProperty Column39 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column40 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column41 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column42 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column43 
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column46 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column47 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   2970.142
         EndProperty
         BeginProperty Column50 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column51 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column52 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column53 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column54 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column55 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column56 
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column57 
            ColumnWidth     =   2115.213
         EndProperty
         BeginProperty Column58 
            Locked          =   -1  'True
         EndProperty
      EndProperty
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
      Left            =   5400
      TabIndex        =   18
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton cmdVerAyuda 
         Caption         =   "Prefijos"
         Height          =   375
         Left            =   4920
         TabIndex        =   39
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optMayores 
         Caption         =   "Mayor 50"
         Height          =   255
         Left            =   4800
         TabIndex        =   35
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton optVerCorreo 
         Caption         =   "Rev Correos"
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   560
         Width           =   1215
      End
      Begin VB.OptionButton OptDNIDup 
         Caption         =   "DNI Duplicado"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Solo si la consulta es realizada por Fecha"
         Top             =   1040
         Width           =   1575
      End
      Begin VB.OptionButton OptUbigeoB 
         Caption         =   "Ubigeo = b"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   800
         Width           =   1215
      End
      Begin VB.OptionButton OptErrTel 
         Caption         =   "Error Telefono"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   560
         Width           =   1455
      End
      Begin VB.OptionButton OptDirB 
         Caption         =   "Direcc = blanco"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1575
      End
      Begin VB.OptionButton OptErrCuspp 
         Caption         =   "Error en CUSPP"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   560
         Width           =   1575
      End
      Begin VB.OptionButton OptCodPromo 
         Caption         =   "Error Promotor"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   800
         Width           =   1335
      End
      Begin VB.OptionButton OptVarEmple 
         Caption         =   "Varios Empleadores"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1040
         Width           =   1695
      End
      Begin VB.OptionButton OptTodosF 
         Caption         =   "Todos "
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptCusppB 
         Caption         =   "CUSPP = blanco"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   300
         Width           =   1695
      End
      Begin VB.OptionButton OptCusppSol 
         Caption         =   "= Cuspp en DIF Sol"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   800
         Width           =   1815
      End
      Begin VB.OptionButton OptErrUbigeo 
         Caption         =   "Error Ubigeo"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   1040
         Width           =   1215
      End
   End
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
      Left            =   11400
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
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
      Left            =   13800
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   4935
      Left            =   120
      TabIndex        =   30
      Top             =   5520
      Width           =   14775
      _Version        =   131074
      _ExtentX        =   26061
      _ExtentY        =   8705
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tot Registros"
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
      Left            =   11400
      TabIndex        =   21
      Top             =   120
      Width           =   1155
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
      Left            =   11400
      TabIndex        =   20
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmRepSolic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsConsSolic As New Recordset        'TempConsulta
Public RsRepoSolic As New Recordset        'TempConsulta
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
Dim xprimer As String
Dim SQLstr As String
Dim B1 As String
Dim Activado As Boolean
Dim X1 As Single
Dim Y1 As Single

Private Sub cmdCerrar_Click()
frmPrefijos.Visible = False
End Sub

Private Sub cmdVerAyuda_Click()
frmPrefijos.Visible = True
End Sub

Private Sub DtGSolic_HeadClick(ByVal ColIndex As Integer)
    RsConsSolic.Sort = DtGSolic.Columns(ColIndex).DataField
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
MovImg KeyCode
KeyCode = 0
End Sub

Private Sub Form_Load()
    cmdGrabar.Enabled = False
    CmdVerImag.Enabled = False
    xfrmrep = "S"
    TxtFecha = gDate
    B1 = ""
    'DtGSolic.Columns(26). = True
    'DtGSolic.Columns(26).DataFormat = "Si/No"
    
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
'= isnull(convert(char(10),a.fecnac,103),'')

    B1_1 = " SELECT a.fecsal,a.lote,a.secu,a.numsol,a.inddoc,a.apepat,a.apemat,a.nombre1,a.nombre2,a.sexo,a.ubigeonac,a.ubigeoren, fecnac = convert(char(10), isnull(a.fecnac,''),103) ,a.tipodoc,a.numdoc,a.ipss,a.cuspp,a.nac,a.direccion,a.numero,a.urb,a.distrito,a.provincia,a.departamento,a.ubigeo,a.telefono,a.empleadores,a.cantemple,a.TipTrabajador,a.Proc19990,a.Proc20530,a.PrimEmpleo,a.TipAfil,a.CodPromo,a.bolpago,a.pagoultapo,a.copiasim,a.otrodoc,a.descotrodoc,a.fecregpro,a.codafpori,a.TipoFonOri,a.TipoFonDes,a.Fecregpro,a.fecfirmsol,a.indenvmail,a.indmail,a.status,a.idimagen,a.codusu,a.error,IPSS, a.e_mail, a.IndDirCorr, IndCSBSErr, IndCodGeoDoc, IndCodGeoEmp,prefdirafi, preflocafi, UsuRepSol "
    B1_2 = " FROM SOLICITUD a (NOLOCK) join Ctrldoc b on a.Lote = b.Lote and a.fecsal = b.fecsal  "
'*****
    If OptTodos.Value Then
        B1_3 = "WHERE a.indDoc is not NULL and a.secu=0 "
    End If
    If OptSoloAfi.Value Then
        B1_3 = "WHERE a.indDoc = 'C' and a.secu=0 "
    End If
    If OptSoloTra.Value Then
        B1_3 = "WHERE a.indDoc = 'T' and a.secu=0 "
    End If
'*****
    If OptFecha.Value Then
        If TxtFecha <> "  /  /   " And TxtFecha <> "" Then
            B1_4 = " and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
        Else
            MsgBox "Ingresar Fecha a buscar", vbCritical, "Error en la busqueda "
            TxtFecha.SetFocus
            Exit Sub
        End If
    Else
       If OptLote Then
            If val(TxtLote) > 0 Then
                B1_4 = " and a.Lote = '" & TxtLote & "' "
            Else
                MsgBox "Ingresar Lote a buscar", vbCritical, "Error en la busqueda "
                TxtLote.SetFocus
                Exit Sub
            End If
       Else
            If Trim(txtEnvio) <> "" Then
                B1_4 = " and b.Envio = '" & txtEnvio & "' "
            Else
                MsgBox "Ingresar Envio a buscar", vbCritical, "Búsqueda "
                TxtLote.SetFocus
                Exit Sub
            End If
            
       End If
    End If
'*****
    If OptTodosF.Value Then
        B1_5 = ""
    End If
    If optVerCorreo.Value Then
        B1_5 = ""
    End If
    If OptDirB.Value Then
        B1_5 = " and a.direccion='' "
    End If
    If OptErrTel.Value Then
        B1_5 = " and ((len(a.telefono)<7 and a.telefono<>'') or a.telefono like '%-%' or a.telefono='' or a.telefono like '%11111%' or a.telefono is null ) "
    End If
    If OptCodPromo.Value Then
        B1_5 = " and a.codpromo not in (select codpromo from tpromotor b) "
    End If
    If OptCusppB.Value Then
        B1_5 = " and (a.cuspp='' or a.cuspp is null) "
    End If
    If OptErrCuspp.Value Then
        B1_5 = " and ( a.cuspp='' or a.cuspp is null or substring(a.cuspp,7,1)<> substring(a.nombre1,1,1) " & _
                "or substring(a.cuspp,8,1)<> substring(a.apepat,1,1) or substring(a.cuspp,9,1)<> substring(a.apemat,1,1) " & _
                "or substring(a.cuspp,10,1)<> (case when len(a.apepat)>3 then substring(a.apepat,4,1) else substring(a.apepat,len(a.apepat),1) end) " & _
                "or substring(a.cuspp,11,1)<> (case when len(a.apemat)>3 then substring(a.apemat,4,1) else substring(a.apemat,len(a.apemat),1) end) ) "
    End If
    If OptVarEmple.Value Then
        B1_5 = " and (a.cantemple>1 or a.cantemple=0) "
    End If
    If OptUbigeoB.Value Then
        B1_5 = " and a.ubigeo = '' "
    End If
    If OptErrUbigeo.Value Then
        B1_5 = " and ( a.ubigeo is null or ( a.ubigeo like '__00__' ) or ( a.ubigeo like '____00' ) or a.distrito is null or a.provincia is null or a.departamento is null )  "
    End If
       
    If OptCusppSol.Value Then
        B1_2 = " FROM solicitud a, (SELECT cuspp, num=count(*) FROM solicitud WHERE fecsal = CONVERT(datetime,'" & TxtFecha & "',103) and cuspp<>'' GROUP BY cuspp) b "
        
        B1_5 = " and a.cuspp = b.cuspp and num > 1 and a.cuspp<>''  "
    End If
    If OptFecha Then
        If OptDNIDup Then
            B1_5 = " and numdoc in (select numdoc from solicitud where fecsal = '" & TxtFecha & "' group by numdoc having count(numdoc)>1 ) "
        End If
    End If

'*****

    If optMayores Then
        B1_5 = " and  year(getdate()) - year(Fecnac) >= 50 "
    End If
    
    B1_6 = " order by a.lote,a.numsol,a.secu "

    B1 = B1_1 & B1_2 & B1_3 & B1_4 & B1_5 & B1_6
    
    Call CargarSOL

End Sub

Sub CargarSOL()
    LblTotReg = 0
    If RsConsSolic.State = 1 Then RsConsSolic.Close
    RsConsSolic.CursorLocation = adUseClient
    RsConsSolic.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic

    With RsConsSolic
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Solicitudes"
            DtGSolic.Refresh
            cmdGrabar.Enabled = False
            Exit Sub
        Else
            Set DtGSolic.DataSource = RsConsSolic
            DtGSolic.Refresh
            cmdGrabar.Enabled = True
            LblTotReg.Caption = RsConsSolic.RecordCount
            CargaImag
        End If
    End With

End Sub

Private Sub CmdRepo_Click()
    If RsRepoSolic.State = 1 Then RsRepoSolic.Close
    If OptTodosF.Value Then
        SQLstr = "SELECT * FROM solicitud, ctrldoc WHERE solicitud.lote = ctrldoc.lote and secu='0' and solicitud.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote, numsol, secu"
    End If
    If OptCusppSol.Value Then
        SQLstr = "SELECT a.fecsal,c.envio,a.lote,c.agencia,a.numsol,a.secu,a.cuspp,a.apepat,a.apemat,a.nombre1,a.nombre2,a.tipodoc,a.numdoc,a.direccion,a.numero,a.urb,a.distrito,a.provincia,a.departamento,a.ubigeo,a.telefono,a.empleadores,a.cantemple,a.codpromo,a.bolpago,a.pagoultapo,a.copiasim,a.otrodoc,a.descotrodoc,a.fecregpro,a.codafpori,a.fecfirmsol,a.riaafil,a.fecnac,a.sexo,a.e_mail,a.status,a.idimagen,a.codusu FROM solicitud a, (SELECT cuspp, num=count(*) FROM solicitud WHERE fecsal = CONVERT(datetime,'" & TxtFecha & "',103) GROUP BY cuspp) b, ctrldoc c WHERE a.lote = c.lote and a.cuspp = b.cuspp and num > 1 and a.cuspp<>'' and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY a.lote, numsol, secu"
    End If
    If OptVarEmple.Value Then
        SQLstr = "SELECT * FROM solicitud, ctrldoc WHERE solicitud.lote = ctrldoc.lote and (cantemple>1 or cantemple=0) and solicitud.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote, numsol, secu"
    End If
    If OptCusppB.Value Then
        SQLstr = "SELECT * FROM solicitud, ctrldoc WHERE solicitud.lote = ctrldoc.lote and secu='0' and (cuspp='' or cuspp is null) and solicitud.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote, numsol, secu"
    End If
    If OptErrCuspp.Value Then
        SQLstr = "SELECT * FROM solicitud, ctrldoc WHERE solicitud.lote = ctrldoc.lote and secu='0' and ( substring(cuspp,7,1)<> substring(nombre1,1,1) " & _
                "or substring(cuspp,8,1)<> substring(apepat,1,1) or substring(cuspp,9,1)<> substring(apemat,1,1) or substring(cuspp,10,1)<> substring(apepat,4,1) " & _
                "or substring(cuspp,11,1)<> substring(apemat,4,1) ) and solicitud.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote, numsol, secu "
    End If
    If OptDirB.Value Then
        SQLstr = "SELECT * FROM solicitud, ctrldoc WHERE solicitud.lote = ctrldoc.lote and secu='0' and solicitud.direccion='' and solicitud.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote, numsol, secu"
    End If
    If OptErrTel.Value Then
        SQLstr = "SELECT * FROM solicitud, ctrldoc WHERE solicitud.lote = ctrldoc.lote and secu='0' and len(telefono)<7 and telefono<>'' and solicitud.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote, numsol, secu"
    End If
    If OptCodPromo.Value Then
        SQLstr = "SELECT * FROM solicitud a, ctrldoc c WHERE a.lote = c.lote and secu='0' and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) and a.codpromo not in (select codpromo from tpromotor b) ORDER BY a.lote, numsol, secu"
    End If
    If OptFNacBla.Value Then
        SQLstr = "SELECT * FROM solicitud, ctrldoc WHERE solicitud.lote = ctrldoc.lote and secu='0' and (fecnac is null or sexo='' or riaafil<=1 ) and solicitud.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY solicitud.lote, numsol, secu"
    End If
    If OptErrUbigeo.Value Then
        SQLstr = "SELECT * FROM solicitud a, ctrldoc b, ubigeo c WHERE a.lote = b.lote and a.ubigeo = c.codgeo and ((a.distrito <> c.distrito and a.provincia = c.provincia and a.departamento = c.departamento) or (a.distrito = c.distrito and a.provincia <> c.provincia and a.departamento = c.departamento) or (a.distrito = c.distrito and a.provincia = c.provincia and a.departamento <> c.departamento) or ( ubigeo like '__00__' ) or ( ubigeo like '____00' ) or a.distrito is null or a.provincia is null or a.departamento is null) and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY a.lote,numsol, secu"
    End If

    RsRepoSolic.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtRSolic.DataSource = RsRepoSolic
    DtRSolic.Show
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
    'Grabar Usuario
    On Error GoTo errores
    
    RsConsSolic.Fields("UsuRepSol") = gNomUsu
    RsConsSolic.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
    
    Exit Sub
errores:
    If Err.Number = -2147217864 Then
        MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
    End If
    
End Sub

Private Sub CmdSalir_Click()
    xfrmrep = ""
    Unload Me
End Sub

Private Sub CmdPromo_Click()
    FrmConsPromo.Show
End Sub

Private Sub CmdUbigeo_Click()
    FrmConsUbi.Show
End Sub

Private Sub DtGSolic_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "S"
    If Activado Then
        On Error GoTo Salir
        With RsConsSolic
            xlote = .Fields("lote")
            xnumsol = .Fields("numsol")
            xidimagen = .Fields("idimagen")
            FrmMuestraImg.Show
        End With
    Else
        MsgBox "Seleccione la solicitud a buscar", vbCritical, "Error en la busqueda "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CargaImag()
    With RsConsSolic
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
    "WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and substring(idimagen,3,9)= '" & xnumsol & "' order by posicion "
    
    'and ( TipImag = '" & xtipimag1 & "' or TipImag = '" & xtipimag2 & "' or TipImag = 'D' )
    
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
    
    If optVerCorreo Then
       ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 3)
        ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 1)
    End If
    
    ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    ImgEdit1.Refresh
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
' alto
    ImgEdit1.Height = FrmRepSolic2.Height - 7800
' ancho
    ImgEdit1.Width = FrmRepSolic2.Width - 300
' alto
'    DtGHV.Height = 1200 - ImgEdit1.Height
    ImgEdit1.Height = FrmRepSolic2.Height - 6500
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


Private Sub frmPrefijos_DragDrop(Source As Control, X As Single, Y As Single)
    'Copy this code into the dragDrop Event of the Target object
    StopDrag Source, frmPrefijos, X, Y
    Source.ZOrder 0 'Add if you want the control moved to the front
End Sub

Private Sub frmPrefijos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Copy this code into the MouseDown Event of the drag object
    If Shift = 1 Then StartDrag frmPrefijos, X, Y
End Sub


 Sub StartDrag(Source As Control, X As Single, Y As Single)
        Source.Tag = "Drag"
        'Stores the initial X and Y of the mouse on the drag object
        X1 = X
        Y1 = Y
        Source.Drag
End Sub
Sub StopDrag(Source As Control, Target As Object, X As Single, Y As Single)
    On Error Resume Next
    If Source.Tag <> Target.Tag Then 'not dropped on itself
        'Check target for valid container type
        If TypeOf Target Is Form Or TypeOf Target Is Frame Or TypeOf Target Is PictureBox Then
            'Set the container to target object
            Set Source.Container = Target
            Source.Left = X - X1
            Source.Top = Y - Y1
        Else 'Target is not a Container
            'Is the target an object within itself
            If Source.Tag <> Target.Container.Tag Then
                'Set the container to target object's container
                Set Source.Container = Target.Container
                Source.Left = Target.Left + (X - X1)
                Source.Top = Target.Top + (Y - Y1)
            Else
                If X > X1 Then Source.Left = Source.Left + Target.Left + (X - X1) Else Source.Left = Source.Left - (X1 - X)
                If Y > Y1 Then Source.Top = Source.Top + Target.Top + (Y - Y1) Else Source.Top = Source.Top - (Y1 - Y)
           End If
        End If
    Else 'Dropped on itself
        If X > X1 Then Source.Left = Source.Left + (X - X1) Else Source.Left = Source.Left - (X1 - X)
        If Y > Y1 Then Source.Top = Source.Top + (Y - Y1) Else Source.Top = Source.Top - (Y1 - Y)
    End If
    Source.Tag = ""
End Sub

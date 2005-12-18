VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSucursal 
   Caption         =   "Consulta Sucursales"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   1245
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14610
   Begin VB.Frame FraCriterio 
      Caption         =   "Criterio de Busqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton OptDescri 
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton OptRuc 
         Caption         =   "RUC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptUbigeo 
         Caption         =   "Ubigeo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptDirec 
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox TxtBusqueda 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   19
      Top             =   340
      Width           =   4335
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
      Height          =   440
      Left            =   6000
      TabIndex        =   18
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Left            =   13320
      TabIndex        =   17
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agregar Sucursal"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   14295
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   12720
         TabIndex        =   37
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtAnxFax 
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
         Height          =   375
         Left            =   10200
         MaxLength       =   7
         TabIndex        =   32
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtFax 
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
         Height          =   375
         Left            =   8880
         MaxLength       =   9
         TabIndex        =   31
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtAnxTel 
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
         Height          =   375
         Left            =   7560
         MaxLength       =   7
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtTelef 
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
         Height          =   375
         Left            =   6240
         MaxLength       =   9
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtDir 
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
         Height          =   375
         Left            =   7200
         MaxLength       =   60
         TabIndex        =   28
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox TxtRuc 
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
         Height          =   375
         Left            =   240
         MaxLength       =   11
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtNSuc 
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
         Height          =   375
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtDescri 
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
         Height          =   375
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   8
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox TxtUbigeo 
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
         Height          =   375
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TxtNumero 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         MaxLength       =   20
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtUrba 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Anx FAX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10320
         TabIndex        =   36
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9120
         TabIndex        =   35
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Anx Telef"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7680
         TabIndex        =   34
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6360
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7320
         TabIndex        =   27
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N Suc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ubigeo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5280
         TabIndex        =   13
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Numero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11880
         TabIndex        =   12
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Urbanizacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1110
      End
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
      Height          =   440
      Left            =   13320
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtCaptura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   7920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton CmdVisualizar 
      Caption         =   "Visualizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DtGSucursal 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   8281
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "ruc"
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
      BeginProperty Column01 
         DataField       =   "n_sucursal"
         Caption         =   "N_Suc"
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
         DataField       =   "descripcion"
         Caption         =   "Descripcion"
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
      BeginProperty Column04 
         DataField       =   "numero"
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
      BeginProperty Column05 
         DataField       =   "Urbanizacion"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "anx_telef"
         Caption         =   "Anx_telef"
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
         DataField       =   "fax"
         Caption         =   "Fax"
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
         DataField       =   "anx_fax"
         Caption         =   "Anx_Fax"
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
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3000.189
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCriterio 
      AutoSize        =   -1  'True
      Caption         =   "Campo de Busqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3360
      TabIndex        =   20
      Top             =   100
      Width           =   1980
   End
End
Attribute VB_Name = "FrmSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsSucursal As New Recordset      'Agencia
Public RsAgrega As New Recordset
Dim Activado As Boolean
Dim xcodusu As String
Dim B1 As String

Private Sub Form_Load()
    xcodusu = gNomUsu
    CmdAgregar.Enabled = False
    CmdEliminar.Enabled = False
'    CmdVisualizar.Enabled = True
    TxtBusqueda = ""
    TxtRuc = ""
    TxtNSuc = ""
    TxtDescri = ""
    TxtDir = ""
    TxtNumero = ""
    TxtUrba = ""
    TxtUbigeo = ""
    TxtTelef = ""
    TxtAnxTel = ""
    TxtFax = ""
    TxtAnxFax = ""
End Sub

Private Sub DtGSucursal_HeadClick(ByVal ColIndex As Integer)
    RsSucursal.Sort = DtGSucursal.Columns(ColIndex).DataField
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub DtGSucursal_SelChange(Cancel As Integer)
    Activado = True
    CmdEliminar.Enabled = True
End Sub

Private Sub TxtBusqueda_Change()
'
'    If OptBcoT.Value Then
'        SQLstrC02 = " "
'    End If
'
'    If OptCriT.Value Then
'        SQLstrC01 = "SELECT * FROM Sucursal WHERE codage <> '000' "
'        SQLstrC03 = "ORDER BY CodBco,CodAge,CodOfi,ubic "
'    End If
'    If OptDesAge.Value Then
'        SQLstrC01 = "SELECT * FROM Sucursal WHERE descri like '%" & Trim(TxtBusqueda) & "%' "
'        SQLstrC03 = "ORDER BY codbco,ubic,descri,codage "
'    End If
'    If OptCodAge.Value Then
'        SQLstrC01 = "SELECT * FROM Sucursal WHERE CodAge like '%" & Trim(TxtBusqueda) & "%' or CodOfi like '" + Trim(TxtBusqueda) & "%' "
'        SQLstrC03 = "ORDER BY CodBco,CodAge,CodOfi,ubic "
'    End If
'    If OptDistriD.Value Then
'        SQLstrC01 = "SELECT * FROM Sucursal WHERE disage like '%" & Trim(TxtBusqueda) & "%' "
'        SQLstrC03 = "ORDER BY codbco,disage,ubic,codage "
'    End If
'    If OptUbi.Value Then
'        SQLstrC01 = "SELECT * FROM Sucursal WHERE ubic like '%" & Trim(TxtBusqueda) & "%' "
'        SQLstrC03 = "ORDER BY CodBco,Ubic,codage "
'    End If
'
'    SQLstr = SQLstrC01 + SQLstrC02 + SQLstrC03
'
'    If RsSucursal.State = 1 Then RsSucursal.Close
'    RsSucursal.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
'    Set DtGSucursal.DataSource = RsSucursal
'    DtGSucursal.Refresh
'
'    If DtGSucursal.ApproxCount <> 0 Then
'        Dim Comienzo, Final, numReg
'        Comienzo = Trim(DtGSucursal.Columns(1).CellText(DtGSucursal.FirstRow))
'        Final = Trim(DtGSucursal.Columns(1).CellText(DtGSucursal.ApproxCount))
'        numReg = DtGSucursal.ApproxCount
'        CmdVisualizar.Enabled = True
'        txtCaptura.Text = "Se Filtraron las Agencias desde " & Comienzo & " a " & Final & " con un total de " & numReg & " Registros."
'    Else
'        txtCaptura.Text = " No hay registros con esa letra o codigo "
'    End If
End Sub

Private Sub CmdVisualizar_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""

    B1_3 = "ORDER BY descripcion,nsucursal "
    
    If OptTodos.Value Then
        B1_1 = "SELECT * FROM Sucursal WHERE RUC <> '00000000000' "
    End If

    If OptRuc.Value Then
        B1_2 = " and ruc like '%" & Trim(TxtBusqueda) & "%'  "
        B1_3 = "ORDER BY RUC,nsucursal "
    End If
    If OptDescri.Value Then
        B1_2 = " and descripcion like '%" & Trim(TxtBusqueda) & "%' "
    End If
    If OptDirec.Value Then
        B1_2 = " and direccion like '%" & Trim(TxtBusqueda) & "%' "
    End If
    If OptUbigeo.Value Then
        B1_2 = " and Ubigeo like '" & Trim(TxtBusqueda) & "%' "
    End If
    
    B1 = B1_1 + B1_2 + B1_3
        
    If RsSucursal.State = 1 Then RsSucursal.Close
        
    RsSucursal.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    Set DtGSucursal.DataSource = RsSucursal
    DtGSucursal.Refresh
    
    If DtGSucursal.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGSucursal.Columns(0).CellText(DtGSucursal.FirstRow))
        Final = Trim(DtGSucursal.Columns(0).CellText(DtGSucursal.ApproxCount))
        numReg = DtGSucursal.ApproxCount
        txtCaptura.Text = "Se Filtraron las Agencias desde " & Comienzo & " a " & Final & " con un total de " & numReg & " Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
    TxtBusqueda.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    TxtNSuc = Format(TxtNSuc, "0#")
    
    If Len(Trim(TxtRuc)) <> 11 Then
        MsgBox "Debe ingresar RUC ", vbCritical, "Error en Ingreso de RUC"
        TxtRuc.SetFocus
        Exit Sub
    End If
    If TxtUbic = "L" Or TxtUbic = "P" Then
    Else
        MsgBox "Debe ingresar Ubicacion (L)ima, (P)rovincia ", vbCritical, "Error en Ingreso de Ubicacion"
        TxtUbic.SetFocus
        Exit Sub
    End If
    If TxtDAge = "" Then
        MsgBox "Debe ingresar nombre de Agencia ", vbCritical, "Error en Ingreso de Agencia"
        TxtDAge.SetFocus
        Exit Sub
    End If
    If TxtDist = "" Then
        MsgBox "Debe ingresar Distrito o Departamento ", vbCritical, "Error en Ingreso de Distrito o Departamento"
        TxtDist.SetFocus
        Exit Sub
    End If
    Dim SQLstrA As String
    
    If RsAgrega.State = 1 Then RsAgrega.Close
    SQLstrA = "SELECT * FROM Sucursal where CodBco = '" & TxtCBco & "' and CodAge = '" & TxtCAge & "' and CodOfi = '" & TxtCOfi & "' ORDER BY CodBco,CodAge,CodOfi"
    
    RsAgrega.Open SQLstrA, Con.Cnx, adOpenStatic, adLockReadOnly

    If RsAgrega.EOF Then
    Else
        MsgBox "La Agencia YA EXISTE"
        TxtCBco = ""
        TxtCAge = ""
        TxtCOfi = ""
        TxtDAge = ""
        TxtUbic = ""
        TxtDist = ""
        TxtBusqueda.SetFocus
        Exit Sub
    End If

    With RsSucursal
        .AddNew
        .Fields("CodBco") = Format(TxtCBco, "00#")
        .Fields("CodAge") = Format(TxtCAge, "00#")
        .Fields("CodOfi") = Format(TxtCOfi, "00#")
        .Fields("Descri") = TxtDAge
        .Fields("Ubic") = TxtUbic
        .Fields("Disage") = TxtDist
        .Fields("CodUsu") = xcodusu
        .Fields("Fecdig") = gDate
        .Fields("Horini") = ObtenerHora
        .Fields("Horfin") = ObtenerHora
        .UpdateBatch
    End With
    TxtCBco = ""
    TxtCAge = ""
    TxtCOfi = ""
    TxtDAge = ""
    TxtUbic = ""
    TxtDist = ""
    
    CmdGrabar.Enabled = True
    CmdGrabar.SetFocus
End Sub

Private Sub CmdGrabar_Click()
    RsSucursal.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub CmdEliminar_Click()
    If Activado Then
        On Error GoTo Salir
        Dim Resp As Integer
        Resp = MsgBox("¿Esta seguro que desea eliminar el registro?", 4 + 32, "Confirma Eliminación de Registro ")
        If Resp = vbYes Then
            RsSucursal.Delete
            RsSucursal.UpdateBatch
            MsgBox "Registro Eliminado", vbInformation, "Proceso de Eliminacion de Registros"
        End If
    Else
        MsgBox "Selecciona el registro a eliminar", vbCritical, "Error en Eliminación "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

'Private Sub TxtCBco_GotFocus()
'    CmdAgregar.Enabled = True
'    Marcado TxtCBco
'    TxtCBco.BackColor = &HC0FFFF
'End Sub

'Private Sub TxtCBco_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 48 To 57
'        Case 8
'        Case 13
'            TxtCAge.SetFocus
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub

'Private Sub TxtCBco_LostFocus()
'    If TxtCBco = "" Then
'        TxtCBco = "000"
'    End If
'    TxtCBco.BackColor = &H80000005
'    TxtCBco = Format(TxtCBco, "00#")
'End Sub

'Private Sub TxtCAge_GotFocus()
'    Marcado TxtCAge
'    TxtCAge.BackColor = &HC0FFFF
'End Sub
'
'Private Sub TxtCAge_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 48 To 57
'        Case 8
'        Case 13
'            TxtCOfi.SetFocus
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub
'
'Private Sub TxtCAge_LostFocus()
'    If TxtCAge = "" Then
'        TxtCAge = "000"
'    End If
'    TxtCAge.BackColor = &H80000005
'    TxtCAge = Format(TxtCAge, "00#")
'End Sub
'
'Private Sub TxtCOfi_GotFocus()
'    Marcado TxtCOfi
'    TxtCOfi.BackColor = &HC0FFFF
'End Sub

'Private Sub TxtCOfi_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 48 To 57
'        Case 8
'        Case 13
'            TxtUbic.SetFocus
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub
'
'Private Sub TxtCOfi_LostFocus()
'    If TxtCOfi = "" Then
'        TxtCOfi = "000"
'    End If
'    TxtCOfi.BackColor = &H80000005
'    TxtCOfi = Format(TxtCOfi, "00#")
'End Sub
'
'Private Sub TxtUbic_GotFocus()
'    Marcado TxtUbic
'    TxtUbic.BackColor = &HC0FFFF
'End Sub
'
'Private Sub TxtUbic_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 13 Then
'        TxtDAge.SetFocus
'    End If
'End Sub
'
'Private Sub TxtUbic_LostFocus()
'    TxtUbic.BackColor = &H80000005
'End Sub
'
'Private Sub TxtDAge_GotFocus()
'    Marcado TxtDAge
'    TxtDAge.BackColor = &HC0FFFF
'End Sub
'
'Private Sub TxtDAge_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Select Case KeyAscii
'        Case 65 To 90
'        Case 8
'        Case 165    'Ñ
'        Case 209    '
'        Case 13
'            If Len(TxtDAge) = 0 Then
'                txtCaptura.Text = "Debe ingresar el nombre de la Agencia "
'                TxtDAge.SetFocus
'            Else
'                TxtDist.SetFocus
'            End If
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub

'Private Sub TxtDAge_LostFocus()
'    TxtDAge.BackColor = &H80000005
'    txtCaptura.Text = ""
'End Sub
'
'Private Sub TxtDist_GotFocus()
'    Marcado TxtDist
'    TxtDist.BackColor = &HC0FFFF
'End Sub
'
'Private Sub TxtDist_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Select Case KeyAscii
'        Case 65 To 90
'        Case 8
'        Case 165    'Ñ
'        Case 209    '
'        Case 13
'            If Len(TxtDist) = 0 Then
'                txtCaptura.Text = "Debe ingresar el Distrito o Departamento "
'                TxtDist.SetFocus
'            Else
'                CmdAgregar.SetFocus
'            End If
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub
'
'Private Sub TxtDist_LostFocus()
'    TxtDist.BackColor = &H80000005
'    txtCaptura.Text = ""
'End Sub



VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "Imgedit.ocx"
Begin VB.Form FrmVeriRiaFec2 
   Caption         =   "Verifica RIA y Fechas"
   ClientHeight    =   10485
   ClientLeft      =   1005
   ClientTop       =   1125
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   14655
   Begin VB.CommandButton cmdNacionalidad 
      Caption         =   "Nacionalidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   12360
      TabIndex        =   38
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdVerPromotor 
      Caption         =   "Promotor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11160
      TabIndex        =   37
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdNacSxB 
      Caption         =   "&Nac/Sx vacios"
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
      Left            =   12360
      TabIndex        =   29
      Top             =   120
      Width           =   975
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
      Height          =   525
      Left            =   12360
      TabIndex        =   28
      Top             =   650
      Width           =   975
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
      Height          =   1575
      Left            =   3120
      TabIndex        =   18
      Top             =   0
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
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   280
         Value           =   -1  'True
         Width           =   975
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   620
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
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1335
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
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   0
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   17
         Top             =   600
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
         TabIndex        =   16
         Top             =   280
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
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   15
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
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
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
      Height          =   525
      Left            =   11280
      TabIndex        =   4
      Top             =   650
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
      Height          =   525
      Left            =   11280
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DtGVeri 
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   14415
      _ExtentX        =   25426
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
      ColumnCount     =   27
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
         DataField       =   "fecnac"
         Caption         =   "Fec Nac 1"
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
         DataField       =   "fecnaci"
         Caption         =   "Fec Nac HV"
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
         DataField       =   "ubigeonac"
         Caption         =   "INEI nac"
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
         DataField       =   "ubigeonacine"
         Caption         =   "INEI HV"
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
         DataField       =   "ubigeoren"
         Caption         =   "Reniec nac"
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
         DataField       =   "ubigeonacren"
         Caption         =   "Reniec HV"
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
         DataField       =   "renta"
         Caption         =   "RIA HV"
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
      BeginProperty Column11 
         DataField       =   "nac"
         Caption         =   "Nac Sol"
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
         DataField       =   "nacion"
         Caption         =   "Nac HV"
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
         DataField       =   "sexo"
         Caption         =   "Sx Sol"
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
         DataField       =   "sexo2"
         Caption         =   "Sx HV"
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
         DataField       =   "FecRegPro"
         Caption         =   "Fec. Ver / Suscrip"
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
         DataField       =   "FecFirmSol"
         Caption         =   "Fec. Dec"
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
         DataField       =   "IndFecVer"
         Caption         =   "Sin F.Verif"
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
         DataField       =   "IndFecDec"
         Caption         =   "Sin F.Decl"
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
         DataField       =   "IndFecsus"
         Caption         =   "Sin.F.Susc"
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
      BeginProperty Column21 
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
      BeginProperty Column22 
         DataField       =   "codgeocor"
         Caption         =   "Ubi. Corr."
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
         DataField       =   "IndErrPro"
         Caption         =   "Error Cod Adm"
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
         DataField       =   "IndCodGeoNac"
         Caption         =   "Ind. Ubigeo Nac."
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
      BeginProperty Column26 
         DataField       =   "UsuRepVida"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1200.189
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
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column14 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column15 
            Alignment       =   2
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column16 
            Alignment       =   2
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column17 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column18 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column19 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column20 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column25 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column26 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtRia 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10080
      TabIndex        =   7
      Top             =   840
      Width           =   1095
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
      Height          =   400
      Left            =   10080
      TabIndex        =   0
      Top             =   120
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
      Height          =   525
      Left            =   13440
      TabIndex        =   5
      Top             =   650
      Width           =   1095
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   2415
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   14415
      _Version        =   131074
      _ExtentX        =   25426
      _ExtentY        =   4260
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   350711296
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
   Begin ImgeditLibCtl.ImgEdit ImgEdit2 
      Height          =   2415
      Left            =   120
      TabIndex        =   30
      Top             =   7920
      Width           =   14415
      _Version        =   131074
      _ExtentX        =   25426
      _ExtentY        =   4260
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   350711296
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
   Begin VB.Frame FraRep 
      Caption         =   "Criterios"
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
      Height          =   1575
      Left            =   4800
      TabIndex        =   12
      Top             =   0
      Width           =   5175
      Begin VB.OptionButton optIndErrAdm 
         Caption         =   "Errores Cod. Interno"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optDif 
         Caption         =   "Diferencia entre Empleador y Hoja Vida"
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   960
         Width           =   3255
      End
      Begin VB.OptionButton optUbicor 
         Caption         =   "Ubi <> GeoCorr"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optFecs 
         Caption         =   "Fecha Suscrip / Verif/ Dec."
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton OptNacionSx 
         Caption         =   "Nacionalidad <> / Sx <>"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton OptUbiRenDif 
         Caption         =   "Ubigeo RENIEC <>"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   500
         Width           =   1695
      End
      Begin VB.OptionButton OptUbiINEIDif 
         Caption         =   "Ubigeo INEI <>"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptFecNacDif 
         Caption         =   "Fec Nac <>"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   500
         Width           =   1215
      End
      Begin VB.OptionButton OptRiaMenor 
         Caption         =   "Ria < Valor RIA "
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   500
         Width           =   1455
      End
      Begin VB.OptionButton OptRiasMay 
         Caption         =   "Ria > Valor RIA "
         Height          =   255
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptFecNac 
         Caption         =   "Fec Nac Rango"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "|******** UBIGEOS  DE NACIMIENTO ********|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   6120
      TabIndex        =   27
      Top             =   1560
      Width           =   3165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "RIA compara"
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
      TabIndex        =   10
      Top             =   600
      Width           =   1035
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
      Left            =   13320
      TabIndex        =   9
      Top             =   0
      Width           =   1305
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
      Left            =   13440
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmVeriRiaFec2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVeri As New Recordset
Dim RsImag1 As New Recordset
Dim RsImag2 As New Recordset
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
Dim xinddoc As String
Dim Activado As Boolean
Dim B1 As String

Private Sub cmdNacionalidad_Click()
frmConsNacion.Show 1
End Sub

Private Sub cmdVerPromotor_Click()
FrmConsPromo.Show 1
End Sub

Private Sub DtGVeri_HeadClick(ByVal ColIndex As Integer)
    RsVeri.Sort = DtGVeri.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
    xfrmrep = "V"
    TxtFecha = gDate
    TxtRia = 10000
    B1 = ""
    xinddoc = "C"
    CmdGrabar.Enabled = False
    CmdVerImag.Enabled = False
End Sub

Private Sub CmdCons_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""
CEmp = "(Select count(numsol) from empleadores where numsol = a.numsol and fecsal = a.fecsal) "
cHJ = " (Select count(numsol) from HojaVida where numsol = a.numsol and fecsal = a.fecsal) "
    If OptTodos.Value Then
        B1_1 = "SELECT a.fecsal,a.lote,a.numsol,a.inddoc,a.secu,a.riaafil,b.renta,a.fecnac,b.fecnaci,a.ubigeonac,a.ubigeoren,b.ubigeonacren,b.ubigeonacine,a.nac,b.nacion,a.sexo,sexo2=b.sexo,a.Fecregpro,a.Fecfirmsol,IndFecDec, IndFecVer, IndFecSus,a.ubigeo , b.codgeocor, a.posicion, b.PosIma, a.codusu , IndErrPro , a.IndCodGeoNac  ,CEmp = " & CEmp & ", CHV = " & cHJ & " ,UsuRepVida  " & _
        "FROM SOLICITUD a left join HOJAVIDA b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.sechv join Ctrldoc c on a.lote = c.lote and a.fecsal= c.fecsal " & _
        "WHERE a.indDoc is not NULL "
    End If
    If OptSoloAfi.Value Then
        B1_1 = "SELECT a.fecsal,a.lote,a.numsol,a.inddoc,a.secu,a.riaafil,b.renta,a.fecnac,b.fecnaci,a.ubigeonac,a.ubigeoren,b.ubigeonacren,b.ubigeonacine,a.nac,b.nacion,a.sexo,sexo2=b.sexo,a.Fecregpro,a.Fecfirmsol,IndFecDec, IndFecVer, IndFecSus, a.ubigeo , b.codgeocor,a.codusu, a.posicion, b.PosIma, IndErrPro, a.IndCodGeoNac, UsuRepVida  " & _
        "FROM SOLICITUD a left join HOJAVIDA b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.sechv join Ctrldoc c on a.lote = c.lote and a.fecsal= c.fecsal " & _
        "WHERE a.indDoc = 'C' "
    End If
    If OptSoloTra.Value Then
        B1_1 = "SELECT a.fecsal,a.lote,a.numsol,a.inddoc,a.secu,a.riaafil,b.renta,a.fecnac,b.fecnaci,a.ubigeonac,a.ubigeoren,b.ubigeonacren,b.ubigeonacine,a.nac,b.nacion,a.sexo,sexo2=b.sexo,a.Fecregpro,a.Fecfirmsol,IndFecDec, IndFecVer, IndFecSus, a.ubigeo , b.codgeocor, a.codusu, a.posicion, b.PosIma, IndErrPro, a.IndCodGeoNac, UsuRepVida  " & _
        "FROM SOLICITUD a left join HOJAVIDA b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.sechv join Ctrldoc c on a.lote = c.lote and a.fecsal= c.fecsal " & _
        "WHERE a.indDoc = 'T' "
    End If
'*****
    If OptFecha.Value Then
        B1_2 = " and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
    Else
        If OptLote Then
            B1_2 = " and a.Lote = '" & TxtLote & "' "
        Else
            B1_2 = " and c.Envio = '" & txtEnvio & "' "
        End If
    End If
'*****
    If OptRiasMay.Value Then
        B1_3 = " and ( b.renta >= " & TxtRia & " ) "
    End If
    If OptRiaMenor.Value Then
        B1_3 = " and ( b.renta = 0 or  b.renta <= " & TxtRia & " ) "
    End If
    If OptFecNac.Value Then
        B1_3 = " and ( year(a.Fecnac)>=year(getdate())-18 or year(a.Fecnac)<=year(getdate())-74 or year(b.Fecnaci)>=year(getdate())-18 or year(b.Fecnaci)<=year(getdate())-74 or a.fecnac >= CONVERT(datetime,'" & TxtFecha & "',103) or b.fecnaci >= CONVERT(datetime,'" & TxtFecha & "',103) ) "
    End If
    If OptFecNacDif.Value Then
        B1_3 = " and ( a.fecnac<>b.Fecnaci or a.fecnac is null or b.Fecnaci is null ) "
    End If
    If OptUbiINEIDif.Value Then
        B1_3 = " and ( a.ubigeonac<>b.ubigeonacine or a.ubigeonac='' or b.ubigeonacine='' ) "
    End If
    If OptUbiRenDif.Value Then
        B1_3 = " and ( a.ubigeoren<>b.ubigeonacren or a.ubigeoren='' or b.ubigeonacren='' ) "
    End If
    If OptNacionSx.Value Then
        B1_3 = " and ( a.nac is null or b.nacion is null or a.nac<>b.nacion or a.sexo is null or b.sexo is null or a.sexo<>b.sexo ) "
    End If
    If optFecs.Value Then
        B1_3 = " and (IndFecSus = 1 or  IndFecVer = 1 or IndFecDec = 1 )  "
    End If
    If optUbicor.Value Then
        B1_3 = " and (a.ubigeo <> b.codgeocor )  "
    End If
    
    If optIndErrAdm.Value Then
        B1_3 = " and (a.IndErrPro = 1)  "
    End If
    
    
    B1_4 = " order by a.lote,a.numsol,a.secu "

    B1 = B1_1 + B1_2 + B1_3 + B1_4
    
    LlenaControles
        
End Sub

Private Sub LlenaControles()
    LblTotReg = 0
    
    If RsVeri.State = 1 Then RsVeri.Close
    
    If optDif Then
        B1 = " select * from VT_SolicHJ where CEmp <> CHV "
        If OptFecha.Value Then
            B1_2 = " and fecsal = '" & Me.TxtFecha & "' "
        Else
            If OptLote.Value Then
                B1_2 = " and lote = '" & TxtLote & "' "
            End If
        End If
        
        B1 = " select * from VT_SolicHJ where CEmp <> CHV " & B1_2
        
        RsVeri.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    Else
        RsVeri.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    End If
    
    
    With RsVeri
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            Set DtGVeri.DataSource = Nothing
            DtGVeri.Refresh
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Verificacion"
            Exit Sub
        Else
            Set DtGVeri.DataSource = RsVeri
            DtGVeri.Refresh
            LblTotReg.Caption = RsVeri.RecordCount
            CmdGrabar.Enabled = True
            CmdVerImag.Enabled = True
            CargaImag
        End If
    End With
End Sub

Private Sub DtGVeri_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "V"
    If Activado Then
        On Error GoTo Salir
        With RsVeri
            xlote = .Fields("lote")
            xnumsol = .Fields("numsol")
            xidimagen = "01" + xnumsol
            FrmMuestraImg.Show
        End With
    Else
        MsgBox "Selecciona la solicitud a buscar", vbCritical, "Error en la busqueda "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CmdGrabar_Click()
    RsVeri.Fields("UsuRepVida") = gNomUsu
    RsVeri.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub CmdNacSxB_Click()
    If OptFecha.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update SOLICITUD " & _
            "set nac=b.nacion " & _
            "from SOLICITUD a join HOJAVIDA b on a.NUMSOL=b.NUMSOL and a.fecsal=b.fecsal and a.secu=b.sechv " & _
            "WHERE a.fecsal= '" & TxtFecha & "' and a.nac is null and b.nacion is not null "
            .Execute
            .CommandText = "Update SOLICITUD " & _
            "set sexo=b.sexo " & _
            "from SOLICITUD a join HOJAVIDA b on a.NUMSOL=b.NUMSOL and a.fecsal=b.fecsal and a.secu=b.sechv " & _
            "WHERE a.fecsal= '" & TxtFecha & "' and a.sexo='' and b.sexo<>'' "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Reemplaza Nac y Sx"
    Else
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update SOLICITUD " & _
            "set nac=b.nacion " & _
            "from SOLICITUD a join HOJAVIDA b on a.NUMSOL=b.NUMSOL and a.fecsal=b.fecsal and a.secu=b.sechv " & _
            "WHERE a.lote= '" & TxtLote & "' and a.nac is null and b.nacion is not null "
            .Execute
            .CommandText = "Update SOLICITUD " & _
            "set sexo=b.sexo " & _
            "from SOLICITUD a join HOJAVIDA b on a.NUMSOL=b.NUMSOL and a.fecsal=b.fecsal and a.secu=b.sechv " & _
            "WHERE a.lote= '" & TxtLote & "' and a.sexo='' and b.sexo<>'' "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Reemplaza Nac y Sx"
    End If
End Sub

Private Sub CmdUbigeo_Click()
    FrmConsUbi.Show
End Sub

Private Sub txtEnvio_LostFocus()
   txtEnvio.Text = Format(txtEnvio.Text, "000000#")
End Sub

Private Sub TxtFecha_GotFocus()
    Marcado TxtFecha
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
    If KeyAscii = 13 Then Me.CmdCons.SetFocus
End Sub

Private Sub TxtFecha_LostFocus()
    ValidaFecha TxtFecha
End Sub



Private Sub TxtLote_LostFocus()
    TxtLote = Format(TxtLote, "0000000#")
End Sub

Private Sub TxtRia_KeyPress(KeyAscii As Integer)
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
    xfrmveriRia = ""
    Unload Me
End Sub

Private Sub CargaImag()
On Error Resume Next
    With RsVeri
        xlote = Format(!Lote, "0000000#")
        xlote2 = Format(!Lote, "0000000#")
        xnumsol = Format(!Numsol, "000000#")
        If !Secu = 0 Then
            xtipimag1 = "S"
        Else
            xtipimag1 = "E"
        End If
        xtipimag2 = "V"
        xidimagen = "01" + xnumsol
        xinddoc = !IndDoc
    End With
    xnumpag = 1
      
    If RsImag1.State = 1 Then RsImag1.Close
    RsImag1.CursorLocation = adUseClient
    
    SQLImag1 = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and ( TipImag = '" & xtipimag1 & "' ) and substring(idimagen,3,9)= '" & xnumsol & "' and posicion = " & RsVeri.Fields("posicion") & " order by posicion "
    
    RsImag1.Open SQLImag1, Con.Cnx, adOpenStatic, adLockReadOnly
'******
    If RsImag2.State = 1 Then RsImag2.Close
    RsImag2.CursorLocation = adUseClient
    
    SQLImag2 = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and ( TipImag = '" & xtipimag2 & "' ) and substring(idimagen,3,9)= '" & xnumsol & "' and posicion = " & RsVeri.Fields("posima") & "   order by posicion "
    
    RsImag2.Open SQLImag2, Con.Cnx, adOpenStatic, adLockReadOnly
'******
    
    With RsImag1          'coneccion a _Imagenes
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            ImgEdit1.ClearDisplay
            Exit Sub
        Else
            .MoveFirst
'******
            With RsImag2          'coneccion a _Imagenes
                If .BOF And .EOF Then   'pregunta si tabla esta vacia
                    MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
                    ImgEdit2.ClearDisplay
                    Exit Sub
                Else
                    .MoveFirst
                End If
            End With
'******
            MuestraImg
        End If
    End With
'******
End Sub

Private Sub MuestraImg()
    
    With RsImag1  'coneccion a _Imagenes
'        .Find "[posicion]='" & xnumpag & "'"
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
    
    If xinddoc = "C" Then
        ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 15) * 3)
    Else
        ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 15) * 2)
    End If
    ImgEdit1.ScrollPositionX = 0
    
    
'    ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    ImgEdit1.Refresh
'******
    With RsImag2  'coneccion a _Imagenes
'        .Find "[posicion]='" & xnumpag & "'"
'        .Find "[idimagen]='" & xidimagen & "'"
        
        If .EOF Then .MovePrevious
        ximagen = IIf(IsNull(!Imagen), "", !Imagen)
        xrutavol = .Fields("ruta")
        xrutavol = Replace(xrutavol, "\\", "\")
        xidvolumen = IIf(IsNull(!idvolumen), "", !idvolumen)
    End With
      
    ImgEdit2.Image = xrutavol & "\" & xlote2 & "\" & ximagen & ".TIF"
    ImgEdit2.Enabled = True

    ImgEdit2.FitTo 1
'    ImgEdit2.Zoom = 50
    ImgEdit2.Display
    
    If OptRiasMay.Value Or OptRiaMenor.Value Then
        ImgEdit2.ScrollPositionY = (Int(ImgEdit2.ImageScaleHeight / 4) * 4)
    Else
        ImgEdit2.ScrollPositionY = (Int(ImgEdit2.ImageScaleHeight / 15) * 2)
    End If
    
    If optUbicor.Value Then
        ImgEdit2.ScrollPositionY = (Int(ImgEdit2.ImageScaleHeight / 7) * 2)
    End If
   
    
    ImgEdit2.ScrollPositionX = 0
   
'    ImgEdit2.ScrollPositionX = (Int((ImgEdit2.ImageScaleWidth / 3)) * 2)
    ImgEdit2.Refresh

End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'' alto
'    ImgEdit1.Height = FrmVeriRiaFec.Height - 7800
'' ancho
'    ImgEdit1.Width = FrmVeriRiaFec.Width - 300
'' alto
''    DtGHV.Height = 1200 - ImgEdit1.Height
'    ImgEdit1.Height = FrmVeriRiaFec.Height - 6500
'End Sub
'**************

Sub MovImg(KeyCode As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF12 Then
        If Not (RsImag1.EOF And RsImag1.BOF) Then
            RsImag1.MoveNext
            If RsImag1.EOF Then RsImag1.MoveLast
            'ImgEdit1.Image = Rutaimag(RsImag1!Ruta, txtLoteAfp, RsImag1!imagen)
            ImgEdit1.Image = Rutaimag(RsImag1!Ruta, RsImag1!lotefisico, RsImag1!Imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If

    If KeyCode = vbKeyF11 Then
        If Not (RsImag1.EOF And RsImag1.BOF) Then
            RsImag1.MovePrevious
            If RsImag1.BOF Then RsImag1.MoveFirst
            'ImgEdit1.Image = Rutaimag(RsImag1!Ruta, txtLoteAfp, RsImag1!imagen)
            ImgEdit1.Image = Rutaimag(RsImag1!Ruta, RsImag1!lotefisico, RsImag1!Imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If
End Sub

Sub MovImg2(KeyCode As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF12 Then
        If Not (RsImag2.EOF And RsImag2.BOF) Then
            RsImag2.MoveNext
            If RsImag2.EOF Then RsImag2.MoveLast
            'ImgEdit2.Image = Rutaimag(RsImag2!Ruta, txtLoteAfp, RsImag2!imagen)
            ImgEdit2.Image = Rutaimag(RsImag2!Ruta, RsImag2!lotefisico, RsImag2!Imagen)
      
            ImgEdit2.Display
            ImgEdit2.Refresh
        End If
    End If

    If KeyCode = vbKeyF11 Then
        If Not (RsImag2.EOF And RsImag2.BOF) Then
            RsImag2.MovePrevious
            If RsImag2.BOF Then RsImag2.MoveFirst
            'ImgEdit2.Image = Rutaimag(RsImag2!Ruta, txtLoteAfp, RsImag2!imagen)
            ImgEdit2.Image = Rutaimag(RsImag2!Ruta, RsImag2!lotefisico, RsImag2!Imagen)
      
            ImgEdit2.Display
            ImgEdit2.Refresh
        End If
    End If
End Sub

Private Sub ImgEdit1_DblClick()
    On Error Resume Next
    ImgEdit1.FitTo 1
End Sub

Private Sub ImgEdit2_DblClick()
    On Error Resume Next
    ImgEdit2.FitTo 1
End Sub

Private Sub ImgEdit1_KeyDown(KeyCode As Integer, Shift As Integer)
    MovImg KeyCode
End Sub

Private Sub ImgEdit2_KeyDown(KeyCode As Integer, Shift As Integer)
    MovImg2 KeyCode
End Sub

Private Sub ImgEdit1_MarkEnd(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MarkType As Integer, ByVal GroupName As String)
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

Private Sub ImgEdit2_MarkEnd(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MarkType As Integer, ByVal GroupName As String)
    ImgEdit2.ZoomToSelection
    ImgEdit2.RemoveAllOCRMarks
    ImgEdit2.Refresh
End Sub

Private Sub ImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

Private Sub ImgEdit2_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    ImgEdit2.ZoomToSelection
    ImgEdit2.RemoveAllOCRMarks
    ImgEdit2.Refresh
End Sub

'********************
Private Sub ImgEdit1_Click()
    On Error Resume Next
End Sub

Private Sub ImgEdit2_Click()
    On Error Resume Next
End Sub

'Private Sub UpDown1_DownClick()
'    On Error Resume Next
'    If Not (RsImag1.EOF And RsImag1.BOF) Then
'        RsImag1.MoveNext
'        If RsImag1.EOF Then RsImag1.MoveLast
'        'ImgEdit1.Image = Rutaimag(RsImag1!Ruta, txtLoteAfp, RsImag1!imagen)
'        ImgEdit1.Image = Rutaimag(xrutavol, RsImag1!lotefisico, RsImag1!Imagen)
'        ImgEdit1.Display
'        ImgEdit1.Refresh
'    End If
'End Sub

'Private Sub UpDown2_DownClick()
'    On Error Resume Next
'    If Not (RsImag2.EOF And RsImag2.BOF) Then
'        RsImag2.MoveNext
'        If RsImag2.EOF Then RsImag2.MoveLast
'        'ImgEdit2.Image = Rutaimag(RsImag2!Ruta, txtLoteAfp, RsImag2!imagen)
'        ImgEdit2.Image = Rutaimag(xrutavo2, RsImag2!lotefisico, RsImag2!Imagen)
'        ImgEdit2.Display
'        ImgEdit2.Refresh
'    End If
'End Sub
'
'Private Sub UpDown1_UpClick()
'    On Error Resume Next
'    If Not (RsImag1.EOF And RsImag1.BOF) Then
'        RsImag1.MovePrevious
'        If RsImag1.BOF Then RsImag1.MoveFirst
'        'ImgEdit1.Image = Rutaimag(RsImag1!Ruta, txtLoteAfp, RsImag1!imagen)
'        ImgEdit1.Image = Rutaimag(xrutavol, RsImag1!lotefisico, RsImag1!Imagen)
'        ImgEdit1.Display
'        ImgEdit1.Refresh
'    End If
'End Sub

'Private Sub UpDown2_UpClick()
'    On Error Resume Next
'    If Not (RsImag2.EOF And RsImag2.BOF) Then
'        RsImag2.MovePrevious
'        If RsImag2.BOF Then RsImag2.MoveFirst
'        'ImgEdit2.Image = Rutaimag(RsImag2!Ruta, txtLoteAfp, RsImag2!imagen)
'        ImgEdit2.Image = Rutaimag(xrutavol, RsImag2!lotefisico, RsImag2!Imagen)
'        ImgEdit2.Display
'        ImgEdit2.Refresh
'    End If
'End Sub


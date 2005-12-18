VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmVerificacionD2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación : Segunda Digitación"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15225
   Begin VB.Frame FraAyuda 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   6600
      TabIndex        =   66
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdCopiar 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Copiar Dato"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdCerrar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblAyuda 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   855
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   240
      TabIndex        =   57
      Top             =   -120
      Width           =   14895
      Begin VB.CheckBox chkIndicador 
         Caption         =   "Ver Todos"
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
         Left            =   6720
         TabIndex        =   71
         ToolTipText     =   "Ver todos inclusive los que ya se hicieron verificación"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
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
         Left            =   5040
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox txtLote 
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
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   195
         Width           =   1215
      End
      Begin VB.TextBox txtFecsal 
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
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   1
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Total"
         Height          =   255
         Left            =   8760
         TabIndex        =   62
         Top             =   285
         Width           =   495
      End
      Begin VB.Label lblcont 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   14160
         TabIndex        =   61
         Top             =   240
         Width           =   15
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "0"
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
         Left            =   9480
         TabIndex        =   60
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   59
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   240
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5640
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "numsol"
         Caption         =   "Nro Solicitud"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Secu"
         Caption         =   "Secuencia"
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
         DataField       =   "apepat"
         Caption         =   "Ape. Pat"
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
         DataField       =   "apemat"
         Caption         =   "Ape. Mat."
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
         DataField       =   "Nombre1"
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
      BeginProperty Column05 
         DataField       =   "Nombre2"
         Caption         =   "Nombre2"
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
         DataField       =   "tipdoc"
         Caption         =   "Tipo Doc"
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
         DataField       =   "numdoc"
         Caption         =   "Num. Doc."
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
         DataField       =   "numcus"
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "Codusu"
         Caption         =   "Usuario 1er"
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
         DataField       =   "Codmodd2"
         Caption         =   "Usuario 2da"
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
         DataField       =   "eUrb"
         Caption         =   "Urbanización"
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
         DataField       =   "eNumero"
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
      BeginProperty Column14 
         DataField       =   "eDireccion"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2880
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   4935.118
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "cmdAnterior"
      Height          =   5295
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   15135
      Begin Traspasos27.UserControl1 ucUbigeoEmp 
         Height          =   375
         Left            =   4680
         TabIndex        =   23
         Top             =   3840
         Width           =   5775
         _extentx        =   10610
         _extenty        =   661
      End
      Begin VB.Frame frmAfiliado 
         Caption         =   "Identificación del Afiliado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2775
         Left            =   240
         TabIndex        =   26
         Top             =   0
         Width           =   14175
         Begin Traspasos27.UserControl1 ucUbigeoAfi 
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Top             =   2280
            Width           =   8055
            _extentx        =   14208
            _extenty        =   661
         End
         Begin VB.ComboBox cboPrefLocAfi 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmVerificacionD2.frx":0000
            Left            =   8640
            List            =   "frmVerificacionD2.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox cboPrefDirAfi 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmVerificacionD2.frx":0004
            Left            =   240
            List            =   "frmVerificacionD2.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtUrba 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   10200
            MaxLength       =   100
            TabIndex        =   17
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox txtNroDir 
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
            Left            =   6120
            MaxLength       =   40
            TabIndex        =   15
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txtAv 
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
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   14
            Top             =   1680
            Width           =   4215
         End
         Begin VB.TextBox txtcorreo 
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
            Left            =   8520
            MaxLength       =   50
            TabIndex        =   12
            Top             =   1080
            Width           =   4455
         End
         Begin VB.TextBox txtTipFO 
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
            Left            =   6600
            MaxLength       =   12
            TabIndex        =   10
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtTipFD 
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
            Left            =   7560
            MaxLength       =   12
            TabIndex        =   11
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton cmdReniec 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ver &Reniec"
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
            Left            =   12240
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtTipDoc 
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
            Left            =   240
            MaxLength       =   3
            TabIndex        =   6
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtNumDoc 
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
            TabIndex        =   7
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtApePat 
            BackColor       =   &H00FFFFFF&
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
            Left            =   240
            MaxLength       =   30
            TabIndex        =   2
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtApeMat 
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
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   3
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtNombreuno 
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
            Left            =   5760
            MaxLength       =   20
            TabIndex        =   4
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtCUSPP 
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
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   8
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtFecNac 
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
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   9
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtNombredos 
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
            Left            =   8520
            MaxLength       =   20
            TabIndex        =   5
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento / Provincia / Distrito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Urba."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8760
            TabIndex        =   83
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Vía"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Urbanización"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11040
            TabIndex        =   81
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Número (N) / Departamento (Dpto) "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   80
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Avenida (Av.) / Calle (CL)/ Pasaje (Pj) / Jirón (Jr)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   79
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "E- mail"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8640
            TabIndex        =   74
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblOrigen 
            BackStyle       =   0  'Transparent
            Caption         =   "F.Origen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6600
            TabIndex        =   73
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblDestino 
            BackStyle       =   0  'Transparent
            Caption         =   "F.Destino"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   72
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Doc "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Doc. "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblCUSPP 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSPP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   32
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblApePat 
            BackStyle       =   0  'Transparent
            Caption         =   "Apell. Paterno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblApeMat 
            BackStyle       =   0  'Transparent
            Caption         =   "Apell. Materno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   30
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblNomuno 
            BackStyle       =   0  'Transparent
            Caption         =   "1er Nombre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblNomdos 
            BackStyle       =   0  'Transparent
            Caption         =   "2do Nombre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8640
            TabIndex        =   28
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nac"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   27
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.TextBox txtFactor 
         Height          =   375
         Left            =   14520
         TabIndex        =   39
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salir"
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
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Grabar"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdrotIzq 
         Caption         =   "<"
         Height          =   495
         Left            =   12960
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Rotar Izquierda"
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton cmdrotDer 
         Caption         =   ">"
         Height          =   495
         Left            =   14400
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Rotar Derecha"
         Top             =   4440
         Width           =   495
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   14895
         TabIndex        =   40
         Top             =   720
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         OrigLeft        =   14760
         OrigTop         =   720
         OrigRight       =   15000
         OrigBottom      =   1095
         Increment       =   20
         Max             =   100
         Min             =   -100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit2 
         Height          =   1455
         Left            =   12840
         TabIndex        =   41
         Top             =   2880
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   2566
         _StockProps     =   0
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   13560
         TabIndex        =   51
         Text            =   "Text4"
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   13560
         TabIndex        =   52
         Text            =   "Text3"
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   13560
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   13440
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame frmEmpleador 
         Caption         =   "Datos del Empleador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1575
         Left            =   240
         TabIndex        =   42
         Top             =   2760
         Width           =   12495
         Begin VB.ComboBox cboPrefLocEmp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmVerificacionD2.frx":0008
            Left            =   120
            List            =   "frmVerificacionD2.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1080
            Width           =   1575
         End
         Begin VB.ComboBox cboPrefDirEmp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmVerificacionD2.frx":000C
            Left            =   6240
            List            =   "frmVerificacionD2.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtTelefEmp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10320
            MaxLength       =   10
            TabIndex        =   24
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox lblRazonSoc 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   19
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtUrbaEmp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   22
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtRuc 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   11
            TabIndex        =   18
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtNroEmp 
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
            Left            =   10800
            MaxLength       =   40
            TabIndex        =   21
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtDireEmp 
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
            Left            =   7440
            MaxLength       =   100
            TabIndex        =   20
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Vía"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   78
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Urba"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Razon Social"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento / Provincia / Distrito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   49
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10680
            TabIndex        =   48
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Urbanización"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   47
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Av /CL/ Pasaje (Pj) / Jirón (Jr)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   46
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Ruc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   44
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Número (N) / Departamento (Dpto) / Interior(Int) / Manzana (Mz)  / Lote (Lte.)  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11040
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label26 
         Caption         =   "Ajuste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14520
         TabIndex        =   56
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblMensaje 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1920
         TabIndex        =   55
         Top             =   4560
         Width           =   6615
      End
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   2295
      Left            =   240
      TabIndex        =   63
      Top             =   7560
      Width           =   14895
      _Version        =   65536
      _ExtentX        =   26273
      _ExtentY        =   4048
      _StockProps     =   0
      ImageControl    =   "ImgEdit"
      AnnotationBackColor=   4210752
      AnnotationFillColor=   65535
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
   End
End
Attribute VB_Name = "frmVerificacionD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSolicitud As Recordset

Dim rsUbigeoEmp As Recordset
Dim rsUbigeoAfi As Recordset
Public ControlOrig As Control

Dim RsImag As Recordset
Private IndDoc As String

Sub CargarDatosSolicitud()
Dim lcCadena As String
Dim Condicion As String

lcCadena = "Select a.CodAfp , Secu, a.FecSal , a.Lote , a.Inddoc, a.NumSol , a.ApePat , a.ApeMat, a.Nombre1 , a.Nombre2  , " & _
" a.TipoDoc , a.NumDoc, CUSPP, TipoFonOri, TipoFonDes, e_mail ,a.Empleadores , a.Codusu , a.fecdig, a.ApePatD2 , a.ApeMatD2, a.Nombre1D2 , a.Nombre2D2, CUSPPD2 , Tipodocd2, numdocd2,fecnac=isnull(fecnac,''),fecnacD2 = isnull(fecnacD2,''), a.direccion, a.Urb, a.Numero, " & _
" imagen, d.Ruta, c.LoteFisico, Secuencial, NumRuc, e.Razsoc, a.CodModD2, " & _
" eDireccion = e.Direccion, eNumero = e.Numero, eUrb = e.Urb , eDistrito = f.Distrito, eProvincia = f.Provincia, " & _
" eDepartamento = f.Departamento, eUbigeo = e.Ubigeo, eTelefono = e.Telefono, eDesgeo = f.desgeo, a.posicion, TipoFonOriD2=isnull(TipoFonOriD2,''), TipoFonDesD2=isnull(TipoFonDesD2,''), " & _
" e.NumRucD2, e.DireccionD2, e.NumeroD2 , e.UrbD2, e.UbigeoD2, e.TelefonoD2, e.PrefDirEmp, e.PrefLocEmp, e.PrefDirEmpD2, e.PrefLocEmpD2, pl.AbreLoc, pd.AbreDir, AbreLocD2 = pl2.AbreLoc, AbreDirD2 = pd2.AbreDir, plf.AbreLoc, pdf.AbreDir, fAbreLocD2 = plf2.AbreLoc, fAbreDirD2 = pdf2.AbreDir, eDesgeoD2 = (select Desgeo from ubigeo where codgeo = e.ubigeoD2 ), a.PrefDirAfi , a.PrefLocAfi, a.c_PrefDirAfiD2 , a.c_PrefLocAfiD2, a.c_direccionD2, a.c_UrbD2, a.c_NumeroD2, a.c_UbigeoD2, c_DesgeoD2 = (select Desgeo from ubigeo where codgeo = a.c_UbigeoD2 ), c_Desgeo = (select Desgeo from ubigeo where codgeo = a.Ubigeo ), a.Ubigeo  " & _
" from solicitud a join _Imagenes c on c.idImagen = a.idImagen and  a.posicion = c.posicion " & _
" join _Volumenes d on d.idvolumen = c.idvolumen " & _
" left join Empleadores e on e.codafp = a.codafp and e.numsol = a.numsol  and e.fecsal = a.fecsal and e.Secuencial = a.secu " & _
" left join Ubigeo f on E.ubigeo = F.codgeo  left join _Scan s on a.lote = s.lotefisico " & _
" left join PrefLoc pl on pl.CodLoc = e.PrefLocEmp  left join PrefDir pd on pd.CodDir = e.PrefDirEmp left join PrefLoc plf on plf.CodLoc = a.PrefLocAfi  left join PrefDir pdf on pdf.CodDir = a.PrefDirAfi " & _
" left join PrefLoc pl2 on pl2.CodLoc = e.PrefLocEmpD2 left join PrefDir pd2 on pd2.CodDir = e.PrefDirEmpD2  left join PrefLoc plf2 on plf2.CodLoc = a.c_PrefLocAfiD2 left join PrefDir pdf2 on pdf2.CodDir = a.c_PrefDirAfiD2  " & _
" where indD2 = 1 and (TipImag = 'S' or TipImag = 'E') and a.secu = 0 and s.estado in ('5','6')  and (a.ApePat <> a.ApePatD2 or a.ApeMat <> a.ApeMatD2 or  " & _
" a.Nombre1 <> a.Nombre1D2 or  a.Nombre2 <> a.Nombre2D2 or  a.Tipodoc <> a.TipoDocD2 or a.Numdoc <> a.NumdocD2 or a.direccion<>a.c_direccionD2 or a.Urb<>a.c_UrbD2 or a.Numero<>a.c_NumeroD2 or a.Ubigeo<>a.c_UbigeoD2 or" & _
" a.CUSPP <> a.CUSPPD2 or PrefDirEmp <> PrefDirEmpD2 or PrefLocEmp <> PrefLocEmpD2 or (isnull(a.fecnac,'') <> isnull(a.fecnacD2,'') " & _
" and a.inddoc = 'C') or TipoFonOri <> TipoFonOriD2 or  TipoFonDes <> TipoFonDesD2 " & _
" or e.NumRuc <> e.NumRucD2 or  e.Direccion <> e.DireccionD2 or " & _
" e.Numero <> e.NumeroD2 or  e.Urb <> e.UrbD2 or " & _
" e.Telefono <> e.TelefonoD2 or  PrefDirEmp <> PrefDirEmpD2 or " & _
" PrefLocEmp <> PrefLocEmpD2 ) " & IIf(chkIndicador.Value = 1, "", " and isnull(a.indVerD2,0) = 0")
     
If Trim(txtFecsal.Text) <> "" Then Condicion = Condicion & " and convert(char,a.fecsal,103) = '" & Trim(txtFecsal.Text) & "' "
If Trim(txtLote.Text) <> "" Then Condicion = Condicion & " and a.Lote = '" & Trim(txtLote.Text) & "'"
'If Trim(txtEnvio.Text) <> "" Then Condicion = Condicion & " and Envio = '" & Trim(txtEnvio.Text) & "'"
    
lcCadena = lcCadena & Condicion & " order by a.Lote, a.Numsol, a.Secu "

If rsSolicitud.State = 1 Then rsSolicitud.Close
rsSolicitud.Open lcCadena, Con.Cnx, adOpenStatic, adLockReadOnly
 
 
With rsSolicitud

    If (.EOF And .BOF) Then
       cmdGrabar.Enabled = False
       MsgBox "No existen Registros", 64, Caption
       ImgEdit1.ClearDisplay
       ImgEdit2.ClearDisplay
       RefrescarSol False
    Else
        RefrescarSol True
    End If
    
Set DataGrid1.DataSource = rsSolicitud
lblTotal.Caption = .RecordCount

End With
End Sub

Private Sub cboPrefDirAfi_DblClick()
    ControlAyuda cboPrefDirAfi, rsSolicitud.Fields("c_PrefDirAfiD2"), Frame1.Top + cboPrefDirAfi.Top
    Set ControlOrig = cboPrefDirAfi
End Sub

Private Sub cboPrefDirAfi_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 90, 560
End Sub

Private Sub cboPrefDirEmp_DblClick()
   ControlAyuda cboPrefDirEmp, rsSolicitud.Fields("AbreDirD2"), cboPrefDirEmp.Top + Frame1.Top + frmEmpleador.Top
   Set ControlOrig = cboPrefDirEmp
End Sub

Private Sub cboPrefDirEmp_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 846
Else
    Resaltado 0, 0, 1200, 150, 57, 1161
End If
End Sub

Private Sub cboPrefLocAfi_DblClick()
    ControlAyuda cboPrefLocAfi, rsSolicitud.Fields("c_PrefLocAfiD2"), Frame1.Top + cboPrefLocAfi.Top
    Set ControlOrig = cboPrefLocAfi
End Sub
Private Sub cboPrefLocAfi_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1600, 560
End Sub
Private Sub cboPrefLocEmp_DblClick()
    ControlAyuda cboPrefLocEmp, rsSolicitud.Fields("AbreLocD2"), cboPrefLocEmp.Top + Frame1.Top + frmEmpleador.Top
    Set ControlOrig = cboPrefLocEmp
End Sub

Private Sub cboPrefLocEmp_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1769, 846
Else
    Resaltado 0, 0, 1200, 150, 1695, 1161
End If

End Sub

Private Sub cmdBuscar_Click()
CargarDatosSolicitud
End Sub

Private Sub cmdCerrar_Click()
FraAyuda.Visible = False
ControlOrig.SetFocus
End Sub

Private Sub cmdCopiar_Click()
On Error Resume Next
If TypeOf ControlOrig Is TextBox Then
    ControlOrig.Text = lblAyuda.Caption
End If
If TypeOf ControlOrig Is UserControl1 Then
    ControlOrig.cTexto = lblAyuda.Caption
End If

End Sub

Private Sub CmdGrabar_Click()
GrabarVerificacion
End Sub

Private Sub cmdReniec_Click()
SiempreVisible frmReniec, True
End Sub

Private Sub DataGrid1_Click()
RefrescarSol True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
RefrescarSol True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
MovImg KeyCode
End Sub

Private Sub Form_Load()
'CenterForm Me

Set rsSolicitud = New ADODB.Recordset

Set ucUbigeoAfi.cConexion = Con.Cnx
    ucUbigeoAfi.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    ucUbigeoAfi.Campocondicion = "desgeo"
    ucUbigeoAfi.Busqinicial = False
    ucUbigeoAfi.Enabled = True
    ucUbigeoAfi.nroCaracteres = 3
    Set rsUbigeoAfi = New Recordset

Set ucUbigeoEmp.cConexion = Con.Cnx
    ucUbigeoEmp.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    ucUbigeoEmp.Campocondicion = "desgeo"
    ucUbigeoEmp.Busqinicial = False
    ucUbigeoEmp.Enabled = True
    ucUbigeoEmp.nroCaracteres = 3
    Set rsUbigeoEmp = New Recordset
    
CargarComboSql cboPrefDirAfi, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocAfi, "select AbreLoc, CodLoc  from prefloc"

CargarComboSql cboPrefDirEmp, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocEmp, "select AbreLoc, CodLoc  from prefloc"


'IndDoc = "T"
'
'Me.ImgEdit1.Image = "c:\12229_00000003.tif"
'ImgEdit1.Display
End Sub



Sub RefrescarSol(vbResImag As Boolean)
On Error Resume Next
Call InicioControl
With rsSolicitud
If Not (.EOF And .BOF) Then
'        txtFecsal = IIf(IsNull(.Fields("Fecsal")), "", .Fields("Fecsal"))
'        txtLote = IIf(IsNull(.Fields("Lote")), "", .Fields("Lote"))
'        txtNroSolicitud = IIf(IsNull(.Fields("Numsol")), "", .Fields("Numsol"))
        IndDoc = .Fields("Inddoc")
        
        If IndDoc = "T" Then
                txtFecNac.Visible = False
                Label17.Visible = False
        Else
                txtFecNac.Visible = True
                Label17.Visible = True
        End If
        
        txtApePat.Text = IIf(IsNull(.Fields("Apepat")), "", Trim(.Fields("Apepat")))
        txtApeMat.Text = IIf(IsNull(.Fields("Apemat")), "", Trim(.Fields("Apemat")))
        txtNombreuno = IIf(IsNull(.Fields("Nombre1")), "", Trim(.Fields("Nombre1")))
        txtNombredos = IIf(IsNull(.Fields("Nombre2")), "", Trim(.Fields("Nombre2")))
        txtTipDoc = IIf(IsNull(.Fields("TipoDoc")), "", Trim(.Fields("TipoDoc")))
        txtNumDoc = IIf(IsNull(.Fields("Numdoc")), "", Trim(.Fields("Numdoc")))
        txtCUSPP = IIf(IsNull(.Fields("CUSPP")), "", Trim(.Fields("CUSPP")))
        
        
        txtTipFO.Text = IIf(IsNull(.Fields("TipoFonOri")), "", Trim(.Fields("TipoFonOri")))
        txtTipFD.Text = IIf(IsNull(.Fields("TipoFonDes")), "", Trim(.Fields("TipoFonDes")))
        txtcorreo.Text = IIf(IsNull(.Fields("e_mail")), "", Trim(.Fields("e_mail")))
        
        txtAv = IIf(IsNull(.Fields("Direccion")), "", Trim(.Fields("Direccion")))
        txtNroDir = IIf(IsNull(.Fields("Numero")), "", Trim(.Fields("Numero")))
        txtUrba = IIf(IsNull(.Fields("Urb")), "", Trim(.Fields("Urb")))
        
        ucUbigeoAfi.cTexto = IIf(IsNull(.Fields("c_Desgeo")), "", Trim(.Fields("c_Desgeo") & " / " & .Fields("Ubigeo")))
         
        pgBuscarEnCombo cboPrefDirAfi, IIf(IsNull(.Fields("PrefDirAfi")), "", Trim(.Fields("PrefDirAfi"))), 2
        pgBuscarEnCombo cboPrefLocAfi, IIf(IsNull(.Fields("PrefLocAfi")), "", Trim(.Fields("PrefLocAfi"))), 2

        pgBuscarEnCombo cboPrefDirEmp, IIf(IsNull(.Fields("PrefDirEmp")), "", Trim(.Fields("PrefDirEmp"))), 2
        pgBuscarEnCombo cboPrefLocEmp, IIf(IsNull(.Fields("PrefLocEmp")), "", Trim(.Fields("PrefLocEmp"))), 2
        
        '-------------------Datos de Empleadores --------------
        
         txtRuc = IIf(IsNull(.Fields("NumRuc")), "", Trim(.Fields("NumRuc")))
         txtDireEmp = IIf(IsNull(.Fields("eDireccion")), "", Trim(.Fields("eDireccion")))
         txtNroEmp = IIf(IsNull(.Fields("eNumero")), "", Trim(.Fields("eNumero")))
         txtUrbaEmp = IIf(IsNull(.Fields("eUrb")), "", Trim(.Fields("eUrb")))
         ucUbigeoEmp.cTexto = IIf(IsNull(.Fields("eDesgeo")), "", Trim(.Fields("eDesgeo") & " / " & .Fields("eUbigeo")))
         txtTelefEmp = IIf(IsNull(.Fields("eTelefono")), "", Trim(.Fields("eTelefono")))
         lblRazonSoc.Text = IIf(IsNull(.Fields("Razsoc")), "", Trim(.Fields("Razsoc")))
         txtFecNac.Text = IIf(IsNull(.Fields("Fecnac")), "", Trim(.Fields("Fecnac")))
         
        '---------------------------------------
        '--- Visualizar Imagen correspondiente
        If vbResImag Then
            ImgEdit1.Image = Rutaimag(.Fields("Ruta"), .Fields("LoteFisico"), .Fields("Imagen"))
            ImgEdit1.Display
        End If
                
        ImgEdit2.Image = ImgEdit1.Image
        ImgEdit2.FitTo 0
        ImgEdit2.Display
        lblcont.Caption = .AbsolutePosition

        Call Comparar
        Set RsImag = CargarImagenes(IIf(IsNull(.Fields("numsol")), "", Trim(.Fields("numsol"))), .Fields("LoteFisico"))
        
Else
        txtApePat.Text = ""
        txtApeMat.Text = ""
        txtNombreuno = ""
        txtNombredos = ""
        txtTipDoc = ""
        txtNumDoc = ""
        txtCUSPP = ""
        txtFecNac.Text = ""
                                                        
        txtAv = ""
        txtNroDir = ""
        txtUrba = ""
        cboPrefDirAfi.ListIndex = -1
        cboPrefLocAfi.ListIndex = -1
                                                        
        txtRuc = ""
        lblRazonSoc.Text = ""
        txtDireEmp.Text = ""
        txtNroEmp.Text = ""
        txtUrbaEmp.Text = ""
        ucUbigeoAfi.cTexto = ""
        ucUbigeoEmp.cTexto = ""
        txtTelefEmp.Text = ""
        txtTipFO.Text = ""
        txtTipFD.Text = ""
        txtcorreo.Text = ""
        cboPrefDirEmp.ListIndex = -1
        cboPrefLocEmp.ListIndex = -1
        
        
End If

End With

End Sub

Sub LimpiarPantalla()

        txtApePat.Text = ""
        txtApeMat.Text = ""
        txtNombreuno = ""
        txtNombredos = ""
        txtTipDoc = ""
        txtNumDoc = ""
        txtCUSPP = ""
        txtFecNac.Text = ""
        
        txtAv = ""
        txtNroDir = ""
        txtUrba = ""
                                                        
        txtRuc = ""
        lblRazonSoc.Text = ""
        txtDireEmp.Text = ""
        txtNroEmp.Text = ""
        txtUrbaEmp.Text = ""
        ucUbigeoEmp.cTexto = ""
        txtTelefEmp.Text = ""
        ImgEdit1.ClearDisplay
        ImgEdit2.ClearDisplay
        
End Sub
Sub InicioControl()

  txtApePat.BackColor = -2147483643
  txtApeMat.BackColor = -2147483643
  txtNombreuno.BackColor = -2147483643
  txtNombredos.BackColor = -2147483643
  txtTipDoc.BackColor = -2147483643
  txtNumDoc.BackColor = -2147483643
  txtCUSPP.BackColor = -2147483643
  txtFecNac.BackColor = -2147483643
  txtTipFO.BackColor = -2147483643
  txtTipFD.BackColor = -2147483643
        
  txtAv.BackColor = -2147483643
  txtNroDir.BackColor = -2147483643
  txtUrba.BackColor = -2147483643
  
  ucUbigeoAfi.ColorFondo = -2147483643
   
  cboPrefDirAfi.BackColor = -2147483643
  cboPrefLocAfi.BackColor = -2147483643
        
        '-------------------Datos de Empleadores --------------
    
  txtRuc.BackColor = -2147483643
  txtDireEmp.BackColor = -2147483643
  txtNroEmp.BackColor = -2147483643
  txtUrbaEmp.BackColor = -2147483643
  txtUrbaEmp.BackColor = -2147483643
  ucUbigeoEmp.ColorFondo = -2147483643
  txtUrbaEmp.BackColor = -2147483643
  cboPrefDirEmp.BackColor = -2147483643
  cboPrefLocEmp.BackColor = -2147483643
         
        
End Sub
Sub Comparar()
With rsSolicitud

        If Trim(txtApePat.Text) <> IIf(IsNull(.Fields("ApepatD2")), "", Trim(.Fields("ApepatD2"))) Then
            txtApePat.BackColor = &HC0E0FF
        End If
        
        If Trim(txtApeMat.Text) <> IIf(IsNull(.Fields("ApematD2")), "", Trim(.Fields("ApematD2"))) Then
            txtApeMat.BackColor = &HC0E0FF
        End If
        
        If Trim(txtNombreuno.Text) <> IIf(IsNull(.Fields("Nombre1D2")), "", Trim(.Fields("Nombre1D2"))) Then
            txtNombreuno.BackColor = &HC0E0FF
        End If
        
        If Trim(txtNombredos.Text) <> IIf(IsNull(.Fields("Nombre2D2")), "", Trim(.Fields("Nombre2D2"))) Then
            txtNombredos.BackColor = &HC0E0FF
        End If
        
        If Trim(txtTipDoc.Text) <> IIf(IsNull(.Fields("TipoDocD2")), "", Trim(.Fields("TipoDocD2"))) Then
            txtTipDoc.BackColor = &HC0E0FF
        End If
                                                        
        If Trim(txtNumDoc.Text) <> IIf(IsNull(.Fields("NumdocD2")), "", Trim(.Fields("NumdocD2"))) Then
            txtNumDoc.BackColor = &HC0E0FF
        End If
                                
        If Trim(txtCUSPP.Text) <> IIf(IsNull(.Fields("CUSPPD2")), "", Trim(.Fields("CUSPPD2"))) Then
            txtCUSPP.BackColor = &HC0E0FF
        End If
        
        If Trim(txtTipFO.Text) <> IIf(IsNull(.Fields("TipoFonOriD2")), "", Trim(.Fields("TipoFonOriD2"))) Then
            txtTipFO.BackColor = &HC0E0FF
        End If
                                
        If Trim(txtTipFD.Text) <> IIf(IsNull(.Fields("TipoFonDesD2")), "", Trim(.Fields("TipoFonDesD2"))) Then
            txtTipFD.BackColor = &HC0E0FF
        End If
        
         If txtFecNac.Visible Then
            If Trim(txtFecNac.Text) <> IIf(IsNull(.Fields("FecnacD2")), "", Trim(.Fields("FecnacD2"))) Then
                txtFecNac.BackColor = &HC0E0FF
             End If
         End If
         
         If txtAv.Visible Then
            If Trim(txtAv.Text) <> IIf(IsNull(.Fields("c_DireccionD2")), "", Trim(.Fields("c_DireccionD2"))) Then
                txtAv.BackColor = &HC0E0FF
             End If
         End If
         
         If txtNroDir.Visible Then
            If Trim(txtNroDir.Text) <> IIf(IsNull(.Fields("c_NumeroD2")), "", Trim(.Fields("c_NumeroD2"))) Then
                txtNroDir.BackColor = &HC0E0FF
             End If
         End If
          
         If txtUrba.Visible Then
            If Trim(txtUrba.Text) <> IIf(IsNull(.Fields("c_UrbD2")), "", Trim(.Fields("c_UrbD2"))) Then
                txtUrba.BackColor = &HC0E0FF
             End If
         End If
        

                         
        '-------------------Datos de Empleadores --------------
                          
         If Trim(txtRuc.Text) <> IIf(IsNull(.Fields("NumRucD2")), "", Trim(.Fields("NumRucD2"))) Then
            txtRuc.BackColor = &HC0E0FF
         End If
         
         If Trim(txtDireEmp.Text) <> IIf(IsNull(.Fields("DireccionD2")), "", Trim(.Fields("DireccionD2"))) Then
            txtDireEmp.BackColor = &HC0E0FF
         End If
         
         If Trim(txtNroEmp.Text) <> IIf(IsNull(.Fields("NumeroD2")), "", Trim(.Fields("NumeroD2"))) Then
            txtNroEmp.BackColor = &HC0E0FF
         End If
         
         If Trim(txtUrbaEmp.Text) <> IIf(IsNull(.Fields("UrbD2")), "", Trim(.Fields("UrbD2"))) Then
            txtUrbaEmp.BackColor = &HC0E0FF
         End If
         
         If Trim(txtUrbaEmp.Text) <> IIf(IsNull(.Fields("UrbD2")), "", Trim(.Fields("UrbD2"))) Then
            txtUrbaEmp.BackColor = &HC0E0FF
         End If
                  
                  
         If Trim(.Fields("Ubigeo")) <> IIf(IsNull(.Fields("c_UbigeoD2")), "", Trim(.Fields("c_UbigeoD2"))) Then
            ucUbigeoAfi.ColorFondo = &HC0E0FF
         End If
         
         If Trim(.Fields("eUbigeo")) <> IIf(IsNull(.Fields("UbigeoD2")), "", Trim(.Fields("UbigeoD2"))) Then
            ucUbigeoEmp.ColorFondo = &HC0E0FF
         End If
         
         
         If Trim(txtTelefEmp.Text) <> IIf(IsNull(.Fields("TelefonoD2")), "", Trim(.Fields("TelefonoD2"))) Then
            txtTelefEmp.BackColor = &HC0E0FF
         End If
         
         If Right(Trim(cboPrefDirEmp.Text), 2) <> IIf(IsNull(.Fields("PrefDirEmpD2")), "", Trim(.Fields("PrefDirEmpD2"))) Then
            cboPrefDirEmp.BackColor = &HC0E0FF
         End If
         
                  
         If Right(Trim(cboPrefLocEmp.Text), 2) <> IIf(IsNull(.Fields("PrefLocEmpD2")), "", Trim(.Fields("PrefLocEmpD2"))) Then
            cboPrefLocEmp.BackColor = &HC0E0FF
         End If
         
         
        If Right(Trim(cboPrefDirAfi.Text), 2) <> IIf(IsNull(.Fields("c_PrefDirAfiD2")), "", Trim(.Fields("c_PrefDirAfiD2"))) Then
           cboPrefDirAfi.BackColor = &HC0E0FF
        End If
        
        If Right(Trim(cboPrefLocAfi.Text), 2) <> IIf(IsNull(.Fields("c_PrefLocAfiD2")), "", Trim(.Fields("c_PrefLocAfiD2"))) Then
           cboPrefLocAfi.BackColor = &HC0E0FF
        End If
             
End With
End Sub

Sub ControlAyuda(ControlOrigen As Control, TextoMostrar As String, Coordenaday As Integer)


With frmFraAyuda
 .Move ControlOrigen.Left + frmAfiliado.Left, MDIPrincipal.Height - Me.Height + Coordenaday - .Height, ControlOrigen.Width, .Height
.Shape1.Move 0, 0, .Width, .Height
.lblAyuda.Move .lblAyuda.Left, lblAyuda.Top, .Width, .Height
.cmdCerrar.Left = .Width - cmdCerrar.Width
.cmdCopiar.Left = .Width - cmdCopiar.Width
.lblAyuda.Caption = TextoMostrar
 SiempreVisible frmFraAyuda, True
 'frmFraAyuda.Show
End With



'FraAyuda.Visible = True
'FraAyuda.Move ControlOrigen.Left + frmAfiliado.Left, Coordenaday - FraAyuda.Height, ControlOrigen.Width, FraAyuda.Height
'Shape1.Move 0, 0, FraAyuda.Width, FraAyuda.Height
'lblAyuda.Move lblAyuda.Left, lblAyuda.Top, FraAyuda.Width, FraAyuda.Height
'cmdCerrar.Left = FraAyuda.Width - cmdCerrar.Width
'cmdCopiar.Left = FraAyuda.Width - cmdCopiar.Width
'lblAyuda.Caption = TextoMostrar



End Sub

Private Sub ImgEdit2_KeyDown(KeyCode As Integer, Shift As Integer)
 SiempreVisible frmVisorImag, True
 frmVisorImag.ImgEdit1.Image = ImgEdit2.Image
 frmVisorImag.ImgEdit1.Display
 frmVisorImag.ImgEdit1.FitTo 1
 frmVisorImag.Show
End Sub

Private Sub txtApeMat_DblClick()
        ControlAyuda txtApeMat, rsSolicitud.Fields("ApematD2"), Frame1.Top + txtApeMat.Top
        Set ControlOrig = txtApeMat
End Sub

Private Sub txtApePat_DblClick()
        ControlAyuda txtApePat, rsSolicitud.Fields("ApepatD2"), Frame1.Top + txtApePat.Top
        Set ControlOrig = txtApePat
End Sub

Private Sub TxtApePat_LostFocus()
'    FraAyuda.Visible = False
End Sub

Private Sub txtAv_DblClick()
    ControlAyuda txtAv, rsSolicitud.Fields("c_DireccionD2"), Frame1.Top + txtAv.Top
    Set ControlOrig = txtAv
End Sub
Private Sub txtAv_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 100, 560

End Sub

Private Sub txtcorreo_GotFocus()
ImgEdit1.Zoom = 75
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 121, 1458
Else
    Resaltado 0, 0, 1200, 150, 121, 1458
End If
txtcorreo.BackColor = &HC0FFFF
End Sub

Private Sub txtcorreo_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
End Sub

Private Sub txtCUSPP_DblClick()
    ControlAyuda txtCUSPP, rsSolicitud.Fields("CUSPPD2"), txtCUSPP.Top + Frame1.Top + frmAfiliado.Top
     Set ControlOrig = txtCUSPP
End Sub

Private Sub txtDireEmp_DblClick()
   ControlAyuda txtDireEmp, IIf(IsNull(rsSolicitud.Fields("DireccionD2")), "", Trim(rsSolicitud.Fields("DireccionD2"))), txtDireEmp.Top + Frame1.Top + frmEmpleador.Top
   Set ControlOrig = txtDireEmp
End Sub

Private Sub txtFecNac_DblClick()
    ControlAyuda txtNombredos, rsSolicitud.Fields("fecnacD2"), Frame1.Top + txtNombredos.Top
    Set ControlOrig = txtFecNac
End Sub

Private Sub txtFecNAc_GotFocus()
    If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
    Resaltado 0, 0, 1200, 150, 0, 680
'    If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
'    Resaltado 0, 0, 1200, 150, 337, 633
'    txtFecNac.BackColor = &HC0FFFF
End Sub

Private Sub TxtLote_LostFocus()
txtLote = Format(txtLote, "0000000#")
End Sub

Private Sub txtNombredos_DblClick()
    ControlAyuda txtNombredos, rsSolicitud.Fields("Nombre2D2"), Frame1.Top + txtNombredos.Top
    Set ControlOrig = txtNombredos
End Sub

Private Sub txtNombreuno_DblClick()
    ControlAyuda txtNombreuno, rsSolicitud.Fields("Nombre1D2"), Frame1.Top + txtNombreuno.Top
     Set ControlOrig = txtNombreuno
End Sub

Private Sub txtNroDir_DblClick()
    ControlAyuda txtNroDir, rsSolicitud.Fields("c_NumeroD2"), Frame1.Top + txtNroDir.Top
    Set ControlOrig = txtNroDir
End Sub
Private Sub txtNroDir_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 600, 560
End Sub

Private Sub txtNroEmp_DblClick()
    ControlAyuda txtNroEmp, rsSolicitud.Fields("NumeroD2"), txtNroEmp.Top + Frame1.Top + frmEmpleador.Top
    Set ControlOrig = txtNroEmp
End Sub

Private Sub txtNumDoc_DblClick()
    ControlAyuda txtNumDoc, rsSolicitud.Fields("Numdocd2"), Frame1.Top + txtNumDoc.Top
    Set ControlOrig = txtNumDoc
End Sub

Sub Resaltado(Izq As Integer, Top As Integer, Width As Integer, Height As Integer, X As Integer, Y As Integer)
On Local Error Resume Next
        ImgEdit1.DeleteAnnotationGroup ("Resaltado")
        ImgEdit1.AddAnnotationGroup ("Resaltado")
        ImgEdit1.AnnotationType = 4  '4
    
        ImgEdit1.ScrollPositionX = X
    
        ImgEdit1.ScrollPositionY = Y + txtFactor
        ImgEdit1.Draw Izq, Top, Width, Height
        ImgEdit1.ShowAnnotationGroup ("Resaltado")
        ImgEdit1.AnnotationType = wiNone
        ImgEdit1.Enabled = False
        ImgEdit1.Enabled = True
        ImgEdit1.MousePointer = 0
        ImgEdit1.DrawSelectionRect 1, 1, 1, 1
        ImgEdit1.Refresh
    
End Sub



Function ValDNI(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto de apellido paterno

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtNumDoc) = "" Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ValDNI = False
             txtNumDoc.SetFocus
             Exit Function
    End If

    If Not vbDNI Then
      vbDNI = True
      vbDNI2 = True  '2da digitacion
      vtmpDNI = txtNumDoc.Text
      txtNumDoc.Text = ""
      'txtNumDoc.SetFocus
      'txtCUSPP.SetFocus
      txtCUSPP.SetFocus
      
    End If
 End If
 
 If KeyAscii = vbKeyF9 Then
   If vbDNI2 Then '2da Digitacion
    vbnoCompDNI = True
    'El primer valor se hace igual al 2do
    lblMensaje.Caption = ""
   Else
    ValDNI = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If
 End If
 
' If vbDNI Then
'
'    If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
'        If Not vbnoCompDNI Then  'si el falso
'
'            If Not CompararTexto(vtmpDNI, txtNumDoc.Text, KeyAscii, Len(txtNumDoc) + 1) Then
'              KeyAscii = 0
'              Beep
'              lblMensaje.Caption = " Caracter no Corresponde"
'              txtNumDoc.SetFocus
'              ValDNI = False
'
'            Else
'              lblMensaje.Caption = ""
'            End If
'
'        Else
'           vbnoCompDNI = False
'        End If
'    Else
'       If KeyAscii = 13 And (Len(vtmpDNI) <= Len(txtNumDoc)) Then
'            ValDNI = True
'            txtCUSPP.SetFocus
'       End If
'    End If
' End If
'
End Function
Function ValCuss(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto de apellido paterno

   
If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtCUSPP) = "" Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ValCuss = False
             txtCUSPP.SetFocus
             Exit Function
    End If

    'If Not vbCuss Then
         If Not ValCusp2(txtCUSPP) Then
           ValCuss = False
           lblMensaje.Caption = "CUSPP No cumple estructura establecida ######AAAAA#"
           txtCUSPP.SetFocus
           Exit Function
         Else
           ValCuss = True
           lblMensaje.Caption = ""
           
           If IndDoc = "T" Then
                txtTipFO.SetFocus
           Else
                txtFecNac.SetFocus
           End If
         End If
      
     ' vtmpCuss = txtCUSPP.Text
      'txtCUSPP.Text =
    'End If
 End If
 
 
 If KeyAscii = vbKeyF9 Then
 
  
'   If vbCuss2 Then '2da Digitacion
'    vbnoCuss = True
'    'El primer valor se hace igual al 2do
'    lblMensaje.Caption = ""
'   Else
'    vbCuss = False
'    lblMensaje = "Debe digitar 2 veces el campo"
'    Beep
'   End If
   
'   If Not ValCusp2(txtCUSPP) Then
'     ValCuss = False
'     lblMensaje.Caption = "CUSPP No cumple estructura establecida ######AAAAA#"
'     txtCUSPP.SetFocus
'     'Exit Function
'   End If
   
 End If
 
 
'
' If vbCuss Then
'
'    If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
'        If Not vbnoCuss Then  'si el falso
'
'            If Not CompararTexto(vtmpCuss, txtCUSPP.Text, KeyAscii, Len(txtCUSPP) + 1) Then
'              KeyAscii = 0
'              Beep
'              lblMensaje.Caption = " Caracter no Corresponde"
'              txtCUSPP.SetFocus
'              ValCuss = False
'
'            Else
'              lblMensaje.Caption = ""
'            End If
'
'        Else
'           vbnoCuss = False
'        End If
'    Else
'       If KeyAscii = 13 And (Len(vtmpCuss) <= Len(txtCUSPP.Text)) Then vbCuss = True: lblMensaje = "": txtFTipoO.SetFocus
'    End If
' End If
 
End Function


Function ValCusp2(texto As String) As Boolean
ValCusp2 = True
If Len(Trim(texto)) < 11 Then ValCusp2 = False: Exit Function


For i = 1 To 6
  If Not ((Asc(Mid(Trim(texto), i, 1)) >= 48 And Asc(Mid(Trim(texto), i, 1)) <= 57)) Then
     ValCusp2 = False
     Exit Function
  End If
Next i


If IsNumeric(Left(Trim(texto), 6)) Then
    ValCusp2 = True
 Else
    ValCusp2 = False
    Exit Function
End If

For i = 7 To 11
  If Not ((Asc(Mid(Trim(texto), i, 1)) >= 65 And Asc(Mid(Trim(texto), i, 1)) <= 90) Or Asc(Mid(Trim(texto), i, 1)) = 165 Or Asc(Mid(Trim(texto), i, 1)) = 209) Then
     ValCusp2 = False
     Exit Function
  End If
Next i

If IsNumeric(Mid(Trim(texto), Len(Trim(texto)), 1)) Then
    ValCusp2 = True
 Else
    ValCusp2 = False
    Exit Function
End If

End Function

Private Sub cmdrotDer_Click()
On Local Error Resume Next
ImgEdit1.RotateRight
End Sub

Private Sub cmdrotIzq_Click()
On Local Error Resume Next
ImgEdit1.RotateLeft
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub
Private Sub lblRazonSoc_GotFocus()
ImgEdit1.Zoom = 50

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 129, 360
Else
    Resaltado 0, 0, 1200, 150, 0, 517
End If
End Sub

Private Sub lblRazonSoc_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub txtApeMat_GotFocus()
'If ImgEdit1.Zoom = 50 Then
ImgEdit1.Zoom = 100


If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 938, 380
Else
    Resaltado 0, 0, 1200, 150, 525, 580
End If

'If txtApeMat.BackColor = &HC0E0FF Then
'    ControlAyuda txtApeMat, rsSolicitud.Fields("ApematD2"), txtApeMat.Top + Frame1.Top + frmAfiliado.Top
'    Set ControlOrig = txtApeMat
'End If

End Sub

Private Sub txtApeMat_LostFocus()
'    FraAyuda.Visible = False
End Sub


Private Sub TxtApePat_GotFocus()

'If ImgEdit1.Zoom = 50 Then
ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 380
Else
    Resaltado 0, 0, 1200, 150, 0, 580
End If

'If txtApePat.BackColor = &HC0E0FF Then
'    ControlAyuda txtApePat, rsSolicitud.Fields("ApepatD2"), txtApePat.Top + Frame1.Top + frmAfiliado.Top
'    Set ControlOrig = txtApePat
'End If

End Sub

Private Sub txtApePat_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
Else
KeyCode = 0
End If

End Sub

Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 18 Or KeyAscii = 209) Then
'vbDigApepat = ApePat(KeyAscii)

If KeyAscii = 13 Then txtApeMat.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub txtApePat_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then

' If (GetKeyState(vbKeyTab) < 0) Then
'    If Not ApePat(13) Then
'     Cancel = True
'    End If
'    Exit Sub
'  End If
'
'  If vbDigApepat Then Cancel = False: Exit Sub
'
'  If txtApePat = rsDetalle.Fields("apepat") Then
'     Cancel = False
'  Else
'     vbDigApepat = ApePat(13)
'     If Not vbDigApepat Then Cancel = True
'  End If
'Else
'    If Not vbDigApepat Then Cancel = True
'End If

End Sub

Private Sub txtApeMat_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then

Else
KeyCode = 0
End If

End Sub

Private Sub TxtApeMat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
'vbDigApeMat = ApeMat(KeyAscii)
If KeyAscii = 13 Then txtNombreuno.SetFocus
Else
KeyAscii = 0
End If

End Sub

Private Sub txtApeMat_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then

'  If (GetKeyState(vbKeyTab) < 0) Then
'    If Not ApeMat(13) Then
'     Cancel = True
'    End If
'    Exit Sub
'  End If
'
'  If vbDigApeMat Then Cancel = False: Exit Sub
'
'  If val(txtApeMat) = val(rsDetalle.Fields("apemat")) Then
'     Cancel = False
'  Else
'     vbDigApeMat = ApeMat(13)
'     If Not vbDigApeMat Then Cancel = True
'  End If
'Else
'    If Not vbDigApeMat Then Cancel = True
'End If

End Sub

Private Sub TxtCUSPP_GotFocus()
    'If ImgEdit1.Zoom = 50 Then
    ImgEdit1.Zoom = 100
      
    If IndDoc = "T" Then
        'Resaltado 0, 0, 1200, 150, 1600, 510
        Resaltado 0, 0, 1200, 150, 1809, 480
    Else
        'Resaltado 0, 0, 1200, 150, 1126, 633
        Resaltado 0, 0, 1200, 150, 1713, 633
    End If
      
   
'   If txtCUSPP.BackColor = &HC0E0FF Then
'      ControlAyuda txtCUSPP, rsSolicitud.Fields("CUSPPD2"), txtCUSPP.Top + Frame1.Top + frmAfiliado.Top
'      Set ControlOrig = txtCUSPP
'   End If

   'Resaltado 0, 0, 1200, 150, 1809, 530
   
   
End Sub

Private Sub TxtCUSPP_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18 Or KeyCode = 209 Then
    If KeyCode = vbKeyTab Then
     vbDigCuss = ValCuss(13)
    Exit Sub
    End If
    If KeyCode = vbKeyF9 Or KeyCode = vbKeyF8 Then
     vbDigCuss = ValCuss(KeyCode)
    End If
Else
    KeyCode = 0
End If

End Sub

Private Sub TxtCuspp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
  vbDigCusPP = ValCuss(KeyAscii)
  Else
  KeyAscii = 0
End If


End Sub

Private Sub txtCUSPP_KeyUp(KeyCode As Integer, Shift As Integer)
If txtCUSPP.MaxLength = Len(Trim(txtCUSPP)) Then
      Call ValCuss(13)
  End If

End Sub

Private Sub TxtCUSPP_LostFocus()
FraAyuda.Visible = False
End Sub

Private Sub TxtCUSPP_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
    If (GetKeyState(vbKeyTab) < 0) Then
        If Not ValCuss(13) Then
         Cancel = True
        End If
        Exit Sub
    End If

End Sub

Private Sub txtDireEmp_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100


If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 846
Else
    Resaltado 0, 0, 1200, 150, 57, 1161
End If

'If txtDireEmp.BackColor = &HC0E0FF Then
'    ControlAyuda txtDireEmp, rsSolicitud.Fields("DireccionD2"), txtDireEmp.Top + Frame1.Top + frmEmpleador.Top
'      Set ControlOrig = txtDireEmp
'End If

End Sub

Private Sub txtDireEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii
End Sub

Private Sub txtDireEmp_LostFocus()
'txtDireEmp.BackColor = &H80000005
FraAyuda.Visible = False
End Sub

Private Sub txtDireEmp_Validate(Cancel As Boolean)
'If MsgBox("¿Verificó la dirección contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
'    Cancel = False
'Else
'    Cancel = True
'End If
End Sub

Private Sub txtFecNac_KeyPress(KeyAscii As Integer)
FVFecha txtFecNac, KeyAscii
End Sub

Private Sub txtFecNac_LostFocus()
ValidaFecha txtFecNac
End Sub

Private Sub txtFecsal_KeyPress(KeyAscii As Integer)
    FVFecha txtFecsal, KeyAscii
End Sub

Private Sub txtFecsal_LostFocus()
    ValidaFecha txtFecsal
End Sub


Private Sub txtNombreDos_GotFocus()
'If ImgEdit1.Zoom = 50 Then
ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1841, 380
Else
    Resaltado 0, 0, 1200, 150, 1745, 580
End If
'If txtNombredos.BackColor = &HC0E0FF Then
''    ControlAyuda txtNombredos, rsSolicitud.Fields("Nombre2D2"), txtNombredos.Top + Frame1.Top + frmAfiliado.Top
''    Set ControlOrig = txtNombredos
'End If

End Sub

Private Sub txtNombreDos_LostFocus()
'    FraAyuda.Visible = False
End Sub

Private Sub txtNombreUno_GotFocus()
'If ImgEdit1.Zoom = 50 Then
ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1618, 380
Else
    Resaltado 0, 0, 1200, 150, 1080, 580
End If

'If txtNombreuno.Backcolor = &HC0E0FF Then
'    ControlAyuda txtNombreuno, rsSolicitud.Fields("Nombre1D2"), txtNombreuno.Top + Frame1.Top + frmAfiliado.Top
'    Set ControlOrig = txtNombreuno
'End If

End Sub

Private Sub txtNombreuno_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18 Or KeyCode = 209) Then
Else
    KeyCode = 0
End If
End Sub

Private Sub txtNombreUno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
'vbDigNomb1 = NombreUno(KeyAscii)
If KeyAscii = 13 Then txtNombredos.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNombredos_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18 Or KeyCode = 209) Then
' If KeyCode = vbKeyTab Then
'  vbDigNomb2 = NombreDos(13)
'  Exit Sub
' End If
' If KeyCode = vbKeyF9 Then
'  vbDigNomb2 = NombreDos(KeyCode)
' End If
Else
KeyCode = 0
End If
End Sub

Private Sub txtNombreDos_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
'vbDigNomb2 = NombreDos(KeyAscii)
If KeyAscii = 13 Then txtTipDoc.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNombreUno_LostFocus()
'    FraAyuda.Visible = False
End Sub

Private Sub txtNombreuno_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
'    If (GetKeyState(vbKeyTab) < 0) Then
'        If Not NombreUno(13) Then
'         Cancel = True
'        End If
'        Exit Sub
'    End If
'
'  If vbDigNomb1 Then Cancel = False: Exit Sub
'
'  If txtNombreuno = val(rsDetalle.Fields("nomuno")) Then
'     'If Err Then vbDigNomb1 = ValRemuAsig(13): Exit Sub
'     Cancel = False
'  Else
'     vbDigNomb1 = NombreUno(13)
'     If Not vbDigNomb1 Then Cancel = True
'  End If
'Else
'    If Not vbDigNomb1 Then Cancel = True
'End If
End Sub


Private Sub txtNroEmp_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 892, 846
Else
    Resaltado 0, 0, 1200, 150, 876, 1161
End If
'If txtNroEmp.BackColor = &HC0E0FF Then
'    ControlAyuda txtNroEmp, rsSolicitud.Fields("NumeroD2"), txtNroEmp.Top + Frame1.Top + frmEmpleador.Top
'    Set ControlOrig = txtNroEmp
'End If
End Sub

Private Sub txtNroEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtNroEmp_LostFocus()
FraAyuda.Visible = False
End Sub

Private Sub txtNroEmp_Validate(Cancel As Boolean)
'If MsgBox("¿Verificó el número contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
'    Cancel = False
'Else
'    Cancel = True
'End If
End Sub

'Private Sub txtNroSolicitud_GotFocus()
''On Local Error Resume Next
'' If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
'' Marcado txtNroSolicitud
'' Resaltado 0, 0, 1200, 150, (Int((ImgEdit1.ImageScaleWidth / 3))), 176
'' txtNroSolicitud.Backcolor = &HC0FFFF
'End Sub

'Private Sub txtNroSolicitud_LostFocus()
''txtNroSolicitud.Text = Format(txtNroSolicitud.Text, "000000#")
''txtNroSolicitud.Backcolor = &H80000005
'End Sub


Private Sub txtNumDoc_GotFocus()
Marcado txtNumDoc
'If ImgEdit1.Zoom = 50 Then
ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 950, 510
Else
    Resaltado 0, 0, 1200, 150, 1175, 680
End If

'If txtNumDoc.BackColor = &HC0E0FF Then
'    ControlAyuda txtNumDoc, rsSolicitud.Fields("numDocD2"), txtNumDoc.Top + Frame1.Top + frmAfiliado.Top
'    Set ControlOrig = txtNumDoc
'End If

End Sub

Private Sub txtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)

If ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
' If KeyCode = vbKeyTab Then
'  vbDigDNI = ValDNI(13)
' Exit Sub
' End If
'
' If KeyCode = vbKeyF9 Then
'  'vbDigDNI = ValDNI(KeyCode)
' End If
Else
KeyCode = 0
End If
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 18) Then
    If KeyAscii = 13 Then Me.txtCUSPP.SetFocus
End If
End Sub

Private Sub txtNumDoc_LostFocus()
'FraAyuda.Visible = False
End Sub

Private Sub txtNumDoc_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
'
'  If (GetKeyState(vbKeyTab) < 0) Then
'    If Not ValDNI(13) Then
'     Cancel = True
'    End If
'    Exit Sub
'  End If
'
'  If vbDigDNI Then Cancel = False: Exit Sub
'
'  If val(txtNumDoc) = val(rsDetalle.Fields("numdoc")) Then
'     Cancel = False
'  Else
'     vbDigDNI = ValDNI(13)
'     If Not vbDigDNI Then Cancel = True
'  End If
''Else
''    If Not vbDigApeMat Then Cancel = True
''End If

End Sub


Private Sub txtRuc_DblClick()
ControlAyuda txtRuc, rsSolicitud.Fields("numRUCD2"), Frame1.Top + frmEmpleador.Top + txtRuc.Top
Set ControlOrig = txtRuc
End Sub

Private Sub txtRuc_LostFocus()
FraAyuda.Visible = False
End Sub

Private Sub txtTelefEmp_DblClick()
    ControlAyuda txtTelefEmp, rsSolicitud.Fields("TelefonoD2"), txtTelefEmp.Top + Frame1.Top + frmEmpleador.Top
    Set ControlOrig = txtTelefEmp
End Sub

Private Sub txtTelefEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1819, 936
Else
    Resaltado 0, 0, 1200, 150, 1656, 1245
End If

' If txtTelefEmp.BackColor = &HC0E0FF Then
'    ControlAyuda txtTelefEmp, rsSolicitud.Fields("TelefonoD2"), txtTelefEmp.Top + Frame1.Top + frmEmpleador.Top
'    Set ControlOrig = txtTelefEmp
' End If

End Sub

Private Sub txtTelefEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtTelefEmp_LostFocus()
txtTelefEmp.BackColor = &H80000005
End Sub

Private Sub txtTipDoc_DblClick()
    ControlAyuda txtTipDoc, rsSolicitud.Fields("TipodocD2"), Frame1.Top + txtTipDoc.Top
    Set ControlOrig = txtTipDoc
End Sub

Private Sub txtTipDoc_GotFocus()
Marcado txtTipDoc
'If ImgEdit1.Zoom = 50 Then
ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 510
Else
    Resaltado 0, 0, 1200, 150, 666, 680
End If

End Sub

Private Sub txtTipDoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtTipDoc_LostFocus()
FraAyuda.Visible = False
End Sub

Private Sub txtTipFD_GotFocus()
ImgEdit1.Zoom = 60
'Resaltado 0, 0, 1200, 150, 0, 923
Resaltado 0, 0, 1200, 150, 0, 732
End Sub

Private Sub txtTipFO_DblClick()
    ControlAyuda txtTipFO, rsSolicitud.Fields("TipoFonOriD2"), txtTipFO.Top + Frame1.Top + frmAfiliado.Top
     Set ControlOrig = txtTipFO
End Sub

Private Sub txtTipFD_DblClick()
    ControlAyuda txtTipFD, rsSolicitud.Fields("TipoFonDesD2"), txtTipFD.Top + Frame1.Top + frmAfiliado.Top
     Set ControlOrig = txtTipFD
End Sub

Private Sub txtTipFO_GotFocus()
ImgEdit1.Zoom = 60
'Resaltado 0, 0, 1200, 150, 0, 923
Resaltado 0, 0, 1200, 150, 0, 732
End Sub


Private Sub txtTipFD_KeyPress(KeyAscii As Integer)
Dim vbtemkey As Integer

vbtemkey = KeyAscii

If Not (KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or KeyAscii = vbKey3 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyF9 Or KeyAscii = 13) Then
    KeyAscii = 0
Else
    txtcorreo.SetFocus
End If

'TabSig vbtemkey
End Sub

Private Sub txtTipFD_Validate(Cancel As Boolean)
If txtTipFD = "1" Or txtTipFD = "2" Or txtTipFD = "3" Or txtTipFD = "" Then
    Cancel = False
  Else
    Cancel = True
End If
End Sub


Private Sub txtTipFO_KeyPress(KeyAscii As Integer)
Dim vbtemkey As Integer

vbtemkey = KeyAscii

If Not (KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or KeyAscii = vbKey3 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyF9 Or KeyAscii = 13) Then
    KeyAscii = 0
Else
    txtTipFD.SetFocus
End If

'TabSig vbtemkey
End Sub

Private Sub txtTipFO_Validate(Cancel As Boolean)
If txtTipFO = "1" Or txtTipFO = "2" Or txtTipFO = "3" Or txtTipFO = "" Then
    Cancel = False
  Else
    Cancel = True
End If
End Sub

Private Sub txtUrba_DblClick()
    ControlAyuda txtUrba, rsSolicitud.Fields("c_UrbD2"), Frame1.Top + txtUrba.Top
    Set ControlOrig = txtUrba
End Sub
Private Sub txtUrba_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1600, 560
End Sub

Private Sub txtUrbaEmp_DblClick()
    ControlAyuda txtUrbaEmp, rsSolicitud.Fields("UrbD2"), txtUrbaEmp.Top + Frame1.Top + frmEmpleador.Top
    Set ControlOrig = txtUrbaEmp
End Sub

Private Sub txtUrbaEmp_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100


If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1769, 846
Else
    Resaltado 0, 0, 1200, 150, 1695, 1161
End If

'If txtUrbaEmp.BackColor = &HC0E0FF Then
'    ControlAyuda txtUrbaEmp, rsSolicitud.Fields("UrbD2"), txtUrbaEmp.Top + Frame1.Top + frmEmpleador.Top
'    Set ControlOrig = txtUrbaEmp
'End If
End Sub

Private Sub txtUrbaEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtUrbaEmp_LostFocus()
FraAyuda.Visible = False
End Sub

Private Sub txtUrbaEmp_Validate(Cancel As Boolean)
'If MsgBox("¿Verificó la Urbanización contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
'    Cancel = False
'Else
'    Cancel = True
'End If
End Sub


Function validar() As Boolean
validar = True
'If Trim(txtNroSolicitud) = "" Then
'    MsgBox "Debe Ingresar el Nro de Solicitud", vbCritical, Caption
'    txtNroSolicitud.SetFocus
'    validar = False
'    Exit Function
'End If

'If Trim(txtRuc.Text) = "" Then
'
'    If MsgBox("¿El Afiliado tiene Empleador ?", vbQuestion + vbYesNo, Caption) = vbNo Then
'        vbindEmp2 = False
'
'    Else
'        vbindEmp2 = True
'        'txtRuc.SetFocus
'        'Exit Function
'    End If
'Else
'    vbindEmp2 = True
'
'End If



If Trim(txtRuc) = "" Then
    If MsgBox(" No ingresó el Ruc ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        validar = False
        txtRuc.SetFocus
        Exit Function
    End If
End If


If Trim(txtRuc) <> "" And Trim(lblRazonSoc) = "" Then
     Call MsgBox(" Debe ingresar la Razon Social ", vbQuestion + vbYesNo + vbDefaultButton2, Caption)
     validar = False
     lblRazonSoc.SetFocus
     Exit Function
End If

If Trim(txtDireEmp) = "" Then
    If MsgBox(" No ingresó la dirección del empleador ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        validar = False
        txtDireEmp.SetFocus
        Exit Function
    End If
End If

If Trim(txtNroEmp) = "" Then
    If MsgBox(" No ingresó el nro de la dirección del empleador ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        validar = False
        txtNroEmp.SetFocus
        Exit Function
    End If
End If

If Trim(txtTelefEmp) = "" Then
    If MsgBox(" No ingresó telefono del empleador ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        validar = False
        txtTelefEmp.SetFocus
        Exit Function
    End If
End If

'If Trim(vbCodgeoEmp) = "" Then
'    If MsgBox(" No ingresó el Departamento-Provincia-Distrito del empleador ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
'        validar = False
'        ucUbigeoEmp.SetFocus
'        Exit Function
'    End If
'End If



If Trim(txtTipDoc.Text) = "" Then
        If MsgBox("No ingresó el tipo de documento ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtTipDoc.SetFocus
            Exit Function
        End If
End If
    
    
If Trim(txtNumDoc.Text) = "" Then
        If MsgBox("No ingresó el número de documento ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtNumDoc.SetFocus
            Exit Function
        End If
End If
  
  
If IndDoc = "T" Then
    
    If Trim(txtCUSPP.Text) = "" Then
            If MsgBox("No ingresó el Cuspp ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
                validar = False
                txtCUSPP.SetFocus
                Exit Function
            End If
    End If

End If
      
    Select Case txtTipDoc
    Case "DNI"
        If Len(Trim(txtNumDoc)) <> 8 Then
            MsgBox "El DNI debe tener 8 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
        
    Case "CE"
        If Len(Trim(txtNumDoc)) < 5 Then
            MsgBox "El CE debe tener de 5 a 8 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
    
    Case "CIA"
        If Len(Trim(txtNumDoc)) < 2 And Len(Trim(txtNumDoc)) > 5 Then
            MsgBox "El CE debe tener de 2 a 5 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
    
    Case "CFA"
        If Len(Trim(txtNumDoc)) < 6 And Len(Trim(txtNumDoc)) > 8 Then
            MsgBox "El CE debe tener de 6 a 8 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
    
    Case Else
    
    End Select
    
    
    




End Function

Private Sub ucUbigeoAfi_DblClick()
    ControlAyuda ucUbigeoAfi, rsSolicitud.Fields("c_UbigeoD2"), Frame1.Top + ucUbigeoAfi.Top
    Set ControlOrig = ucUbigeoAfi
End Sub

Private Sub ucUbigeoAfi_GotFocus()
ImgEdit1.Zoom = 85
Resaltado 0, 0, 1200, 150, 100, 560
End Sub
Private Sub ucUbigeoAfi_DobleClick()
    ControlAyuda ucUbigeoAfi, rsSolicitud.Fields("c_DesgeoD2") & " / " & rsSolicitud.Fields("c_UbigeoD2"), ucUbigeoAfi.Top + frmAfiliado.Top
    Set ControlOrig = ucUbigeoAfi
End Sub

Private Sub ucUbigeoAfi_Pulsacion(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If LLenarDatosUbigeoAfi Then TabSig KeyAscii
Else
   If KeyAscii = vbKeyF9 Then TabSig 13
End If

End Sub
Private Sub ucUbigeoAfi_KeyUP(KeyAscii As Integer)
If KeyAscii = vbKeyF9 Then TabSig 13
End Sub

Private Sub ucUbigeoAfi_LostFocus()
    LLenarDatosUbigeoAfi
End Sub


Private Sub ucUbigeoEmp_GotFocus()
 ImgEdit1.Zoom = 50
 
 If IndDoc = "T" Then
    Resaltado 0, 15, 1200, 50, 38, 460
 Else
    Resaltado 0, 15, 1200, 50, 38, 609
 End If
End Sub

Private Sub ucUbigeoEmp_DobleClick()
    ControlAyuda ucUbigeoEmp, rsSolicitud.Fields("eDesgeoD2") & " / " & rsSolicitud.Fields("UbigeoD2"), txtUrbaEmp.Top + Frame1.Top + frmEmpleador.Top
    Set ControlOrig = ucUbigeoEmp
End Sub

Private Sub ucUbigeoEmp_KeyUP(KeyAscii As Integer)
If KeyAscii = vbKeyF9 Then TabSig 13
End Sub

Private Sub ucUbigeoEmp_LostFocus()

'If vbCodDistEmp = "" Or vbCodProvEmp = "" Or vbDesDistEmp = "" Then
    LLenarDatosUbigeoEmp
'End If

End Sub

Private Sub ucUbigeoEmp_Pulsacion(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If LLenarDatosUbigeoEmp Then TabSig KeyAscii
Else
   If KeyAscii = vbKeyF9 Then TabSig 13
End If

End Sub

Function LLenarDatosUbigeoEmp() As Boolean
'Ubigeo de Afiliados
On Error Resume Next
If rsUbigeoEmp.State = 1 Then rsUbigeoEmp.Close

rsUbigeoEmp.Open "select * from ubigeo where codgeo = '" & Trim(Mid(Trim(ucUbigeoEmp.cTexto), InStr(ucUbigeoEmp.cTexto, "/") + 1, 10)) & "'", Con.Cnx, adOpenStatic
    If (rsUbigeoEmp.EOF And rsUbigeoEmp.BOF) Then
        MsgBox "El Dato ingresado no existe", 64, "Mensaje"
        LLenarDatosUbigeoEmp = False
        vbCodDistEmp = ""
        vbCodProvEmp = ""
        vbCodDptoEmp = ""
        vbDesDistEmp = ""
        vbDesProvEmp = ""
        vbDesDptoEmp = ""
        vbCodgeoEmp = ""
    Else
        LLenarDatosUbigeoEmp = True
        vbCodDistEmp = rsUbigeoEmp.Fields("CodDist")
        vbCodProvEmp = rsUbigeoEmp.Fields("Codprov")
        vbCodDptoEmp = rsUbigeoEmp.Fields("CodDpto")
        vbDesDistEmp = rsUbigeoEmp.Fields("Distrito")
        vbDesProvEmp = rsUbigeoEmp.Fields("Provincia")
        vbDesDptoEmp = rsUbigeoEmp.Fields("Departamento")
        vbCodgeoEmp = rsUbigeoEmp.Fields("codgeo")
         'TabSig 13
    End If


'Set rsAfiliado = Nothing
End Function

Function LLenarDatosUbigeoAfi() As Boolean
'Ubigeo de Afiliados
On Error Resume Next
If rsUbigeoAfi.State = 1 Then rsUbigeoAfi.Close

rsUbigeoAfi.Open "select * from ubigeo where codgeo = '" & Trim(Mid(Trim(ucUbigeoAfi.cTexto), InStr(ucUbigeoAfi.cTexto, "/") + 1, 10)) & "'", Con.Cnx, adOpenStatic
    If (rsUbigeoAfi.EOF And rsUbigeoAfi.BOF) Then
        MsgBox "El Dato ingresado no existe", 64, "Mensaje"
        LLenarDatosUbigeoAfi = False
        vbCodDistAfi = ""
        vbCodProvAfi = ""
        vbCodDptoAfi = ""
        vbDesDistAfi = ""
        vbDesProvAfi = ""
        vbDesDptoAfi = ""
        vbCodgeoAfi = ""
    Else
        LLenarDatosUbigeoAfi = True
        vbCodDistAfi = rsUbigeoAfi.Fields("CodDist")
        vbCodProvAfi = rsUbigeoAfi.Fields("Codprov")
        vbCodDptoAfi = rsUbigeoAfi.Fields("CodDpto")
        vbDesDistAfi = rsUbigeoAfi.Fields("Distrito")
        vbDesProvAfi = rsUbigeoAfi.Fields("Provincia")
        vbDesDptoAfi = rsUbigeoAfi.Fields("Departamento")
        vbCodgeoAfi = rsUbigeoAfi.Fields("codgeo")
         'TabSig 13
    End If


'Set rsAfiliado = Nothing
End Function

Private Sub txtRuc_GotFocus()
'Set controlf = ActiveControl
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 777
Else
    Resaltado 0, 0, 1200, 150, 0, 1089
End If

'If txtRuc.BackColor = &HC0E0FF Then
'    ControlAyuda txtRuc, rsSolicitud.Fields("NumRucD2"), txtRuc.Top + Frame1.Top + frmEmpleador.Top
'    Set ControlOrig = txtRuc
'End If

End Sub

Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If KeyCode = vbKeyTab Then
 Call ValRuc(13)
End If

If KeyCode = vbKeyF9 Then
 Call ValRuc(KeyCode)
End If
End Sub

Private Sub txtRUC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Call ValRuc(KeyAscii)
FEntDec txtRuc, KeyAscii, 11, 0
End Sub

Private Sub txtRUC_KeyUp(KeyCode As Integer, Shift As Integer)
  If txtRuc.MaxLength = Len(Trim(txtRuc)) Then
      Call ValRuc(13)
  End If
End Sub

Private Sub txtRuc_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ValRuc(13) Then
      Cancel = True
    End If
    Exit Sub
  End If
End Sub
Function ValRuc(KeyAscii) As Boolean

Dim rsDatosRuc As Recordset
Set rsDatosRuc = New ADODB.Recordset
Dim sqltexto As String
If KeyAscii = 13 Then

If Trim(txtRuc) = "" Then
         lblMensaje.Caption = "Debe Ingresar RUC"
         Beep
         ValRuc = False
         txtRuc.SetFocus
         Exit Function
End If
    'Busqueda en la Tabla Magne Si es Pago Magnetico
        'Recupera DAtos de Razon Social
        
            If Modulo11(txtRuc.Text) Then
                'vRuc = fgDatosql("Select NUMRUC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                sqltexto = "select razsoc, direcc, a.codgeo, telefono, numero, urba, ANEXO_EL, CODUSUC=isnull(CODUSUC,''), Feccar , Desgeo  from Truc a left join ubigeo b on" & _
                " a.codgeo = b.codgeo where NUMRUC = '" & txtRuc & "'"
                
                rsDatosRuc.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

                If rsDatosRuc.RecordCount > 0 Then
                'Consultar no existe razon en truc
                   'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValRuc = True
                   vbindRuc = True
                   
                   If CDate(rsDatosRuc.Fields("Feccar")) = CDate(gDate) Then vbindRuc = False
                   'lblMensaje.Caption = ""
                   lblRazonSoc.SetFocus
                   lblRazonSoc.Text = IIf(IsNull(rsDatosRuc.Fields("razsoc")), "", rsDatosRuc.Fields("razsoc"))
                   txtDireEmp = IIf(IsNull(rsDatosRuc.Fields("direcc")), "", rsDatosRuc.Fields("direcc"))
                   txtNroEmp = IIf(IsNull(rsDatosRuc.Fields("numero")), "", rsDatosRuc.Fields("numero"))
                   txtUrbaEmp = IIf(IsNull(rsDatosRuc.Fields("urba")), "", rsDatosRuc.Fields("urba"))
                   txtTeleOfi = IIf(IsNull(rsDatosRuc.Fields("telefono")), "", rsDatosRuc.Fields("telefono"))
                   txtAnxOfi = IIf(IsNull(rsDatosRuc.Fields("ANEXO_EL")), "", rsDatosRuc.Fields("ANEXO_EL"))
                   ucUbigeoEmp.cTexto = IIf(IsNull(rsDatosRuc.Fields("Desgeo")), "", rsDatosRuc.Fields("Desgeo")) & " / " & IIf(IsNull(rsDatosRuc.Fields("codgeo")), "", rsDatosRuc.Fields("codgeo"))
                   
                Else
                   lblMensaje.Caption = "RUC no se encuentra en la tabla TRUC"
                   ValRuc = True
                   vbindRuc = False
                   lblRazonSoc.SetFocus
                   lblRazonSoc.Text = ""
                   txtDireEmp = ""
                   txtNroEmp = ""
                   txtUrbaEmp = ""
                   txtTeleOfi = ""
                   txtAnxOfi = ""
                   ucUbigeoEmp.cTexto = ""
                   ' Beep
                End If
            Else
                lblMensaje.Caption = "Ruc No Válido"
                Beep
            End If
        
End If
 If KeyAscii = vbKeyF9 Then
   'Fuerza el pase al sgte campo
   ValRuc = True
   lblMensaje.Caption = ""
   vbindRuc = False
   txtDireEmp.SetFocus ' TabSig 13
 End If

End Function

Private Sub ucUbigeoEmp_Validate(Cancel As Boolean)
'If MsgBox("¿Verificó Distrito, Provincia y Departamento contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
'    Cancel = False
'Else
'    Cancel = True
'End If
End Sub
Sub GrabarVerificacion()

If Not validar Then Exit Sub

    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = " sp_Verificar2Dig4 ?,?,?,?,?, ?,?,?,?,?, " & _
                                              " ?,?,?,?,?, ?,?,?,?,? ,?,?,?,?,?,?,?,?,?,?,?,?"
                                                                                                                      
        .Parameters("@CodAfp") = "01"
        .Parameters("@FecSal") = Trim(rsSolicitud.Fields("fecsal"))
        .Parameters("@Lote") = Trim(rsSolicitud.Fields("lote"))
        .Parameters("@NumSol") = Trim(rsSolicitud.Fields("numsol"))
        .Parameters("@Inddoc") = IndDoc
        
        .Parameters("@ApePat") = Trim(txtApePat)
        .Parameters("@ApeMat") = Trim(txtApeMat)
        .Parameters("@Nombre1") = Trim(txtNombreuno)
        .Parameters("@Nombre2") = Trim(txtNombredos)
        .Parameters("@TipoDoc") = Trim(txtTipDoc)
        
        .Parameters("@NumDoc") = Trim(txtNumDoc)
        .Parameters("@CUSPP") = Trim(txtCUSPP)
        .Parameters("@Ruc") = Trim(txtRuc)
        .Parameters("@DireEmp") = Trim(txtDireEmp)
        .Parameters("@NroEmp") = Trim(txtNroEmp)
        
        .Parameters("@UrbaEmp") = Trim(txtUrbaEmp)
        .Parameters("@UbigeoEmp") = Trim(Left(Mid(ucUbigeoEmp.cTexto, InStr(ucUbigeoEmp.cTexto, "/") + 1), 7))
        .Parameters("@TelefEmp") = Trim(txtTelefEmp)
        .Parameters("@Secuencial") = IIf(IsNull(rsSolicitud.Fields("Secuencial")), 0, rsSolicitud.Fields("Secuencial"))
        .Parameters("@CodmodD2") = gNomUsu
        
        .Parameters("@FecNac") = IIf((Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//"), Null, txtFecNac)
        .Parameters("@TipoFonOri") = Trim(txtTipFO.Text)
        .Parameters("@TipoFonDes") = Trim(txtTipFD.Text)
        .Parameters("@e_mail") = Trim(txtcorreo.Text)
        .Parameters("@PrefDirEmp") = Right(Trim(cboPrefDirEmp.Text), 2)
        .Parameters("@PrefLocEmp") = Right(Trim(cboPrefLocEmp.Text), 2)
        
        .Parameters("@Direccion") = Trim(txtAv)
        .Parameters("@Numero") = Trim(txtNroDir)
        .Parameters("@Urb") = Trim(txtUrba)
        .Parameters("@PrefDirAfi") = Right(Trim(cboPrefDirAfi.Text), 2)
        .Parameters("@PrefLocAfi") = Right(Trim(cboPrefLocAfi.Text), 2)
        .Parameters("@Ubigeo") = Trim(Left(Mid(ucUbigeoAfi.cTexto, InStr(ucUbigeoAfi.cTexto, "/") + 1), 7))
        
                     
        If Con.GetFEjecutarSQL(comando) Then
            If LCase(Con.vbdev) = "exito" Then
                'Call NuevoRegistro
                
                Call VerificarRevision(rsSolicitud.Fields("lote"))
                
                If Not rsSolicitud.EOF Then rsSolicitud.MoveNext
                
                If rsSolicitud.EOF Then
                    LimpiarPantalla
                Else
                    Call RefrescarSol(True)
                End If
                                                                
            End If
            
        End If
                                                                  

    End With


End Sub



Sub MovImg(ByRef KeyCode As Integer)
On Error Resume Next

If KeyCode = vbKeyF12 Then
    If Not (RsImag.EOF And RsImag.BOF) Then
      RsImag.MoveNext
      If RsImag.EOF Then RsImag.MoveLast

      ImgEdit2.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      txtnumpag = RsImag!Posicion
      ImgEdit2.Display
            
      ImgEdit2.FitTo 1
      ImgEdit2.Refresh
      
      KeyCode = 0

    End If
End If

If KeyCode = vbKeyF11 Then
    If Not (RsImag.EOF And RsImag.BOF) Then
      RsImag.MovePrevious
      If RsImag.BOF Then RsImag.MoveFirst

      ImgEdit2.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      txtnumpag = RsImag!Posicion
      ImgEdit2.Display
      
      ImgEdit2.FitTo 1
      ImgEdit2.Refresh
      
      KeyCode = 0
      
    End If
End If


End Sub


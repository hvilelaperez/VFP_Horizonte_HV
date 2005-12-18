VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form frmHojaVida 
   Caption         =   "Hoja de Vida"
   ClientHeight    =   9585
   ClientLeft      =   180
   ClientTop       =   720
   ClientWidth     =   14430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   14430
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Afiliado"
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   14415
      Begin VB.ComboBox cboDominios2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cboDominios 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboNacionalidad 
         Height          =   345
         Left            =   12000
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtSexo 
         Height          =   330
         Left            =   13560
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtLugNac 
         Height          =   330
         Left            =   10680
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtFecNaci 
         Height          =   330
         Left            =   10680
         MaxLength       =   10
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtNomCorres 
         Height          =   330
         Left            =   8400
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3735
      End
      Begin VB.ComboBox cboCargos 
         Height          =   345
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtCelular 
         Height          =   330
         Left            =   8400
         MaxLength       =   9
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtTelfCasa 
         Height          =   330
         Left            =   8400
         MaxLength       =   9
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtEmailTrab 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtEmailCasa 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   3
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtNombresComp 
         Height          =   330
         Left            =   2040
         MaxLength       =   150
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   10095
      End
      Begin VB.Label Label13 
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   92
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   91
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblGeoInei 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11640
         TabIndex        =   78
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
         Height          =   255
         Left            =   12240
         TabIndex        =   77
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo"
         Height          =   225
         Left            =   12960
         TabIndex        =   76
         Top             =   1395
         Width           =   495
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar Nac"
         Height          =   225
         Left            =   9720
         TabIndex        =   75
         Top             =   1400
         Width           =   975
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Nac."
         Height          =   225
         Left            =   9720
         TabIndex        =   74
         Top             =   1040
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre en Comunicaciones"
         Height          =   255
         Left            =   5880
         TabIndex        =   45
         Top             =   645
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
         Height          =   255
         Left            =   7680
         TabIndex        =   40
         Top             =   1400
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Telf. Casa"
         Height          =   225
         Left            =   7440
         TabIndex        =   39
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "e-mail (Trabajo)"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "e-mail (Casa)"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres Completos"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblApeMat 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         Height          =   255
         Left            =   720
         TabIndex        =   36
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   12600
      TabIndex        =   80
      Top             =   1680
      Width           =   1815
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   3
         Left            =   240
         TabIndex        =   81
         Top             =   3840
         Width           =   1335
         _Version        =   131074
         _ExtentX        =   2355
         _ExtentY        =   1931
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   387940096
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   45
         lReserved2      =   45
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   2
         Left            =   240
         TabIndex        =   82
         Top             =   2640
         Width           =   1335
         _Version        =   131074
         _ExtentX        =   2355
         _ExtentY        =   1931
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   387940096
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   45
         lReserved2      =   45
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   1
         Left            =   240
         TabIndex        =   83
         Top             =   1440
         Width           =   1335
         _Version        =   131074
         _ExtentX        =   2355
         _ExtentY        =   1931
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   387940096
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   45
         lReserved2      =   45
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   0
         Left            =   240
         TabIndex        =   84
         Top             =   240
         Width           =   1335
         _Version        =   131074
         _ExtentX        =   2355
         _ExtentY        =   1931
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   387940096
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   45
         lReserved2      =   45
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
   End
   Begin VB.CommandButton cmdGrabarHoja 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1215
   End
   Begin Traspasos27.UserControl1 UcUbigeoCorres 
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   2750
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
   End
   Begin Traspasos27.UserControl1 ucUbigeoEmp 
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   5400
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      Height          =   375
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1215
   End
   Begin VB.ComboBox cboEstadoCivil 
      Height          =   345
      ItemData        =   "frmHojaVida.frx":0000
      Left            =   1800
      List            =   "frmHojaVida.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6300
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Height          =   940
      Left            =   120
      TabIndex        =   42
      Top             =   6600
      Width           =   10815
      Begin VB.CommandButton cmdVerPromotor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Promotor"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   520
         Width           =   1575
      End
      Begin VB.CheckBox chkCodErr 
         Caption         =   "Codigo errado"
         Height          =   255
         Left            =   9000
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtFecSus 
         Enabled         =   0   'False
         Height          =   330
         Left            =   960
         MaxLength       =   20
         TabIndex        =   30
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox txtRenta 
         Height          =   330
         Left            =   960
         MaxLength       =   20
         TabIndex        =   32
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtCodInterno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   31
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblPromotor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3960
         TabIndex        =   73
         Top             =   150
         Width           =   4935
      End
      Begin VB.Label Label17 
         Caption         =   "Fec. Visita"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Renta"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Cod Int."
         Height          =   255
         Left            =   2160
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   1935
      Left            =   120
      TabIndex        =   55
      Top             =   7560
      Width           =   14175
      _Version        =   131074
      _ExtentX        =   25003
      _ExtentY        =   3413
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AnnotationBackColor=   4210752
      AnnotationFillColor=   65535
      AutoRefresh     =   -1  'True
      UndoBufferSize  =   387854592
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
   Begin VB.Frame Frame2 
      Caption         =   "Dirección Correspondencia"
      ForeColor       =   &H00C00000&
      Height          =   1365
      Left            =   120
      TabIndex        =   41
      Top             =   1800
      Width           =   12375
      Begin VB.ComboBox cboPrefLocCor 
         Height          =   345
         ItemData        =   "frmHojaVida.frx":0166
         Left            =   8040
         List            =   "frmHojaVida.frx":0168
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cboPrefDirCor 
         Height          =   345
         ItemData        =   "frmHojaVida.frx":016A
         Left            =   120
         List            =   "frmHojaVida.frx":016C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtUrbaCorres 
         Height          =   330
         Left            =   9600
         MaxLength       =   100
         TabIndex        =   17
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtNroCorres 
         Height          =   330
         Left            =   5760
         MaxLength       =   40
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtDireCorres 
         Height          =   330
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   14
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Urba"
         Height          =   255
         Left            =   8160
         TabIndex        =   86
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Via"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Avenida (Av.) / Calle (CL)/ Pasaje (Pj) / Jirón (Jr)"
         Height          =   255
         Left            =   1680
         TabIndex        =   64
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento / Provincia / Distrito"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Urbanización"
         Height          =   255
         Left            =   9600
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Número (N) / Departamento (Dpto) / Interior(Int) / Manzana (Mz)  / Lote (Lte.)  "
         Height          =   255
         Left            =   5640
         TabIndex        =   61
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   3015
      Left            =   120
      TabIndex        =   46
      Top             =   3240
      Width           =   12375
      Begin VB.ComboBox cboPrefLocEmp 
         Height          =   345
         ItemData        =   "frmHojaVida.frx":016E
         Left            =   240
         List            =   "frmHojaVida.frx":0170
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cboPrefDirEmp 
         Height          =   345
         ItemData        =   "frmHojaVida.frx":0172
         Left            =   240
         List            =   "frmHojaVida.frx":0174
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Direcciones"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   0
         TabIndex        =   66
         Top             =   -2280
         Width           =   14175
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cance&lar"
            Height          =   375
            Left            =   12840
            TabIndex        =   69
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdAcepDire 
            Caption         =   "Ace&ptar"
            Height          =   375
            Left            =   12840
            TabIndex        =   68
            Top             =   360
            Width           =   1215
         End
         Begin VB.ListBox lstDirecc 
            Height          =   1185
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   12495
         End
      End
      Begin VB.CommandButton cmdSucursales 
         Caption         =   "Sucu&rsales"
         Height          =   375
         Left            =   9960
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtReferencia 
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   28
         Top             =   2595
         Width           =   6735
      End
      Begin VB.TextBox txtAnxOfi 
         Height          =   315
         Left            =   10320
         MaxLength       =   5
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtTeleOfi 
         Height          =   315
         Left            =   8520
         MaxLength       =   7
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtDireEmp 
         Height          =   330
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   21
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox txtNroEmp 
         Height          =   330
         Left            =   7560
         MaxLength       =   40
         TabIndex        =   22
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtRuc 
         Height          =   315
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtUrbaEmp 
         Height          =   315
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   24
         Top             =   1560
         Width           =   5775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Urba"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   1250
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Via"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo"
         Height          =   255
         Left            =   10320
         TabIndex        =   59
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono Oficina"
         Height          =   255
         Left            =   8520
         TabIndex        =   58
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblY 
         Caption         =   "lbly"
         Height          =   375
         Left            =   12840
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblX 
         Caption         =   "lblx"
         Height          =   375
         Left            =   12480
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Número (N) / Departamento (Dpto) / Interior(Int) / Manzana (Mz)  / Lote (Lte.)  "
         Height          =   255
         Left            =   7560
         TabIndex        =   53
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblRazonSoc 
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3720
         TabIndex        =   52
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label21 
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
         Left            =   12720
         TabIndex        =   51
         Top             =   -1080
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Ruc"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Avenida (Av.) / Calle (CL)/ Pasaje (Pj) / Jirón (Jr)"
         Height          =   255
         Left            =   1680
         TabIndex        =   49
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Urbanización"
         Height          =   255
         Left            =   1680
         TabIndex        =   48
         Top             =   1250
         Width           =   2655
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento / Provincia / Distrito"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1920
         Width           =   3615
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Civil"
      Height          =   255
      Left            =   480
      TabIndex        =   89
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   9840
      TabIndex        =   71
      Top             =   8520
      Width           =   3015
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Civil"
      Height          =   255
      Left            =   480
      TabIndex        =   65
      Top             =   9720
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmHojaVida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NombresComp As String
Public vbDireCorrs As String
Public Numsol As String
Public vbRuc As String
Public vbAve As String
Public vbNro As String
Public vbUrba As String
Public vbUbigeo As String
Public vbRazSoc As String
Public vbPrefDirEmp As String
Public vbPrefLocEmp As String

Public vbAveCR As String
Public vbNroCR As String
Public vbUrbaCR As String
Public vbUbigeoCR As String
Public vbPrefDirCR As String
Public vbPrefLocCR As String

Public vbTelefAfi As String
Public vbTelefEmp As String

Public TipDoc As String '
Public Fecsal As String '
Public Secu As Integer
Public rutaImagen As String
Public Imagen As String
Private SecHV As Integer
Public Fecnac As String
Public UbigeoNacUbi As String
Public Sexo As String
Public CodAdmDoc As String 'Codigo Adm proveniente del documento anterior

Private rsFamiliar As ADODB.Recordset
Private Secfam As Integer 'Ultimo  nro de familiares solo referencial
Private vbCodgeoEmp As String
Private vbCodgeoAfi As String
Private rsUbigeoAfi As New Recordset
Private rsUbigeoEmp As New Recordset
Private vbUGeoInei As String
Dim rsSuc As Recordset
Private rsImagAlt As New Recordset

Private vbRenta As Boolean
Private vbRenta2 As Boolean
Private vbnoCompRenta As Boolean
Private vbDigRenta As Boolean
Private vtmpRenta As String
Private vbindPro As Boolean

Private vbPromI As Boolean
Private vbPromI2 As Boolean
Private vbnoCompPromI As Boolean
Private vbDigPromI As Boolean
Private vtmpPromI As String



Sub CrearTmpFamiliares()
'With rsFamiliar
'    .Fields.Append "Nro", adBigInt
'    .Fields.Append "Nombres", adChar, 120
'    .Fields.Append "Apepat", adChar, 30
'    .Fields.Append "Apemat", adChar, 30
'    .Fields.Append "Nomb1", adChar, 20
'    .Fields.Append "Nomb2", adChar, 20
'    .Fields.Append "CodPar", adChar, 1
'    .Fields.Append "Parentesco", adChar, 20
'    .Fields.Append "Sexo", adChar, 1
'    .Fields.Append "FecNac", adChar, 10
'    .Open
'    Set dtgFamilia.DataSource = rsFamiliar
'
'End With
End Sub


Private Sub cboCargos_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1020, 300
cboCargos.BackColor = &HC0FFFF

End Sub

Private Sub cboCargos_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboCargos_LostFocus()
cboCargos.BackColor = &H80000005
End Sub

Private Sub cboDominios_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then cboDominios.ListIndex = -1
End Sub



Private Sub cboDominios2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then cboDominios2.ListIndex = -1
End Sub

Private Sub cboEstadoCivil_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 38, 793
 cboEstadoCivil.BackColor = &HC0FFFF
End Sub

Private Sub cboEstadoCivil_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboEstadoCivil_LostFocus()
cboEstadoCivil.BackColor = &H80000005
End Sub

Private Sub cboNacionalidad_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1070, 527
cboNacionalidad.BackColor = &HC0FFFF
End Sub

Private Sub cboNacionalidad_LostFocus()
cboNacionalidad.BackColor = &H80000005
End Sub





Private Sub cboPrefDirCor_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 309
End Sub

Private Sub cboPrefDirCor_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefDirCor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefDirCor.ListIndex = -1
End Sub

Private Sub cboPrefDirCor_LostFocus()
cboPrefDirCor.BackColor = vbWhite
End Sub

Private Sub cboPrefDirEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 222, 961
End Sub

Private Sub cboPrefDirEmp_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefDirEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefDirEmp.ListIndex = -1
End Sub

Private Sub cboPrefDirEmp_LostFocus()
cboPrefDirEmp.BackColor = vbWhite
End Sub

Private Sub cboPrefLocCor_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 317
cboPrefLocCor.BackColor = &HC0FFFF
End Sub

Private Sub cboPrefLocCor_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefLocCor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefLocCor.ListIndex = -1
End Sub

Private Sub cboPrefLocCor_LostFocus()
cboPrefLocCor.BackColor = vbWhite
End Sub

Private Sub cboPrefLocEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1021, 978
End Sub

Private Sub cboPrefLocEmp_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefLocEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefLocEmp.ListIndex = -1
End Sub

Private Sub cboPrefLocEmp_LostFocus()
cboPrefLocEmp.BackColor = vbWhite
End Sub

Private Sub cmdAcepDire_Click()
Frame5.Visible = False
ucUbigeoEmp.Visible = True
If lstDirecc.Text <> "" Then
    CopiarDireccion
End If
txtDireEmp.SetFocus
End Sub

'Private Sub CmdAgregar_Click()
'AgregarFamiliar
'txtApePat.SetFocus
'End Sub

Private Sub cmdCancelar_Click()
Frame5.Visible = False
ucUbigeoEmp.Visible = True
txtDireEmp.SetFocus
End Sub

'Private Sub CmdEliminar_Click()
'On Error Resume Next
'rsFamiliar.Delete adAffectCurrent
'If val(lblNunFam) - 1 < 0 Then lblNunFam.Caption = "0"
'End Sub

Private Sub cmdGrabarHoja_Click()
GrabarHoja
End Sub

'Private Sub cmdModificar_Click()
'ModificarFamiliar
'End Sub

'Private Sub cmdNuevo_Click()
'NuevoFam
'End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSucursales_Click()
Frame5.Visible = True
ucUbigeoEmp.Visible = False
lstDirecc.SetFocus
End Sub

Private Sub cmdVerPromotor_Click()
FrmConsPromo.Show 1
End Sub

'Private Sub cmdVer_Click()
'If cmdVer.Caption = "Ver parte Superior" Then
'    cmdVer.Caption = "Ver parte Inferior"
'Else
'    cmdVer.Caption = "Ver parte Superior"
'End If
'txtApePat.SetFocus
'End Sub

'Private Sub dtgFamilia_Click()
'On Error Resume Next
'Call RefrescarFamiliar
'End Sub

Private Sub Form_Load()
'
'ImgEdit1.Image = "c:\SCN.tif"
'ImgEdit1.Display
CodAdmDoc = fgDatosql(" Select CodProAd from solicitud a join Tpromotor b " & _
                      " on a.codpromo = b.codpromo where numsol = '" & Numsol & "'and fecsal = '" & Fecsal & "' ")

Set rsFamiliar = New ADODB.Recordset
Call InicializaVar
Call CrearTmpFamiliares

'ccombosql cboParen, "Select [CodParen] + '-' + [DesParen] from PARENTESCO "
ccombosql cboDominios, "Select DesDom from Dominio"
ccombosql cboDominios2, "Select DesDom from Dominio"

CargarComboSql Me.cboCargos, "Select Descargo, CodCargo from Cargo "

Set rsSuc = New ADODB.Recordset
Set UcUbigeoCorres.cConexion = Con.Cnx
    UcUbigeoCorres.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    UcUbigeoCorres.Campocondicion = "desgeo"
    UcUbigeoCorres.Busqinicial = False
    UcUbigeoCorres.Enabled = True
    UcUbigeoCorres.nroCaracteres = 3

Set ucUbigeoEmp.cConexion = Con.Cnx
    ucUbigeoEmp.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    ucUbigeoEmp.Campocondicion = "desgeo"
    ucUbigeoEmp.Busqinicial = False
    ucUbigeoEmp.Enabled = True
    ucUbigeoEmp.nroCaracteres = 3

Call CargarImagen
Secfam = 0
txtNombresComp.Text = Trim(NombresComp)

CargarComboSql cboNacionalidad, "select descri_corta, codigo from nacionalidad "
CargarComboSql cboPrefDirCor, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocCor, "select AbreLoc, CodLoc  from prefloc"
CargarComboSql cboPrefDirEmp, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocEmp, "select AbreLoc, CodLoc  from prefloc"

pgBuscarEnCombo cboNacionalidad, "000114", 6

'Public vbDireCorrs As String
'-------Direccion Correspondencia--
txtDireCorres.Text = vbAveCR
txtNroCorres.Text = vbNroCR
txtUrbaCorres.Text = vbUrbaCR
UcUbigeoCorres.cTexto = vbUbigeoCR
vbCodgeoAfi = Trim(Mid(Trim(vbUbigeoCR), InStr(vbUbigeoCR, "/") + 1, 10))

pgBuscarEnCombo cboPrefDirCor, vbPrefDirCR, 2
pgBuscarEnCombo cboPrefLocCor, vbPrefLocCR, 2
pgBuscarEnCombo cboPrefDirEmp, vbPrefDirEmp, 2
pgBuscarEnCombo cboPrefLocEmp, vbPrefLocEmp, 2
'----------------------------------

If Len(vbTelefAfi) > 7 Then
    txtCelular = vbTelefAfi
Else
    txtTelfCasa.Text = vbTelefAfi
End If


txtTeleOfi.Text = vbTelefEmp
txtRuc.Text = vbRuc
txtDireEmp.Text = vbAve
txtNroEmp.Text = vbNro
txtUrbaEmp.Text = vbUrba
ucUbigeoEmp.cTexto = vbUbigeo
lblRazonSoc.Caption = " " & vbRazSoc
vbCodgeoEmp = Trim(Mid(Trim(vbUbigeo), InStr(vbUbigeo, "/") + 1, 10))

txtFecNaci = Fecnac
txtLugNac.Text = UbigeoNacUbi
txtSexo.Text = Sexo

Frame5.Visible = False
Call ObtenerDireccSuc(vbRuc)




End Sub

'Sub AgregarFamiliar()
'On Error GoTo errores
'
'If Not ValidarFam Then Exit Sub
'With rsFamiliar
'    .AddNew
'    .Fields("Nro") = Secfam + 1
'    .Fields("Nombres") = Trim(txtApePat.Text) & " " & Trim(txtApeMat.Text) & " " & Trim(txtNombreuno.Text) & " " & Trim(txtNombredos.Text)
'    .Fields("Apepat").Value = Trim(txtApePat.Text)
'    .Fields("ApeMat").Value = Trim(txtApeMat.Text)
'    .Fields("Nomb1").Value = Trim(txtNombreuno.Text)
'
'    .Fields("Nomb2").Value = Trim(txtNombredos.Text)
'    .Fields("CodPar").Value = Trim(Left(Trim(cboParen), 1))
'    .Fields("Parentesco").Value = Trim(Right(Trim(cboParen), 20))
'    .Fields("Sexo").Value = Trim(Left(cboSexo.Text, 1))
'    .Fields("Fecnac").Value = IIf(Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//", "", txtFecNac)
'    Secfam = Secfam + 1
'    lblNunFam = val(lblNunFam) + 1
'    Call NuevoFam
'
'End With
'Exit Sub
'errores:
'MsgBox Err.Description, vbCritical, Caption
'End Sub

Private Sub ImgEdit1_Scroll()
Me.lblx = Me.ImgEdit1.ScrollPositionX
Me.lbly = Me.ImgEdit1.ScrollPositionY
End Sub


Private Sub ImgEdit3_Click(Index As Integer)
On Error Resume Next
 SiempreVisible frmVisorImag, True
 
 frmVisorImag.ImgEdit1.Image = ImgEdit3(Index).Image
 frmVisorImag.ImgEdit1.Display
 frmVisorImag.ImgEdit1.FitTo 1
 frmVisorImag.Show
End Sub

Private Sub lstDirecc_GotFocus()
txtAnxOfi.BackColor = &HC0FFFF
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 90, 646
End Sub

Private Sub txtAnxOfi_GotFocus()
txtAnxOfi.BackColor = &HC0FFFF
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 392, 893
End Sub

Private Sub txtAnxOfi_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
FEntDec txtAnxOfi, KeyAscii, 10, 0
End Sub

Private Sub txtAnxOfi_LostFocus()
txtAnxOfi.BackColor = &H80000005
End Sub

'Private Sub txtApeMat_GotFocus()
'ImgEdit1.Zoom = 50
'If cmdVer.Caption = "Ver parte Inferior" Then
'    Resaltado 0, 0, 1200, 150, 65, 880
'Else
'    Resaltado 0, 0, 1200, 150, 65, 990
'End If
'txtApeMat.BackColor = &HC0FFFF
'End Sub
'
'Private Sub TxtApeMat_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig (KeyAscii)
'End Sub
'
'Private Sub txtApeMat_LostFocus()
'txtApeMat.BackColor = &H80000005
'End Sub
'
'Private Sub TxtApePat_GotFocus()
'ImgEdit1.Zoom = 50
'If cmdVer.Caption = "Ver parte Inferior" Then
'    Resaltado 0, 0, 1200, 150, 65, 880
'Else
'    Resaltado 0, 0, 1200, 150, 65, 990
'End If
'
'txtApePat.BackColor = &HC0FFFF
'End Sub
'
'Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig (KeyAscii)
'End Sub
'
'Private Sub TxtApePat_LostFocus()
'txtApePat.BackColor = &H80000005
'End Sub

Private Sub txtCelular_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1001, 379
txtCelular.BackColor = &HC0FFFF
End Sub

Private Sub txtCelular_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
FEntDec txtCelular, KeyAscii, 10, 0

End Sub

Private Sub txtCelular_LostFocus()
txtCelular.BackColor = &H80000005
End Sub

Private Sub txtCodInterno_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 620
txtCodInterno.BackColor = &HC0FFFF
End Sub

Private Sub txtCodInterno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then
 Call PromDig(13)
End If

If KeyCode = vbKeyF9 Then
 Call PromDig(KeyCode)
End If
End Sub

Private Sub txtCodInterno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtCodInterno, KeyAscii, vbCaracteresNumero)
'TabSig KeyAscii
If KeyAscii = 13 Then Call PromDig(KeyAscii)
End Sub

Private Sub txtCodInterno_LostFocus()
 txtCodInterno.BackColor = &H80000005
End Sub

Private Sub txtCodInterno_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
'  If (GetKeyState(vbKeyTab) < 0) Then
'    If Not ValProm(13) Then
'      Cancel = True
'    End If
'    Exit Sub
'  End If
  
  
  On Error Resume Next


  If (GetKeyState(vbKeyTab) < 0) Then
    If Not PromDig(13) Then
       Cancel = True
    End If
    Exit Sub
  End If

  If vbDigPromI Then Cancel = False: Exit Sub
  
  If val(txtCodInterno) = val(rsDetalle.Fields("remase")) Then
     Cancel = False
  Else
     vbDigPromI = PromDig(13)
     If Not vbDigPromI Then Cancel = True
  End If

  
End Sub

Private Sub txtDireCorres_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 309
txtDireCorres.BackColor = &HC0FFFF
End Sub

Private Sub txtDireCorres_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
KeyAscii = 0
End If
TabSig KeyAscii
End Sub

Private Sub txtDireCorres_LostFocus()
txtDireCorres.BackColor = &H80000005
End Sub

Private Sub txtDireEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 222, 961
txtDireEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtDireEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtDireEmp_LostFocus()
txtDireEmp.BackColor = &H80000005
End Sub

Private Sub txtEmailCasa_GotFocus()
txtEmailCasa.BackColor = &HC0FFFF
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 269, 460

End Sub

Private Sub txtEmailCasa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtEmailCasa, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ" & "_-." & vbCaracteresNumero)

TabSig KeyAscii
End Sub


Private Sub txtEmailCasa_LostFocus()
txtEmailCasa.BackColor = &H80000005
End Sub

Private Sub txtEmailTrab_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 269, 516
txtEmailTrab.BackColor = &HC0FFFF
End Sub

Private Sub txtEmailTrab_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
''If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
KeyAscii = ValidarDigitacion(txtEmailTrab, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ" & "_-." & vbCaracteresNumero)
TabSig KeyAscii
End Sub

Private Sub txtEmailTrab_LostFocus()
txtEmailTrab.BackColor = &H80000005
End Sub

'Private Sub txtFecNAc_GotFocus()
'ImgEdit1.Zoom = 50
'If cmdVer.Caption = "Ver parte Inferior" Then
'    Resaltado 0, 0, 1200, 150, 65, 880
'Else
'    Resaltado 0, 0, 1200, 150, 65, 990
'End If
'txtFecNac.BackColor = &HC0FFFF
'End Sub

'Private Sub txtFecNac_KeyPress(KeyAscii As Integer)
'TabSig (KeyAscii)
'FVFecha txtFecNac, KeyAscii
'
''If KeyAscii = 13 Then
''    AgregarFamiliar
''End If
'
'End Sub

'Private Sub txtFecNac_LostFocus()
'ValidaFecha txtFecNac
'txtFecNac.BackColor = &H80000005
'End Sub
'Sub NuevoFam()
'txtApePat.Enabled = True
'txtApeMat.Enabled = True
'txtNombreuno.Enabled = True
'txtNombredos.Enabled = True
'cboParen.Enabled = True
'cboSexo.Enabled = True
'txtFecNac.Enabled = True
'
'txtApePat.Text = ""
'txtApeMat.Text = ""
'txtNombreuno.Text = ""
'txtNombredos.Text = ""
'cboParen.ListIndex = -1
'cboSexo.ListIndex = -1
'txtFecNac.Text = ""
'End Sub
'Function ValidarFam() As Boolean
'ValidarFam = True
'
'If Trim(txtApePat) = "" Then
'    ValidarFam = False
'    MsgBox "Debe ingresar el apellido paterno", 64, Caption
'    txtApePat.SetFocus
'    Exit Function
'End If
'
'If Trim(txtApeMat) = "" Then
'    ValidarFam = False
'    MsgBox "Debe ingresar el apellido materno", 64, Caption
'    txtApeMat.SetFocus
'    Exit Function
'End If
'
''If Trim(txtNombreuno) = "" Then
''    ValidarFam = False
''    MsgBox "Debe ingresar el primer nombre", 64, Caption
''    txtNombreuno.SetFocus
''    Exit Function
''End If
'
''If cboParen.ListIndex = -1 Then
''    ValidarFam = False
''    MsgBox "Debe seleccionar el parentesco", 64, Caption
''    cboParen.SetFocus
''    Exit Function
''End If
'
''If cboSexo.ListIndex = -1 Then
''    ValidarFam = False
''    MsgBox "Debe seleccionar el sexo", 64, Caption
''    cboSexo.SetFocus
''    Exit Function
''End If
'
'
'If fgDatosql("select numsol from familiar where numsol = '" & Numsol & "' and fecsal = '" & Fecsal & "' and indDoc = '" & TipDoc & "'") <> "" Then
'    Dim rstmpFam As New ADODB.Recordset
'    Set rstmpFam = New ADODB.Recordset
'    Dim strFam As String
'
'    rstmpFam.Open "SET CONCAT_NULL_YIELDS_NULL OFF " & _
'        " Select convert(varchar, Secuencia) + '- ' +  ApePat +' '+ ApeMat+' '+ NomUno+ ' '+ NomDos " & _
'        " from familiar where numsol = '" & Numsol & "' and fecsal = '" & Fecsal & "' and indDoc = '" & TipDoc & "'", Con.Cnx, adOpenStatic, adLockReadOnly
'    If rstmpFam.RecordCount > 0 Then
'        Do While Not rstmpFam.EOF
'            strFam = strFam & rstmpFam.Fields(0) & Chr(10)
'            rstmpFam.MoveNext
'        Loop
'    End If
'
'    strFam = " Existen Familiares registrados:  " & Chr(10) & strFam & " Verifique los Familiares ¿ Desea continuar ? "
'
'    If MsgBox(strFam, vbYesNo + vbQuestion, Caption) = vbNo Then
'        ValidarFam = False
'
'        txtApePat.Text = ""
'        txtApeMat.Text = ""
'        txtNombreuno.Text = ""
'        txtNombredos.Text = ""
'        cboParen.ListIndex = -1
'        cboSexo.ListIndex = -1
'
'        Exit Function
'    End If
'
'    Set rstmpFam = Nothing
'
'End If
'
'End Function

Private Sub txtFecNaci_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 244, 529
txtFecNaci.BackColor = &HC0FFFF
End Sub

Private Sub txtFecNaci_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
FVFecha txtFecNaci, KeyAscii
End Sub

Private Sub txtFecNaci_LostFocus()
ValidaFecha txtFecNaci
txtFecNaci.BackColor = &H80000005
End Sub

Private Sub txtFecNaci_Validate(Cancel As Boolean)
If IsDate(txtFecNaci) Then
If txtFecNaci <> "" Then
    If CDate(txtFecNaci) < "01/01/1920" Then
        If MsgBox("La fecha de nacimiento es menor a 1920 ¿Desea Continuar?", 64, Caption) = vbYes Then
            Cancel = True
        End If
    End If
End If
End If
End Sub

Private Sub txtFecSus_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 2135
txtFecSus.BackColor = &HC0FFFF
End Sub

Private Sub txtFecSus_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
FVFecha txtFecSus, KeyAscii

End Sub

Private Sub txtFecSus_LostFocus()
ValidaFecha txtFecSus
txtFecSus.BackColor = &H80000005
End Sub

Private Sub txtFecSus_Validate(Cancel As Boolean)
If IsDate(txtFecSus) Then
    If DateDiff("d", txtFecSus, gDate) > 7 Then
        If MsgBox("La Fecha de Declaración es anterior a una semana, Desea Continuar? ", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        Else
            Cancel = False
        End If
    End If
End If
End Sub

Private Sub txtLugNac_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 244, 585
txtLugNac.BackColor = &HC0FFFF
End Sub

Private Sub txtLugNac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
 Call ValLugarNac(KeyCode)
End If
End Sub

Private Sub txtLugNac_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig KeyAscii
If KeyAscii = 13 Then Call ValLugarNac(KeyAscii)
End Sub

Private Sub txtLugNac_LostFocus()
txtLugNac.BackColor = &H80000005
End Sub

Private Sub txtLugNac_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ValLugarNac(13) Then
      Cancel = True
    End If
    Exit Sub
  End If
End Sub


'Private Sub txtNombreDos_GotFocus()
'ImgEdit1.Zoom = 50
'If cmdVer.Caption = "Ver parte Inferior" Then
'    Resaltado 0, 0, 1200, 150, 65, 880
'Else
'    Resaltado 0, 0, 1200, 150, 65, 990
'End If
'txtNombredos.BackColor = &HC0FFFF
'End Sub
'
'Private Sub txtNombreDos_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig (KeyAscii)
'End Sub
'
'Private Sub txtNombreDos_LostFocus()
'txtNombredos.BackColor = &H80000005
'End Sub

Private Sub txtNombresComp_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
    KeyAscii = 0
End If

TabSig KeyAscii
End Sub

'Private Sub txtNombreUno_GotFocus()
'ImgEdit1.Zoom = 50
'If cmdVer.Caption = "Ver parte Inferior" Then
'    Resaltado 0, 0, 1200, 150, 65, 880
'Else
'    Resaltado 0, 0, 1200, 150, 65, 990
'End If
'txtNombreuno.BackColor = &HC0FFFF
'End Sub
'
'Private Sub txtNombreUno_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig (KeyAscii)
'End Sub
'
'Private Sub txtNombreUno_LostFocus()
'txtNombreuno.BackColor = &H80000005
'End Sub

Private Sub txtNomCorres_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 1000
Resaltado 0, 0, 1200, 150, 325, 319
txtNomCorres.BackColor = &HC0FFFF
End Sub

Private Sub txtNomCorres_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
KeyAscii = 0
End If
TabSig KeyAscii
End Sub

Private Sub txtNomCorres_LostFocus()
txtNomCorres.BackColor = &H80000005
End Sub


Private Sub txtNroCorres_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 309
txtNroCorres.BackColor = &HC0FFFF
End Sub

Private Sub txtNroCorres_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtNroCorres_LostFocus()
txtNroCorres.BackColor = &H80000005
End Sub

Private Sub txtNroEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 859, 978
txtNroEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtNroEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtNroEmp_LostFocus()
txtNroEmp.BackColor = &H80000005
End Sub

Private Sub txtReferencia_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 561
txtReferencia.BackColor = &HC0FFFF
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtRenta_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 680
txtCodInterno.BackColor = &HC0FFFF
End Sub

Private Sub txtRenta_KeyDown(KeyCode As Integer, Shift As Integer)
SuprimirPunto txtRenta, KeyCode, 10, 2
End Sub

Private Sub txtRenta_KeyPress(KeyAscii As Integer)
vbDigRenta = ValRenta(KeyAscii)
FEntDec txtRenta, KeyAscii, 10, 2
End Sub

Private Sub txtRenta_Validate(Cancel As Boolean)
On Error Resume Next


  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ValRenta(13) Then
       Cancel = True
    End If
    Exit Sub
  End If

  If vbDigRenta Then Cancel = False: Exit Sub
  
  If val(txtRenta) = val(rsDetalle.Fields("remase")) Then
     Cancel = False
  Else
     vbDigRenta = ValRenta(13)
     If Not vbDigRenta Then Cancel = True
  End If

End Sub

Private Sub txtRuc_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 800, 916
txtRuc.BackColor = &HC0FFFF
End Sub

Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txtRuc_LostFocus()
txtRuc.BackColor = &H80000005
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

Private Sub txtSexo_GotFocus()
Marcado txtSexo
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1070, 559
txtSexo.BackColor = &HC0FFFF
End Sub

Private Sub txtSexo_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
KeyAscii = Asc(UCase(Chr(KeyAscii)))

'If Not vbdigTipFondo Then
''si es el segundo empleador ir defrente a grabar
'    'If KeyAscii = 13 Then
'     If Frame4.Enabled = False Then
'        cmdGrabar.SetFocus
'     Else
'        RevisarTipFondo
'     End If
'    'End If
'Else
'    TabSig KeyAscii
'End If


End Sub

Private Sub txtSexo_KeyUp(KeyCode As Integer, Shift As Integer)
If txtSexo.MaxLength = Len(Trim(txtSexo)) Then TabSig 13
End Sub

Private Sub txtSexo_LostFocus()
txtSexo.BackColor = &H80000005
End Sub

Private Sub txtSexo_Validate(Cancel As Boolean)
If txtSexo = "F" Or txtSexo = "M" Then
 Cancel = False
 Else
 Cancel = True
End If

End Sub


Private Sub txtTeleOfi_GotFocus()
txtTeleOfi.BackColor = &HC0FFFF
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 246, 893
End Sub

Private Sub txtTeleOfi_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
FEntDec txtTeleOfi, KeyAscii, 10, 0
End Sub

Private Sub txtTeleOfi_LostFocus()
txtTeleOfi.BackColor = &H80000005
End Sub

Private Sub txtTelfCasa_GotFocus()
txtTelfCasa.BackColor = &HC0FFFF
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1001, 330

End Sub

Private Sub txtTelfCasa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
FEntDec txtTelfCasa, KeyAscii, 10, 0

End Sub


Sub GrabarHoja()
On Error GoTo ERRORES
Dim cmdHoja As ADODB.Command
Set cmdHoja = New ADODB.Command


If Not ValidarHoja Then Exit Sub
With cmdHoja
    .ActiveConnection = Con.Cnx
    .CommandText = "pGrabarHojaVida4 ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,? "
    .Parameters("@Numsol") = Numsol
    .Parameters("@IndDoc") = TipDoc
    .Parameters("@Fecsal") = Fecsal
    .Parameters("@Secu") = Secu
    .Parameters("@NomCorres") = Trim(txtNomCorres)
    .Parameters("@CodCargo") = Trim(Right(cboCargos, 5))
    .Parameters("@EmailDoc") = "" 'IIf(txtEmailCasa = "", "", Trim(txtEmailCasa) & "@" & cboDominios.Text)
    .Parameters("@EmailTra") = "" 'IIf(txtEmailTrab = "", "", Trim(txtEmailTrab) & "@" & cboDominios.Text)
    .Parameters("@Celular") = Trim(txtCelular)
    .Parameters("@DesDomCor") = Trim(txtDireCorres) 'Avenida
    
    .Parameters("@NroDomCor") = Trim(txtNroCorres) 'Avenida
    .Parameters("@UrbaDomCor") = Trim(txtUrbaCorres)
    .Parameters("@CodgeoCor") = Trim(vbCodgeoAfi)
    .Parameters("@DireEmp") = Trim(txtDireEmp)
    .Parameters("@NroEmp") = Trim(txtNroEmp)
    .Parameters("@UrbaEmp") = Trim(txtUrbaEmp)
    .Parameters("@Codgeo") = vbCodgeoEmp
    .Parameters("@CodEstCiv") = Right(Trim(cboEstadoCivil.Text), 1)
    If Trim(txtFecSus) <> "" Then
    .Parameters("@FecSus") = CDate(txtFecSus)
    Else
    .Parameters("@FecSus") = Null
    End If
    .Parameters("@CodInPro") = Trim(txtCodInterno.Text)
    
    .Parameters("@Renta") = val(txtRenta)
    .Parameters("@nomImag") = Imagen
    .Parameters("@IDimagen") = vbGAFP & Numsol
    .Parameters("@TeleOfi") = txtTeleOfi
    .Parameters("@AnexOfi") = txtAnxOfi
    .Parameters("@Refdirec") = txtReferencia
    If lstDirecc.ListIndex = -1 Then
        .Parameters("@CodSuc") = "000"
    Else
        .Parameters("@CodSuc") = lstDirecc.ItemData(lstDirecc.ListIndex)
    End If
    .Parameters("@FecNaci") = IIf(Trim(txtFecNaci) = "" Or Trim(txtFecNaci) = "//", Null, Trim(txtFecNaci))
    .Parameters("@UbigeoNacRen") = IIf(Trim(Right(cboNacionalidad, 3)) = 114, Trim(txtLugNac), "330" & Trim(Right(cboNacionalidad, 3)))
    .Parameters("@UbigeoNacINE") = IIf(Trim(Right(cboNacionalidad, 3)) = 114, Trim(vbUGeoInei), "330" & Trim(Right(cboNacionalidad, 3)))
    
    .Parameters("@Nacion") = Trim(Right(cboNacionalidad, 3)) 'Por ver si son los ultimos 3 digitos
    .Parameters("@Sexo") = Trim(txtSexo.Text)
    .Parameters("@IndDirCorr") = IndDirecCorr
    .Parameters("@IndErrProadm") = IIf(Trim(CodAdmDoc) <> Trim(txtCodInterno), 1, 0)
    .Parameters("@PrefDirCor") = Right(Trim(cboPrefDirCor.Text), 2)
    .Parameters("@PrefLocCor") = Right(Trim(cboPrefLocCor.Text), 2)
    .Parameters("@PrefDirEmp") = Right(Trim(cboPrefDirEmp.Text), 2)
    .Parameters("@PrefLocEmp") = Right(Trim(cboPrefLocEmp.Text), 2)
    .Parameters("@TelfCasaHV") = Trim(txtTelfCasa)
    
    '@FecNaci ,@UbigeoNacRen , @UbigeoNacINE , @Nacion , @Sexo

    
    If Con.GetFEjecutarSQL(cmdHoja) Then
        If Con.vbdev = "Exito" Then
            SecHV = fgDatosql("Select max(SecHV) from Hojavida where numsol = '" & Numsol & "' and IndDoc = '" & TipDoc & "' and Fecsal = '" & Fecsal & "' and Secu = '" & Secu & "' ")
                                    
                ActIndRAM
                If TipDoc = "T" Then
                    frmTraspasos.NuevoRegistro
                Else
                    frmAfiliacion.NuevoRegistro
                End If
                
                Unload Me
        End If
            
    End If
    
    
End With


Exit Sub
ERRORES:
MsgBox Err.Description, vbCritical, Caption
End Sub

Private Sub txtTelfCasa_LostFocus()
txtTelfCasa.BackColor = &H80000005
End Sub

Private Sub txtUrbaCorres_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 317
txtUrbaCorres.BackColor = &HC0FFFF
End Sub

Private Sub txtUrbaCorres_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtUrbaCorres_LostFocus()
txtUrbaCorres.BackColor = &H80000005
End Sub

Private Sub txtUrbaEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1021, 978
txtUrbaEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtUrbaEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii

End Sub

Private Sub txtUrbaEmp_LostFocus()
txtUrbaEmp.BackColor = &H80000005
End Sub

Private Sub UcUbigeoCorres_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 38, 317
'txtEmailCasa.BackColor = &HC0FFFF
End Sub

Private Sub UcUbigeoCorres_LostFocus()
LLenarDatosUbigeoAfil
End Sub

Private Sub UcUbigeoCorres_Pulsacion(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtRuc, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ- " & vbCaracteresNumero)

If KeyAscii = 13 Then
   If LLenarDatosUbigeoAfil Then TabSig KeyAscii
Else
   If KeyAscii = vbKeyF9 Then TabSig 13
End If
End Sub

Private Sub ucUbigeoEmp_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 38, 520
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
KeyAscii = ValidarDigitacion(txtRuc, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ- " & vbCaracteresNumero)
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
        vbCodgeoEmp = ""
    Else
        LLenarDatosUbigeoEmp = True
        vbCodgeoEmp = rsUbigeoEmp.Fields("codgeo")
         'TabSig 13
    End If


'Set rsAfiliado = Nothing
End Function
Function LLenarDatosUbigeoAfil() As Boolean
'Ubigeo de Afiliados
On Error Resume Next
If rsUbigeoAfi.State = 1 Then rsUbigeoAfi.Close

rsUbigeoAfi.Open "select * from ubigeo where codgeo = '" & Trim(Mid(Trim(UcUbigeoCorres.cTexto), InStr(UcUbigeoCorres.cTexto, "/") + 1, 10)) & "'", Con.Cnx, adOpenStatic
    If (rsUbigeoAfi.EOF And rsUbigeoAfi.BOF) Then
        MsgBox "El Dato ingresado no existe", 64, "Mensaje"
        LLenarDatosUbigeoAfil = False
        vbCodgeoAfi = ""
    Else
        LLenarDatosUbigeoAfil = True
        vbCodgeoAfi = rsUbigeoAfi.Fields("codgeo")
         'TabSig 13
    End If


'Set rsAfiliado = Nothing
End Function

Function GrabarFamiliares() As Boolean
On Error GoTo ERRORES
Dim cmdFamilia As Command
Set cmdFamilia = New ADODB.Command

With cmdFamilia
    If rsFamiliar.RecordCount > 0 Then
        rsFamiliar.MoveFirst
       Do While Not rsFamiliar.EOF
            .ActiveConnection = Con.Cnx
            .CommandText = "exec pGrabarFamiliar ?,?,?,?,?, ?,?,?,?,?, ?,? "
            
            .Parameters("@Numsol") = Numsol
            .Parameters("@indDoc") = TipDoc
            .Parameters("@Fecsal") = Fecsal
            .Parameters("@Secu") = Secu
            .Parameters("@SecHV") = SecHV
            .Parameters("@ApePat") = Trim(rsFamiliar.Fields("Apepat"))
            .Parameters("@ApeMat") = Trim(rsFamiliar.Fields("Apemat"))
            .Parameters("@NomUno") = Trim(rsFamiliar.Fields("Nomb1"))
            .Parameters("@NomDos") = Trim(rsFamiliar.Fields("Nomb2"))
            .Parameters("@CodPare") = rsFamiliar.Fields("Codpar")
            .Parameters("@Sexo") = rsFamiliar.Fields("Sexo")
            .Parameters("@Fecnac") = IIf(Trim(rsFamiliar.Fields("Fecnac")) = "" Or Trim(rsFamiliar.Fields("Fecnac")) = "//", Null, rsFamiliar.Fields("Fecnac"))
            
            If Con.GetFEjecutarSQL(cmdFamilia) Then
                If Con.vbdev <> "Exito" Then
                    GrabarFamiliares = False
                    Exit Function
                End If
                
            End If
            
            rsFamiliar.MoveNext
            
            
       Loop
        
    End If
End With

GrabarFamiliares = True
Exit Function
ERRORES:
GrabarFamiliares = False

End Function


Sub CargarImagen()
      
      ImgEdit1.Image = rutaImagen
      ImgEdit1.Display
      Resaltado 0, 0, 1200, 150, (Int((ImgEdit1.ImageScaleWidth / 3))), 100
      
      ImagenesAlternas Imagen


End Sub

Sub Resaltado(Izq As Integer, Top As Integer, Width As Integer, Height As Integer, X As Integer, Y As Integer)
On Local Error Resume Next
        ImgEdit1.DeleteAnnotationGroup ("Resaltado")
        ImgEdit1.AddAnnotationGroup ("Resaltado")
        ImgEdit1.AnnotationType = 4  '4
        'ImgEdit1.Draw 700, 100, 1200, 150
        ImgEdit1.ScrollPositionX = X
        'ImgEdit1.ScrollPositionY = 100
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
      
Function ValProm(KeyAscii) As Boolean
Dim vbDesPromo As String

If KeyAscii = 13 Then

    If Trim(txtCodInterno) = "" Then
             lblMensaje.Caption = "Debe Ingresar Código de Promotor Interno"
             Beep
             ValProm = False
             txtCodInterno.SetFocus
             Exit Function
    End If
    'Busqueda en la Tabla Magne Si es Pago Magnetico
        'Recupera DAtos de Razon Social
        
      
                vbDesPromo = fgDatosql("select isnull(desPromo,'Sin Descripción') from tpromotor where CodProAd  = '" & Trim(txtCodInterno.Text) & "'")
                If vbDesPromo <> "" Then
                'Consultar no existe razon en truc
                'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValProm = True
                   vbindPro = True
                   lblMensaje.Caption = ""
                   lblPromotor = " " & vbDesPromo
                   'txtRenta.SetFocus
                Else
                   lblMensaje.Caption = "Código de Promotor Interno no se encuentra en la Tabla TPromotor"
                   ValProm = False
                   vbindPro = False
                   'txtCodInterno.SetFocus
                   Beep
                End If
                              
End If
 If KeyAscii = vbKeyF9 Then
   'Fuerza el pase al sgte campo
   ValProm = True
   lblMensaje.Caption = ""
   vbindPro = False
   'txtRenta.SetFocus
 End If

End Function
     
Function ValidarHoja() As Boolean
ValidarHoja = True

If cboCargos.ListIndex = -1 Then
    If MsgBox("No seleccionó el cargo del afiliado ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        cboCargos.SetFocus
        ValidarHoja = False
        Exit Function
    End If
End If

If Trim(txtNomCorres.Text) = "" Then
    If MsgBox("No ingresó el nombre de comunicaciones ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        txtNomCorres.SetFocus
        ValidarHoja = False
        Exit Function
    End If
End If


If Trim(txtDireCorres.Text) = "" Then
    If MsgBox("No ingresó la dirección del afiliado ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        txtDireCorres.SetFocus
        ValidarHoja = False
         Exit Function
    End If
End If
'
'
'If Trim(txtUrbaCorres.Text) = "" Then
'    If MsgBox("No ingresó la urbanización del afiliado ¿Desea continuar?", vbQuestion + vbDefaultButton2, Caption) = vbNo Then
'        txtUrbaCorres.SetFocus
'        ValidarHoja = False
'    End If
'End If

If Trim(vbCodgeoAfi) = "" Then
    If MsgBox("No ingresó ubigeo del afiliado ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        UcUbigeoCorres.SetFocus
        ValidarHoja = False
        Exit Function
    End If
End If

'If Trim(txtFecSus.Text) = "" Or Trim(txtFecSus.Text) = "//" Then
'    If MsgBox("No ingresó la fecha de visita ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
'        txtFecSus.SetFocus
'        ValidarHoja = False
'        Exit Function
'    End If
'End If

If Trim(txtFecSus.Text) <> "" Then
    If CDate(txtFecSus.Text) > gDate Then
            MsgBox "La fecha de visita no puede ser mayor a la fecha actual", vbCritical, Caption
            txtFecSus.SetFocus
            ValidarHoja = False
            Exit Function
        
    End If
End If

If Trim(txtCodInterno.Text) = "" Then
    If MsgBox("No ingresó el código del promotor ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        txtCodInterno.SetFocus
        ValidarHoja = False
        Exit Function
    End If
End If

If Trim(txtCodInterno.Text) <> Trim(CodAdmDoc) Then
    If MsgBox("El código interno del promotor no corresponde con el codigo sbs digitado en la solicitud  ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        txtCodInterno.SetFocus
        ValidarHoja = False
        Exit Function
    End If
End If


If val(txtRenta.Text) = 0 Then
    If MsgBox("La renta ingresada es cero ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        txtRenta.SetFocus
        ValidarHoja = False
        Exit Function
    End If
End If

    If cboEstadoCivil.ListIndex = -1 Then
        MsgBox "Debe ingresar el Estado Civil, Debe obtenerlo del DNI", 64, Caption
        cboEstadoCivil.SetFocus
        ValidarHoja = False
        Exit Function
    End If
    
    
    If Trim(txtEmailCasa.Text & cboDominios.Text) <> "" Then
        Call Validar_Email(txtEmailCasa.Text & cboDominios.Text)
    End If

    If Trim(txtEmailTrab.Text & cboDominios2.Text) <> "" Then
        Call Validar_Email(txtEmailTrab.Text & cboDominios2.Text)
    End If


End Function

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
                sqltexto = "select razsoc, direcc, a.codgeo, telefono, numero, urba, ANEXO_EL, Desgeo  from Truc a left join ubigeo b on" & _
                " a.codgeo = b.codgeo where NUMRUC = '" & txtRuc & "'"
                
                rsDatosRuc.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

                If rsDatosRuc.RecordCount > 0 Then
                'Consultar no existe razon en truc
                   'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValRuc = True
                   vbindRuc = True
                   'lblMensaje.Caption = ""
                   txtDireEmp.SetFocus
'                   lblRazonSoc.Caption = IIf(IsNull(rsDatosRuc.Fields("razsoc")), "", rsDatosRuc.Fields("razsoc"))
'                   txtDireEmp = IIf(IsNull(rsDatosRuc.Fields("direcc")), "", rsDatosRuc.Fields("direcc"))
'                   txtNroEmp = IIf(IsNull(rsDatosRuc.Fields("numero")), "", rsDatosRuc.Fields("numero"))
'                   txtUrbaEmp = IIf(IsNull(rsDatosRuc.Fields("urba")), "", rsDatosRuc.Fields("urba"))
'                   txtTeleOfi = IIf(IsNull(rsDatosRuc.Fields("telefono")), "", rsDatosRuc.Fields("telefono"))
'                   txtAnxOfi = IIf(IsNull(rsDatosRuc.Fields("ANEXO_EL")), "", rsDatosRuc.Fields("ANEXO_EL"))
'                   ucUbigeoEmp.cTexto = IIf(IsNull(rsDatosRuc.Fields("Desgeo")), "", rsDatosRuc.Fields("Desgeo")) & " / " & IIf(IsNull(rsDatosRuc.Fields("codgeo")), "", rsDatosRuc.Fields("codgeo"))
                   
                Else
                   lblMensaje.Caption = "RUC no se encuentra en la tabla TRUC"
                   ValRuc = True
                   vbindRuc = False
'                   txtDireEmp.SetFocus
'                   lblRazonSoc.Caption = ""
'                   txtDireEmp = ""
'                   txtNroEmp = ""
'                   txtUrbaEmp = ""
'                   txtTeleOfi = ""
'                   txtAnxOfi = ""
                 '  ucUbigeoEmp.cTexto = ""
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

Sub ObtenerDireccSuc(Ruc As String)
'Obtener Direcciones de la tabla sucursales

Dim sqltexto As String
Dim Index As Integer
sqltexto = "Select a.*, desgeo from sucursal a left join  ubigeo b on a.ubigeo = codgeo where ruc = '" & Ruc & "'"


With rsSuc
    If rsSuc.State = 1 Then .Close
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
    If .RecordCount > 0 Then
        lstDirecc.Clear
        Index = 0
        Do While Not .EOF
             texto = IIf(IsNull(.Fields("DIRECCION")), "", .Fields("DIRECCION")) & " " & _
                     IIf(IsNull(.Fields("NUMERO")), "", .Fields("NUMERO")) & " " & _
                     IIf(IsNull(.Fields("URBANIZACION")), "", .Fields("URBANIZACION")) & " " & _
                     IIf(IsNull(.Fields("desgeo")), "", .Fields("desgeo"))
                     

             lstDirecc.AddItem texto, Index
             lstDirecc.ItemData(Index) = .Fields("N_SUCURSAL").Value
             Index = Index + 1
             
             
            .MoveNext
        Loop
    End If
End With

End Sub
Sub CopiarDireccion()

With rsSuc
    txtDireEmp = IIf(IsNull(rsDatosRuc.Fields("direcc")), "", rsDatosRuc.Fields("direcc"))
    txtNroEmp = IIf(IsNull(rsDatosRuc.Fields("numero")), "", rsDatosRuc.Fields("numero"))
    txtUrbaEmp = IIf(IsNull(rsDatosRuc.Fields("URBANIZACION")), "", rsDatosRuc.Fields("URBANIZACION"))
    txtTeleOfi = IIf(IsNull(rsDatosRuc.Fields("TELEFONO")), "", rsDatosRuc.Fields("TELEFONO"))
    txtAnxOfi = IIf(IsNull(rsDatosRuc.Fields("ANX_TELEF")), "", rsDatosRuc.Fields("ANX_TELEF"))
    ucUbigeoEmp.cTexto = IIf(IsNull(rsDatosRuc.Fields("desgeo")), "", rsDatosRuc.Fields("desgeo")) & " / " & IIf(IsNull(rsDatosRuc.Fields("codgeo")), "", rsDatosRuc.Fields("codgeo"))
End With

End Sub

'Sub RefrescarFamiliar()
'With rsFamiliar
'     txtApePat.Text = .Fields("Apepat")
'     txtApeMat.Text = .Fields("ApeMat")
'     txtNombreuno.Text = .Fields("Nomb1")
'     txtNombredos.Text = .Fields("Nomb2")
'     If .Fields("CodPar") <> "" Then IzqBuscarEnCombo cboParen, .Fields("CodPar"), 1
'     IzqBuscarEnCombo Me.cboSexo, .Fields("Sexo"), 1
'     If .Fields("Fecnac") = "" Then .Fields("Fecnac") = ""
'End With
'End Sub
'
'Sub ModificarFamiliar()
'
'With rsFamiliar
'    If Not ValidarFam Then Exit Sub
'    .Fields("Nombres") = Trim(txtApePat.Text) & " " & Trim(txtApeMat.Text) & " " & Trim(txtNombreuno.Text) & " " & Trim(txtNombredos.Text)
'    .Fields("Apepat") = Trim(txtApePat.Text)
'    .Fields("ApeMat") = Trim(txtApeMat.Text)
'    .Fields("Nomb1") = Trim(txtNombreuno.Text)
'    .Fields("Nomb2") = Trim(txtNombredos.Text)
'    .Fields("CodPar") = Trim(Left(Trim(cboParen), 1))
'    .Fields("Parentesco") = Trim(Right(Trim(cboParen), 20))
'    .Fields("Sexo") = Left(cboSexo.Text, 1)
'    .Fields("Fecnac") = IIf(Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//", "", txtFecNac)
'End With
'End Sub

Function ValLugarNac(KeyAscii) As Boolean
Dim vCodGeo As String

If KeyAscii = 13 Then

If Trim(txtLugNac) = "" Then
         lblMensaje.Caption = "Debe Ingresar Código de Reniec"
         Beep
         ValLugarNac = False
         txtLugNac.SetFocus
         Exit Function
End If
    
        
      
                vCodGeo = fgDatosql("select codgeo from ubigeo where codubic = '" & Trim(txtLugNac.Text) & "'")
                If vCodGeo <> "" Then
                'Consultar no existe razon en truc
                'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValLugarNac = True
                   'vbindLugNac = True
                   lblMensaje.Caption = ""
                   lblGeoInei = vCodGeo
                   cboNacionalidad.SetFocus
                   vbUGeoInei = vCodGeo
                   
                Else
                   lblMensaje.Caption = "Código de Lugar Reniec no se encuentra en la Tabla Ubigeo"
                   ValLugarNac = False
                   vbindLugNac = False
                   txtLugNac.SetFocus
                    vbUGeoInei = ""
                    lblGeoInei = ""
                   Beep
                End If
                              
End If
 If KeyAscii = vbKeyF9 Then
   'Fuerza el pase al sgte campo
    ValLugarNac = True
    lblMensaje.Caption = ""
    vbindLugNac = False

    cboNacionalidad.SetFocus
 End If

End Function


Function ValRenta(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto ingreso de la remuneracion asegurada
If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
  If Not vbRenta Then  'Si no se dio la primera digitacion
    If val(Trim(txtRenta)) = 0 Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ValRenta = False
             txtRenta.SetFocus
             Exit Function
    End If
  End If


    If vbRenta Then
            If Trim(txtRenta.Text) <> "" Then
                vbRenta2 = True  '2da digitacion
            End If
        If Comparar(vtmpRenta, txtRenta.Text) Then
        'Pasa al sgte campo
         lblMensaje.Caption = ""
         ValRenta = True
                           
         'txtVolcFin.SetFocus
         cmdGrabarHoja.SetFocus
        Else
         lblMensaje.Caption = " Valores no coinciden"
         Beep
         ValRenta = False
         txtRenta.SetFocus
        End If
    Else
      vbRenta = True
      
      vtmpRenta = txtRenta.Text
      txtRenta.Text = ""
    ' txtRenta.SetFocus
      txtCodInterno.SetFocus
            
      
    End If
End If
 
 If KeyAscii = vbKeyF9 Then
    If vbRenta2 Then '2da Digitacion
       'Fuerza el pase al sgte campo
       ValRenta = True
       'El primer valor se hace igual al 2do
         
       If val(txtRemAseg.Text) = 0 Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             ValRenta = False
             txtRenta.SetFocus
             Exit Function
       Else
             vtmpRenta = txtRemAseg.Text
             cmdGrabarHoja.SetFocus
       End If
             
    Else
       ValRenta = False
       lblMensaje = "Debe digitar 2 veces el campo"
       Beep
    End If
 End If

End Function


Function PromDig(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto ingreso de la remuneracion asegurada
If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
  If Not vbPromI Then  'Si no se dio la primera digitacion
    If Trim(txtCodInterno) = "" Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             PromDig = False
             txtCodInterno.SetFocus
             Exit Function
    End If
  End If


    If vbPromI Then
            If Trim(txtCodInterno.Text) <> "" Then
                vbPromI2 = True  '2da digitacion
            End If
        If Comparar(vtmpPromI, txtCodInterno.Text) Then
        'Pasa al sgte campo
         lblMensaje.Caption = ""
         PromDig = True
         txtRenta.SetFocus
         Call ValProm(13)
         
        Else
         lblMensaje.Caption = " Valores no coinciden"
         Call ValProm(13)
         Beep
         PromDig = False
         txtCodInterno.SetFocus
        End If
    Else
      
      If ValProm(13) Then
        vbPromI = True
        vtmpPromI = txtCodInterno.Text
        txtCodInterno.Text = ""
        lblPromotor.Caption = ""
        Call ValProm(13)
        txtRenta.SetFocus
      End If
    End If
End If
 
 If KeyAscii = vbKeyF9 Then
    If vbPromI2 Then '2da Digitacion
   'Fuerza el pase al sgte campo
    PromDig = True
    lblMensaje.Caption = ""
    txtRenta.SetFocus
    'TabSig 13
   Else
    PromDig = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If

 End If



End Function


Sub ActIndRAM()

Dim cmdFed As Command
Set cmdFed = New Command

Dim rsDatosProm As Recordset
Set rsDatosProm = New Recordset


If val(txtRenta) = 0 Then

    With cmdFed
        .ActiveConnection = Con.Cnx
        .CommandText = " Update solicitud set IndRam = '" & IIf(val(txtRenta) = 0, 1, 0) & "', IndCSBSErr = '" & IIf(chkCodErr.Value = 1 Or vbindPro = False, "1", "0") & "'  " & _
                       " where numsol = '" & Numsol & "' and Fecsal = '" & Fecsal & "' and Inddoc = '" & TipDoc & "'  "
        .Execute
    End With

End If

Set cmdFed = Nothing
Set rsDatosProm = Nothing

End Sub

Sub InicializaVar()
vbRenta = False
vbRenta2 = False
vbnoCompRenta = False
vbDigRenta = False
vtmpRenta = ""
vbindPro = False

vbPromI = False
vbPromI2 = False
vbnoCompPromI = False
vbDigPromI = False
vtmpPromI = ""

End Sub

Function IndDirecCorr() As Boolean

If Trim(vbAveCR) = Trim(txtDireCorres) And Trim(vbNroCR) = Trim(txtNroCorres) And Trim(vbUrbaCR) = Trim(txtUrbaCorres) And Trim(vbUbigeoCR) = Trim(UcUbigeoCorres.cTexto) Then
    IndDirecCorr = True
Else
    IndDirecCorr = False
End If
End Function


Sub ImagenesAlternas(nomimag As String)
   With rsImagAlt
   
    sqltexto = "select top 4 Ruta, LoteFisico, Imagen from _imagenes a join " & _
    " _Volumenes b on a.IdVolumen = b.IdVolumen where  imagen > '" & nomimag & "' and a.Lotefisico = '" & vbGLte & "' order by imagen"
    If .State = 1 Then .Close
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    Cont = 0
    While Not .EOF
        ImgEdit3(Cont).Image = Rutaimag(.Fields("Ruta"), .Fields("LoteFisico"), .Fields("imagen"))
        ImgEdit3(Cont).Display
        ImgEdit3(Cont).FitTo 0
        Cont = Cont + 1
        .MoveNext
    Wend
               
   End With
   
End Sub

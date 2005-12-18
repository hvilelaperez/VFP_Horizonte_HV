VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAfiliacion 
   Caption         =   "Afiliación"
   ClientHeight    =   11010
   ClientLeft      =   585
   ClientTop       =   1050
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   13680
      TabIndex        =   117
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   14400
      TabIndex        =   116
      Text            =   "Text3"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   1695
      Left            =   120
      TabIndex        =   75
      Top             =   9240
      Width           =   15015
      _Version        =   131074
      _ExtentX        =   26485
      _ExtentY        =   2990
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AnnotationBackColor=   4210752
      AnnotationFillColor=   65535
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   447761664
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   33
      lReserved2      =   33
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   120
      TabIndex        =   67
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo Contrato"
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
         Left            =   6240
         TabIndex        =   114
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdOtroEmpleador 
         Caption         =   "Otro Empleador"
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
         Left            =   7920
         TabIndex        =   94
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtNroSolicitud 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   38
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtLote 
         Enabled         =   0   'False
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
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   37
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtFecsal 
         Enabled         =   0   'False
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
         Left            =   960
         MaxLength       =   15
         TabIndex        =   36
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Traspasos"
         Height          =   255
         Left            =   10440
         TabIndex        =   91
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Total"
         Height          =   255
         Left            =   12000
         TabIndex        =   90
         Top             =   120
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
         Left            =   11280
         TabIndex        =   89
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblTotal 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12480
         TabIndex        =   88
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Sec."
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
         TabIndex        =   70
         Top             =   120
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
         Left            =   2160
         TabIndex        =   69
         Top             =   120
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
         Left            =   360
         TabIndex        =   68
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "cmdAnterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   -120
      TabIndex        =   45
      Top             =   600
      Width           =   15255
      Begin VB.CommandButton cmdVerPromotor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Promotor"
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   6840
         Width           =   1575
      End
      Begin VB.TextBox txtRIAAfil 
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
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   6240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin Traspasos27.UserControl1 ucUbigeoEmp 
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   5520
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdHojaVida 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hoja de Vida"
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdrotDer 
         Caption         =   ">"
         Height          =   495
         Left            =   11160
         TabIndex        =   93
         Top             =   7080
         Width           =   495
      End
      Begin VB.CommandButton cmdrotIzq 
         Caption         =   "<"
         Height          =   495
         Left            =   10560
         TabIndex        =   92
         Top             =   7080
         Width           =   495
      End
      Begin Traspasos27.UserControl1 ucUbigeoAfi 
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   2280
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdEnlOtros 
         Caption         =   "Enl. &Otros."
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
         Left            =   7200
         TabIndex        =   43
         Top             =   7800
         Width           =   1095
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   3
         Left            =   13680
         TabIndex        =   86
         Top             =   4560
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
         UndoBufferSize  =   446532608
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   33
         lReserved2      =   33
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   2
         Left            =   13680
         TabIndex        =   85
         Top             =   3360
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
         UndoBufferSize  =   446532608
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   33
         lReserved2      =   33
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   1
         Left            =   13680
         TabIndex        =   84
         Top             =   2160
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
         UndoBufferSize  =   446532864
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   33
         lReserved2      =   33
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   0
         Left            =   13680
         TabIndex        =   83
         Top             =   960
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
         UndoBufferSize  =   446532608
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   33
         lReserved2      =   33
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin VB.CommandButton cmdEnlazarMec 
         Caption         =   "Enl. &Bol."
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
         Left            =   6000
         TabIndex        =   42
         Top             =   7800
         Width           =   1095
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   14055
         TabIndex        =   82
         Top             =   480
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
      Begin VB.TextBox txtFactor 
         Height          =   375
         Left            =   13680
         TabIndex        =   81
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   ">"
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
         Left            =   1320
         TabIndex        =   80
         Top             =   7800
         Width           =   975
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "<"
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
         Left            =   240
         TabIndex        =   79
         Top             =   7800
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
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
         Left            =   10560
         TabIndex        =   44
         Top             =   7800
         Width           =   1095
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit2 
         Height          =   2415
         Left            =   11760
         TabIndex        =   76
         Top             =   6120
         Width           =   3495
         _Version        =   131074
         _ExtentX        =   6165
         _ExtentY        =   4260
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   446527744
         OcrZoneVisibility=   -3624
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   33
         lReserved2      =   33
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin VB.CommandButton cmdMarcaHoja 
         Caption         =   "&Hoja Ctrl"
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
         Left            =   3600
         TabIndex        =   40
         Top             =   7800
         Width           =   1095
      End
      Begin VB.CommandButton cmdEnlazar 
         Caption         =   "&Enl. DNI"
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
         Left            =   4800
         TabIndex        =   41
         Top             =   7800
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   2400
         TabIndex        =   39
         Top             =   7800
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         Caption         =   "Antecedentes Previsionales"
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
         Height          =   1095
         Left            =   240
         TabIndex        =   66
         Top             =   2760
         Width           =   13335
         Begin VB.ComboBox cboTipAfi 
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
            ItemData        =   "frmAfiliacion.frx":0000
            Left            =   7560
            List            =   "frmAfiliacion.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox cboPriEmp 
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
            ItemData        =   "frmAfiliacion.frx":0024
            Left            =   5880
            List            =   "frmAfiliacion.frx":002E
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox cboLey25530 
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
            ItemData        =   "frmAfiliacion.frx":003A
            Left            =   4200
            List            =   "frmAfiliacion.frx":0044
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox cboLey19990 
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
            ItemData        =   "frmAfiliacion.frx":0050
            Left            =   2520
            List            =   "frmAfiliacion.frx":005A
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox cboTipTrabajador 
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
            ItemData        =   "frmAfiliacion.frx":0066
            Left            =   240
            List            =   "frmAfiliacion.frx":0070
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblX 
            Caption         =   "lblx"
            Height          =   375
            Left            =   10320
            TabIndex        =   108
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblY 
            Caption         =   "lbly"
            Height          =   375
            Left            =   11280
            TabIndex        =   107
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label35 
            Caption         =   "Tipo de Afiliacio"
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
            TabIndex        =   106
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label34 
            Caption         =   "Primer Empleo"
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
            TabIndex        =   105
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label33 
            Caption         =   "Ley 20530"
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
            Left            =   4200
            TabIndex        =   104
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label32 
            Caption         =   " Ley 19990"
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
            Left            =   2520
            TabIndex        =   103
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Trabajador"
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
            TabIndex        =   102
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Suscripción"
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
         Height          =   855
         Left            =   240
         TabIndex        =   64
         Top             =   6840
         Width           =   3735
         Begin VB.TextBox txtFechaSol 
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
            Left            =   840
            MaxLength       =   15
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label14 
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
            Left            =   240
            TabIndex        =   65
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   2055
         Left            =   240
         TabIndex        =   54
         Top             =   3960
         Width           =   13335
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
            ItemData        =   "frmAfiliacion.frx":0099
            Left            =   8040
            List            =   "frmAfiliacion.frx":009B
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   960
            Width           =   1455
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
            ItemData        =   "frmAfiliacion.frx":009D
            Left            =   240
            List            =   "frmAfiliacion.frx":009F
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   960
            Width           =   1215
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
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   25
            Top             =   360
            Width           =   6015
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
            Left            =   8640
            MaxLength       =   10
            TabIndex        =   32
            Top             =   1560
            Width           =   2055
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
            Left            =   9600
            MaxLength       =   100
            TabIndex        =   30
            Top             =   960
            Width           =   3255
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
            Left            =   1080
            MaxLength       =   11
            TabIndex        =   24
            Top             =   360
            Width           =   2295
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
            Left            =   5520
            MaxLength       =   40
            TabIndex        =   28
            Top             =   960
            Width           =   2295
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
            Height          =   315
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   27
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label Label38 
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
            Left            =   8160
            TabIndex        =   121
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Via"
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
            TabIndex        =   120
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social"
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
            TabIndex        =   113
            Top             =   360
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
            Left            =   240
            TabIndex        =   73
            Top             =   1320
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
            Left            =   8760
            TabIndex        =   63
            Top             =   1320
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
            Left            =   9600
            TabIndex        =   62
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label13 
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
            Left            =   1560
            TabIndex        =   61
            Top             =   720
            Width           =   3735
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
            Left            =   240
            TabIndex        =   57
            Top             =   360
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
            TabIndex        =   56
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
            Left            =   5520
            TabIndex        =   55
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Promotor"
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
         Height          =   735
         Left            =   240
         TabIndex        =   95
         Top             =   6000
         Width           =   11415
         Begin VB.CheckBox chkCodErr 
            Caption         =   "Código errado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   9840
            TabIndex        =   115
            Top             =   280
            Width           =   1540
         End
         Begin VB.TextBox lblPromotor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   6255
         End
         Begin VB.TextBox txtCodProm 
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
            MaxLength       =   5
            TabIndex        =   33
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Código Promotor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   330
            Width           =   1575
         End
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   12480
         TabIndex        =   78
         Text            =   "Text4"
         Top             =   3960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame2 
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
         TabIndex        =   46
         Top             =   0
         Width           =   13335
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
            ItemData        =   "frmAfiliacion.frx":00A1
            Left            =   8280
            List            =   "frmAfiliacion.frx":00A3
            Style           =   2  'Dropdown List
            TabIndex        =   15
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
            ItemData        =   "frmAfiliacion.frx":00A5
            Left            =   240
            List            =   "frmAfiliacion.frx":00A7
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtNumdoc 
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
            MaxLength       =   12
            TabIndex        =   8
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cboNacionalidad 
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
            Left            =   9360
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtIPSS 
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
            MaxLength       =   15
            TabIndex        =   9
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtSexo 
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
            MaxLength       =   1
            TabIndex        =   4
            Top             =   1080
            Width           =   495
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
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtUbiNacRe 
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
            Left            =   960
            MaxLength       =   6
            TabIndex        =   5
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtTelefAfi 
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
            Left            =   9000
            MaxLength       =   10
            TabIndex        =   18
            Top             =   2280
            Width           =   2055
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
            Left            =   9840
            MaxLength       =   100
            TabIndex        =   16
            Top             =   1680
            Width           =   3375
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
            Left            =   5640
            MaxLength       =   40
            TabIndex        =   14
            Top             =   1680
            Width           =   2535
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
            Height          =   315
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   13
            Top             =   1680
            Width           =   3735
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
            Left            =   5160
            MaxLength       =   3
            TabIndex        =   7
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtApePat 
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
            TabIndex        =   0
            Top             =   480
            Width           =   2655
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
            Left            =   11520
            MaxLength       =   12
            TabIndex        =   11
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
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
            Left            =   3720
            MaxLength       =   30
            TabIndex        =   1
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
            Left            =   6840
            MaxLength       =   20
            TabIndex        =   2
            Top             =   480
            Width           =   2535
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
            Left            =   9720
            MaxLength       =   20
            TabIndex        =   3
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label29 
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
            Left            =   8400
            TabIndex        =   119
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Via"
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
            TabIndex        =   118
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Nacionalidad"
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
            Left            =   9360
            TabIndex        =   110
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "IPSS"
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
            TabIndex        =   109
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
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
            TabIndex        =   101
            Top             =   840
            Width           =   495
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
            Left            =   3720
            TabIndex        =   100
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblLugarNac 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2040
            TabIndex        =   99
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar Nac."
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
            Left            =   960
            TabIndex        =   98
            Top             =   840
            Width           =   975
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
            TabIndex        =   72
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono"
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
            Left            =   9240
            TabIndex        =   71
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label12 
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
            Left            =   9840
            TabIndex        =   60
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Número / Departamento "
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
            Left            =   5640
            TabIndex        =   59
            Top             =   1440
            Width           =   4815
         End
         Begin VB.Label Label10 
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
            Left            =   1800
            TabIndex        =   58
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Doc  Ide"
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
            TabIndex        =   53
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tip. Doc."
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
            Left            =   5280
            TabIndex        =   52
            Top             =   840
            Width           =   975
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
            Left            =   11520
            TabIndex        =   51
            Top             =   840
            Visible         =   0   'False
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
            TabIndex        =   50
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
            Left            =   4320
            TabIndex        =   49
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
            Left            =   6960
            TabIndex        =   48
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
            Left            =   9720
            TabIndex        =   47
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   12360
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "RIA"
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
         Left            =   7800
         TabIndex        =   112
         Top             =   6240
         Width           =   375
      End
      Begin VB.Label lblMensaje 
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4920
         TabIndex        =   74
         Top             =   7080
         Width           =   4455
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
         Left            =   13680
         TabIndex        =   87
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAfiliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vbAccion As String
Dim vbmarca As Variant
Dim rsCtrldoc As New Recordset
Dim rsSolicitud As New Recordset
Dim RsEmpleadores As New Recordset
Dim sqltexto As String

Private vbApepat As Boolean
Private vbApepat2 As Boolean
Private vbnoCompApepat As Boolean
Private vbDigApepat As Boolean
Private vtmpApepat As String

Private vbApeMat As Boolean
Private vbApeMat2 As Boolean
Private vbnoCompApeMat As Boolean
Private vbDigApeMat As Boolean
Private vtmpApeMat As String

Private vbNomb1 As Boolean
Private vbNomb12 As Boolean
Private vbnoCompNomb1 As Boolean
Private vbDigNomb1 As Boolean
Private vtmpNomb1 As String

Private vbNomb2 As Boolean
Private vbNomb22 As Boolean
Private vbnoCompNomb2 As Boolean
Private vbDigNomb2 As Boolean
Private vtmpNomb2 As String

Private vbDNI As Boolean
Private vbDNI2 As Boolean
Private vbnoCompDNI As Boolean
Private vbDigDNI As Boolean
Private vtmpDNI As String

Private vbCuss As Boolean
Private vbCuss2 As Boolean
Private vbnoCuss As Boolean
Private vbDigCuss As Boolean
Private vtmpCuss As String

'Variable Obligatorio
Private vbNroSol As Boolean
Private vbNroSol2 As Boolean
Private vtmpNroSol As String
Private vbDigNroSol As Boolean

'Fecha de nacimiento
Private vbFecnac As Boolean
Private vbFecnac2 As Boolean
Private vtmpFecnac As String
Private vbDigFecnac As Boolean


Private rsUbigeoAfi As New Recordset
Private rsUbigeoEmp As New Recordset
Private RsImag As New Recordset

'Afiliado

Private vbCodDist As String
Private vbCodProv As String
Private vbCodDpto As String

Private vbDesDist As String
Private vbDesProv As String
Private vbDesDpto As String

'Empl

Private vbCodDistEmp As String
Private vbCodProvEmp As String
Private vbCodDptoEmp As String

Private vbDesDistEmp As String
Private vbDesProvEmp As String
Private vbDesDptoEmp As String

Private vbCodgeo As String
Private vbCodgeoEmp As String
Private vbindRuc As Boolean
Private vbindPro As Boolean 'Indicador de Codigo de Promotor
Private vbindEmp As Boolean 'Indicador de Nuevo Empleador para la misma solicitud
Private vbindEmp2 As Boolean 'Indicaador para saber si tiene o no empleador

Private rsImagAlt As New Recordset
Private vbNumsolant As String 'Validar solicitud anterior

Private vbtmpFTipoO As Integer  'Origen
Private vbtmpFTipoD As Integer  'Destino
Private vbdigTipFondo As Boolean ' Indica si ya se digito El Tipo de Fondo
Private vbdigF9FO As Boolean 'Indica q presiono f9 Fondo Origen
Private vbdigF9FD As Boolean 'Indica q presiono f9 Fondo destino
Private RutaImagUlt As String 'Ruta completa de la ultima imagen
Private ImagUlt As String 'Nombre de la ultima imagen
Private vbRucUlt As String
Private vbRazSocUlt As String
Private vbAveUlt As String
Private vbNroUlt As String
Private vbUrbaUlt As String
Private vbUbigeoUlt As String
Private vbNomcUlt As String
Private vbFecnacUlt As String
Private vbUbigeoNacUbiUlt As String
Private vbSexoUlt As String

Private vbAveUltCR As String
Private vbNroUltCR As String
Private vbUrbaUltCR As String
Private vbUbigeoUltCR As String
Private indCambioFec As Boolean
Private indFecSus As Boolean
Private vbTelefAfi As String
Private vbTelefEmp As String

Private vbPrefDirAfi As String
Private vbPrefLocAfi As String
Private vbPrefDirEmp As String
Private vbPrefLocEmp As String

Private vbcodAdm As String


Sub GrabarTraspaso()
If Not validar Then Exit Sub

    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = " sp_GrabarAfiliacion2 ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,? "
                                          
                                                                            
        .Parameters("@CodAfp") = vbGAFP
        .Parameters("@FecSal") = Trim(txtFecsal)
        .Parameters("@Lote") = Trim(txtLote)
        .Parameters("@NumSol") = Trim(txtNroSolicitud)
        .Parameters("@ApePat") = Trim(txtApePat)
        .Parameters("@ApeMat") = Trim(txtApeMat)
        .Parameters("@Nombre1") = Trim(txtNombreuno)
        .Parameters("@Nombre2") = Trim(txtNombredos)
        .Parameters("@TipoDoc") = Trim(txtTipDoc)
        .Parameters("@NumDoc") = Trim(txtNumdoc)
        .Parameters("@CUSPP") = Trim(txtCUSPP)
        .Parameters("@Direccion") = Trim(txtAv)
        .Parameters("@Numero") = Trim(txtNroDir)
        .Parameters("@Urb") = Trim(txtUrba)
        .Parameters("@Distrito") = vbDesDist
        .Parameters("@Provincia") = vbDesProv
        .Parameters("@Departamento") = vbDesDpto
        .Parameters("@Ubigeo") = vbCodgeo
        .Parameters("@Telefono") = IIf(Trim(txtTelefAfi) = "", "1111111", Trim(txtTelefAfi))
        .Parameters("@Empleadores") = IIf(vbindEmp, "1", "0") 'Revisar luego
        .Parameters("@Cantemple") = "0"  'Luego se calcula
        .Parameters("@CodPromo") = Trim(txtCodProm)
        .Parameters("@BolPago") = ""
        .Parameters("@PagoUltApo") = ""
        .Parameters("@CopiaSim") = ""
        .Parameters("@OtroDoc") = ""
        .Parameters("@DescOtroDoc") = ""
        .Parameters("@FecRegPro") = IIf(Trim(txtFechaSol) = "" Or Trim(txtFechaSol) = "//", Null, Trim(txtFechaSol))
        .Parameters("@CodAFPOri") = ""
        .Parameters("@TipoFonOri") = ""
        .Parameters("@TipoFonDes") = ""
        .Parameters("@FecFirmSol") = IIf(Trim(txtFechaSol) = "" Or Trim(txtFechaSol) = "//", Null, Trim(txtFechaSol))
        .Parameters("@Aceptacion") = ""
        .Parameters("@Texto") = ""
        .Parameters("@RIAafil") = IIf(val(txtRIAAfil) = 0, 1#, val(txtRIAAfil))
        .Parameters("@FecNac") = IIf((Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//"), Null, txtFecNac)
        .Parameters("@Sexo") = Trim(txtSexo) 'IIf(Trim(txtSexo) = "", "M", "F")
        .Parameters("@e_mail") = ""
        .Parameters("@Status") = "5"
        .Parameters("@IDimagen") = vbGAFP & Trim(txtNroSolicitud)
        .Parameters("@CodUsu") = gNomUsu
        .Parameters("@Fecdig") = gDate
        .Parameters("@Horini") = ""
        .Parameters("@Horfin") = ObtenerHora
        .Parameters("@Numpul") = 0
        .Parameters("@CodMod") = gNomUsu
        .Parameters("@Fecmod") = Null
        .Parameters("@Horinim") = ""
        .Parameters("@Horfinm") = ""
        .Parameters("@Ruc") = Trim(txtRuc)
        .Parameters("@DireEmp") = Trim(txtDireEmp)
        .Parameters("@NroEmp") = Trim(txtNroEmp)
        .Parameters("@UrbaEmp") = Trim(txtUrbaEmp)
        .Parameters("@DistritoEmp") = vbDesDistEmp
        .Parameters("@ProvinciaEmp") = vbDesProvEmp
        .Parameters("@Departamentoemp") = vbDesDptoEmp
        .Parameters("@UbigeoEmp") = vbCodgeoEmp
        .Parameters("@TelefEmp") = IIf(Trim(txtTelefEmp) = "", "1111111", Trim(txtTelefEmp))
        .Parameters("@ExisteBD") = IIf(vbindRuc, "1", "0")
        .Parameters("@nomImag") = RsImag!Imagen
        .Parameters("@indEmp") = IIf(vbindEmp, "1", "0")
        .Parameters("@indExisEmp") = IIf(vbindEmp2, "1", "0")
        .Parameters("@TipImag") = IIf(vbindEmp, "E", "S")
        .Parameters("@posicion") = RsImag!Posicion
        .Parameters("@Cuenta") = ""
        .Parameters("@Boletin") = ""
        .Parameters("@e_mail2") = ""
        .Parameters("@indDoc") = "C"
        .Parameters("@UbigeoNac") = Trim(txtUbiNacRe)
        .Parameters("@UbigeoRen") = Trim(lblLugarNac.Caption)
        .Parameters("@IPSS") = LimpiarCadena(Trim(txtIPSS))
        .Parameters("@Nac") = Trim(Right(cboNacionalidad, 3)) 'Por ver si son los ultimos 3 digitos
        .Parameters("@TipTrab") = Left(cboTipTrabajador, 1)
        .Parameters("@Ley19990") = Left(cboLey19990, 1)
        .Parameters("@Ley25530") = Left(cboLey25530, 1)
        .Parameters("@PriEmp") = Left(cboPriEmp, 1)
        .Parameters("@TipAfi") = Left(cboTipAfi, 1)
        .Parameters("@RazSoc") = Trim(lblRazonSoc)
        .Parameters("@PrefDirAfi") = Right(Trim(cboPrefDirAfi.Text), 2)
        .Parameters("@PrefLocAfi") = Right(Trim(cboPrefLocAfi.Text), 2)
        .Parameters("@PrefDirEmp") = Right(Trim(cboPrefDirEmp.Text), 2)
        .Parameters("@PrefLocEmp") = Right(Trim(cboPrefLocEmp.Text), 2)
                     
        If Con.GetFEjecutarSQL(comando) Then
            If LCase(Con.vbdev) = "exito" Then
                Call ActIndFedatario
                vbNumsolant = txtNroSolicitud.Text
                lblcont.Caption = val(lblcont.Caption) + 1
                Call NuevoRegistro
                
            End If
            
        End If
    
                                          

    End With

End Sub

Sub ModificarTraspaso()

vbmarca = rsSolicitud.Bookmark

If Not validar Then Exit Sub

    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = " sp_ModificarSolicitud1 ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,? "
                                                                            
        .Parameters("@CodAfp") = vbGAFP
        .Parameters("@FecSal") = Trim(txtFecsal)
        .Parameters("@Lote") = Trim(txtLote)
        .Parameters("@NumSol") = Trim(txtNroSolicitud)
        .Parameters("@ApePat") = Trim(txtApePat)
        .Parameters("@ApeMat") = Trim(txtApeMat)
        .Parameters("@Nombre1") = Trim(txtNombreuno)
        .Parameters("@Nombre2") = Trim(txtNombredos)
        .Parameters("@TipoDoc") = Trim(txtTipDoc)
        .Parameters("@NumDoc") = Trim(txtNumdoc)
        .Parameters("@CUSPP") = Trim(txtCUSPP)
        .Parameters("@Direccion") = Trim(txtAv)
        .Parameters("@Numero") = Trim(txtNroDir)
        .Parameters("@Urb") = Trim(txtUrba)
        .Parameters("@Distrito") = vbDesDist
        .Parameters("@Provincia") = vbDesProv
        .Parameters("@Departamento") = vbDesDpto
        .Parameters("@Ubigeo") = Trim(Left(Mid(ucUbigeoAfi.cTexto, InStr(ucUbigeoAfi.cTexto, "/") + 1), 7))
        .Parameters("@Telefono") = Trim(txtTelefAfi)
        .Parameters("@Empleadores") = "0"  'Revisar luego
        .Parameters("@Cantemple") = "0"  'Luego se calcula
        .Parameters("@CodPromo") = Trim(txtCodProm)
        .Parameters("@BolPago") = IIf(chkBolPag.Value = 1, "S", "N")
        .Parameters("@PagoUltApo") = IIf(chkPagUlt.Value = 1, "S", "N")   '
        .Parameters("@CopiaSim") = IIf(chkCopSim.Value = 1, "S", "N")   '
        .Parameters("@OtroDoc") = IIf(chkOtroDoc.Value = 1, "S", "N")    '
        .Parameters("@DescOtroDoc") = Trim(txtEspecificar)
        .Parameters("@FecRegPro") = IIf(Trim(txtFechaSol) = "", Null, Trim(txtFechaSol))
        .Parameters("@CodAFPOri") = Left(Trim(cboAfp.Text), 2)
        .Parameters("@TipoFonOri") = Trim(txtFTipoO.Text)
        .Parameters("@TipoFonDes") = Trim(txtFTipoD.Text)
        .Parameters("@FecFirmSol") = IIf(Trim(txtFecTras) = "", Null, Trim(txtFecTras))
        .Parameters("@Aceptacion") = IIf(chkProc.Value = 1, "S", "N")
        .Parameters("@Texto") = Trim(txtMotiv)
        .Parameters("@RIAafil") = val(txtRIAAfil)
        .Parameters("@FecNac") = IIf((Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//"), Null, txtFecNac)
        .Parameters("@Sexo") = IIf(Trim(txtSexo) = "", "M", "F")
        .Parameters("@e_mail") = Trim(txtemail)
        .Parameters("@Status") = "5"
        .Parameters("@IDimagen") = vbGAFP & Trim(txtNroSolicitud)
        .Parameters("@CodUsu") = gNomUsu
        .Parameters("@Fecdig") = gDate
        .Parameters("@Horini") = ""
        .Parameters("@Horfin") = ObtenerHora
        .Parameters("@Numpul") = 0
        .Parameters("@CodMod") = gNomUsu
        .Parameters("@Fecmod") = Null
        .Parameters("@Horinim") = ""
        .Parameters("@Horfinm") = ""
        .Parameters("@Ruc") = Trim(txtRuc)
        .Parameters("@DireEmp") = Trim(txtDireEmp)
        .Parameters("@NroEmp") = Trim(txtNroEmp)
        .Parameters("@UrbaEmp") = Trim(txtUrbaEmp)
        .Parameters("@DistritoEmp") = vbDesDistEmp
        .Parameters("@ProvinciaEmp") = vbDesProvEmp
        .Parameters("@Departamentoemp") = vbDesDptoEmp
        .Parameters("@UbigeoEmp") = Trim(Left(Mid(ucUbigeoEmp.cTexto, InStr(ucUbigeoEmp.cTexto, "/") + 1), 7))
        .Parameters("@TelefEmp") = Trim(txtTelefEmp)
        .Parameters("@ExisteBD") = IIf(vbindRuc, "1", "0")
        .Parameters("@nomImag") = rsSolicitud!Imagen
        .Parameters("@indEmp") = IIf(vbindEmp, "1", "0")
        .Parameters("@indExisEmp") = IIf(vbindEmp2, "1", "0")
        .Parameters("@TipImag") = IIf(vbindEmp, "E", "S")
        .Parameters("@posicion") = rsSolicitud!Posicion
        .Parameters("@Secuencial") = rsSolicitud!Secu
        .Parameters("@Cuenta") = IIf(chkCuenta.Value, "X", "")
        .Parameters("@Boletin") = IIf(chkBol.Value, "X", "")
        .Parameters("@e_mail2") = Trim(txtemail2)
        
                     
        If Con.GetFEjecutarSQL(comando) Then
            If LCase(Con.vbdev) = "exito" Then
                'Call NuevoRegistro
                CargarDatosSolicitud
                rsSolicitud.Bookmark = vbmarca
                Call RefrescarSol(False)
            End If
            
        End If
                                                                  

    End With


End Sub






Private Sub cboLey19990_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 351, 966
cboLey19990.BackColor = &HC0FFFF
End Sub

Private Sub cboLey19990_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboLey19990.ListIndex = -1
End If
End Sub

Private Sub cboLey19990_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboLey19990_LostFocus()
cboLey19990.BackColor = &H80000005
End Sub

Private Sub cboLey19990_Validate(Cancel As Boolean)
If Trim(txtIPSS) = "" Then
    If UCase(cboLey19990.Text) = "SI" Then
        If MsgBox("El valor de Ley 19990 debe ser NO, ¿Desea Continuar?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        End If
    End If
Else
    If UCase(cboLey19990.Text) = "NO" Then
        If MsgBox("El valor de Ley 19990 debe ser SI, ¿Desea Continuar?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub cboLey25530_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 857, 966
cboLey25530.BackColor = &HC0FFFF
End Sub

Private Sub cboLey25530_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboLey25530.ListIndex = -1
End If

End Sub

Private Sub cboLey25530_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboLey25530_LostFocus()
cboLey25530.BackColor = &H80000005
End Sub

Private Sub cboLey25530_Validate(Cancel As Boolean)
    If UCase(cboLey25530.Text) = "SI" Then
        If MsgBox("El valor de Ley 25530 debe ser NO, ¿Desea Continuar?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub cboNacionalidad_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1772, 680
cboNacionalidad.BackColor = &HC0FFFF
End Sub

Private Sub cboNacionalidad_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboNacionalidad_LostFocus()
cboNacionalidad.BackColor = &H80000005
End Sub


Private Sub cboNacionalidad_Validate(Cancel As Boolean)

End Sub

Private Sub cboPrefDirAfi_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 760
End Sub

Private Sub cboPrefDirAfi_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefDirAfi_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefDirAfi.ListIndex = -1
End Sub

Private Sub cboPrefDirEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 57, 1161
End Sub

Private Sub cboPrefDirEmp_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefDirEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefDirEmp.ListIndex = -1
End Sub

Private Sub cboPrefLocAfi_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1716, 760
End Sub

Private Sub cboPrefLocAfi_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefLocAfi_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefLocAfi.ListIndex = -1
End Sub

Private Sub cboPrefLocEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1695, 1161
End Sub

Private Sub cboPrefLocEmp_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefLocEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefLocEmp.ListIndex = -1
End Sub

Private Sub cboPriEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1365, 966
cboPriEmp.BackColor = &HC0FFFF
End Sub

Private Sub cboPriEmp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboPriEmp.ListIndex = -1
End If

End Sub

Private Sub cboPriEmp_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPriEmp_LostFocus()
cboPriEmp.BackColor = &H80000005
End Sub

Private Sub cboPriEmp_Validate(Cancel As Boolean)
If Trim(txtIPSS) = "" Then
    If UCase(cboPriEmp.Text) = "NO" Then
        If MsgBox("El valor de Primer Empleo debe ser SI, ¿Desea Continuar?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        End If
    End If
Else
    If UCase(cboPriEmp.Text) = "SI" Then
        If MsgBox("El valor de Primer Empleo debe ser NO, ¿Desea Continuar?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub cboTipAfi_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1767, 966
cboTipAfi.BackColor = &HC0FFFF
End Sub

Private Sub cboTipAfi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboTipAfi.ListIndex = -1
End If
End Sub

Private Sub cboTipAfi_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboTipAfi_LostFocus()
cboTipAfi.BackColor = &H80000005
End Sub

Private Sub cboTipAfi_Validate(Cancel As Boolean)
If Trim(txtIPSS) = "" Then
    If UCase(cboTipAfi.Text) = "SI" Then
        If MsgBox("El valor de Tipo de Afiliación debe ser NO, ¿Desea Continuar?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        End If
    End If
Else
    If UCase(cboTipAfi.Text) = "NO" Then
        If MsgBox("El valor de Tipo de Afiliación debe ser SI, ¿Desea Continuar?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub cboTipTrabajador_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 966
cboTipTrabajador.BackColor = &HC0FFFF
End Sub

Private Sub cboTipTrabajador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboTipTrabajador.ListIndex = -1
End If

End Sub

Private Sub cboTipTrabajador_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboTipTrabajador_LostFocus()
cboTipTrabajador.BackColor = &H80000005
End Sub


Private Sub cmdAnterior_Click()
On Local Error Resume Next
With rsSolicitud
    .MovePrevious
    If .BOF Then .MoveFirst
End With
RefrescarSol True
 Resaltado 0, 0, 1200, 150, (Int((ImgEdit1.ImageScaleWidth / 3))), 100
End Sub

Private Sub cmdEnlazar_Click()
On Error GoTo ERRORES

If MsgBox("¿Esta seguro de que la imagen es un DNI ?", vbQuestion + vbYesNo, Caption) = vbNo Then Exit Sub

If vbNumsolant = "" Then
'   If Trim(txtNroSolicitud) = "" Then
'      MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
'      Exit Sub
'   Else
'     vbNumsolant = txtNroSolicitud
'   End If
RecuperacionAntRegistro
End If

bloqueo:

If RsImag.EOF Then Exit Sub

If Trim(vbNumsolant) <> "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set idimagen = '" & vbGAFP & Trim(vbNumsolant) & "', Status = '5', TipImag =  'D', " & _
            " posicion = isnull((select max(convert(int,posicion)) + 1  from _imagenes where idimagen = '" & vbGAFP & Trim(vbNumsolant) & "'),1) where Imagen = '" & RsImag!Imagen & "'"
            .Execute , , adExecuteNoRecords
            If Not CargarImagen Then Unload Me: Exit Sub
      End With
Else
    MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
End If
txtApePat.SetFocus
Exit Sub
ERRORES:
If Err.Number = -2147467259 Then
    GoTo bloqueo
End If

End Sub

Private Sub cmdEnlazarMec_Click()
On Error GoTo ERRORES
If MsgBox("¿Esta seguro que la imagen es un Boleta ?", vbQuestion + vbYesNo, Caption) = vbNo Then Exit Sub
If vbNumsolant = "" Then
'   If Trim(txtNroSolicitud) = "" Then
'      MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
'      Exit Sub
'   Else
'     vbNumsolant = txtNroSolicitud
'   End If
   RecuperacionAntRegistro
End If

bloqueo:


If RsImag.EOF Then Exit Sub
If Trim(vbNumsolant) <> "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set idimagen = '" & vbGAFP & Trim(vbNumsolant) & "', Status = '5', TipImag =  'B', " & _
            " posicion = isnull((select max(convert(int,posicion)) + 1  from _imagenes where idimagen = '" & vbGAFP & Trim(vbNumsolant) & "'),1) where Imagen = '" & RsImag!Imagen & "'"
            .Execute
            If Not CargarImagen Then Unload Me: Exit Sub
      End With
Else
    MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
End If
'txtNroSolicitud.SetFocus
txtApePat.SetFocus
Exit Sub
ERRORES:
If Err.Number = -2147467259 Then
    GoTo bloqueo
End If

End Sub


Private Sub cmdEnlOtros_Click()
On Error GoTo ERRORES:

If MsgBox("¿Esta seguro de que la imagen es un documento no clasificado (Otro)?", vbQuestion + vbYesNo, Caption) = vbNo Then Exit Sub

If vbNumsolant = "" Then
   RecuperacionAntRegistro
'   If Trim(txtNroSolicitud) = "" Then
'      MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
'      Exit Sub
'   Else
'     vbNumsolant = txtNroSolicitud
'   End If
End If
bloqueo:
If RsImag.EOF Then Exit Sub
If Trim(vbNumsolant) <> "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set idimagen = '" & vbGAFP & Trim(vbNumsolant) & "', Status = '5', TipImag =  'O', " & _
            " posicion = isnull((select max(convert(int,posicion)) + 1  from _imagenes where idimagen = '" & vbGAFP & Trim(vbNumsolant) & "'),1) where Imagen = '" & RsImag!Imagen & "'"
            .Execute
            If Not CargarImagen Then Unload Me: Exit Sub
      End With
Else
    MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
End If
txtApePat.SetFocus
Exit Sub
ERRORES:
If Err.Number = -2147467259 Then
    GoTo bloqueo
End If

End Sub


Private Sub CmdGrabar_Click()

If Frame2.Enabled Then
'Si se trata de una nueva solicitud almacena los nuevos valores
'para la posterior Hoja de Vida
    
    vbNomcUlt = Trim(txtApePat) & " " & Trim(txtApeMat) & " " & Trim(txtNombreuno) & " " & Trim(txtNombredos)
    'Afiliado
    vbAveUltCR = Trim(txtAv)
    vbNroUltCR = Trim(txtNroDir)
    vbUrbaUltCR = Trim(txtUrba)
    vbUbigeoUltCR = Trim(ucUbigeoAfi.cTexto)
    
    vbFecnacUlt = txtFecNac.Text
    vbUbigeoNacUbiUlt = Trim(lblLugarNac.Caption)
    vbTelefAfi = Trim(txtTelefAfi)
    vbSexoUlt = txtSexo
    
    vbPrefDirAfi = Right(Trim(cboPrefDirAfi.Text), 2)
    vbPrefLocAfi = Right(Trim(cboPrefLocAfi.Text), 2)
    
End If

    'Empleador
    vbRucUlt = Trim(txtRuc)
    vbRazSocUlt = Trim(lblRazonSoc.Text)
    vbAveUlt = Trim(txtDireEmp)
    vbNroUlt = Trim(txtNroEmp)
    vbUrbaUlt = Trim(txtUrbaEmp)
    vbUbigeoUlt = Trim(ucUbigeoEmp.cTexto)
    vbTelefEmp = Trim(txtTelefEmp)
    vbPrefDirEmp = Right(Trim(cboPrefDirEmp.Text), 2)
    vbPrefLocEmp = Right(Trim(cboPrefLocEmp.Text), 2)
    

Call VeriFecSus
If cmdGrabar.Caption = "&Grabar" Then
   Call GrabarTraspaso
  Else
   Call ModificarTraspaso
End If

chkCodErr.Value = 0

End Sub

Function CargarImagen() As Boolean
On Error Resume Next
Dim strsql As String
'Lectura de 1 Imagen
strsql = " Select top 1 imagen, c.Ruta, LoteFisico, fecsal,posicion " & _
         " from _Imagenes b join _Volumenes c on c.idvolumen = b.idvolumen " & _
         " join CtrlDoc d on d.Lote = b.Lotefisico " & _
         " where barcode = 0 and b.status = '1' and b.Lotefisico = '" & vbGLte & "' order by  imagen"
                 'idImagen is null
    If RsImag.State = 1 Then RsImag.Close
    RsImag.Open strsql, Con.Cnx, adOpenStatic, adLockReadOnly
    If Not (RsImag.EOF And RsImag.BOF) Then
            
        
        txtNroSolicitud.SetFocus
        
        Frame2.Enabled = True 'Datos Afiliado
        Frame5.Enabled = True  'Datos Verificacion
        Frame4.Enabled = True  'Declaracion de Traspaso
        Frame6.Enabled = True  'Resultado de la Solicitud
        
        txtFecNac.Enabled = True '
            
            
      'Se Toma la imagen y se le cambia el Status
      Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set Status = '4' where Imagen = '" & RsImag!Imagen & "'"
            .Execute
      End With
      
      RsImag.MoveFirst
      
      ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      RutaImagUlt = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      ImagUlt = RsImag!Imagen
      
      ImagenesAlternas RsImag!Imagen
      
      txtLote.Text = RsImag!lotefisico
      txtFecsal.Text = RsImag!Fecsal
     'Copia la imagen al image 2
      ImgEdit2.Image = ImgEdit1.Image
      ImgEdit2.FitTo 0
      ImgEdit2.DisplayScaleAlgorithm = wiScaleGray8
      ImgEdit2.Display
      'ImgEdit1.FitTo 1
      ImgEdit1.Display
    '  If Err.Number > 0 Then ImgEdit1.Enabled = False: MsgBox Err.Description
       Resaltado 0, 0, 1200, 150, (Int((ImgEdit1.ImageScaleWidth / 3))), 100
     '  ImgEdit1.Display
      
    '  lblTotpag.Caption = rsImag!imagfin
      'Call VerPlanillaScala
      'Call MovLinea
      CargarImagen = True
      
      Else
        If fgDatosql("Select Estado from _scan where lotefisico = '" & Format(vbGLte, "0000000#") & "' and Estado < 3 ") <> "" Then
            Call CambiarEstadoLote(Trim(Format(vbGLte, "0000000#")), "3")
        End If
      
        MsgBox "Lote Concluido (no hay más imágenes para el Lote)", 64, Caption
      
        CargarImagen = False
      
    End If
        
        
End Function

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


Function ApePat(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto de apellido paterno

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtApePat) = "" Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ApePat = False
             txtApePat.SetFocus
             Exit Function
    End If

    If Not vbApepat Then
      vbApepat = True
      vbApepat2 = True  '2da digitacion
      vtmpApepat = txtApePat.Text
      txtApePat.Text = ""
      txtApeMat.SetFocus
    End If
 End If
 
 If KeyAscii = vbKeyF9 Then
   If vbApepat2 Then '2da Digitacion
    vbnoCompApepat = True
    'El primer valor se hace igual al 2do
    lblMensaje.Caption = ""
   Else
    ApePat = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If
 End If
 
 If vbApepat Then
        
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
        If Not vbnoCompApepat Then  'si el falso
          
            If Not CompararTexto(vtmpApepat, txtApePat.Text, KeyAscii, Len(txtApePat) + 1) Then
              KeyAscii = 0
              Beep
              lblMensaje.Caption = " Caracter no Corresponde"
              txtApePat.SetFocus
              ApePat = False
    
            Else
               lblMensaje.Caption = ""
            End If
            
        Else
           vbnoCompApepat = False
        End If
    Else
       
       If KeyAscii = 13 And (Len(vtmpApepat) <= Len(txtApePat)) Then ApePat = True: txtApeMat.SetFocus
    End If
 End If
 
End Function

Function ApeMat(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto de apellido paterno

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtApeMat) = "" Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ApeMat = False
             txtApeMat.SetFocus
             Exit Function
    End If

    If Not vbApeMat Then
      vbApeMat = True
      vbApeMat2 = True  '2da digitacion
      vtmpApeMat = txtApeMat.Text
      txtApeMat.Text = ""
      txtNombreuno.SetFocus
    End If
 End If
 
 If KeyAscii = vbKeyF9 Then
   If vbApeMat2 Then '2da Digitacion
    vbnoCompApeMat = True
    'El primer valor se hace igual al 2do
    lblMensaje.Caption = ""
   Else
    ApeMat = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If
 End If
 
 If vbApeMat Then
        
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
        If Not vbnoCompApeMat Then  'si el falso
          
            If Not CompararTexto(vtmpApeMat, txtApeMat.Text, KeyAscii, Len(txtApeMat) + 1) Then
              KeyAscii = 0
              Beep
              lblMensaje.Caption = " Caracter no Corresponde"
              txtApeMat.SetFocus
              ApeMat = False
    
            Else
              lblMensaje.Caption = ""
            End If
            
        Else
           vbnoCompApeMat = False
        End If
    Else
       If KeyAscii = 13 And (Len(vtmpApeMat) <= Len(txtApeMat)) Then ApeMat = True: txtNombreuno.SetFocus
    End If
 End If
 
End Function

Function ValDNI(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto de apellido paterno

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtNumdoc) = "" Then
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ValDNI = False
             txtNumdoc.SetFocus
             Exit Function
    End If

    If Not vbDNI Then
      vbDNI = True
      vbDNI2 = True  '2da digitacion
      vtmpDNI = txtNumdoc.Text
      txtNumdoc.Text = ""
      txtIPSS.SetFocus
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
 
 If vbDNI Then
        
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
        If Not vbnoCompDNI Then  'si el falso
          
            If Not CompararTexto(vtmpDNI, txtNumdoc.Text, KeyAscii, Len(txtNumdoc) + 1) Then
              KeyAscii = 0
              Beep
              lblMensaje.Caption = " Caracter no Corresponde"
              txtNumdoc.SetFocus
              ValDNI = False
     
            Else
              lblMensaje.Caption = ""
            End If
            
        Else
           vbnoCompDNI = False
        End If
    Else
       If KeyAscii = 13 And (Len(vtmpDNI) <= Len(txtNumdoc)) Then
            ValDNI = True
            txtCUSPP.SetFocus
       End If
    End If
 End If
 
End Function
Function ValCuss(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto de apellido paterno

   
If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtCUSPP) = "" Then
             'lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ValCuss = False
             'txtCUSPP.SetFocus
             If Not vbCuss Then
                vbCuss = True
                cboPrefDirAfi.SetFocus
                ValCuss = True
             Else
                cmdGrabar.SetFocus
             End If
             
             Exit Function
    End If

    If Not vbCuss Then
         If Not ValCusp2(txtCUSPP) Then
           ValCuss = False
           lblMensaje.Caption = "CUSPP No cumple estructura establecida ######AAAAA#"
           txtCUSPP.SetFocus
           Exit Function
         End If

      vbCuss = True
      vbCuss2 = True  '2da digitacion
      vtmpCuss = txtCUSPP.Text
      txtCUSPP.Text = ""
      'txtAv.SetFocus
      cboPrefDirAfi.SetFocus
    End If
 End If
 
 
 If KeyAscii = vbKeyF9 Then
 
  
  
   If vbCuss2 Then '2da Digitacion
    vbnoCuss = True
    'El primer valor se hace igual al 2do
    lblMensaje.Caption = ""
    cmdGrabar.SetFocus
   Else
    vbCuss = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If
   
    If Not ValCusp2(txtCUSPP) Then
    ValCuss = False
    lblMensaje.Caption = "CUSPP No cumple estructura establecida ######AAAAA#"
    txtCUSPP.SetFocus
    'Exit Function
   End If
   
 End If
 
 
 
 If vbCuss Then
        
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
        If Not vbnoCuss Then  'si el falso
          
            If Not CompararTexto(vtmpCuss, txtCUSPP.Text, KeyAscii, Len(txtCUSPP) + 1) Then
              KeyAscii = 0
              Beep
              lblMensaje.Caption = " Caracter no Corresponde"
              txtCUSPP.SetFocus
              ValCuss = False
    
            Else
              lblMensaje.Caption = ""
            End If
            
        Else
           vbnoCuss = False
        End If
    Else
       If KeyAscii = 13 And (Len(vtmpCuss) <= Len(txtCUSPP.Text)) Then vbCuss = True: lblMensaje = "": cmdGrabar.SetFocus
    End If
 End If
 
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


Function NombreUno(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto  apellido paterno

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtNombreuno) = "" Then
             NombreUno = True
             txtNombredos.SetFocus
             Exit Function
    End If

    If Not vbNomb1 Then
      vbNomb1 = True
      vbNomb12 = True  '2da digitacion
      vtmpNomb1 = txtNombreuno.Text
      txtNombreuno.Text = ""
      txtNombredos.SetFocus
    End If
 End If
 
 If KeyAscii = vbKeyF9 Then
   If vbNomb12 Then '2da Digitacion
    vbnoCompNomb1 = True
    'El primer valor se hace igual al 2do
    lblMensaje.Caption = ""
   Else
    NombreUno = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If
 End If
 
 If vbNomb1 Then
        
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
        If Not vbnoCompNomb1 Then  'si el falso
          
            If Not CompararTexto(vtmpNomb1, txtNombreuno.Text, KeyAscii, Len(txtNombreuno) + 1) Then
              KeyAscii = 0
              Beep
              lblMensaje.Caption = " Caracter no Corresponde"
              txtNombreuno.SetFocus
              NombreUno = False
    
            Else
              lblMensaje.Caption = ""
            End If
            
        Else
           vbnoCompNomb1 = False
        End If
    Else
       If KeyAscii = 13 And (Len(vtmpNomb1) <= Len(txtNombreuno)) Then NombreUno = True: txtNombredos.SetFocus
    End If
 End If
 
End Function

Function NombreDos(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto de apellido paterno

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(txtNombredos) = "" Then
             NombreDos = True
             txtSexo.SetFocus
             Exit Function
    End If

    If Not vbNomb2 Then
      vbNomb2 = True
      vbNomb22 = True  '2da digitacion
      vtmpNomb2 = txtNombredos.Text
      txtNombredos.Text = ""
      txtSexo.SetFocus
    End If
 End If
 
 If KeyAscii = vbKeyF9 Then
   If vbNomb22 Then '2da Digitacion
    vbnoCompNomb2 = True
    'El primer valor se hace igual al 2do
    lblMensaje.Caption = ""
   Else
    NombreDos = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If
 End If
 
 If vbNomb2 Then
        
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 209) Then
        If Not vbnoCompNomb2 Then  'si el falso
          
            If Not CompararTexto(vtmpNomb2, txtNombredos.Text, KeyAscii, Len(txtNombredos) + 1) Then
              KeyAscii = 0
              Beep
              lblMensaje.Caption = " Caracter no Corresponde"
              txtNombredos.SetFocus
              NombreDos = False
            Else
              lblMensaje.Caption = ""
            End If
            
        Else
           vbnoCompNomb2 = False
        End If
    Else
       If KeyAscii = 13 And (Len(vtmpNomb2) <= Len(txtNombredos)) Then NombreDos = True: Me.txtFecNac.SetFocus
    End If
 End If
 
End Function


Private Sub cmdGrabar_GotFocus()
'If txtFTipoO.Text = "" Then
' If txtFTipoO.Enabled Then
' txtFTipoO.SetFocus
' End If
'End If
End Sub

Private Sub cmdHojaVida_Click()
If vbNumsolant = "" Then ' MsgBox "Debe digitar primero el Documento del Traspaso", 64, Caption: Exit Sub

    Call RecuperacionAntRegistro
    
    If vbNumsolant = "" Then MsgBox "Debe digitar primero el Contrato de Afiliación", 64, Caption: Exit Sub
        
End If

With frmHojaVida
    
        .Numsol = vbNumsolant
        .TipDoc = "C"
        .Fecsal = txtFecsal
        .Secu = 0
        .rutaImagen = RutaImagUlt
        .Imagen = ImagUlt
        .NombresComp = vbNomcUlt
        .vbRuc = vbRucUlt
        .vbAve = vbAveUlt
        .vbNro = vbNroUlt
        .vbUrba = vbUrbaUlt
        .vbUbigeo = vbUbigeoUlt
        .vbRazSoc = vbRazSocUlt
        .Fecnac = vbFecnacUlt
        .UbigeoNacUbi = vbUbigeoNacUbiUlt
        .Sexo = vbSexoUlt
        
        .vbAveCR = vbAveUltCR
        .vbNroCR = vbNroUltCR
        .vbUrbaCR = vbUrbaUltCR
        .vbUbigeoCR = vbUbigeoUltCR
        
        .vbTelefAfi = vbTelefAfi
        .vbTelefEmp = vbTelefEmp
        .CodAdmDoc = vbcodAdm
        
    
        .vbPrefDirCR = vbPrefDirAfi
        .vbPrefLocCR = vbPrefLocAfi
        .vbPrefDirEmp = vbPrefDirEmp
        .vbPrefLocEmp = vbPrefLocEmp
        
        .Show 1
End With

      
End Sub

Private Sub cmdMarcaHoja_Click()

If MsgBox("¿Esta seguro de que la imagen es una Hoja de Control?", vbQuestion + vbYesNo, Caption) = vbNo Then Exit Sub
Set comando = New Command

With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "Update _Imagenes set Barcode = 1, TipImag =  'H' where imagen = '" & RsImag!Imagen & "' "
    .Execute
End With
If Not CargarImagen Then Unload Me: Exit Sub
End Sub

Private Sub cmdNuevo_Click()
If MsgBox("¿Esta seguro de que se trata de un nuevo contrato?", vbQuestion + vbYesNo + vbDefaultButton1, Caption) = vbYes Then
    Call NuevoContrato
End If
End Sub

Private Sub cmdOtroEmpleador_Click()
Call BuscarSolicitud
End Sub

'Private Sub cmdRecAnterior_Click()
'Call RecuperacionAntRegistro
'End Sub

Private Sub cmdrotDer_Click()
On Local Error Resume Next
ImgEdit1.RotateRight
End Sub

Private Sub cmdrotIzq_Click()
On Local Error Resume Next
ImgEdit1.RotateLeft
End Sub

Private Sub CmdSalir_Click()
'Liberar la imagen actual
Unload Me
End Sub

Private Sub cmdSiguiente_Click()
On Local Error Resume Next
With rsSolicitud
    .MoveNext
    If .EOF Then .MoveLast
End With
RefrescarSol True
 Resaltado 0, 0, 1200, 150, (Int((ImgEdit1.ImageScaleWidth / 3))), 100
End Sub

Private Sub cmdVerPromotor_Click()
FrmConsPromo.Show 1
End Sub

'Private Sub cmdSuelto_Click()
''Imagen suelta
''If Trim(txtNumDoc) = "" Then
'    Set comando = New Command
'      With comando
'            .ActiveConnection = Con.Cnx
'            .CommandText = "Update _Imagenes set Status = '2', TipImag = 'N' where Imagen = '" & rsImag!Imagen & "'"
'            .Execute
'      End With
'      If Not CargarImagen Then Unload Me: Exit Sub
''End If
'End Sub


Private Sub Form_Load()
''''CargarComboSql cboAfp, "select codafp + ' - ' + descri, codafp from tafp"
Inicializar
vbGAFP = "01" ''CAmbiar--------
vbNumsolant = ""
CargarComboSql cboNacionalidad, "select descri_corta, codigo from nacionalidad "
pgBuscarEnCombo cboNacionalidad, "000114", 6

Set ucUbigeoAfi.cConexion = Con.Cnx
    ucUbigeoAfi.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    ucUbigeoAfi.Campocondicion = "desgeo"
    ucUbigeoAfi.Busqinicial = False
    ucUbigeoAfi.Enabled = True
    ucUbigeoAfi.nroCaracteres = 3
    'ucUbigeoAfi.RecordBusqueda

 Set ucUbigeoEmp.cConexion = Con.Cnx
    ucUbigeoEmp.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    ucUbigeoEmp.Campocondicion = "desgeo"
    ucUbigeoEmp.Busqinicial = False
    ucUbigeoEmp.Enabled = True
    ucUbigeoEmp.nroCaracteres = 3


If vbAccion = "Nuevo" Then
    cmdGrabar.Caption = "&Grabar"
    cmdAnterior.Enabled = False
    cmdSiguiente.Enabled = False
    Call CargarImagen

Else
    cmdAnterior.Enabled = True
    cmdSiguiente.Enabled = True
    cmdGrabar.Caption = "&Modificar"
    Call CargarDatosSolicitud
    RefrescarSol True
End If



CargarComboSql cboPrefDirAfi, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocAfi, "select AbreLoc, CodLoc  from prefloc"
CargarComboSql cboPrefDirEmp, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocEmp, "select AbreLoc, CodLoc  from prefloc"

lblTotal.Caption = frmTraspasos1.dtgLotes.Columns(4)
txtNroSolicitud = "C" & Format(IIf(val(ObtenerSecuencial) = 0, 0, ObtenerSecuencial), "0000000#")
txtNroSolicitud.Enabled = False

'
'ImgEdit1.Image = "C:\NEW AFILIACION 1.tif"
'ImgEdit1.Display



End Sub
Sub Inicializar()
    'Afiliados
        vbCodDist = ""
        vbCodProv = ""
        vbCodDpto = ""
        vbDesDist = ""
        vbDesProv = ""
        vbDesDpto = ""
        vbCodgeo = ""
      '  vbcodAdm = ""
        
     txtFactor = 0
    'Empleadores
        vbCodDistEmp = ""
        vbCodProvEmp = ""
        vbCodDptoEmp = ""
        vbDesDistEmp = ""
        vbDesProvEmp = ""
        vbDesDptoEmp = ""
        vbCodgeoEmp = ""
        
        txtCodProm = ""
        vbindRuc = False
        txtNroSolicitud = ""
        
        vbApepat = False
        vbApepat2 = False
        vbnoCompApepat = False
        vbDigApepat = False
        vtmpApepat = ""
        
        vbApeMat = False
        vbApeMat2 = False
        vbnoCompApeMat = False
        vbDigApeMat = False
        vtmpApeMat = ""
        
        vbNomb1 = False
        vbNomb12 = False
        vbnoCompNomb1 = False
        vbDigNomb1 = False
        vtmpNomb1 = ""
        
        vbNomb2 = False
        vbNomb22 = False
        vbnoCompNomb2 = False
        vbDigNomb2 = False
        vtmpNomb2 = ""
        
        vbDNI = False
        vbDNI2 = False
        vbnoCompDNI = False
        vbDigDNI = False
        vtmpDNI = ""
    
        vbCuss = False
        vbCuss2 = False
        vbnoCuss = False
        vbDigCuss = False
        vtmpCuss = ""
        
        'Variable Nro Solicitud
        vbNroSol = False
        vbNroSol2 = False
        vtmpNroSol = ""
        vbDigNroSol = False
        
        vbFecnac = False
        vbFecnac2 = False
        vtmpFecnac = ""
        vbDigFecnac = False
        
        
        vbindEmp2 = True  'Por defecto tiene empleador
        
        'Ind de Empleador
        vbindEmp = False

                
        txtApePat.Text = ""
        txtApeMat.Text = ""
        txtNombreuno = ""
        txtNombredos = ""
        'txtTipDoc = ""
        txtNumdoc = ""
        txtCUSPP = ""
        txtAv = ""
        txtNroDir = ""
        txtUrba = ""
        ucUbigeoAfi.cTexto = ""
                
        cboPrefDirAfi.ListIndex = -1
        cboPrefLocAfi.ListIndex = -1

        cboPrefDirEmp.ListIndex = -1
        cboPrefLocEmp.ListIndex = -1
        
        txtTelefAfi = ""
        'x defecto
        txtTipDoc = "00"
        
        txtRuc = ""
        txtDireEmp = ""
        txtNroEmp = ""
        txtUrbaEmp = ""
        ucUbigeoEmp.cTexto = ""
        txtTelefEmp = ""
        chkBolPag = 0
        chkPagUlt = 0
        chkCopSim = 0
        chkOtroDoc = 0
        'txtEspecificar = ""
        txtFechaSol = ""
        'cboAfp.ListIndex = -1
        txtFecTras = ""
        'chkProc.Value = 0
        'txtMotiv = ""
        txtFecNac = ""
        txtRIAAfil = ""
        
        txtSexo = ""
        txtUbiNacRe.Text = ""
        txtIPSS.Text = ""
        cboNacionalidad.ListIndex = -1
        cboTipTrabajador.ListIndex = -1
        cboLey19990.ListIndex = -1
        cboLey25530.ListIndex = -1
        cboPriEmp.ListIndex = -1
        cboTipAfi.ListIndex = -1
        lblLugarNac.Caption = ""
        
        pgBuscarEnCombo cboNacionalidad, "000114", 6
        
        
'        vbUbigeoUlt = Trim(ucUbigeoEmp.cTexto)
'        vbFecnacUlt = txtFecNac.Text
'        vbUbigeoNacUbiUlt = Trim(lblLugarNac.Caption)
'        vbSexoUlt = txtSexo
'
'        vbNomcUlt = ""
'        vbRucUlt = ""
'        vbRazSocUlt = ""
'        vbAveUlt = ""
'        vbNroUlt = ""
'        vbUrbaUlt = ""
'        vbUbigeoUlt = ""
'        vbFecnacUlt = ""
'        vbUbigeoNacUbiUlt = ""
'        vbSexoUlt = ""
                               
      
End Sub
Sub NuevoRegistro()
Call Inicializar
'Cargar la nueva imagen
If Not CargarImagen Then Unload Me: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set comando = New Command

If RsImag.State <> 0 Then

    If Not RsImag.EOF Then
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set Status = '1' where Imagen = '" & RsImag!Imagen & "'"
            .Execute
        End With
    End If
End If


Set frmTraspasos = Nothing

End Sub

Private Sub ImgEdit1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2 = X
Text3 = Y
End Sub

Private Sub ImgEdit1_Scroll()
lblX = ImgEdit1.ScrollPositionX
lblY = ImgEdit1.ScrollPositionY
End Sub

Private Sub ImgEdit2_Click()
 SiempreVisible frmVisorImag, True
 frmVisorImag.ImgEdit1.Image = ImgEdit2.Image
 frmVisorImag.ImgEdit1.Display
 frmVisorImag.ImgEdit1.FitTo 1
 frmVisorImag.Show
End Sub

Private Sub ImgEdit3_Click(Index As Integer)
 SiempreVisible frmVisorImag, True
 
 frmVisorImag.ImgEdit1.Image = ImgEdit3(Index).Image
 frmVisorImag.ImgEdit1.Display
 frmVisorImag.ImgEdit1.FitTo 1
 frmVisorImag.Show
End Sub

Private Sub lblPromotor_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 39, 653
End Sub

Private Sub lblPromotor_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
End Sub

Private Sub lblRazonSoc_GotFocus()
ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 0, 517
End Sub

Private Sub lblRazonSoc_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub txtApeMat_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 525, 580
txtApeMat.BackColor = &HC0FFFF
End Sub

Private Sub txtApeMat_LostFocus()
txtApeMat.BackColor = &H80000005
End Sub

Private Sub TxtApePat_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 580
txtApePat.BackColor = &HC0FFFF
End Sub

Private Sub txtApePat_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
' If KeyCode = vbKeyTab Then
'  vbDigApepat = ApePat(13)
' Exit Sub
' End If
' If KeyCode = vbKeyF9 Then
'  vbDigApepat = ApePat(KeyCode)
' End If
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

Private Sub TxtApePat_LostFocus()
txtApePat.BackColor = &H80000005
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
''Else
''    If Not vbDigApepat Then Cancel = True
''End If

End Sub

Private Sub txtApeMat_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
' If KeyCode = vbKeyTab Then
'  vbDigApeMat = ApeMat(13)
' Exit Sub
' End If
' If KeyCode = vbKeyF9 Then
'  vbDigApeMat = ApeMat(KeyCode)
' End If
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
''Else
''    If Not vbDigApeMat Then Cancel = True
''End If

End Sub




Private Sub txtAv_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 760
txtAv.BackColor = &HC0FFFF
End Sub

Private Sub txtAv_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then 'KeyAscii = 39 Or
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii

End Sub

Private Sub txtAv_LostFocus()
txtAv.BackColor = &H80000005
End Sub

Private Sub txtCodProm_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
'ImgEdit1.Zoom = 50
Resaltado 0, 0, 1200, 150, 0, 1358
txtCodProm.BackColor = &HC0FFFF
End Sub

Private Sub txtCodProm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then
 Call ValProm(13)
End If

If KeyCode = vbKeyF9 Then
 Call ValProm(KeyCode)
End If

End Sub

Private Sub txtCodProm_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtUbiNacRe, KeyAscii, vbCaracteresNumero)
'TabSig KeyAscii
If KeyAscii = 13 Then Call ValProm(KeyAscii)
'FEntDec txtRuc, KeyAscii, 11, 0
End Sub

Private Sub txtCodProm_KeyUp(KeyCode As Integer, Shift As Integer)
If txtCodProm.MaxLength = Len(Trim(txtCodProm)) Then
    ValProm (13)
End If
End Sub

Private Sub txtCodProm_LostFocus()
txtCodProm.BackColor = &H80000005
End Sub

Private Sub txtCodProm_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ValProm(13) Then
      Cancel = True
    End If
    Exit Sub
  End If

End Sub

Private Sub TxtCUSPP_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
   Resaltado 0, 0, 1200, 150, 1713, 416
   Marcado txtCUSPP
   txtCUSPP.BackColor = &HC0FFFF
End Sub

Private Sub TxtCUSPP_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 32 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18 Or KeyCode = 209 Then
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
txtCUSPP.BackColor = &H80000005
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

  If vbDigCuss Then Cancel = False: Exit Sub
  
  If txtCUSPP = val(rsDetalle.Fields("nomuno")) Then
     'If Err Then vbDigNomb1 = ValRemuAsig(13): Exit Sub
     Cancel = False
  Else
     vbDigCuss = ValCuss(13)
     If Not vbDigCuss Then Cancel = True
  End If

End Sub

Private Sub txtDireEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 57, 1161
txtDireEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtDireEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then 'Or KeyAscii = 39
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii


End Sub

Private Sub txtDireEmp_LostFocus()
txtDireEmp.BackColor = &H80000005
End Sub

Private Sub txtDireEmp_Validate(Cancel As Boolean)
If MsgBox("¿Verificó la dirección contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
    Cancel = False
Else
    Cancel = True
End If
End Sub

Private Sub txtFechaSol_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 719, 1584
txtFechaSol.BackColor = &HC0FFFF
End Sub

Private Sub txtFechaSol_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
FVFecha txtFechaSol, KeyAscii

End Sub

Private Sub txtFechaSol_LostFocus()
ValidaFecha txtFechaSol
txtFechaSol.BackColor = &H80000005
cmdGrabar.SetFocus
End Sub

Private Sub txtFechaSol_Validate(Cancel As Boolean)
If IsDate(txtFechaSol) Then
    If DateDiff("d", txtFechaSol, gDate) > 7 Then
        If MsgBox("La Fecha de Suscripción es anterior a una semana, Desea Continuar? ", vbQuestion + vbYesNo, Caption) = vbNo Then
            Cancel = True
        Else
            Cancel = False
        End If
    End If
End If
End Sub

Private Sub txtFecNAc_GotFocus()
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 680
txtFecNac.BackColor = &HC0FFFF
End Sub

Private Sub txtFecNac_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF9 Then
'  vbDigFecnac = ValFecNac(KeyCode)
' End If
End Sub

Private Sub txtFecNac_KeyPress(KeyAscii As Integer)
    TabSig KeyAscii
    KeyAscii = ValidarDigitacion(txtUbiNacRe, KeyAscii, "0123456789")
    FVFecha txtFecNac, KeyAscii
 '   vbDigFecnac = ValFecNac(KeyAscii)
End Sub

Private Sub txtFecNac_LostFocus()
    txtFecNac.BackColor = &H80000005
    ValidaFecha txtFecNac
End Sub

Private Sub txtFecNac_Validate(Cancel As Boolean)

If IsDate(txtFecNac) Then
    If (Year(gDate) - Year(txtFecNac)) < 17 Or (Year(gDate) - Year(txtFecNac)) > 102 Then
        MsgBox "Fecha no válida", vbCritical, Caption
        Cancel = True
    End If
End If

'On Error Resume Next
''If vbAccion = "Modificar" Then
'
'  If (GetKeyState(vbKeyTab) < 0) Then
'    If Not ValFecNac(13) Then
'     Cancel = True
'    End If
'    Exit Sub
'  End If
'
'  If vbDigFecnac Then Cancel = False: Exit Sub
'
End Sub

Private Sub txtFecsal_KeyPress(KeyAscii As Integer)

    FVFecha txtFecsal, KeyAscii
    
End Sub

Private Sub txtFecsal_LostFocus()
    ValidaFecha txtFecsal
End Sub

Private Sub txtIPSS_GotFocus()
Marcado txtIPSS
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1724, 680
txtIPSS.BackColor = &HC0FFFF
End Sub

Private Sub txtIPSS_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtIPSS_LostFocus()
txtIPSS.BackColor = &H80000005
End Sub

Private Sub txtNombreDos_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1745, 580
txtNombredos.BackColor = &HC0FFFF
End Sub

Private Sub txtNombreDos_LostFocus()
txtNombredos.BackColor = &H80000005
End Sub

Private Sub txtNombreUno_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1080, 580
txtNombreuno.BackColor = &HC0FFFF
End Sub

Private Sub txtNombreuno_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18 Or KeyCode = 209) Then
' If KeyCode = vbKeyTab Then
'  vbDigNomb1 = NombreUno(13)
' Exit Sub
' End If
' If KeyCode = vbKeyF9 Then
'  vbDigNomb1 = NombreUno(KeyCode)
' End If
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
    If KeyAscii = 13 Then txtSexo.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNombredos_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
'
'    If (GetKeyState(vbKeyTab) < 0) Then
'        If Not NombreDos(13) Then
'         Cancel = True
'        End If
'        Exit Sub
'    End If
'
'
'  If vbDigNomb2 Then Cancel = False: Exit Sub
'
'  If txtNombredos = val(rsDetalle.Fields("nomdos")) Then
'     'If Err Then vbDigNomb1 = ValRemuAsig(13): Exit Sub
'     Cancel = False
'  Else
'     vbDigNomb2 = NombreDos(13)
'     If Not vbDigNomb2 Then Cancel = True
'  End If
'Else
'    If Not vbDigNomb2 Then Cancel = True
'End If

End Sub

Private Sub txtNombreUno_LostFocus()
txtNombreuno.BackColor = &H80000005
End Sub

Private Sub txtNombreuno_Validate(Cancel As Boolean)
On Error Resume Next
''If vbAccion = "Modificar" Then
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
''Else
''    If Not vbDigNomb1 Then Cancel = True
''End If
End Sub

Private Sub txtNroDir_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 585, 760
txtNroDir.BackColor = &HC0FFFF
End Sub

Private Sub txtNroDir_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209 Or KeyAscii = 78 Or KeyAscii = 83 Or KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 45) Then  'Or KeyAscii = 39
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii

End Sub

Private Sub txtNroDir_LostFocus()
txtNroDir.BackColor = &H80000005
End Sub

Private Sub txtNroEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 876, 1161
txtNroEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtNroEmp_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig KeyAscii

KeyAscii = Asc(UCase(Chr(KeyAscii)))

If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209 Or KeyAscii = 78 Or KeyAscii = 83 Or KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 45) Then  'Or KeyAscii = 39
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii

End Sub

Private Sub txtNroEmp_LostFocus()
txtNroEmp.BackColor = &H80000005
End Sub

Private Sub txtNroEmp_Validate(Cancel As Boolean)

If MsgBox("¿Verificó el numero contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
    Cancel = False
Else
    Cancel = True
End If

End Sub

Private Sub txtNroSolicitud_GotFocus()
On Local Error Resume Next
 If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
 Marcado txtNroSolicitud
 Resaltado 0, 0, 1200, 150, (Int((ImgEdit1.ImageScaleWidth / 3))), 176
 txtNroSolicitud.BackColor = &HC0FFFF
End Sub

Private Sub txtNroSolicitud_KeyPress(KeyAscii As Integer)
vbDigNroSol = ValNroSol(KeyAscii)
FEntDec txtNroSolicitud, KeyAscii, 7, 0
If vbAccion = "Modificar" Then Exit Sub

'If Trim(txtNroSolicitud) <> "" Then
'    Call BuscarSolicitud
'End If

End Sub

Private Sub txtNroSolicitud_KeyUp(KeyCode As Integer, Shift As Integer)
  If txtNroSolicitud.MaxLength = Len(Trim(txtNroSolicitud)) Then
      Call ValNroSol(13)
  End If

End Sub

Private Sub txtNroSolicitud_LostFocus()
txtNroSolicitud.Text = Format(txtNroSolicitud.Text, "000000#")
txtNroSolicitud.BackColor = &H80000005
End Sub
Sub BuscarSolicitud()

'vbNumsolant variable temporal q contiene el anterior
'numero de solicitud


If Trim(vbNumsolant) <> "" Then
    
        Frame2.Enabled = False  'Datos Afiliado
        Frame5.Enabled = False  'Datos Verificacion
                
        txtUbiNacRe.Enabled = False
        txtFecNac.Enabled = False '
        txtNroSolicitud = vbNumsolant
        
        txtRuc.SetFocus
        vbindEmp = True
    Else
        'Limpiar las varibles para q pueda ingresar otro Nro Solicitud
        MsgBox "No hay Nro de Solicitud anterior", vbCritical, Caption
        vbNroSol = False
        vbNroSol2 = False
        vtmpNroSol = ""
        vbDigNroSol = False
        txtNroSolicitud.SetFocus
        Frame2.Enabled = True  'Datos Afiliado
        Frame5.Enabled = True  'Datos Verificacion
                
        txtFecNac.Enabled = True
End If


End Sub

Private Sub txtNroSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyTab Then
 vbDigNroSol = ValNroSol(13)
Exit Sub
End If
If KeyCode = vbKeyF9 Then
 vbDigNroSol = ValNroSol(KeyCode)
End If
End Sub


Private Sub txtNroSolicitud_Validate(Cancel As Boolean)
On Local Error Resume Next
'If vbAccion = "Modificar" Then

  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ValNroSol(13) Then
      Cancel = True
    End If
    Exit Sub
  End If

  If vbDigNroSol Then Cancel = False: Exit Sub
  
  If txtNroSolicitud = RsPlanilla.Fields("Nrosol") Then
     Cancel = False
  Else
     vbDigNroSol = ValNroSol(13)
     If Not vbDigNroSol Then Cancel = True
  End If
End Sub

Private Sub txtNumDoc_GotFocus()
If txtTipDoc.Text = "04" Then
    txtNumdoc.MaxLength = 12
Else
    txtNumdoc.MaxLength = 8
End If

Marcado txtNumdoc
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1224, 680
txtNumdoc.BackColor = &HC0FFFF
End Sub

Private Sub txtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)

If ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
' If KeyCode = vbKeyTab Then
'  vbDigDNI = ValDNI(13)
' Exit Sub
' End If
'
' If KeyCode = vbKeyF9 Then
'  vbDigDNI = ValDNI(KeyCode)
' End If
Else
KeyCode = 0
End If
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
'Dim keyasciitmp As Integer
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
'keyasciitmp = KeyAscii

If txtTipDoc.Text = "00" Then
    FEntDec txtNumdoc, KeyAscii, 8, 0
End If

If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
    'FEntDec txtNumDoc, KeyAscii, 8, 0
    'vbDigDNI = ValDNI(KeyAscii)
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 Or KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
Else
    KeyAscii = 0
    'If keyasciitmp = 13 Then vbDigDNI = ValDNI(keyasciitmp)
End If
End Sub

Private Sub txtNumDoc_LostFocus()
txtNumdoc.BackColor = &H80000005
End Sub

Private Sub txtNumDoc_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then

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

Private Sub txtRIAAfil_GotFocus()
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 2402
txtRIAAfil.BackColor = &HC0FFFF
'txtFecNac &HC0FFFF
End Sub

Private Sub txtRIAAfil_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
FEntDec txtRIAAfil, KeyAscii, 13, 2
'txtRIAAfil.BackColor = &H80000005
End Sub

Private Sub txtRIAAfil_LostFocus()
txtRIAAfil.BackColor = &H80000005
  
End Sub

Private Sub txtRIAAfil_Validate(Cancel As Boolean)
    If val(txtRIAAfil) >= 50000 Then
      If MsgBox("RIA de Afiliado es mayor a 50,000.00, El RIA es Correcto ?", vbYesNo, Caption) = vbYes Then
        txtRIAAfil.Text = 1
        Cancel = False
      Else
        Cancel = True
        'txtRIAAfil.SetFocus
      End If
    End If

End Sub

Private Sub txtRuc_LostFocus()
txtRuc.BackColor = &H80000005
End Sub

Private Sub txtSexo_GotFocus()
Marcado txtSexo
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1745, 600
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

Private Sub txtTelefAfi_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1689, 840
txtTelefAfi.BackColor = &HC0FFFF
End Sub

Private Sub txtTelefAfi_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
FEntDec txtTelefAfi, KeyAscii, 10, 0
End Sub

Private Sub txtTelefAfi_LostFocus()
txtTelefAfi.BackColor = &H80000005
End Sub

Private Sub txtTelefEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1656, 1245
txtTelefEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtTelefEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
FEntDec txtTelefEmp, KeyAscii, 10, 0
End Sub

Private Sub txtTelefEmp_LostFocus()
txtTelefEmp.BackColor = &H80000005
End Sub

Private Sub txtTipDoc_GotFocus()
Marcado txtTipDoc
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 715, 680
txtTipDoc.BackColor = &HC0FFFF
End Sub

Private Sub txtTipDoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtTipDoc, KeyAscii, "0123456789")
TabSig KeyAscii
FEntDec txtTipDoc, KeyAscii, 2, 0
End Sub

Private Sub txtTipDoc_LostFocus()
txtTipDoc.BackColor = &H80000005
End Sub

Private Sub txtTipDoc_Validate(Cancel As Boolean)

Select Case txtTipDoc
Case "00"
    
    CargarComboSql cboNacionalidad, "select descri_corta, codigo from nacionalidad where codigo = '000114'"
    pgBuscarEnCombo cboNacionalidad, "000114", 6
        
Case "01"
    CargarComboSql cboNacionalidad, "select descri_corta, codigo from nacionalidad where codigo <> '000114'"
    'cboNacionalidad.ListIndex = 0
End Select



End Sub

Private Sub txtUbiNacRe_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 244, 680
txtUbiNacRe.BackColor = &HC0FFFF
End Sub

Private Sub txtUbiNacRe_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
 Call ValLugarNac(KeyCode)
End If
End Sub

Private Sub txtUbiNacRe_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtUbiNacRe, KeyAscii, "0123456789")
'TabSig KeyAscii
If KeyAscii = 13 Then Call ValLugarNac(KeyAscii)
End Sub

Private Sub txtUbiNacRe_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ValLugarNac(13) Then
      Cancel = True
    End If
    Exit Sub
  End If
End Sub

Private Sub txtUrba_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1716, 760
txtUrba.BackColor = &HC0FFFF
End Sub

Private Sub txtUrba_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then 'Or KeyAscii = 39
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii

End Sub

Private Sub txtUrba_LostFocus()
txtUrba.BackColor = &H80000005
End Sub

Private Sub txtUrbaEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1695, 1161
txtUrbaEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtUrbaEmp_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig KeyAscii

KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then 'Or KeyAscii = 39
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii


End Sub

Private Sub txtUrbaEmp_LostFocus()
txtUrbaEmp.BackColor = &H80000005
End Sub

Private Sub txtUrbaEmp_Validate(Cancel As Boolean)
If MsgBox("¿Verificó la urbanización contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
    Cancel = False
Else
    Cancel = True
End If
End Sub

Private Sub ucUbigeoAfi_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 38, 412
' txtUrba.BackColor = &H80000005
End Sub

Private Sub ucUbigeoAfi_KeyUP(KeyAscii As Integer)
If KeyAscii = vbKeyF9 Then TabSig 13
End Sub

Private Sub ucUbigeoAfi_LostFocus()
' If vbCodDist = "" Or vbCodProv = "" Or vbCodDist = "" Then
    LLenarDatosUbigeo
 'End If
End Sub

Private Sub ucUbigeoAfi_Pulsacion(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtUbiNacRe, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ- " & vbCaracteresNumero)

If KeyAscii = 13 Then
    If LLenarDatosUbigeo Then TabSig KeyAscii
Else
    If KeyAscii = vbKeyF9 Then TabSig 13
End If
End Sub
Function LLenarDatosUbigeo() As Boolean
'Ubigeo de Afiliados
On Error Resume Next
If rsUbigeoAfi.State = 1 Then rsUbigeoAfi.Close

rsUbigeoAfi.Open "select * from ubigeo where codgeo = '" & Trim(Mid(Trim(ucUbigeoAfi.cTexto), InStr(ucUbigeoAfi.cTexto, "/") + 1, 10)) & "'", Con.Cnx, adOpenStatic
    If (rsUbigeoAfi.EOF And rsUbigeoAfi.BOF) Then
        MsgBox "El Dato ingresado no existe", 64, "Mensaje"
        LLenarDatosUbigeo = False
        'rsUbigeoAfi.SetFocus
        vbCodDist = ""
        vbCodProv = ""
        vbCodDpto = ""
        vbDesDist = ""
        vbDesProv = ""
        vbDesDpto = ""
        vbCodgeo = ""
        
    Else
        LLenarDatosUbigeo = True
        vbCodDist = rsUbigeoAfi.Fields("CodDist")
        vbCodProv = rsUbigeoAfi.Fields("Codprov")
        vbCodDpto = rsUbigeoAfi.Fields("CodDpto")
        vbDesDist = rsUbigeoAfi.Fields("Distrito")
        vbDesProv = rsUbigeoAfi.Fields("Provincia")
        vbDesDpto = rsUbigeoAfi.Fields("Departamento")
        vbCodgeo = rsUbigeoAfi.Fields("codgeo")
        'TabSig 13
    End If


'Set rsAfiliado = Nothing
End Function
Function validar() As Boolean
validar = True
If Trim(txtNroSolicitud) = "" Then
    Call cmdNuevo_Click
    
    If Trim(txtNroSolicitud) = "" Then
        validar = False
        txtApePat.SetFocus
        Exit Function
    End If
End If

If Trim(txtRuc.Text) = "" Then

    If MsgBox("¿El Afiliado tiene Empleador ?", vbQuestion + vbYesNo, Caption) = vbNo Then
        vbindEmp2 = False
        
    Else
        vbindEmp2 = True
        'txtRuc.SetFocus
        'Exit Function
    End If
Else
    vbindEmp2 = True
    
End If




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

If Trim(vbCodgeoEmp) = "" Then
    If MsgBox(" No ingresó el Departamento-Provincia-Distrito del empleador ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        validar = False
        ucUbigeoEmp.SetFocus
        Exit Function
    End If
End If

'
'If Not vbindEmp Then
'    If Trim(cboAfp.Text) = "" Then
'        MsgBox "Debe seleccionar la AFP origen", vbCritical, Caption
'        Me.cboAfp.SetFocus
'        Validar = False
'        Exit Function
'    End If
'End If

If txtFechaSol <> "" Then
    If CDate(txtFechaSol.Text) > gDate Then
            MsgBox "La fecha de firma del Contrato no puede ser mayor a " & gDate, vbCritical, Caption
            txtFechaSol.SetFocus
            validar = False
            Exit Function
        End If
End If



If Frame2.Enabled Then

    If Trim(txtTipDoc.Text) = "" Then
        If MsgBox("No ingresó el tipo de documento ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtTipDoc.SetFocus
            Exit Function
        End If
    End If
    
    
    If Trim(txtNumdoc.Text) = "" Then
        If MsgBox("No ingresó el número de documento ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtNumdoc.SetFocus
            Exit Function
        End If
    End If
    

    
    
    If Trim(txtAv.Text) = "" Then
        If MsgBox("No ingresó la Avenida / calle  ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtAv.SetFocus
            Exit Function
        End If
    End If
    
    
    If Trim(vbCodgeo) = "" Then
        If MsgBox("No seleccionó el Departamento-Provincia-Distrito  ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            ucUbigeoAfi.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//" Then
        If MsgBox(" No ingresó la fecha de nacimiento del afiliado ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtFecNac.SetFocus
            Exit Function
        End If
    End If

    If Trim(txtSexo) = "" Then
        If MsgBox(" No ingresó el sexo del afiliado ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtSexo.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(lblLugarNac.Caption) = "" Then
          If MsgBox(" No ingresó codigo de lugar de nacimiento ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            txtUbiNacRe.SetFocus
            Exit Function
        End If
    End If

'    If txtFecTras <> "" Then
'        If CDate(txtFecTras.Text) > gDate Then
'            MsgBox "La fecha de declaración de traspaso no puede ser mayor a " & gDate, vbCritical, Caption
'            txtFecTras.SetFocus
'            Validar = False
'            Exit Function
'        End If
'    End If
    
    Select Case txtTipDoc
    Case "DNI"
        If Len(Trim(txtNumdoc)) <> 8 Then
            MsgBox "El DNI debe tener 8 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
        
    Case "CE"
        If Len(Trim(txtNumdoc)) < 5 Then
            MsgBox "El CE debe tener de 5 a 8 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
    
    Case "CIA"
        If Len(Trim(txtNumdoc)) < 2 And Len(Trim(txtNumdoc)) > 5 Then
            MsgBox "El CE debe tener de 2 a 5 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
    
    Case "CFA"
        If Len(Trim(txtNumdoc)) < 6 And Len(Trim(txtNumdoc)) > 8 Then
            MsgBox "El CE debe tener de 6 a 8 caracteres", 64, Caption
            validar = False
            Exit Function
        End If
    
    Case Else
    
    End Select
    
    
    If cboTipTrabajador.ListIndex = -1 Then
        If MsgBox(" No selecciono el tipo de trabajador ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            cboTipTrabajador.SetFocus
            Exit Function
        End If
    End If
    
    If cboLey19990.ListIndex = -1 Then
        If MsgBox(" No selecciono tipo de ley 19990 ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            cboLey19990.SetFocus
            Exit Function
        End If
    End If
    
    If cboLey25530.ListIndex = -1 Then
        If MsgBox(" No selecciono tipo de ley 20530 ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            cboLey20530.SetFocus
            Exit Function
        End If
    End If
    
    If cboPriEmp.ListIndex = -1 Then
        If MsgBox(" No selecciono Primer Empleo ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            cboPriEmp.SetFocus
            Exit Function
        End If
    End If
    
    If cboTipAfi.ListIndex = -1 Then
        If MsgBox(" No selecciono Tipo de Afiliación ¿Desea Continuar  ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            validar = False
            cboTipAfi.SetFocus
            Exit Function
        End If
    End If
    
    If Left(cboTipAfi.Text, 1) = "N" Then
        If Trim(txtCUSPP.Text) <> "" Then
           MsgBox "Cuspp debe estar en blanco", vbCritical, Caption
           txtCUSPP.SetFocus
           validar = False
           Exit Function
        End If
    End If
    
    If Left(cboTipAfi.Text, 1) = "R" Then
        If Trim(txtCUSPP.Text) = "" Then
           MsgBox "Debe ingresar el Cuspp por ser una Reafiliación", vbCritical, Caption
           txtCUSPP.SetFocus
           Exit Function
        End If
    End If
    
    
'    If Trim(txtCUSPP.Text) = "" Then
'        If MsgBox("No ingresó el Cuspp ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
'            Validar = False
'            txtCUSPP.SetFocus
'            Exit Function
'        End If
'    End If
    
End If



End Function

Private Sub ucUbigeoEmp_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 38, 609
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
KeyAscii = ValidarDigitacion(txtUbiNacRe, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ- " & vbCaracteresNumero)
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

Private Sub txtRuc_GotFocus()
'Set controlf = ActiveControl
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 1089
txtRuc.BackColor = &HC0FFFF
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
                sqltexto = "select razsoc, direcc, a.codgeo, telefono, numero, urba, ANEXO_EL, CODUSUC = isnull(CODUSUC,''), Feccar , Desgeo  from Truc a left join ubigeo b on" & _
                " a.codgeo = b.codgeo where NUMRUC = '" & txtRuc & "'"
                
                rsDatosRuc.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

                If rsDatosRuc.RecordCount > 0 Then
                'Consultar no existe razon en truc
                   'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValRuc = True
                   vbindRuc = True
                   'lblMensaje.Caption = ""
                   If CDate(rsDatosRuc.Fields("Feccar")) = CDate(gDate) Then vbindRuc = False
                   lblRazonSoc.SetFocus
                   lblRazonSoc.Text = IIf(IsNull(rsDatosRuc.Fields("razsoc")), "", Trim(rsDatosRuc.Fields("razsoc")))
                   txtDireEmp = IIf(IsNull(rsDatosRuc.Fields("direcc")), "", Trim(rsDatosRuc.Fields("direcc")))
                   txtNroEmp = IIf(IsNull(rsDatosRuc.Fields("numero")), "", Trim(rsDatosRuc.Fields("numero")))
                   txtUrbaEmp = IIf(IsNull(rsDatosRuc.Fields("urba")), "", Trim(rsDatosRuc.Fields("urba")))
                   txtTeleOfi = IIf(IsNull(rsDatosRuc.Fields("telefono")), "", Trim(rsDatosRuc.Fields("telefono")))
                   txtAnxOfi = IIf(IsNull(rsDatosRuc.Fields("ANEXO_EL")), "", Trim(rsDatosRuc.Fields("ANEXO_EL")))
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
Function ValProm(KeyAscii) As Boolean
Dim vRuc As String

If KeyAscii = 13 Then

If Trim(txtCodProm) = "" Then
         lblMensaje.Caption = "Debe Ingresar Código de Promotor"
         Beep
         ValProm = False
         txtCodProm.SetFocus
         Exit Function
End If
    'Busqueda en la Tabla Magne Si es Pago Magnetico
        'Recupera DAtos de Razon Social
                  
                Dim RsPromo As New ADODB.Recordset
                
                Call Con.AbrirTabla("select despromo = isnull(desPromo,'Sin Descripción'), CodProAd = isnull(CodProAd,'') from tpromotor where codpromo = '" & Trim(txtCodProm.Text) & "'", RsPromo, adOpenStatic, adLockReadOnly)
                                        
                
                With RsPromo
                  If .RecordCount > 0 Then
                    vRuc = .Fields("despromo")
                    vbcodAdm = .Fields("CodProAd")
                  End If
                End With
                      
                Set RsPromo = Nothing
                
                If vRuc <> "" Then
                'Consultar no existe razon en truc
                'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValProm = True
                   vbindPro = True
                   lblMensaje.Caption = ""
                   lblPromotor = " " & vRuc
                   'If chkBolPag.Enabled Then chkBolPag.SetFocus
                   lblPromotor.SetFocus
                Else
                   lblMensaje.Caption = "Código de Promotor no se encuentra en la Tabla TPromotor"
                   ValProm = False
                   vbindPro = False
                   txtCodProm.SetFocus
                   vbcodAdm = ""
                   Beep
                End If
                              
End If
 If KeyAscii = vbKeyF9 Then
   'Fuerza el pase al sgte campo
   ValProm = True
   lblMensaje.Caption = ""
   vbindPro = False
   
   txtFechaSol.SetFocus
 End If

End Function


Function ValNroSol(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto ingreso Obligatorio

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
  If Not vbNroSol Then
    If Trim(txtNroSolicitud) = "" Then
             lblMensaje.Caption = "Debe ingresar Nro de Solicitud "
             ValNroSol = False
             txtNroSolicitud.SetFocus  'TabSig 13
             Exit Function
    End If
  End If

    If vbNroSol Then
        If Trim(txtNroSolicitud.Text) <> "" Then
            vbNroSol2 = True  '2da digitacion
        End If
        If Comparar(vtmpNroSol, txtNroSolicitud.Text, True) Then
        'Pasa al sgte campo
         lblMensaje.Caption = ""
         ValNroSol = True
         'txtVolcFin.SetFocus
         'Call BuscarSolicitud
         cmdGrabar.SetFocus
         
         Exit Function
         
         'TabSig KeyAscii
        Else
         lblMensaje.Caption = " Valores no coinciden"
         Beep
         ValNroSol = False
         txtNroSolicitud.SetFocus
        End If
    Else
      vbNroSol = True
      vtmpNroSol = txtNroSolicitud.Text
      txtNroSolicitud.Text = ""
      txtNroSolicitud.SetFocus
      'Verifico si es de una fecha distinta
      If fgDatosql("select numsol from  solicitud where numsol =  '" & vtmpNroSol & "' and  convert(char,fecsal,103) <> '" & txtFecsal & "'") <> "" Then
        If MsgBox("El número de solicitud " & vtmpNroSol & " ya fue digitado, ¿ Desea continuar ?", vbQuestion + vbYesNo, Caption) = vbNo Then
            Unload Me
        End If
      End If
      
    End If
 End If
 
 If KeyAscii = vbKeyF9 Then
   If vbNroSol2 Then '2da Digitacion
   'Fuerza el pase al sgte campo
    ValNroSol = True
    lblMensaje.Caption = ""
    
    txtApePat.SetFocus
    'TabSig 13
   Else
    ValNroSol = False
    lblMensaje = "Debe digitar 2 veces el campo"
    Beep
   End If
 End If

End Function

Sub CargarDatosSolicitud()
Dim Strsqltexto As String

Strsqltexto = "Select a.CodAfp ,Secu, a.FecSal , a.Lote , a.NumSol , a.ApePat , a.ApeMat, a.Nombre1 , a.Nombre2  , " & _
" a.TipoDoc , a.NumDoc, CUSPP, a.Direccion , a.Numero, a.Urb, a.Distrito, a.Provincia,  a.Departamento, " & _
" a.Ubigeo, a.Telefono,  a.Empleadores ,  Cantemple ,  CodPromo ,  BolPago , PagoUltApo , CopiaSim , " & _
" OtroDoc , DescOtroDoc , FecRegPro ,  CodAFPOri , TipoFonOri , TipoFonDes , FecFirmSol ,  Aceptacion , " & _
" Texto , a.RIAafil ,  a.FecNac , Sexo  , e_mail, c.Status, a.IDimagen,  a.Codusu , a.fecdig, a.Horini, a.Horfin, " & _
" a.Numpul, a.CodMod, a.Fecmod, b.desgeo,  a.Horinim , a.Horfinm, imagen, d.Ruta, LoteFisico, Secuencial, NumRuc, " & _
" eDireccion = e.Direccion, eNumero = e.Numero, eUrb = e.Urb , eDistrito = f.Distrito, eProvincia = f.Provincia, " & _
" eDepartamento = f.Departamento, eUbigeo = e.Ubigeo, eTelefono = e.Telefono, eRIAafil = e.RIAafil , eDesgeo = f.desgeo, a.posicion " & _
" from solicitud a left join ubigeo b on a.ubigeo = b.codgeo " & _
" join _Imagenes c on c.idImagen = a.idImagen and  a.posicion = c.posicion " & _
" join _Volumenes d on d.idvolumen = c.idvolumen " & _
" left join Empleadores e on e.codafp = a.codafp and e.numsol = a.numsol and e.Secuencial = a.secu " & _
" left join Ubigeo f on E.ubigeo = F.codgeo " & _
" where a.CodAfp = '" & vbGAFP & "' and a.Lote = ' " & vbGLte & " '"

If rsSolicitud.State = 1 Then rsSolicitud.Close
rsSolicitud.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
 
With rsSolicitud

If (.EOF And .BOF) Then
     'Call CargarDatosEmpleadores
     MsgBox "No existen Registros para este Lote", 64, Caption
End If

End With




End Sub

'Sub CargarDatosEmpleadores()
'
'
'Dim Strsqltexto As String
'
'Strsqltexto = "SELECT CodAfp, NumSol, Secuencial, NumRuc, eDireccion = Direccion, eNumero = Numero," & _
'              "eUrb = Urb , eDistrito = a.Distrito, eProvincia = a.Provincia, eDepartamento = a.Departamento, eUbigeo = a.Ubigeo, eTelefono = f.Telefono, eRIAafil = f.RIAafil , eDesgeo = desgeo " & _
'              "from empleadores a join Ubigeo b on a.ubigeo = b.codgeo   where CodAfp = '" & vbGAFP & "'  and   NumSol = '" & rsSolicitud.Fields("Numsol") & "' "
'
'rsEmpleadores.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
'
'
'End Sub


'Sub RefrescarEmp()
'With rsEmpleadores
'
'
'If Not (.EOF And .BOF) Then
'     txtRuc = IIf(IsNull(.Fields("NumRuc")), "", .Fields("NumRuc"))
'     txtDireEmp = IIf(IsNull(.Fields("Direccion")), "", .Fields("Direccion"))
'     txtNroEmp = IIf(IsNull(.Fields("Numero")), "", .Fields("Numero"))
'     txtUrbaEmp = IIf(IsNull(.Fields("Urb")), "", .Fields("Urb"))
'     ucUbigeoEmp.cTexto = IIf(IsNull(.Fields("Desgeo")), "", .Fields("Desgeo"))
'     txtTelefEmp = IIf(IsNull(.Fields("Telefono")), "", .Fields("Telefono"))
'     txtRIAAfil = IIf(IsNull(.Fields("RIAafil")), 0, .Fields("RIAafil")) '
'
'Else
'
'    txtRuc = ""
'     txtDireEmp = ""
'     txtNroEmp = ""
'     txtUrbaEmp = ""
'     ucUbigeoEmp.cTexto = ""
'     txtTelefEmp = ""
'     txtRIAAfil = ""
'End If
'End With
'
'
'End Sub


Sub RefrescarSol(vbResImag As Boolean)

With rsSolicitud
If Not (.EOF And .BOF) Then
        txtFecsal = IIf(IsNull(.Fields("Fecsal")), "", .Fields("Fecsal"))
        txtLote = IIf(IsNull(.Fields("Lote")), "", .Fields("Lote"))
        Me.txtNroSolicitud = IIf(IsNull(.Fields("Numsol")), "", .Fields("Numsol"))
        txtApePat.Text = IIf(IsNull(.Fields("Apepat")), "", .Fields("Apepat"))
        txtApeMat.Text = IIf(IsNull(.Fields("Apemat")), "", .Fields("Apemat"))
        txtNombreuno = IIf(IsNull(.Fields("Nombre1")), "", .Fields("Nombre1"))
        txtNombredos = IIf(IsNull(.Fields("Nombre2")), "", .Fields("Nombre2"))
        txtTipDoc = IIf(IsNull(.Fields("TipoDoc")), "", .Fields("TipoDoc"))
        txtNumdoc = IIf(IsNull(.Fields("Numdoc")), "", .Fields("Numdoc"))
        txtCUSPP = IIf(IsNull(.Fields("CUSPP")), "", .Fields("CUSPP"))
        txtAv = IIf(IsNull(.Fields("Direccion")), "", .Fields("Direccion"))
        txtNroDir = IIf(IsNull(.Fields("Numero")), "", .Fields("Numero"))
        txtUrba = IIf(IsNull(.Fields("Urb")), "", .Fields("Urb"))
        ucUbigeoAfi.cTexto = IIf(IsNull(.Fields("Desgeo")), "", .Fields("Desgeo") & " / " & .Fields("Ubigeo"))
        txtTelefAfi = IIf(IsNull(.Fields("Telefono")), "", .Fields("Telefono"))
        
        '-*------------------Datos de Empleadores --------------
         txtRuc = IIf(IsNull(.Fields("NumRuc")), "", .Fields("NumRuc"))
         txtDireEmp = IIf(IsNull(.Fields("eDireccion")), "", .Fields("eDireccion"))
         txtNroEmp = IIf(IsNull(.Fields("eNumero")), "", .Fields("eNumero"))
         txtUrbaEmp = IIf(IsNull(.Fields("eUrb")), "", .Fields("eUrb"))
         ucUbigeoEmp.cTexto = IIf(IsNull(.Fields("eDesgeo")), "", .Fields("eDesgeo") & " / " & .Fields("eUbigeo"))
         txtTelefEmp = IIf(IsNull(.Fields("eTelefono")), "", .Fields("eTelefono"))
         txtRIAAfil = IIf(IsNull(.Fields("eRIAafil")), 0, .Fields("eRIAafil")) '
        '----------------------
                  
        chkBolPag.Value = IIf(IsNull(.Fields("BolPago")) Or Trim(.Fields("BolPago")) = "", 0, 1)
        chkPagUlt.Value = IIf(IsNull(.Fields("PagoUltApo")) Or Trim(.Fields("PagoUltApo")) = "", 0, 1)
        chkCopSim.Value = IIf(IsNull(.Fields("CopiaSim")) Or Trim(.Fields("CopiaSim")) = "", 0, 1)
        chkOtroDoc.Value = IIf(IsNull(.Fields("OtroDoc")) Or Trim(.Fields("OtroDoc")) = "", 0, 1)
        txtEspecificar = IIf(IsNull(.Fields("DescOtroDoc")), "", .Fields("DescOtroDoc"))
        txtFechaSol = IIf(IsNull(.Fields("FecRegPro")), "", .Fields("FecRegPro"))
                        
        'IzqBuscarEnCombo cboAfp, IIf(IsNull(.Fields("CodAFPOri")), "", Trim(.Fields("CodAFPOri"))), 2
        
        txtFecTras = IIf(IsNull(.Fields("FecFirmSol")), "", .Fields("FecFirmSol"))
        chkProc.Value = IIf(IsNull(.Fields("Aceptacion")) Or .Fields("Aceptacion") = "" Or .Fields("Aceptacion") = "N", 0, 1)
        txtMotiv = IIf(IsNull(.Fields("Texto")), "", .Fields("Texto"))
        txtFecNac = IIf(IsNull(.Fields("FecNac")), "", .Fields("FecNac"))
        txtRIAAfil = IIf(IsNull(.Fields("RIAafil")), 0, .Fields("RIAafil")) '
        txtSexo = IIf(IsNull(.Fields("Sexo")), 0, .Fields("Sexo")) '
        txtemail = IIf(IsNull(.Fields("e_mail")), 0, .Fields("e_mail")) '
        txtFTipoO = IIf(IsNull(.Fields("TipoFonOri")), 0, .Fields("TipoFonOri")) '
        txtFTipoD = IIf(IsNull(.Fields("TipoFonDes")), 0, .Fields("TipoFonDes")) '
        
        
        'Visualizar Imagen correspondiente
        If vbResImag Then
            ImgEdit1.Image = Rutaimag(.Fields("Ruta"), .Fields("LoteFisico"), .Fields("Imagen"))
            ImgEdit1.Display
        End If
        
        
        ImgEdit2.Image = ImgEdit1.Image
        ImgEdit2.FitTo 0
        ImgEdit2.Display
        'txtNroSolicitud.SetFocus
Else
        txtApePat.Text = ""
        txtApeMat.Text = ""
        txtNombreuno = ""
        txtNombredos = ""
        txtTipDoc = ""
        txtNumdoc = ""
        txtCUSPP = ""
        txtAv = ""
        txtNroDir = ""
        txtUrba = ""
        ucUbigeoAfi.cTexto = ""
        txtTelefAfi = ""
        
        chkBolPag = 0
        chkPagUlt = 0
        chkCopSim = 0
        chkOtroDoc = 0
        txtEspecificar = ""
        txtFechaSol = ""
        cboAfp.ListIndex = -1
        txtFecTras = ""
        'chkProc.Value = 0
        txtMotiv = ""
        txtFecNac = ""
        txtRIAAfil = ""
        txtemail = ""
      
        
End If

End With


End Sub


Private Sub ucUbigeoEmp_Validate(Cancel As Boolean)
If MsgBox("¿Verificó Distrito, Provincia y Departamento contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
    Cancel = False
Else
    Cancel = True
End If

End Sub

Private Sub UpDown1_DownClick()
On Error Resume Next
ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY + txtFactor
End Sub

Private Sub UpDown1_UpClick()
On Error Resume Next
ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY + txtFactor
End Sub

Sub ImagenesAlternas(nomimag As String)
   With rsImagAlt
    sqltexto = "select top 4 Ruta, LoteFisico, Imagen from _imagenes a join " & _
    " _Volumenes b on a.IdVolumen = b.IdVolumen where  imagen > '" & nomimag & "' order by imagen"
    If .State = 1 Then .Close
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    Cont = 0
    While Not .EOF
        ImgEdit3(Cont).Image = Rutaimag(RsImag.Fields("Ruta"), .Fields("LoteFisico"), .Fields("imagen"))
        ImgEdit3(Cont).Display
        ImgEdit3(Cont).FitTo 0
        Cont = Cont + 1
        .MoveNext
    Wend
               
   End With
   
End Sub

Sub RevisarTipFondo()
On Error Resume Next
vbtmpFTipoO = txtFTipoO.Text
txtFTipoO.Text = ""

vbtmpFTipoD = txtFTipoD.Text
txtFTipoD.Text = ""
vbdigTipFondo = True
txtFTipoO.SetFocus
End Sub

Function ValLugarNac(KeyAscii) As Boolean
Dim vCodGeo As String

If KeyAscii = 13 Then

If Trim(txtUbiNacRe) = "" Then
         lblMensaje.Caption = "Debe Ingresar Código de Reniec"
         Beep
         ValLugarNac = False
         txtUbiNacRe.SetFocus
         Exit Function
End If
                  
                vCodGeo = fgDatosql("select codubic from ubigeo where  codgeo  = '" & Trim(txtUbiNacRe.Text) & "'")
                If vCodGeo <> "" Then
                'Consultar no existe razon en truc
                'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValLugarNac = True
                   vbindLugNac = True
                   lblMensaje.Caption = ""
                   lblLugarNac = vCodGeo
                   txtFecNac.SetFocus
                   
                Else
                   lblMensaje.Caption = "Código de Lugar Reniec no se encuentra en la Tabla Ubigeo"
                   ValLugarNac = False
                   vbindLugNac = False
                   txtUbiNacRe.SetFocus
                   Beep
                End If
                              
End If
 If KeyAscii = vbKeyF9 Then
   'Fuerza el pase al sgte campo
    ValLugarNac = True
    lblMensaje.Caption = ""
    vbindLugNac = False

    txtFecNac.SetFocus
 End If

End Function

Function ObtenerSecuencial() As Double
Dim cmdSec As ADODB.Command
Set cmdSec = New ADODB.Command

With cmdSec
    .ActiveConnection = Con.Cnx
    .CommandText = " exec pGetSecuAfi"
    
    If Con.GetFEjecutarSQL(cmdSec) Then
        If Con.vbdev > 0 Then
            ObtenerSecuencial = Con.vbdev
        End If
    End If
    
End With



End Function
Sub NuevoContrato()

Call Inicializar

txtNroSolicitud = "C" & Format(ObtenerSecuencial, "0000000#")
txtApePat.SetFocus

End Sub


Sub RecuperacionAntRegistro()
On Error GoTo ERRORES
Dim rsDatos As ADODB.Recordset
Set rsDatos = New ADODB.Recordset


Dim sqltexto As String
sqltexto = " select a.Numsol, ApePat=isnull(a.apepat,''), ApeMat=isnull(a.ApeMat,'') , Nombre1= isnull(a.Nombre1,''), Nombre2 = isnull(a.Nombre2,''), " & _
           " b.NumRuc, c.Razsoc, " & _
           " Direccion = isnull(b.Direccion,''), Numero = isnull(b.Numero,''), Urb = isnull(b.Urb,''), Ubigeo = isnull(d.DesGeo,'') + '/' + isnull(d.CodGeo,''), " & _
           " DireccionCR = isnull(a.Direccion,''), NumeroCR = isnull(a.Numero,''), UrbCR = isnull(a.Urb,''), UbigeoCR = isnull(e.DesGeo,'') + '/' + isnull(e.CodGeo,'') " & _
           " from solicitud a join empleadores b on " & _
           " a.NumSol = b.NumSol and a.IndDoc = b.IndDoc and  a.secu = b.Secuencial and " & _
           " a.Fecsal = b.Fecsal join truc c on b.numruc = c.numruc left join ubigeo d " & _
           " on b.Ubigeo = d.codgeo join Ubigeo e on a.ubigeo = e.codgeo  Where a.Secu = 0 And a.Lote = '" & txtLote & "' " & _
           " and NumSec = (select max(numsec) from solicitud where lote = '" & txtLote & "' and secu = 0) "
                   
           rsDatos.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

With rsDatos


    If .RecordCount > 0 Then
        vbNumsolant = Trim(.Fields("Numsol"))
        vbNomcUlt = Trim(.Fields("ApePat")) & " " & Trim(.Fields("ApeMat")) & " " & Trim(.Fields("Nombre1")) & " " & Trim(.Fields("Nombre2"))
        vbRucUlt = Trim(.Fields("NumRuc"))
        vbRazSocUlt = Trim(.Fields("Razsoc"))
        vbAveUlt = Trim(.Fields("Direccion"))
        vbNroUlt = Trim(.Fields("Numero"))
        vbUrbaUlt = Trim(.Fields("Urb"))
        vbUbigeoUlt = Trim(.Fields("Ubigeo"))
        
        vbAveUltCR = Trim(.Fields("DireccionCR"))
        vbNroUltCR = Trim(.Fields("NumeroCR"))
        vbUrbaUltCR = Trim(.Fields("UrbCR"))
        vbUbigeoUltCR = Trim(.Fields("UbigeoCR"))
        
    End If

End With

Set rsDatos = Nothing
Exit Sub
ERRORES:
MsgBox Err.Description, vbCritical, Caption
End Sub

Sub ActIndFedatario()

Dim cmdFed As Command
Set cmdFed = New Command

Dim rsDatosProm As Recordset
Set rsDatosProm = New Recordset

Dim indFed As String

'Si Promotor NO existe

If Trim(txtCodProm) <> "" Then

  rsDatosProm.Open "select codpromo,tipsanc = isnull(tipsanc,'') from tpromotor where codpromo = '" & txtCodProm & "'", Con.Cnx, adOpenStatic, adLockReadOnly
  
  If rsDatosProm.RecordCount > 0 Then
     If Trim(rsDatosProm.Fields("TipSanc")) = "" Then
         indFed = "1"
     Else
         indFed = "0"
     End If
  Else
     indFed = "0"
  End If
Else
  indFed = "0"
End If


With cmdFed
    .ActiveConnection = Con.Cnx
    .CommandText = "  Update solicitud set indFirFed = '" & indFed & "', indRucN = '" & IIf(vbindRuc, 0, 1) & "', IndCSBSErr = '" & IIf(chkCodErr.Value = 1 Or vbindPro = False, "1", "0") & "', IndCamFec = '" & IIf(indCambioFec, 1, 0) & "', IndFecSus = '" & IIf(indFecSus, "1", "0") & "' " & _
                   "  where numsol = '" & txtNroSolicitud & "' and Fecsal = '" & txtFecsal & "' "
    .Execute
End With

'vbindPro


Set cmdFed = Nothing
Set rsDatosProm = Nothing


End Sub

Sub VeriFecSus()
indCambioFec = False
indFecSus = False
If (txtFechaSol.Text = "" Or txtFechaSol.Text = "//") Then
    Fechatmp = fgDatosql("Select Ult_Dia from DiaUtil where FecProce = '" & gDate & "' ")
   If Fechatmp = "" Then
     txtFechaSol.Text = gDate
   Else
     txtFechaSol.Text = Fechatmp
   End If
   indCambFec = True
   indFecSus = True
End If


End Sub

Function ValFecNac(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto ingreso Obligatorio

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
  If Not vbFecnac Then
    If Trim(txtFecNac) = "" Then
             lblMensaje.Caption = "Debe ingresar Fecha de Nacimiento "
             ValFecNac = False
             txtFecNac.SetFocus  'TabSig 13
             Exit Function
    End If
  End If

    If vbFecnac Then
        If Trim(txtFecNac.Text) <> "" Then
            vbFecnac2 = True  '2da digitacion
        End If
        If Comparar(vtmpFecnac, txtFecNac.Text, True) Then
         'Pasa al sgte campo
         lblMensaje.Caption = ""
         ValFecNac = True
         txtNumdoc.SetFocus
         Exit Function
                  
        Else
         lblMensaje.Caption = " Valores no coinciden"
         Beep
         ValFecNac = False
         txtFecNac.SetFocus
        End If
    Else
      vbFecnac = True
      vtmpFecnac = txtFecNac.Text
      txtFecNac.Text = ""
      ValFecNac = True
      txtNumdoc.SetFocus
    End If
 End If
 
 If KeyAscii = vbKeyF9 Then
   If vbFecnac2 Then '2da Digitacion
     'Fuerza el pase al sgte campo
      ValFecNac = True
      lblMensaje.Caption = ""
      txtNumdoc.SetFocus
   Else
      ValFecNac = False
      lblMensaje = "Debe digitar 2 veces el campo"
      Beep
   End If
 End If

End Function

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmTraspasosA 
   Caption         =   "Traspasos v2"
   ClientHeight    =   11010
   ClientLeft      =   585
   ClientTop       =   1050
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   120
      TabIndex        =   77
      Top             =   -120
      Width           =   13575
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
         Left            =   7800
         TabIndex        =   108
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtNroSolicitud 
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
         MaxLength       =   7
         TabIndex        =   2
         Top             =   200
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
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   1
         Top             =   200
         Width           =   1095
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
         TabIndex        =   0
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Traspasos"
         Height          =   255
         Left            =   9960
         TabIndex        =   105
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Total"
         Height          =   255
         Left            =   11880
         TabIndex        =   104
         Top             =   240
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
         Left            =   10800
         TabIndex        =   103
         Top             =   240
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
         Left            =   12360
         TabIndex        =   102
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Solicitud"
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
         Left            =   4800
         TabIndex        =   80
         Top             =   240
         Width           =   1215
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
         Left            =   2520
         TabIndex        =   79
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
         Left            =   360
         TabIndex        =   78
         Top             =   240
         Width           =   855
      End
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   1695
      Left            =   120
      TabIndex        =   85
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
      UndoBufferSize  =   213970432
      OcrZoneVisibility=   -3864
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3864
      sReserved2      =   -3864
      lReserved1      =   36
      lReserved2      =   36
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "cmdAnterior"
      Height          =   9015
      Left            =   -120
      TabIndex        =   48
      Top             =   360
      Width           =   15255
      Begin VB.Frame Frame7 
         Caption         =   "Sello"
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
         Height          =   1215
         Left            =   4920
         TabIndex        =   73
         Top             =   6840
         Width           =   6375
         Begin VB.TextBox txtemail2 
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
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1920
            Width           =   4095
         End
         Begin VB.CheckBox chkBol 
            Caption         =   "Boletin Electrónico"
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
            Left            =   120
            TabIndex        =   37
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox chkCuenta 
            Caption         =   "Estado de Cuenta"
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
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtemail 
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
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1440
            Width           =   4095
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
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   34
            Top             =   240
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
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   32
            Top             =   240
            Width           =   1695
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
            Left            =   3480
            MaxLength       =   15
            TabIndex        =   33
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "E - mail"
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
            Left            =   2400
            TabIndex        =   101
            Top             =   1320
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
            Left            =   5280
            TabIndex        =   76
            Top             =   240
            Width           =   495
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
            Left            =   3000
            TabIndex        =   75
            Top             =   240
            Width           =   375
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
            Left            =   240
            TabIndex        =   74
            Top             =   240
            Width           =   975
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
         Height          =   615
         Left            =   240
         TabIndex        =   116
         Top             =   5400
         Width           =   13575
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
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   117
            Top             =   200
            Width           =   1335
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Cod Promotor"
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
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdSuelto 
         Caption         =   "&Doc sin Enlc"
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
         Left            =   9480
         TabIndex        =   44
         Top             =   8120
         Width           =   1335
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
         Height          =   375
         Left            =   8160
         TabIndex        =   43
         Top             =   8120
         Width           =   1335
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
         Height          =   375
         Left            =   6840
         TabIndex        =   42
         Top             =   8120
         Width           =   1335
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
         Height          =   375
         Left            =   5520
         TabIndex        =   41
         Top             =   8120
         Width           =   1335
      End
      Begin VB.CommandButton cmdrotDer 
         Caption         =   ">"
         Height          =   495
         Left            =   11835
         TabIndex        =   107
         Top             =   7200
         Width           =   495
      End
      Begin VB.CommandButton cmdrotIzq 
         Caption         =   "<"
         Height          =   495
         Left            =   11355
         TabIndex        =   106
         Top             =   7200
         Width           =   495
      End
      Begin Traspasos.UserControl1 ucUbigeoAfi 
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   2160
         Width           =   6255
         _extentx        =   13150
         _extenty        =   661
      End
      Begin Traspasos.UserControl1 ucUbigeoEmp 
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   4920
         Width           =   8055
         _extentx        =   14208
         _extenty        =   661
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   3
         Left            =   13920
         TabIndex        =   99
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
         UndoBufferSize  =   213969408
         OcrZoneVisibility=   -3896
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3896
         sReserved2      =   -3896
         lReserved1      =   36
         lReserved2      =   36
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   2
         Left            =   13920
         TabIndex        =   98
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
         UndoBufferSize  =   213969408
         OcrZoneVisibility=   -3896
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3896
         sReserved2      =   -3896
         lReserved1      =   36
         lReserved2      =   36
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   1
         Left            =   13920
         TabIndex        =   97
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
         UndoBufferSize  =   213969408
         OcrZoneVisibility=   -3896
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3896
         sReserved2      =   -3896
         lReserved1      =   36
         lReserved2      =   36
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit3 
         Height          =   1095
         Index           =   0
         Left            =   13920
         TabIndex        =   96
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
         UndoBufferSize  =   213969408
         OcrZoneVisibility=   -3896
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3896
         sReserved2      =   -3896
         lReserved1      =   36
         lReserved2      =   36
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   14415
         TabIndex        =   94
         Top             =   480
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "txtFactor"
         BuddyDispid     =   196631
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
         Left            =   14040
         TabIndex        =   93
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
         Height          =   375
         Left            =   1560
         TabIndex        =   92
         Top             =   8120
         Width           =   1215
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
         Height          =   375
         Left            =   240
         TabIndex        =   91
         Top             =   8120
         Width           =   1215
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
         Height          =   375
         Left            =   10920
         TabIndex        =   45
         Top             =   8120
         Width           =   1335
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit2 
         Height          =   1815
         Left            =   12480
         TabIndex        =   86
         Top             =   6960
         Width           =   2775
         _Version        =   131074
         _ExtentX        =   4895
         _ExtentY        =   3201
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   213964288
         OcrZoneVisibility=   -3896
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3896
         sReserved2      =   -3896
         lReserved1      =   36
         lReserved2      =   36
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
         Height          =   375
         Left            =   4200
         TabIndex        =   40
         Top             =   8120
         Width           =   1335
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
         Height          =   375
         Left            =   2880
         TabIndex        =   39
         Top             =   8120
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Declaración de Traspaso"
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
         Height          =   1215
         Left            =   240
         TabIndex        =   69
         Top             =   6840
         Width           =   4680
         Begin VB.TextBox txtFTipoD 
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
            Height          =   315
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   30
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtFTipoO 
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
            Height          =   315
            Left            =   840
            MaxLength       =   1
            TabIndex        =   29
            Top             =   840
            Width           =   375
         End
         Begin VB.ComboBox cboAfp 
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
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtFecTras 
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
            Left            =   3360
            MaxLength       =   15
            TabIndex        =   31
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "F.Tipo D."
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
            Left            =   1395
            TabIndex        =   110
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "F.Tipo O."
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
            Left            =   110
            TabIndex        =   109
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "AFP"
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
            TabIndex        =   71
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label15 
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
            Left            =   2760
            TabIndex        =   70
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Verificación de la Solicitud"
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
         TabIndex        =   67
         Top             =   6000
         Width           =   13575
         Begin VB.TextBox txtEspecificar 
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
            Left            =   6360
            MaxLength       =   15
            TabIndex        =   26
            Top             =   480
            Width           =   4095
         End
         Begin VB.CheckBox chkPagUlt 
            Caption         =   "Pago de último aporte exigido"
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
            TabIndex        =   23
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox chkCopSim 
            Caption         =   "Copia simple del documento de identidad"
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
            TabIndex        =   24
            Top             =   240
            Width           =   3855
         End
         Begin VB.CheckBox chkOtroDoc 
            Caption         =   "Otros (Especificar)"
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
            TabIndex        =   25
            Top             =   480
            Width           =   1935
         End
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
            Left            =   11400
            MaxLength       =   15
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkBolPag 
            Caption         =   "Comprobación de Boletas"
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
            TabIndex        =   22
            Top             =   240
            Width           =   3015
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
            Left            =   10680
            TabIndex        =   68
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Resultado de la Solicitud"
         Enabled         =   0   'False
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
         Height          =   375
         Left            =   4800
         TabIndex        =   72
         Top             =   6240
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox txtMotiv 
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
            MaxLength       =   15
            TabIndex        =   47
            Top             =   840
            Width           =   4335
         End
         Begin VB.CheckBox chkProc 
            Caption         =   "Procede"
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
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   1695
         End
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
         Height          =   2535
         Left            =   240
         TabIndex        =   49
         Top             =   120
         Width           =   13575
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
            Left            =   11160
            MaxLength       =   10
            TabIndex        =   14
            Top             =   2040
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
            Left            =   9600
            MaxLength       =   25
            TabIndex        =   12
            Top             =   1680
            Width           =   3615
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
            Left            =   5040
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1680
            Width           =   4335
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
            Left            =   240
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1680
            Width           =   4215
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
            TabIndex        =   7
            Top             =   1080
            Width           =   1575
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
            Left            =   3240
            MaxLength       =   8
            TabIndex        =   8
            Top             =   1080
            Width           =   2535
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
            MaxLength       =   20
            TabIndex        =   3
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
            Left            =   6120
            MaxLength       =   12
            TabIndex        =   9
            Top             =   1080
            Width           =   2535
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
            Left            =   3240
            MaxLength       =   20
            TabIndex        =   4
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
            Left            =   6120
            MaxLength       =   10
            TabIndex        =   5
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
            Left            =   9000
            MaxLength       =   10
            TabIndex        =   6
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label lblY 
            Caption         =   "lblY"
            Height          =   255
            Left            =   10200
            TabIndex        =   120
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblx 
            Caption         =   "lblX"
            Height          =   255
            Left            =   9360
            TabIndex        =   119
            Top             =   960
            Width           =   615
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
            TabIndex        =   82
            Top             =   2160
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
            Left            =   10200
            TabIndex        =   81
            Top             =   2160
            Width           =   855
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
            Left            =   9600
            TabIndex        =   63
            Top             =   1440
            Width           =   4815
         End
         Begin VB.Label Label11 
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
            Left            =   5040
            TabIndex        =   62
            Top             =   1440
            Width           =   2415
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
            Left            =   240
            TabIndex        =   61
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Documento de Identidad"
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
            Left            =   3240
            TabIndex        =   56
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Documento de Identidad"
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
            TabIndex        =   55
            Top             =   840
            Width           =   2775
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
            Left            =   6240
            TabIndex        =   54
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
            TabIndex        =   53
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
            Left            =   3360
            TabIndex        =   52
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
            Left            =   6240
            TabIndex        =   51
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
            Left            =   9000
            TabIndex        =   50
            Top             =   240
            Width           =   2055
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
         Height          =   2775
         Left            =   240
         TabIndex        =   57
         Top             =   2640
         Width           =   13575
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "Eliminar"
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
            Left            =   11280
            TabIndex        =   115
            Top             =   1000
            Width           =   1215
         End
         Begin VB.CommandButton cmdGrabaEmp 
            Caption         =   "Grabar"
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
            Left            =   9960
            TabIndex        =   114
            Top             =   1000
            Width           =   1215
         End
         Begin VB.CommandButton cmdModificar 
            Caption         =   "Modificar"
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
            Left            =   8640
            TabIndex        =   113
            Top             =   1000
            Width           =   1215
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "Nuevo"
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
            Left            =   7320
            TabIndex        =   112
            Top             =   1000
            Width           =   1215
         End
         Begin VB.TextBox txtRiaEmp 
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
            Left            =   9960
            MaxLength       =   15
            TabIndex        =   21
            Top             =   2280
            Width           =   1575
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
            Left            =   8400
            MaxLength       =   10
            TabIndex        =   20
            Top             =   2280
            Width           =   1455
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
            Left            =   9960
            MaxLength       =   25
            TabIndex        =   18
            Top             =   1680
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
            Left            =   240
            MaxLength       =   11
            TabIndex        =   15
            Top             =   1680
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
            Height          =   315
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   17
            Top             =   1680
            Width           =   3135
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
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1680
            Width           =   3495
         End
         Begin MSDataGridLib.DataGrid dtgEmpleador 
            Height          =   735
            Left            =   240
            TabIndex        =   121
            Top             =   240
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   1296
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   17
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   "Nro Ruc"
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
               DataField       =   ""
               Caption         =   "Avenida"
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
               DataField       =   ""
               Caption         =   "Nro"
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
               DataField       =   ""
               Caption         =   "Urba"
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
               DataField       =   ""
               Caption         =   "Departamento / Provincia / Distrito"
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
               DataField       =   ""
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
            BeginProperty Column06 
               DataField       =   ""
               Caption         =   "Codigo Promotor"
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
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1725.165
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1874.835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   3284.788
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column06 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label31 
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
            Left            =   9960
            TabIndex        =   111
            Top             =   2040
            Width           =   375
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
            TabIndex        =   83
            Top             =   2040
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
            Left            =   8400
            TabIndex        =   66
            Top             =   2040
            Width           =   1455
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
            Left            =   10920
            TabIndex        =   65
            Top             =   1440
            Width           =   4575
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
            Left            =   2760
            TabIndex        =   64
            Top             =   1440
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
            Left            =   360
            TabIndex        =   60
            Top             =   1440
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
            TabIndex        =   59
            Top             =   1920
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
            Left            =   6600
            TabIndex        =   58
            Top             =   1440
            Width           =   2775
         End
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   12480
         TabIndex        =   90
         Text            =   "Text4"
         Top             =   4200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   12480
         TabIndex        =   89
         Text            =   "Text3"
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   12480
         TabIndex        =   88
         Text            =   "Text2"
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   12360
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdEnlazarBol 
         Caption         =   "Enl. &Bol."
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
         Height          =   495
         Left            =   11160
         TabIndex        =   95
         Top             =   4680
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   14040
         TabIndex        =   100
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblMensaje 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5400
         TabIndex        =   84
         Top             =   8540
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmTraspasosA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vbAccion As String
Dim vbmarca As Variant
Dim rsCtrldoc As New Recordset
Dim rsSolicitud As New Recordset
'Dim RsEmpleadores As New Recordset
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

Private rsEmpleadores As ADODB.Recordset
Private strAccionEmp As String


Sub GrabarTraspaso()
If Not Validar Then Exit Sub
 
    Set comando = New Command
    With comando
        Con.Cnx.BeginTrans
        .ActiveConnection = Con.Cnx
        .CommandText = " sp_GrabarSolicitud2 ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,?,?,?,?, " & _
                                          " ?,?,?,?,?,?,? "
                                                                            
        .Parameters("@CodAfp") = vbGAFP
        .Parameters("@FecSal") = Trim(txtFecsal)
        .Parameters("@Lote") = Trim(txtLote)
        .Parameters("@NumSol") = Trim(txtNroSolicitud)
        .Parameters("@ApePat") = Trim(txtApePat)
        .Parameters("@ApeMat") = Trim(txtApeMat)
        .Parameters("@Nombre1") = Trim(txtNombreuno)
        .Parameters("@Nombre2") = Trim(txtNombredos)
        .Parameters("@TipoDoc") = Trim(txtTipDoc)
        .Parameters("@NumDoc") = Trim(txtNumDoc)
        .Parameters("@CUSPP") = Trim(txtCUSPP)
        .Parameters("@Direccion") = Trim(txtAv)
        .Parameters("@Numero") = Trim(txtNroDir)
        .Parameters("@Urb") = Trim(txtUrba)
        .Parameters("@Distrito") = vbDesDist
        .Parameters("@Provincia") = vbDesProv
        .Parameters("@Departamento") = vbDesDpto
        .Parameters("@Ubigeo") = vbCodgeo
        .Parameters("@Telefono") = Trim(txtTelefAfi)
        .Parameters("@Empleadores") = IIf(vbindEmp, "1", "0") 'Revisar luego
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
        .Parameters("@Aceptacion") = IIf(chkProc.Value = 1, "S", " ")
        .Parameters("@Texto") = Trim(txtMotiv)
        .Parameters("@RIAafil") = IIf(Val(txtRIAAfil) = 0, 1#, Val(txtRIAAfil))
        .Parameters("@FecNac") = IIf((Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//"), Null, txtFecNac)
        .Parameters("@Sexo") = Trim(txtSexo) 'IIf(Trim(txtSexo) = "", "M", "F")
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
        .Parameters("@UbigeoEmp") = vbCodgeoEmp
        .Parameters("@TelefEmp") = Trim(txtTelefEmp)
        .Parameters("@ExisteBD") = IIf(vbindRuc, "1", "0")
        .Parameters("@nomImag") = RsImag!Imagen
        .Parameters("@indEmp") = IIf(vbindEmp, "1", "0")
        .Parameters("@indExisEmp") = IIf(vbindEmp2, "1", "0")
        .Parameters("@TipImag") = IIf(vbindEmp, "E", "S")
        .Parameters("@posicion") = RsImag!posicion
        .Parameters("@Cuenta") = IIf(chkCuenta.Value, "X", "")
        .Parameters("@Boletin") = IIf(chkBol.Value, "X", "")
        .Parameters("@e_mail2") = Trim(txtemail2)
        
                    
        If Con.GetFEjecutarSQL(comando) Then
            If LCase(Con.vbdev) = "exito" Then
                vbNumsolant = txtNroSolicitud.Text
                If rsEmpleadores.RecordCount > 0 Then
                    If GrabarEmpleadores Then
                         Con.Cnx.CommitTrans
                         lblcont.Caption = Val(lblcont.Caption) + 1
                         NuevoRegistro
                    Else
                         Con.Cnx.RollbackTrans
                    End If
                Else
                    Con.Cnx.CommitTrans
                End If
                
            End If
            
        End If
        
        
                                               

    End With

End Sub

Sub ModificarTraspaso()

vbmarca = rsSolicitud.Bookmark

If Not Validar Then Exit Sub

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
        .Parameters("@NumDoc") = Trim(txtNumDoc)
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
        .Parameters("@RIAafil") = Val(txtRIAAfil)
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
        .Parameters("@posicion") = rsSolicitud!posicion
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


Private Sub cboAfp_GotFocus()
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 60
ImgEdit1.Zoom = 60
Resaltado 0, 0, 1200, 150, 0, 923
cboAfp.BackColor = &HC0FFFF
End Sub

Private Sub cboAfp_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
End Sub

Private Sub cboAfp_LostFocus()
cboAfp.BackColor = &H80000005
End Sub

Private Sub chkBol_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 740, 2156
txtemail2.BackColor = &HC0FFFF

If vbdigTipFondo Then
    If txtFTipoO.Text = "" Then txtFTipoO.SetFocus
End If

End Sub

Private Sub chkBolPag_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 1242
End Sub

Private Sub chkCopSim_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1205, 1248
End Sub


Private Sub chkCuenta_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 740, 2156
txtemail.BackColor = &HC0FFFF
End Sub

Private Sub chkPagUlt_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 1242

End Sub

Private Sub chkProc_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1692, 1526

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

Private Sub cmdEliminar_Click()
Call EliminarEmp
End Sub

Private Sub cmdEnlazar_Click()

If vbNumsolant = "" Then
   If Trim(txtNroSolicitud) = "" Then
      MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
      Exit Sub
   Else
     vbNumsolant = txtNroSolicitud
   End If
End If


If RsImag.EOF Then Exit Sub

If Trim(vbNumsolant) <> "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set idimagen = '" & vbGAFP & Trim(vbNumsolant) & "', Status = '5', TipImag =  'D', " & _
            " posicion = isnull((select max(posicion) + 1 from _imagenes where idimagen = '" & vbGAFP & Trim(vbNumsolant) & "'),1) where Imagen = '" & RsImag!Imagen & "'"
            .Execute
            If Not CargarImagen Then Unload Me: Exit Sub
      End With
Else
    MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
End If
txtNroSolicitud.SetFocus
End Sub

Private Sub cmdEnlazarBol_Click()
If RsImag.EOF Then Exit Sub
If Trim(vbNumsolant) <> "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set idimagen = '" & vbGAFP & Trim(vbNumsolant) & "', Status = '5', TipImag =  'B', " & _
            " posicion = isnull((select max(posicion) + 1 from _imagenes where idimagen = '" & vbGAFP & Trim(txtNroSolicitud) & "'),1) where Imagen = '" & RsImag!Imagen & "'"
            .Execute
            If Not CargarImagen Then Unload Me
      End With
Else
    MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
End If
txtNroSolicitud.SetFocus

End Sub

Private Sub cmdEnlazarMec_Click()

If vbNumsolant = "" Then
   If Trim(txtNroSolicitud) = "" Then
      MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
      Exit Sub
   Else
     vbNumsolant = txtNroSolicitud
   End If
End If


If RsImag.EOF Then Exit Sub
If Trim(vbNumsolant) <> "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set idimagen = '" & vbGAFP & Trim(vbNumsolant) & "', Status = '5', TipImag =  'B', " & _
            " posicion = isnull((select max(posicion) + 1 from _imagenes where idimagen = '" & vbGAFP & Trim(vbNumsolant) & "'),1) where Imagen = '" & RsImag!Imagen & "'"
            .Execute
            If Not CargarImagen Then Unload Me: Exit Sub
      End With
Else
    MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
End If
txtNroSolicitud.SetFocus

End Sub

Private Sub cmdEnlOtros_Click()

If vbNumsolant = "" Then
   If Trim(txtNroSolicitud) = "" Then
      MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
      Exit Sub
   Else
     vbNumsolant = txtNroSolicitud
   End If
End If

If RsImag.EOF Then Exit Sub
If Trim(vbNumsolant) <> "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set idimagen = '" & vbGAFP & Trim(vbNumsolant) & "', Status = '5', TipImag =  'O', " & _
            " posicion = isnull((select max(posicion) + 1 from _imagenes where idimagen = '" & vbGAFP & Trim(vbNumsolant) & "'),1) where Imagen = '" & RsImag!Imagen & "'"
            .Execute
            If Not CargarImagen Then Unload Me: Exit Sub
      End With
Else
    MsgBox "No Puede Enlazar sin el Nro de Solicitud ", 64, Caption
End If
txtNroSolicitud.SetFocus
End Sub

Private Sub cmdGrabaEmp_Click()
If strAccionEmp = "Nuevo" Then
    AgregarTraspEmpleador
Else
    ModificarTraspEmpleador
End If
End Sub

Private Sub cmdGrabar_Click()
If cmdGrabar.Caption = "&Grabar" Then
   Call GrabarTraspaso
  Else
   Call ModificarTraspaso
End If
End Sub

Function CargarImagen() As Boolean
On Error Resume Next
Dim strsql As String
'Lectura de 1 Imagen
strsql = " Select top 1 imagen, c.Ruta, LoteFisico, fecsal,posicion " & _
         " from _Imagenes b join _Volumenes c on c.idvolumen = b.idvolumen " & _
         " join CtrlDoc d on d.Lote = b.Lote " & _
         " where barcode = 0 and b.status = '1' and b.Lote = '" & vbGLte & "' order by  imagen"
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
      
      MsgBox "Lote Concluido (no hay más imágenes para el Lote)", 64, Caption
      'Unload Me
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
      txtApePat.SetFocus
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
      txtApeMat.SetFocus
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
      txtNumDoc.SetFocus
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
          
            If Not CompararTexto(vtmpDNI, txtNumDoc.Text, KeyAscii, Len(txtNumDoc) + 1) Then
              KeyAscii = 0
              Beep
              lblMensaje.Caption = " Caracter no Corresponde"
              txtNumDoc.SetFocus
              ValDNI = False
     
            Else
              lblMensaje.Caption = ""
            End If
            
        Else
           vbnoCompDNI = False
        End If
    Else
       If KeyAscii = 13 And (Len(vtmpDNI) <= Len(txtNumDoc)) Then
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
             lblMensaje.Caption = "Debe Ingresar Valor"
             Beep
             ValCuss = False
             txtCUSPP.SetFocus
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
      txtCUSPP.SetFocus
    End If
 End If
 
 
 If KeyAscii = vbKeyF9 Then
 
  
  
   If vbCuss2 Then '2da Digitacion
    vbnoCuss = True
    'El primer valor se hace igual al 2do
    lblMensaje.Caption = ""
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
       If KeyAscii = 13 And (Len(vtmpCuss) <= Len(txtCUSPP.Text)) Then vbCuss = True: lblMensaje = "": txtAv.SetFocus
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
      txtNombreuno.SetFocus
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
             Me.txtTipDoc.SetFocus
             Exit Function
    End If

    If Not vbNomb2 Then
      vbNomb2 = True
      vbNomb22 = True  '2da digitacion
      vtmpNomb2 = txtNombredos.Text
      txtNombredos.Text = ""
      txtNombredos.SetFocus
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
       If KeyAscii = 13 And (Len(vtmpNomb2) <= Len(txtNombredos)) Then NombreDos = True: Me.txtTipDoc.SetFocus
    End If
 End If
 
End Function


Private Sub cmdGrabar_GotFocus()
If txtFTipoO.Text = "" Then
 If txtFTipoO.Enabled Then
 txtFTipoO.SetFocus
 End If
End If
End Sub

Private Sub cmdMarcaHoja_Click()
Set comando = New Command

With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "Update _Imagenes set Barcode = 1, TipImag =  'H' where imagen = '" & RsImag!Imagen & "' "
    .Execute
End With
If Not CargarImagen Then Unload Me: Exit Sub
End Sub

Private Sub cmdModificar_Click()
strAccionEmp = "Modificar"
End Sub

Private Sub cmdNuevo_Click()
strAccionEmp = "Nuevo"
NuevoEmpleador
End Sub

Private Sub cmdOtroEmpleador_Click()
Call BuscarSolicitud
End Sub

Private Sub cmdrotDer_Click()
On Local Error Resume Next
ImgEdit1.RotateRight
End Sub

Private Sub cmdrotIzq_Click()
On Local Error Resume Next
ImgEdit1.RotateLeft
End Sub

Private Sub cmdSalir_Click()
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

Private Sub cmdSuelto_Click()
'Imagen suelta
If Trim(txtNumDoc) = "" Then
    Set comando = New Command
      With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update _Imagenes set Status = '2', TipImag = 'N' where Imagen = '" & RsImag!Imagen & "'"
            .Execute
      End With
      If Not CargarImagen Then Unload Me: Exit Sub
End If
End Sub


Private Sub Command1_Click()

End Sub

Private Sub DataGrid1_Click()

End Sub


Private Sub dtgEmpleador_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
RefrescarEmpleador
End Sub

Private Sub Form_Load()
CargarComboSql cboAfp, "select codafp + ' - ' + descri, codafp from tafp"
vbGAFP = "01" ''CAmbiar--------

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
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdGrabaEmp.Enabled = False
    cmdEliminar.Enabled = True
    Call CargarImagen
 
Else
    cmdAnterior.Enabled = True
    cmdSiguiente.Enabled = True
    cmdGrabar.Caption = "&Modificar"
    Call CargarDatosSolicitud
    RefrescarSol True
End If
lblTotal.Caption = frmTraspasos1.dtgLotes.Columns(4)

Set rsEmpleadores = New ADODB.Recordset
rsEmpleadores.CursorLocation = adUseClient
rsEmpleadores.LockType = adLockBatchOptimistic

Call CrearTmpEmpleadores
Call ObtenerEmpleadores(Me.txtNroSolicitud)

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
        
        
        vbindEmp2 = True  'Por defecto tiene empleador
        
        'Ind de Empleador
        vbindEmp = False

                
        txtApePat.Text = ""
        txtApeMat.Text = ""
        txtNombreuno = ""
        txtNombredos = ""
        'txtTipDoc = ""
        txtNumDoc = ""
        txtCUSPP = ""
        txtAv = ""
        txtNroDir = ""
        txtUrba = ""
        ucUbigeoAfi.cTexto = ""
        
        txtTelefAfi = ""
        'x defecto
        txtTipDoc = "DNI"
        
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
        txtEspecificar = ""
        txtFechaSol = ""
        cboAfp.ListIndex = -1
        txtFecTras = ""
        chkProc.Value = 0
        txtMotiv = ""
        txtFecNac = ""
        txtRIAAfil = ""
        txtemail = ""
        txtSexo = ""
        txtFTipoO = ""
        txtFTipoD = ""
        txtemail2.Text = ""
        
        
        chkBol.Value = 0
        chkCuenta.Value = 0
        
        
        vbtmpFTipoO = 0  'Origen
        vbtmpFTipoD = 0 'Destino
        vbdigTipFondo = False
        
        vbdigF9FD = False
        vbdigF9FO = False
        
      
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
Text3 = X
Text4 = Y
End Sub

Private Sub ImgEdit1_Scroll()
lblx.Caption = ImgEdit1.ScrollPositionX
lblY.Caption = ImgEdit1.ScrollPositionY

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



Private Sub TxtApeMat_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 769, 335
txtApeMat.BackColor = &HC0FFFF
End Sub

Private Sub TxtApeMat_LostFocus()
txtApeMat.BackColor = &H80000005
End Sub

Private Sub TxtApePat_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 335
txtApePat.BackColor = &HC0FFFF
End Sub

Private Sub txtApePat_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
 If KeyCode = vbKeyTab Then
  vbDigApepat = ApePat(13)
 Exit Sub
 End If
 If KeyCode = vbKeyF9 Then
  vbDigApepat = ApePat(KeyCode)
 End If
Else
KeyCode = 0
End If

End Sub

Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 18 Or KeyAscii = 209) Then
vbDigApepat = ApePat(KeyAscii)
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

 If (GetKeyState(vbKeyTab) < 0) Then
    If Not ApePat(13) Then
     Cancel = True
    End If
    Exit Sub
  End If

  If vbDigApepat Then Cancel = False: Exit Sub
 
  If txtApePat = rsDetalle.Fields("apepat") Then
     Cancel = False
  Else
     vbDigApepat = ApePat(13)
     If Not vbDigApepat Then Cancel = True
  End If
'Else
'    If Not vbDigApepat Then Cancel = True
'End If

End Sub

Private Sub txtApeMat_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
 If KeyCode = vbKeyTab Then
  vbDigApeMat = ApeMat(13)
 Exit Sub
 End If
 If KeyCode = vbKeyF9 Then
  vbDigApeMat = ApeMat(KeyCode)
 End If
Else
KeyCode = 0
End If

End Sub

Private Sub TxtApeMat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
vbDigApeMat = ApeMat(KeyAscii)
Else
KeyAscii = 0
End If

End Sub

Private Sub txtApeMat_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then

  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ApeMat(13) Then
     Cancel = True
    End If
    Exit Sub
  End If

  If vbDigApeMat Then Cancel = False: Exit Sub
 
  If Val(txtApeMat) = Val(rsDetalle.Fields("apemat")) Then
     Cancel = False
  Else
     vbDigApeMat = ApeMat(13)
     If Not vbDigApeMat Then Cancel = True
  End If
'Else
'    If Not vbDigApeMat Then Cancel = True
'End If

End Sub




Private Sub txtAv_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 546
txtAv.BackColor = &HC0FFFF
End Sub

Private Sub txtAv_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
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
Resaltado 0, 0, 1200, 150, 0, 1055
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
  
  If txtCUSPP = Val(rsDetalle.Fields("nomuno")) Then
     'If Err Then vbDigNomb1 = ValRemuAsig(13): Exit Sub
     Cancel = False
  Else
     vbDigCuss = ValCuss(13)
     If Not vbDigCuss Then Cancel = True
  End If

End Sub

Private Sub txtDireEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 846
txtDireEmp.BackColor = &HC0FFFF
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
txtDireEmp.BackColor = &H80000005
End Sub

Private Sub txtemail_GotFocus()
'If ImgEdit1.Zoom = 90 Then ImgEdit1.Zoom = 100
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 700, 2100
txtemail.BackColor = &HC0FFFF
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'If Not vbdigTipFondo Then
''si es el segundo empleador ir defrente a grabar
'    If KeyAscii = 13 Then
'        If Frame4.Enabled = False Then
'           cmdGrabar.SetFocus
'        Else
'           If Trim(txtemail.Text) = "" Then
'               If MsgBox("El email esta en blanco ¿Desea pasar a la revision de fondos?", vbYesNo + vbQuestion, Caption) = vbYes Then
'                       RevisarTipFondo
'               Else
'                       txtemail.SetFocus
'               End If
'           Else
'                chkBol.SetFocus
'           End If
'        End If
'    End If
'Else
'    TabSig KeyAscii
'End If
'TabSig KeyAscii
End Sub

Private Sub txtemail_LostFocus()
txtemail.BackColor = &H80000005
End Sub

Private Sub txtemail_Validate(Cancel As Boolean)
'If Not vbdigTipFondo Then
''si es el segundo empleador ir defrente a grabar
'    'If KeyAscii = 13 Then
'     If Frame4.Enabled = False Then
'        cmdGrabar.SetFocus
'     Else
'        RevisarTipFondo
'     End If
'   ' End If
''Else
''    TabSig KeyAscii
'End If

If (GetKeyState(vbKeyTab) < 0) Then

       If Trim(txtemail.Text) = "" Then
        If Not vbdigTipFondo Then
            If Frame4.Enabled Then
                If MsgBox("El email esta en blanco ¿Desea pasar a la revision de fondos?", vbYesNo + vbQuestion, Caption) = vbYes Then
                    RevisarTipFondo
                Else
                    Cancel = True
                End If
            Else
                 'CmdGrabar.SetFocus
                 chkBol.SetFocus
            End If
        End If
    Else
        chkBol.SetFocus
         'cmdGrabar.SetFocus
    End If
End If

End Sub

Private Sub txtemail2_GotFocus()
'If ImgEdit1.Zoom = 90 Then ImgEdit1.Zoom = 100
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 700, 2250
txtemail2.BackColor = &HC0FFFF
End Sub

Private Sub txtemail2_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'If Not vbdigTipFondo Then
''si es el segundo empleador ir defrente a grabar
'    If KeyAscii = 13 Then
'     If Frame4.Enabled = False Then
'        cmdGrabar.SetFocus
'     Else
'        If Trim(txtemail2.Text) = "" Then
'            If MsgBox("El email2 esta en blanco ¿Desea pasar a la revision de fondos?", vbYesNo + vbQuestion, Caption) = vbYes Then
'                    RevisarTipFondo
'            Else
'                    txtemail2.SetFocus
'            End If
'        Else
'            RevisarTipFondo
'        End If
'     End If
'    End If
'Else
'    TabSig KeyAscii
'End If
End Sub

Private Sub txtemail2_LostFocus()
txtemail2.BackColor = &H80000005
End Sub

Private Sub txtemail2_Validate(Cancel As Boolean)

If (GetKeyState(vbKeyTab) < 0) Then
  If Not vbdigTipFondo Then
     If Frame4.Enabled = False Then
        cmdGrabar.SetFocus
     Else
        If Trim(txtemail2.Text) = "" Then
            If MsgBox("El email2 esta en blanco ¿Desea pasar a la revision de fondos?", vbYesNo + vbQuestion, Caption) = vbYes Then
                    RevisarTipFondo
            Else
                    txtemail2.SetFocus
                    Cancel = True
            End If
        Else
            RevisarTipFondo
            txtFTipoO.SetFocus
            Exit Sub
        End If
     End If
    End If
End If

End Sub

Private Sub txtEspecificar_GotFocus()
txtEspecificar.BackColor = &HC0FFFF
End Sub

Private Sub txtEspecificar_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
    KeyAscii = 0
    Exit Sub
End If
TabSig KeyAscii
End Sub

Private Sub txtEspecificar_LostFocus()
txtEspecificar.BackColor = &H80000005
End Sub

Private Sub txtFechaSol_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 681, 1334
txtFechaSol.BackColor = &HC0FFFF
End Sub

Private Sub txtFechaSol_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
FVFecha txtFechaSol, KeyAscii

End Sub

Private Sub txtFechaSol_LostFocus()

ValidaFecha txtFechaSol
txtFechaSol.BackColor = &H80000005
End Sub

Private Sub txtFecNac_GotFocus()
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 2325
txtFecNac.BackColor = &HC0FFFF
End Sub

Private Sub txtFecNAc_KeyPress(KeyAscii As Integer)
    TabSig KeyAscii
    FVFecha txtFecNac, KeyAscii
End Sub

Private Sub txtFecNAc_LostFocus()
    ValidaFecha txtFecNac
    txtFecNac.BackColor = &H80000005
End Sub

Private Sub txtFecNac_Validate(Cancel As Boolean)
If IsDate(txtFecNac) Then
    If Year(gDate) - Year(txtFecNac) < 17 Then
        MsgBox "Fecha no válida", vbCritical, Caption
        Cancel = True
    End If
End If
End Sub

Private Sub txtFecsal_KeyPress(KeyAscii As Integer)
    FVFecha txtFecsal, KeyAscii
End Sub

Private Sub txtFecsal_LostFocus()
    ValidaFecha txtFecsal
End Sub

Private Sub txtFecTras_GotFocus()
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
'Resaltado 0, 0, 1200, 150, 0, 1592
ImgEdit1.Zoom = 60
Resaltado 0, 0, 1200, 150, 0, 923
txtFecTras.BackColor = &HC0FFFF
End Sub

Private Sub txtFecTras_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
FVFecha txtFecTras, KeyAscii

End Sub

Private Sub txtFecTras_LostFocus()
txtFecTras.BackColor = &H80000005
End Sub

Private Sub txtFecTras_Validate(Cancel As Boolean)
ValidaFecha txtFecTras
End Sub


Private Sub txtFTipoD_GotFocus()
Marcado txtFTipoD
ImgEdit1.Zoom = 60
Resaltado 0, 0, 1200, 150, 0, 923
End Sub



Private Sub txtFTipoD_KeyPress(KeyAscii As Integer)

Dim vbtempKey As Integer
'
vbtempKey = KeyAscii
'TabSig KeyAscii

If Not (KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or KeyAscii = vbKey3 Or KeyAscii = vbKeyBack) Then
    KeyAscii = 0
End If

'If vbdigTipFondo Then
  '  cmdGrabar.SetFocus
'Else
    'TabSig vbtempKey
    If vbtempKey = 13 Then Call txtFTipoD_Validate(False)
'End If
    
'TabSig vbtempKey
End Sub

Private Sub txtFTipoD_KeyUp(KeyCode As Integer, Shift As Integer)
If vbdigTipFondo Then
    If KeyCode = vbKeyF9 Then
        vbdigF9FD = True
        cmdGrabar.SetFocus
    End If
End If

End Sub

Private Sub txtFTipoD_Validate(Cancel As Boolean)
If txtFTipoD = "1" Or txtFTipoD = "2" Or txtFTipoD = "3" Then
    Cancel = False
  Else
    Cancel = True
End If

If vbdigTipFondo Then

    If vbdigF9FD Then
        lblMensaje.Caption = ""
        cmdGrabar.SetFocus
        Exit Sub
    End If

    If Trim(vbtmpFTipoD) <> txtFTipoD.Text Then
        lblMensaje.Caption = " Valores no coinciden"
        Cancel = True
        Exit Sub
    Else
        lblMensaje.Caption = ""
        cmdGrabar.SetFocus
        'Exit Sub
    End If
    
    'vbdigF9FO
Else
    txtFecTras.SetFocus
    
    
End If

End Sub

Private Sub txtFTipoO_GotFocus()
Marcado txtFTipoO
ImgEdit1.Zoom = 60
Resaltado 0, 0, 1200, 150, 0, 923
End Sub


Private Sub txtFTipoO_KeyPress(KeyAscii As Integer)
Dim vbtemkey As Integer

vbtemkey = KeyAscii

If Not (KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or KeyAscii = vbKey3 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyF9) Then
    KeyAscii = 0
End If

TabSig vbtemkey

End Sub

Private Sub txtFTipoO_KeyUp(KeyCode As Integer, Shift As Integer)
'-------Validar--------

If vbdigTipFondo Then
    If KeyCode = vbKeyF9 Then
        vbdigF9FO = True
        txtFTipoD.SetFocus
    End If
End If


End Sub

Private Sub txtFTipoO_Validate(Cancel As Boolean)
If txtFTipoO = "1" Or txtFTipoO = "2" Or txtFTipoO = "3" Then
    Cancel = False
  Else
    Cancel = True
End If

If vbdigTipFondo Then

    
    If vbdigF9FO Then
        lblMensaje.Caption = ""
        txtFTipoD.SetFocus
        Exit Sub
    End If

    If Trim(vbtmpFTipoO) <> txtFTipoO.Text Then
        lblMensaje.Caption = "Valores no coinciden"
        'txtFTipoO.SetFocus
        Cancel = True
        Exit Sub
    Else
        lblMensaje.Caption = ""
        Exit Sub
    End If


End If


End Sub

Private Sub txtMotiv_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1692, 1582
txtMotiv.BackColor = &HC0FFFF
End Sub

Private Sub txtMotiv_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtMotiv_LostFocus()
txtMotiv.BackColor = &H80000005
End Sub

Private Sub txtNombredos_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1841, 335
txtNombredos.BackColor = &HC0FFFF
End Sub

Private Sub txtNombredos_LostFocus()
txtNombredos.BackColor = &H80000005
End Sub

Private Sub txtNombreuno_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1301, 335
txtNombreuno.BackColor = &HC0FFFF
End Sub

Private Sub txtNombreuno_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18 Or KeyCode = 209) Then
 If KeyCode = vbKeyTab Then
  vbDigNomb1 = NombreUno(13)
 Exit Sub
 End If
 If KeyCode = vbKeyF9 Then
  vbDigNomb1 = NombreUno(KeyCode)
 End If
Else
KeyCode = 0
End If
End Sub

Private Sub txtNombreUno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
vbDigNomb1 = NombreUno(KeyAscii)
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNombredos_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18 Or KeyCode = 209) Then
 If KeyCode = vbKeyTab Then
  vbDigNomb2 = NombreDos(13)
  Exit Sub
 End If
 If KeyCode = vbKeyF9 Then
  vbDigNomb2 = NombreDos(KeyCode)
 End If
Else
KeyCode = 0
End If
End Sub

Private Sub txtNombreDos_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
vbDigNomb2 = NombreDos(KeyAscii)
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNombredos_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
 
    If (GetKeyState(vbKeyTab) < 0) Then
        If Not NombreDos(13) Then
         Cancel = True
        End If
        Exit Sub
    End If


  If vbDigNomb2 Then Cancel = False: Exit Sub
  
  If txtNombredos = Val(rsDetalle.Fields("nomdos")) Then
     'If Err Then vbDigNomb1 = ValRemuAsig(13): Exit Sub
     Cancel = False
  Else
     vbDigNomb2 = NombreDos(13)
     If Not vbDigNomb2 Then Cancel = True
  End If
'Else
'    If Not vbDigNomb2 Then Cancel = True
'End If

End Sub

Private Sub txtNombreuno_LostFocus()
txtNombreuno.BackColor = &H80000005
End Sub

Private Sub txtNombreuno_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then
    If (GetKeyState(vbKeyTab) < 0) Then
        If Not NombreUno(13) Then
         Cancel = True
        End If
        Exit Sub
    End If

  If vbDigNomb1 Then Cancel = False: Exit Sub
  
  If txtNombreuno = Val(rsDetalle.Fields("nomuno")) Then
     'If Err Then vbDigNomb1 = ValRemuAsig(13): Exit Sub
     Cancel = False
  Else
     vbDigNomb1 = NombreUno(13)
     If Not vbDigNomb1 Then Cancel = True
  End If
'Else
'    If Not vbDigNomb1 Then Cancel = True
'End If
End Sub

Private Sub txtNroDir_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 683, 546
txtNroDir.BackColor = &HC0FFFF
End Sub

Private Sub txtNroDir_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))


If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
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
Resaltado 0, 0, 1200, 150, 892, 846
txtNroEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtNroEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtNroEmp_LostFocus()
txtNroEmp.BackColor = &H80000005
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
        Frame4.Enabled = False  'Declaracion de Traspaso
        Frame6.Enabled = False  'Resultado de la Solictud
        txtFecNac.Enabled = False '
        txtNroSolicitud = vbNumsolant
        'Frame8.Enabled = False
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
        Frame4.Enabled = True  'Declaracion de Traspaso
        Frame6.Enabled = True  'Resultado de la Solictud
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
Marcado txtNumDoc
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 950, 416
txtNumDoc.BackColor = &HC0FFFF
End Sub

Private Sub txtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)

If ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
 If KeyCode = vbKeyTab Then
  vbDigDNI = ValDNI(13)
 Exit Sub
 End If
 
 If KeyCode = vbKeyF9 Then
  vbDigDNI = ValDNI(KeyCode)
 End If
Else
KeyCode = 0
End If
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig KeyAscii

If txtTipDoc.Text = "DNI" Or txtTipDoc.Text = "CFA" Then
    FEntDec txtNumDoc, KeyAscii, 8, 0
End If

If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
'FEntDec txtNumDoc, KeyAscii, 8, 0
vbDigDNI = ValDNI(KeyAscii)


'KeyAscii = Asc(UCase(Chr(KeyAscii)))

'If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 Or KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
  
Else
  KeyAscii = 0
End If
End Sub

Private Sub txtNumDoc_LostFocus()
txtNumDoc.BackColor = &H80000005
End Sub

Private Sub txtNumDoc_Validate(Cancel As Boolean)
On Error Resume Next
'If vbAccion = "Modificar" Then

  If (GetKeyState(vbKeyTab) < 0) Then
    If Not ValDNI(13) Then
     Cancel = True
    End If
    Exit Sub
  End If

  If vbDigDNI Then Cancel = False: Exit Sub
 
  If Val(txtNumDoc) = Val(rsDetalle.Fields("numdoc")) Then
     Cancel = False
  Else
     vbDigDNI = ValDNI(13)
     If Not vbDigDNI Then Cancel = True
  End If
'Else
'    If Not vbDigApeMat Then Cancel = True
'End If

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
    If Val(txtRIAAfil) >= 50000 Then
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
Resaltado 0, 0, 1200, 150, 0, 2402
txtSexo.BackColor = &HC0FFFF
End Sub

Private Sub txtSexo_KeyPress(KeyAscii As Integer)
'TabSig KeyAscii
KeyAscii = Asc(UCase(Chr(KeyAscii)))


If Not vbdigTipFondo Then
'si es el segundo empleador ir defrente a grabar
    If KeyAscii = 13 Then
     If Frame4.Enabled = False Then
        cmdGrabar.SetFocus
     Else
        RevisarTipFondo
     End If
    End If
Else
    TabSig KeyAscii
End If

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
Resaltado 0, 0, 1200, 150, 1734, 640
txtTelefAfi.BackColor = &HC0FFFF
End Sub


Private Sub txtTelefAfi_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtTelefAfi_LostFocus()
txtTelefAfi.BackColor = &H80000005
End Sub

Private Sub txtTelefEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1734, 936
txtTelefEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtTelefEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtTelefEmp_LostFocus()
txtTelefEmp.BackColor = &H80000005
End Sub

Private Sub txtTipDoc_GotFocus()
Marcado txtTipDoc
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 416
txtTipDoc.BackColor = &HC0FFFF
End Sub

Private Sub txtTipDoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtTipDoc_LostFocus()
txtTipDoc.BackColor = &H80000005
End Sub

Private Sub txtUrba_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1637, 546
txtUrba.BackColor = &HC0FFFF
End Sub

Private Sub txtUrba_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209) Then
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
Resaltado 0, 0, 1200, 150, 1769, 846
txtUrbaEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtUrbaEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtUrbaEmp_LostFocus()
txtUrbaEmp.BackColor = &H80000005
End Sub

Private Sub ucUbigeoAfi_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 38, 295
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
Function Validar() As Boolean
Validar = True
If Trim(txtNroSolicitud) = "" Then
    MsgBox "Debe Ingresar el Nro de Solicitud", vbCritical, Caption
    txtNroSolicitud.SetFocus
    Validar = False
    Exit Function
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


If Not vbindEmp Then
    If Trim(cboAfp.Text) = "" Then
        MsgBox "Debe seleccionar la AFP origen", vbCritical, Caption
        Me.cboAfp.SetFocus
        Validar = False
        Exit Function
    End If
End If

If txtFechaSol <> "" Then
    If CDate(txtFechaSol.Text) > gDate Then
            MsgBox "La fecha de firma de traspaso no puede ser mayor a " & gDate, vbCritical, Caption
            txtFechaSol.SetFocus
            Validar = False
            Exit Function
        End If
End If


If Frame4.Enabled Then

    If Trim(txtFTipoO.Text) = "" Then
            MsgBox "Debe ingresar Tipo de Fondo Origen", vbCritical, Caption
            txtFTipoO.SetFocus
            Validar = False
            Exit Function
    End If
    
    
    If Trim(txtFTipoD.Text) = "" Then
            MsgBox "Debe ingresar Tipo de Fondo Destino", vbCritical, Caption
            txtFTipoD.SetFocus
            Validar = False
            Exit Function
    End If
    
    If txtFecTras <> "" Then
        If CDate(txtFecTras.Text) > gDate Then
            MsgBox "La fecha de declaración de traspaso no puede ser mayor a " & gDate, vbCritical, Caption
            txtFecTras.SetFocus
            Validar = False
            Exit Function
        End If
    End If
    
    Select Case txtTipDoc
    Case "DNI"
        If Len(Trim(txtNumDoc)) <> 8 Then
            MsgBox "El DNI debe tener 8 caracteres", 64, Caption
            Validar = False
            Exit Function
        End If
        
    Case "CE"
        If Len(Trim(txtNumDoc)) < 5 Then
            MsgBox "El CE debe tener de 5 a 8 caracteres", 64, Caption
            Validar = False
            Exit Function
        End If
    
    Case "CIA"
        If Len(Trim(txtNumDoc)) < 2 And Len(Trim(txtNumDoc)) > 5 Then
            MsgBox "El CE debe tener de 2 a 5 caracteres", 64, Caption
            Validar = False
            Exit Function
        End If
    
    Case "CFA"
        If Len(Trim(txtNumDoc)) < 6 And Len(Trim(txtNumDoc)) > 8 Then
            MsgBox "El CE debe tener de 6 a 8 caracteres", 64, Caption
            Validar = False
            Exit Function
        End If
    
    Case Else
    
    End Select
    
    
    
End If

If Trim(txtemail.Text) <> "" Then
  If chkCuenta.Value = 0 Then
        MsgBox "Debe marcar la opción Estado de Cuenta", vbCritical, Caption
        chkCuenta.SetFocus
        Validar = False
        Exit Function
  End If
End If

If Trim(txtemail2.Text) <> "" Then
  If chkBol.Value = 0 Then
        MsgBox "Debe marcar la opción Boletín", vbCritical, Caption
        chkBol.SetFocus
        Validar = False
        Exit Function
  End If
End If

If Trim(txtemail.Text) = "" Then
  If chkCuenta.Value = 1 Then
        MsgBox "Debe ingresar mail", vbCritical, Caption
        txtemail.SetFocus
        Validar = False
        Exit Function
  End If
End If

If Trim(txtemail2.Text) = "" Then
  If chkBol.Value = 1 Then
        MsgBox "Debe ingresar mail2", vbCritical, Caption
        txtemail2.SetFocus
        Validar = False
        Exit Function
  End If
End If



End Function

Private Sub ucUbigeoEmp_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 38, 437
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

Private Sub txtRUC_GotFocus()
'Set controlf = ActiveControl
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 1734, 777
txtRuc.BackColor = &HC0FFFF
End Sub

Private Sub txtRUC_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txtRUC_Validate(Cancel As Boolean)
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
Dim vRuc As String

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
                vRuc = fgDatosql("Select NUMRUC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                If vRuc <> "" Then
                'Consultar no existe razon en truc
                   'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValRuc = True
                   vbindRuc = True
                   lblMensaje.Caption = ""
                   txtDireEmp.SetFocus
                Else
                   lblMensaje.Caption = "RUC no se encuentra en la tabla TRUC"
                   ValRuc = True
                   vbindRuc = False
                   txtDireEmp.SetFocus
                    Beep
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
        
      
                vRuc = fgDatosql("select codpromo from tpromotor where codpromo = '" & Trim(txtCodProm.Text) & "'")
                If vRuc <> "" Then
                'Consultar no existe razon en truc
                'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValProm = True
                   vbindPro = True
                   lblMensaje.Caption = ""
                   If chkBolPag.Enabled Then chkBolPag.SetFocus
                   
                Else
                   lblMensaje.Caption = "Código de Promotor no se encuentra en la Tabla TPromotor"
                   ValProm = False
                   vbindPro = False
                   txtCodProm.SetFocus
                   Beep
                End If
                              
End If
 If KeyAscii = vbKeyF9 Then
   'Fuerza el pase al sgte campo
   ValProm = True
   lblMensaje.Caption = ""
   vbindPro = False
  If chkBolPag.Enabled Then
    chkBolPag.SetFocus
    Else
    'TabSig 13
    txtRIAAfil.SetFocus
  End If
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
         TabSig 13
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
        txtNumDoc = IIf(IsNull(.Fields("Numdoc")), "", .Fields("Numdoc"))
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
                        
        IzqBuscarEnCombo cboAfp, IIf(IsNull(.Fields("CodAFPOri")), "", Trim(.Fields("CodAFPOri"))), 2
        
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
        txtNumDoc = ""
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
        chkProc.Value = 0
        txtMotiv = ""
        txtFecNac = ""
        txtRIAAfil = ""
        txtemail = ""
      
        
End If

End With


End Sub


Private Sub UpDown1_DownClick()
ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY + txtFactor
End Sub

Private Sub UpDown1_UpClick()
ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY + txtFactor
End Sub

Sub ImagenesAlternas(nomimag As String)
   With rsImagAlt
    sqltexto = "select top 4 Ruta, LoteFisico, Imagen from _imagenes a join " & _
    " _Volumenes b on a.IdVolumen = b.IdVolumen where  imagen > '" & nomimag & "' and Lotefisico = '" & vbGLte & "' order by imagen"
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

Sub ObtenerEmpleadores(Numsol As String)
    Dim rsTblEmpleador As ADODB.Recordset
    Set rsTblEmpleador = New ADODB.Recordset
    Dim sqltexto As String
    
    sqltexto = " Select Secuencial, Numruc, Direccion, Urb, Telefono, Ubigeo, RiAafil" & _
               " from Empleadores where Numsol = '" & Numsol & "'"
    
    
    With rsTblEmpleador
    
        .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
     
    End With
    
    With rsEmpleadores
    
        .Open
        
        Do While Not rsTblEmpleador.EOF
            .AddNew
            .Fields("Secuencial") = rsTblEmpleador.Fields("Secuencial")
            .Fields("Numruc") = rsTblEmpleador.Fields("Numruc")
            .Fields("DomDesc") = rsTblEmpleador.Fields("Direccion")
            .Fields("Urb") = rsTblEmpleador.Fields("Urb")
            .Fields("Telefono") = rsTblEmpleador.Fields("Telefono")
            .Fields("Ubigeo") = rsTblEmpleador.Fields("Ubigeo")
            .Fields("Ria") = rsTblEmpleador.Fields("RiAafil")
            
        Loop
    
    End With
    
    Set dtgEmpleador.DataSource = rsEmpleadores

End Sub

Sub AgregarTraspEmpleador()
If Not VerificarRucDup(txtRuc.Text) Then
    With rsEmpleadores
        .AddNew
        .Fields("Secuencial") = rsEmpleadores.RecordCount + 1
        .Fields("Numruc") = Trim(txtRuc)
        .Fields("DomDesc") = Trim(txtDireEmp)
        .Fields("Urb") = Trim(txtUrbaEmp)
        .Fields("Telefono") = Trim(txtTelefEmp)
        .Fields("Ubigeo") = Trim(Left(Mid(ucUbigeoEmp.cTexto, InStr(ucUbigeoEmp.cTexto, "/") + 1), 7))
        .Fields("Ria") = Val(txtRiaEmp)
    End With
  
End If

End Sub
Sub ModificarTraspEmpleador()
If Not VerificarRucDup(txtRuc.Text) Then
    With rsEmpleadores
        
        .Fields("Numruc") = Trim(txtRuc)
        .Fields("DomDesc") = Trim(txtDireEmp)
        .Fields("Urb") = Trim(txtUrbaEmp)
        .Fields("Telefono") = Trim(txtTelefEmp)
        .Fields("Ubigeo") = Trim(Left(Mid(ucUbigeoEmp.cTexto, InStr(ucUbigeoEmp.cTexto, "/") + 1), 7))
        .Fields("Ria") = Val(txtRiaEmp)
  
    End With
    
End If

End Sub
Sub CrearTmpEmpleadores()
With rsEmpleadores
    .Fields.Append "Secuencial", adBigInt
    .Fields.Append "Numruc", adVarChar, 10
    .Fields.Append "DomDesc", adVarChar, 50
    .Fields.Append "Nro", adVarChar, 10
    .Fields.Append "Urb", adVarChar, 20
    .Fields.Append "Telefono", adVarChar, 8
    .Fields.Append "Ubigeo", adVarChar, 6
    .Fields.Append "Ria", adDouble
End With
End Sub

Function VerificarRucDup(numruc As String) As Boolean
With rsEmpleadores
    .MoveFirst
    .Find "Numruc = '" & numruc & "'"
    If .EOF Then
        VerificarRucDup = False
    Else
        VerificarRucDup = True
    End If
End With


End Function


Function GrabarEmpleadores() As Boolean
On Error GoTo errores
Dim cmdEliminaEmpl As ADODB.Command
Set cmdEliminaEmpl = New ADODB.Command


Dim cmdGrabaEmpl As ADODB.Command
Set cmdGrabaEmpl = New ADODB.Command


'With cmdEliminaEmpl
'    .ActiveConnection = Con.Cnx
'    .CommandText = "Delete Empleadores where numsol = '" & txtNroSolicitud.Text & "'"
'    .Execute
'End With


With cmdGrabaEmpl

    'RsEmpleadores.MoveFirst
    'Do While Not RsEmpleadores.EOF
    
        .ActiveConnection = Con.Cnx
        .CommandText = "exec sp_GrabarEmpleador2 ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? "
        .Parameters("@CodAfp") = vbGAFP
        .Parameters("@FecSal") = Trim(txtFecsal)
        .Parameters("@NumSol") = Trim(txtNroSolicitud)
        .Parameters("@secuencial") = rsEmpleadores.Fields("Secuencial")
        .Parameters("@Ruc") = rsEmpleadores.Fields("Numruc")
        .Parameters("@DireEmp") = rsEmpleadores.Fields("DomDesc")
        .Parameters("@NroEmp") = rsEmpleadores.Fields("Nro")
        .Parameters("@UrbaEmp") = rsEmpleadores.Fields("Urb")
        .Parameters("@UbigeoEmp") = rsEmpleadores.Fields("Ubigeo")
        .Parameters("@TelefEmp") = rsEmpleadores.Fields("Telefono")
        .Parameters("@RIAafil") = rsEmpleadores.Fields("Ria")
        .Execute
    'Loop

End With
GrabarEmpleadores = True

errores:
GrabarEmpleadores = False
MsgBox Err.Description, vbCritical, "Grabación de Empleadores"

End Function
Sub NuevoEmpleador()
txtRuc.Text = ""
txtDireEmp.Text = ""
txtNroEmp.Text = ""
ucUbigeoEmp.cTexto = ""
txtTelefEmp.Text = ""
txtRiaEmp.Text = ""
txtRuc.SetFocus
End Sub

Sub RefrescarEmpleador()
    If rsEmpleadores.RecordCount > 0 Then
           txtRuc = .Fields("Numruc")
           txtDireEmp = .Fields("DomDesc")
           txtUrbaEmp = .Fields("Urb")
           txtTelefEmp = .Fields("Telefono")
           ucUbigeoEmp.cTexto = fgDatosql("Select desgeo from ubigeo where codgeo = '" & .Fields("Ubigeo") & "'") & "/" & "" & .Fields("Ubigeo")
           txtRiaEmp = .Fields("Ria")
    End If
End Sub
Sub EliminarEmp()
    
    With rsEmpleadores
        If .RecordCount > 0 Then
        .Delete adAffectCurrent
        End If
    End With
End Sub

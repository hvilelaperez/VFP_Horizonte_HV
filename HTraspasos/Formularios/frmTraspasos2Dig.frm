VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmTraspasos2Dig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspasos 2 Digitación"
   ClientHeight    =   10050
   ClientLeft      =   7260
   ClientTop       =   2265
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   15240
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   2415
      Left            =   120
      TabIndex        =   34
      Top             =   7560
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   4260
      _StockProps     =   0
      ImageControl    =   "ImgEdit"
      AnnotationBackColor=   4210752
      AnnotationFillColor=   65535
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "cmdAnterior"
      Height          =   7095
      Left            =   0
      TabIndex        =   46
      Top             =   720
      Width           =   15375
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
         Height          =   3375
         Left            =   240
         TabIndex        =   52
         Top             =   0
         Width           =   14055
         Begin Traspasos27.UserControl1 ucUbigeoAfi 
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   2280
            Width           =   8055
            _extentx        =   14208
            _extenty        =   661
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
            Left            =   8400
            MaxLength       =   10
            TabIndex        =   16
            Top             =   2280
            Width           =   2055
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
            TabIndex        =   11
            Top             =   1680
            Width           =   4215
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
            TabIndex        =   12
            Top             =   1680
            Width           =   2415
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
            TabIndex        =   14
            Top             =   1680
            Width           =   3255
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
            ItemData        =   "frmTraspasos2Dig.frx":0000
            Left            =   240
            List            =   "frmTraspasos2Dig.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1680
            Width           =   1455
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
            ItemData        =   "frmTraspasos2Dig.frx":0004
            Left            =   8640
            List            =   "frmTraspasos2Dig.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox cboDominios 
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
            Left            =   3000
            TabIndex        =   18
            Top             =   2880
            Width           =   2655
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
            Left            =   11640
            MaxLength       =   1
            TabIndex        =   9
            Top             =   1080
            Width           =   735
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
            Left            =   10560
            MaxLength       =   1
            TabIndex        =   8
            Top             =   1080
            Width           =   735
         End
         Begin VB.CheckBox chkAutoriza 
            Caption         =   "Autoriza a Enviar Estado de cuenta a correo electrónico"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   """S/."" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   8
            EndProperty
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
            TabIndex        =   22
            Top             =   2760
            Width           =   3375
         End
         Begin VB.Frame Frame2 
            Height          =   615
            Left            =   5880
            TabIndex        =   77
            Top             =   2760
            Width           =   4455
            Begin VB.OptionButton optVacio 
               Caption         =   "Vacio"
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
               TabIndex        =   19
               Top             =   160
               Width           =   1215
            End
            Begin VB.OptionButton optTrabajo 
               Caption         =   "Trabajo"
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
               Left            =   3000
               TabIndex        =   21
               Top             =   160
               Width           =   1095
            End
            Begin VB.OptionButton optParticular 
               Caption         =   "Particular"
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
               TabIndex        =   20
               Top             =   160
               Width           =   1215
            End
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
            Height          =   375
            Left            =   240
            MaxLength       =   50
            TabIndex        =   17
            Top             =   2880
            Width           =   2415
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
            Left            =   8280
            MaxLength       =   15
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
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
            Left            =   8280
            MaxLength       =   20
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
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   2
            Top             =   480
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
            Left            =   3000
            MaxLength       =   30
            TabIndex        =   1
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
            Left            =   5640
            MaxLength       =   12
            TabIndex        =   6
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
            MaxLength       =   30
            TabIndex        =   0
            Top             =   480
            Width           =   2655
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
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   5
            Top             =   1080
            Width           =   2535
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
            TabIndex        =   4
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label24 
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
            Left            =   8880
            TabIndex        =   91
            Top             =   2040
            Width           =   3615
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
            TabIndex        =   90
            Top             =   2040
            Width           =   3615
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
            TabIndex        =   89
            Top             =   1440
            Width           =   3735
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
            TabIndex        =   88
            Top             =   1440
            Width           =   2295
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
            Left            =   10200
            TabIndex        =   87
            Top             =   1440
            Width           =   3255
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
            TabIndex        =   86
            Top             =   1440
            Width           =   1335
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
            TabIndex        =   85
            Top             =   1440
            Width           =   1095
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
            Left            =   2760
            TabIndex        =   84
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label lblDestino 
            BackStyle       =   0  'Transparent
            Caption         =   "Destino"
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
            Left            =   11640
            TabIndex        =   81
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblOrigen 
            BackStyle       =   0  'Transparent
            Caption         =   "Origen"
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
            Left            =   10560
            TabIndex        =   80
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lbly 
            Caption         =   "lbly"
            Height          =   495
            Left            =   12600
            TabIndex        =   79
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblx 
            Caption         =   "lblx"
            Height          =   495
            Left            =   11400
            TabIndex        =   78
            Top             =   360
            Visible         =   0   'False
            Width           =   975
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
            Left            =   240
            TabIndex        =   76
            Top             =   2640
            Width           =   975
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
            Left            =   8520
            TabIndex        =   75
            Top             =   840
            Width           =   975
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
            TabIndex        =   59
            Top             =   240
            Width           =   2055
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
            Left            =   6120
            TabIndex        =   58
            Top             =   240
            Width           =   1095
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
            TabIndex        =   57
            Top             =   240
            Width           =   1695
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
            TabIndex        =   56
            Top             =   240
            Width           =   2055
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
            Left            =   6120
            TabIndex        =   55
            Top             =   840
            Width           =   1095
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
            TabIndex        =   54
            Top             =   840
            Width           =   2775
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
            Left            =   3120
            TabIndex        =   53
            Top             =   840
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdrotDer 
         Caption         =   ">"
         Height          =   495
         Left            =   13440
         TabIndex        =   47
         ToolTipText     =   "Rotar Derecha"
         Top             =   6240
         Width           =   495
      End
      Begin VB.CommandButton cmdrotIzq 
         Caption         =   "<"
         Height          =   495
         Left            =   12000
         TabIndex        =   48
         ToolTipText     =   "Rotar Izquierda"
         Top             =   6240
         Width           =   495
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0E0FF&
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0E0FF&
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
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txtFactor 
         Height          =   375
         Left            =   14400
         TabIndex        =   50
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   14775
         TabIndex        =   49
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
         Height          =   2535
         Left            =   11160
         TabIndex        =   51
         Top             =   3600
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   4471
         _StockProps     =   0
         ImageControl    =   "ImgEdit1"
         SelectionRectangleEnabled=   0   'False
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   12480
         TabIndex        =   72
         Text            =   "Text4"
         Top             =   3960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   12480
         TabIndex        =   71
         Text            =   "Text3"
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   12480
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   12360
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   3960
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
         Height          =   2655
         Left            =   240
         TabIndex        =   60
         Top             =   3480
         Width           =   10815
         Begin Traspasos27.UserControl1 ucUbigeoEmp 
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   1680
            Width           =   7575
            _extentx        =   13361
            _extenty        =   661
         End
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
            ItemData        =   "frmTraspasos2Dig.frx":0008
            Left            =   240
            List            =   "frmTraspasos2Dig.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2280
            Width           =   1695
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
            ItemData        =   "frmTraspasos2Dig.frx":000C
            Left            =   240
            List            =   "frmTraspasos2Dig.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1080
            Width           =   1695
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
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1080
            Width           =   3495
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
            Left            =   5760
            MaxLength       =   40
            TabIndex        =   27
            Top             =   1080
            Width           =   2295
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
            TabIndex        =   23
            Top             =   480
            Width           =   1695
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
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   30
            Top             =   2280
            Width           =   5775
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
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   24
            Top             =   480
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
            Left            =   7920
            MaxLength       =   10
            TabIndex        =   31
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Label11 
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
            TabIndex        =   83
            Top             =   2040
            Width           =   3735
         End
         Begin VB.Label Label10 
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
            TabIndex        =   82
            Top             =   840
            Width           =   3735
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
            Left            =   5880
            TabIndex        =   68
            Top             =   840
            Width           =   2775
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
            Left            =   8520
            TabIndex        =   67
            Top             =   1680
            Width           =   1095
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
            TabIndex        =   66
            Top             =   240
            Width           =   1095
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
            Left            =   2040
            TabIndex        =   65
            Top             =   840
            Width           =   3735
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
            TabIndex        =   64
            Top             =   2040
            Width           =   1335
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
            Left            =   8160
            TabIndex        =   63
            Top             =   2040
            Width           =   1575
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
            TabIndex        =   62
            Top             =   1440
            Width           =   3615
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
            TabIndex        =   61
            Top             =   240
            Width           =   1095
         End
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
         TabIndex        =   74
         Top             =   6360
         Width           =   5895
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
         Left            =   14400
         TabIndex        =   73
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   240
      TabIndex        =   35
      Top             =   0
      Width           =   14055
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
         TabIndex        =   38
         Top             =   200
         Width           =   1335
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
         TabIndex        =   37
         Top             =   200
         Width           =   1095
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
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   36
         Top             =   200
         Width           =   1575
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
         TabIndex        =   45
         Top             =   240
         Width           =   855
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
         TabIndex        =   44
         Top             =   240
         Width           =   495
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
         Left            =   4200
         TabIndex        =   43
         Top             =   240
         Width           =   1215
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
         Left            =   12720
         TabIndex        =   42
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
         Left            =   11640
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "Total"
         Height          =   255
         Left            =   12120
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Traspasos"
         Height          =   255
         Left            =   10800
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTraspasos2Dig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsSolicitud As New Recordset

Dim sqltexto As String

Private vbCuss As Boolean
Private vbCuss2 As Boolean
Private vbnoCuss As Boolean
Private vbDigCuss As Boolean
Private vtmpCuss As String

Public IndDoc As String
'Variable Obligatorio

Private rsUbigeoAfi As New Recordset
Private rsUbigeoEmp As New Recordset
Private RsImag As New Recordset

'Empl

Private vbCodDist As String
Private vbCodProv As String
Private vbCodDpto As String
Private vbDesDist As String
Private vbDesProv As String
Private vbDesDpto As String

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

Private strConsulta As String
Private RsBusqueda As New ADODB.Recordset

Private rsImagAlt As New Recordset
Private vbNumsolant As String 'Validar solicitud anterior

Sub ModificarTraspaso()
'vbmarca = rsSolicitud.Bookmark
If Not validar Then Exit Sub

    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = " sp_ModificarSolicitud4 ?,?,?,?,?, ?,?,?,?,?, " & _
                                              " ?,?,?,?,?, ?,?,?,?,?, " & _
                                              " ?,?,?,?,?, ?,?,?,?,?, " & _
                                              " ?,?,?,?,? "
                                                                                                                      
        .Parameters("@CodAfp") = vbGAFP
        .Parameters("@FecSal") = Trim(txtFecsal)
        .Parameters("@Lote") = Trim(TxtLote)
        .Parameters("@NumSol") = Trim(txtNroSolicitud)
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
        .Parameters("@Secuencial") = rsSolicitud.Fields("Secu")
        .Parameters("@CodmodD2") = gNomUsu
        
        .Parameters("@FecNac") = IIf((Trim(txtFecNac) = "" Or Trim(txtFecNac) = "//"), Null, txtFecNac)
        .Parameters("@e_mail") = IIf(Trim(txtcorreo.Text) = "", "", Trim(txtcorreo.Text) & "@" & UCase(Trim(cboDominios.Text)))
        .Parameters("@IndEnvMail") = IIf(chkAutoriza.Value, "1", "0")
        .Parameters("@IndMail") = IIf(optVacio.Value, "0", IIf(optParticular.Value, "1", "2")) ''IIf(optParticular.Value, "1", "2")
        .Parameters("@TipoFonOriD2") = txtTipFO.Text
        .Parameters("@TipoFonDesD2") = txtTipFD.Text
        .Parameters("@PrefDirEmpD2") = Right(Trim(cboPrefDirEmp), 2)
        .Parameters("@PrefLocEmpD2") = Right(Trim(cboPrefLocEmp), 2)
        
        ' -------------------------------------------------------------------
        ' Danes Vilela
        ' 18/Abril/2013
        .Parameters("@PrefDirAfi") = Right(Trim(cboPrefDirAfi), 2)
        .Parameters("@PrefLocAfi") = Right(Trim(cboPrefLocAfi), 2)
        .Parameters("@Direccion") = Trim(txtAv)
        .Parameters("@Numero") = Trim(txtNroDir)
        .Parameters("@Ubigeo") = Trim(Left(Mid(ucUbigeoAfi.cTexto, InStr(ucUbigeoAfi.cTexto, "/") + 1), 7))
        .Parameters("@Telefono") = Trim(txtTelefAfi)
        .Parameters("@Urb") = Trim(txtUrba)
        ' -------------------------------------------------------------------
        
        If Con.GetFEjecutarSQL(comando) Then
            If LCase(Con.vbdev) = "exito" Then
                'Call NuevoRegistro
                rsSolicitud.MoveNext
                
                If Not rsSolicitud.EOF Then
                    Call RefrescarSol(True)
                    Call Inicializar(rsSolicitud.Fields("Secu"))
                Else
                    Call CambiarEstadoLote(Trim(TxtLote), "5")
                    Call ProcesarValidos(TxtLote)
                    Call VerificarRevision(TxtLote)
                    frmSolicitudesD2.CargarDatos
                    Unload Me
                End If
                
            End If
            
        End If
                                                                  

    End With


End Sub

Public Property Let Consulta(texto As String)
     strConsulta = texto
End Property

Public Property Get Consulta() As String
    Consulta = strConsulta
End Property


Private Sub cboDominios_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then cboDominios.ListIndex = -1
End Sub

Private Sub cboPrefDirAfi_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
'Resaltado 0, 0, 1200, 150, 0, 586
Resaltado 0, 0, 1200, 150, 0, 360
End Sub

Private Sub cboPrefDirAfi_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefDirAfi_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboPrefDirAfi.ListIndex = -1
End If
End Sub

Private Sub cboPrefDirEmp_GotFocus()
ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 846
Else
    Resaltado 0, 0, 1200, 150, 57, 1161
End If
'txtDireEmp.BackColor = &HC0FFFF
End Sub

Private Sub cboPrefDirEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboPrefDirEmp.ListIndex = -1
End If
End Sub

Private Sub cboPrefLocAfi_GotFocus()

If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
'Resaltado 0, 0, 1200, 150, 1786, 586
Resaltado 0, 0, 1200, 150, 650, 360
End Sub

Private Sub cboPrefLocAfi_KeyPress(KeyAscii As Integer)
TabSig (KeyAscii)
End Sub

Private Sub cboPrefLocAfi_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cboPrefLocAfi.ListIndex = -1
End Sub

Private Sub cboPrefLocEmp_GotFocus()
ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1769, 846
Else
    Resaltado 0, 0, 1200, 150, 1695, 1161
End If
'txtUrbaEmp.BackColor = &HC0FFFF
End Sub

Private Sub cboPrefLocEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cboPrefLocEmp.ListIndex = -1
End If
End Sub

Private Sub chkAutoriza_GotFocus()
ImgEdit1.Zoom = 50
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 1030
Else
    Resaltado 0, 0, 1200, 150, 0, 1030
End If
End Sub

Private Sub CmdGrabar_Click()
   Call ModificarTraspaso
End Sub

Sub Resaltado(Izq As Integer, Top As Integer, Width As Integer, Height As Integer, X As Integer, Y As Integer)
On Local Error Resume Next
        ImgEdit1.DeleteAnnotationGroup ("Resaltado")
        ImgEdit1.AddAnnotationGroup ("Resaltado")
        ImgEdit1.AnnotationType = 4  '4
        'ImgEdit1.Draw 700, 100, 1200, 150
        ImgEdit1.ScrollPositionX = X
        'ImgEdit1.ScrollPositionY = 100
        ImgEdit1.ScrollPositionY = Y + val(txtFactor.Text)
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
      'txtApePat.SetFocus
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
 
' If vbApepat Then
'
'    If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
'        If Not vbnoCompApepat Then  'si el falso
'
'            If Not CompararTexto(vtmpApepat, txtApePat.Text, KeyAscii, Len(txtApePat) + 1) Then
'              KeyAscii = 0
'              Beep
'              lblMensaje.Caption = " Caracter no Corresponde"
'              txtApePat.SetFocus
'              ApePat = False
'
'            Else
'               lblMensaje.Caption = ""
'            End If
'
'        Else
'           vbnoCompApepat = False
'        End If
'    Else
'
'       If KeyAscii = 13 And (Len(vtmpApepat) <= Len(txtApePat)) Then ApePat = True: txtApeMat.SetFocus
'    End If
' End If
 
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
 
' If vbApeMat Then
'
'    If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
'        If Not vbnoCompApeMat Then  'si el falso
'
'            If Not CompararTexto(vtmpApeMat, txtApeMat.Text, KeyAscii, Len(txtApeMat) + 1) Then
'              KeyAscii = 0
'              Beep
'              lblMensaje.Caption = " Caracter no Corresponde"
'              txtApeMat.SetFocus
'              ApeMat = False
'
'            Else
'              lblMensaje.Caption = ""
'            End If
'
'        Else
'           vbnoCompApeMat = False
'        End If
'    Else
'       If KeyAscii = 13 And (Len(vtmpApeMat) <= Len(txtApeMat)) Then ApeMat = True: txtNombreuno.SetFocus
'    End If
' End If
 
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
 
End Function
Function ValCuss(KeyAscii As Integer) As Boolean
'Funcion que valida el formato del cuspp
If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If IndDoc = "C" Then
        ValCuss = True
        txtFecNac.SetFocus
        Exit Function
    End If

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
                'txtRuc.SetFocus
                txtTipFO.SetFocus
                'txtcorreo.SetFocus
                
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
'       If KeyAscii = 13 And (Len(vtmpCuss) <= Len(txtCUSPP.Text)) Then vbCuss = True: lblMensaje = "": txtTipFO.SetFocus
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

Private Function Busqueda()

With RsBusqueda
 If .State = 1 Then RsBusqueda.Close
 strsql = strConsulta & " where " & campo & " like '%" & Trim(txtCons.Text) & "%' order by " & campo & ""
 RsBusqueda.Open strsql, Con.Cnx, adOpenStatic
 
 If Not (.EOF And .BOF) Then
    Call LlenarLista
 End If
End With
End Function

Private Function LlenarLista()
With lstCons
    .Clear
     RsBusqueda.MoveFirst
     While Not RsBusqueda.EOF
        .AddItem RsBusqueda(0).Value
        RsBusqueda.MoveNext
     Wend
End With
End Function

Public Sub TabSig(i As Integer)
Select Case i
    Case 13
'        SendKeys "{TAB}"
        'Simular que pulsamos la tecla
        keybd_event VK_TAB, 0, 0, 0
        'Simular que soltamos la tecla
                                                                                         keybd_event VK_TAB, 0, KEYEVENTF_KEYUP, 0
End Select

End Sub

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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
MovImg KeyCode
End Sub

Private Sub Form_Load()
''''CargarComboSql cboAfp, "select codafp + ' - ' + descri, codafp from tafp"
'vbGAFP = "01" ''CAmbiar--------
'
CenterForm Me

ccombosql cboDominios, "Select DesDom from Dominio"

'
If IndDoc = "T" Then
    Caption = "Traspasos : Segunda Digitación"
    txtFecNac.Visible = False
    Label17.Visible = False
    txtNroSolicitud.MaxLength = 7

    lblOrigen.Visible = True
    lblDestino.Visible = True
    txtTipFD.Visible = True
    txtTipFO.Visible = True

Else
    Caption = "Afiliación: Segunda Digitación"
    txtFecNac.Visible = True
    Label17.Visible = True
    txtNroSolicitud.MaxLength = 10

    lblOrigen.Visible = False
    lblDestino.Visible = False
    txtTipFD.Visible = False
    txtTipFO.Visible = False
End If


Set ucUbigeoAfi.cConexion = Con.Cnx
    ucUbigeoAfi.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    ucUbigeoAfi.Campocondicion = "desgeo"
    ucUbigeoAfi.Busqinicial = False
    ucUbigeoAfi.Enabled = True
    ucUbigeoAfi.nroCaracteres = 3
    
Set ucUbigeoEmp.cConexion = Con.Cnx
    ucUbigeoEmp.Consulta = "Select desgeo + ' / '+ codgeo, *  from ubigeo "
    ucUbigeoEmp.Campocondicion = "desgeo"
    ucUbigeoEmp.Busqinicial = False
    ucUbigeoEmp.Enabled = True
    ucUbigeoEmp.nroCaracteres = 3

' -------------------------------------------------------------------
' Danes Vilela
' 18/Abril/2013
CargarComboSql cboPrefDirAfi, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocAfi, "select AbreLoc, CodLoc from prefloc"
' -------------------------------------------------------------------
CargarComboSql cboPrefDirEmp, "select abredir, coddir from prefdir"
CargarComboSql cboPrefLocEmp, "select AbreLoc, CodLoc from prefloc"

Call CargarDatosSolicitud
Call RefrescarSol(True)

'lblTotal.Caption = frmTraspasos1.dtgLotes.Columns(4)
'ImgEdit1.Image = "C:\NEW AFILIACION 1.tif"
'ImgEdit1.Display

End Sub
Sub Inicializar(Secu As Integer)
        
         txtFactor = 0
        'Empleadores

        vbCodDist = ""
        vbCodProv = ""
        vbCodDpto = ""
        vbDesDist = ""
        vbDesProv = ""
        vbDesDpto = ""
    
        vbCodDistEmp = ""
        vbCodProvEmp = ""
        vbCodDptoEmp = ""
        vbDesDistEmp = ""
        vbDesProvEmp = ""
        vbDesDptoEmp = ""
        
        
        vbCodgeoEmp = ""
        vbCodgeo = ""
        
        vbindRuc = False
        
        If Secu = 0 Then
            frmAfiliado.Enabled = True
            'txtNroSolicitud = ""
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
            txtFecNac.Text = ""
            lblRazonSoc.Text = ""
            txtApePat.SetFocus
            txtcorreo.Text = ""
            txtTipFO.Text = ""
            txtTipFD.Text = ""
            chkAutoriza.Value = 0
            optParticular.Value = False
            optTrabajo.Value = False
            optVacio.Value = True
            cboDominios.ListIndex = -1
            
            '------------------------------------------------------------------------
            ' Se agregaron nuevos controles a solicitud
            ' Danes Vilela
            ' 18/Abril/2013
            cboPrefDirAfi.ListIndex = -1
            cboPrefLocAfi.ListIndex = -1
            txtAv = ""
            txtNroDir = ""
            txtUrba = ""
            ucUbigeoAfi.cTexto = ""
            txtTelefAfi = ""
            '------------------------------------------------------------------------
        Else
            frmAfiliado.Enabled = False
            txtRuc.SetFocus
        End If

        'x defecto
        txtTipDoc = "00"
        txtRuc = ""
        txtDireEmp = ""
        txtNroEmp = ""
        txtUrbaEmp = ""
        ucUbigeoEmp.cTexto = ""
        txtTelefEmp = ""
        cboPrefDirEmp.ListIndex = -1
        cboPrefLocEmp.ListIndex = -1
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmTraspasos2Dig = Nothing
End Sub

Private Sub ImgEdit1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblx = X
lbly = Y
End Sub

Private Sub ImgEdit1_Scroll()
lblx = ImgEdit1.ScrollPositionX
lbly = ImgEdit1.ScrollPositionY

End Sub

Private Sub ImgEdit2_Click()
 SiempreVisible frmVisorImag, True
 frmVisorImag.ImgEdit1.Image = ImgEdit2.Image
 frmVisorImag.ImgEdit1.Display
 frmVisorImag.ImgEdit1.FitTo 1
 frmVisorImag.Show
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

Private Sub optParticular_GotFocus()
ImgEdit1.Zoom = 75
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1060, 1458
Else
    Resaltado 0, 0, 1200, 150, 1060, 1458
End If

End Sub

Private Sub optTrabajo_GotFocus()
ImgEdit1.Zoom = 75
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1060, 1458
Else
    Resaltado 0, 0, 1200, 150, 1060, 1458
End If

End Sub

Private Sub txtApeMat_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 926, 380
Else
    Resaltado 0, 0, 1200, 150, 525, 563
End If
txtApeMat.BackColor = &HC0FFFF
End Sub

Private Sub txtApeMat_LostFocus()
txtApeMat.BackColor = &H80000005
End Sub

Private Sub TxtApePat_GotFocus()

If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 380
Else
    Resaltado 0, 0, 1200, 150, 0, 563
End If

txtApePat.BackColor = &HC0FFFF
End Sub

Private Sub txtApePat_KeyDown(KeyCode As Integer, Shift As Integer)
'MovImg KeyCode
If ((KeyCode >= 65 And KeyCode <= 90) Or KeyCode = 32 Or KeyCode = 39 Or KeyCode = vbKeyF9 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 165 Or KeyCode = 18) Then
' If KeyCode = vbKeyTab Then
'  vbDigApepat = ApePat(13)
' Exit Sub
 'End If
' If KeyCode = vbKeyF9 Then
''  vbDigApepat = ApePat(KeyCode)
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
'Else
'    If Not vbDigApepat Then Cancel = True
'End If

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
'Else
'    If Not vbDigApeMat Then Cancel = True
'End If

End Sub

Private Sub txtcorreo_GotFocus()
'If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
ImgEdit1.Zoom = 75
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 121, 1458
Else
    Resaltado 0, 0, 1200, 150, 121, 1458
End If

txtcorreo.BackColor = &HC0FFFF
End Sub

Private Sub txtcorreo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
KeyAscii = ValidarDigitacion(txtcorreo, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ" & "_-." & vbCaracteresNumero)


If KeyAscii = 13 Then cboDominios.SetFocus  'optParticular.SetFocus


End Sub

Private Sub TxtCUSPP_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
   'Resaltado 0, 0, 1200, 150, 1713, 633
    Resaltado 0, 0, 1200, 150, 1809, 480
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

'  If vbDigCuss Then Cancel = False: Exit Sub
'
'  If txtCUSPP = val(rsDetalle.Fields("nomuno")) Then
'     'If Err Then vbDigNomb1 = ValRemuAsig(13): Exit Sub
'     Cancel = False
'  Else
'     vbDigCuss = ValCuss(13)
'     If Not vbDigCuss Then Cancel = True
'  End If

End Sub

Private Sub txtDireEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 846
Else
    Resaltado 0, 0, 1200, 150, 57, 1161
End If
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

Private Sub txtDireEmp_Validate(Cancel As Boolean)
If MsgBox("¿Verificó la dirección contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
    Cancel = False
Else
    Cancel = True
End If
End Sub

Private Sub txtFecNAc_GotFocus()
If ImgEdit1.Zoom <> 100 Then ImgEdit1.Zoom = 100
Resaltado 0, 0, 1200, 150, 0, 633
txtFecNac.BackColor = &HC0FFFF
End Sub

Private Sub txtFecNac_KeyPress(KeyAscii As Integer)
    TabSig KeyAscii
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
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    'Resaltado 0, 0, 1200, 150, 1841, 335
    Resaltado 0, 0, 1200, 150, 1841, 380
Else
    Resaltado 0, 0, 1200, 150, 1745, 563
End If
txtNombredos.BackColor = &HC0FFFF
End Sub

Private Sub txtNombreDos_LostFocus()
txtNombredos.BackColor = &H80000005
End Sub

Private Sub txtNombreUno_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100

If IndDoc = "T" Then
    'Resaltado 0, 0, 1200, 150, 1480, 335
    Resaltado 0, 0, 1200, 150, 1618, 380
Else
    Resaltado 0, 0, 1200, 150, 1175, 563
End If

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
If KeyAscii = 13 Then txtTipDoc.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNombreUno_LostFocus()
txtNombreuno.BackColor = &H80000005
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


Private Sub txtNroDir_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
'Resaltado 0, 0, 1200, 150, 941, 586
Resaltado 0, 0, 1200, 150, 300, 360

txtNroDir.BackColor = &HC0FFFF
End Sub

Private Sub txtNroDir_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = vbKeyF9 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 165 Or KeyAscii = 209 Or KeyAscii = 78 Or KeyAscii = 83 Or KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 45) Then
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
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 892, 846
Else
    Resaltado 0, 0, 1200, 150, 876, 1161
End If
txtNroEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtNroEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtNroEmp_LostFocus()
txtNroEmp.BackColor = &H80000005
End Sub

Private Sub txtNroEmp_Validate(Cancel As Boolean)
If MsgBox("¿Verificó el número contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
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

Private Sub txtNroSolicitud_LostFocus()
txtNroSolicitud.Text = Format(txtNroSolicitud.Text, "000000#")
txtNroSolicitud.BackColor = &H80000005
End Sub


Private Sub txtNumDoc_GotFocus()

If txtTipDoc.Text = "04" Then
    txtNumDoc.MaxLength = 12
Else
    txtNumDoc.MaxLength = 8
End If

Marcado txtNumDoc
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 950, 446
Else
    Resaltado 0, 0, 1200, 150, 1175, 633
End If
txtNumDoc.BackColor = &HC0FFFF
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
 
  If val(txtNumDoc) = val(rsDetalle.Fields("numdoc")) Then
     Cancel = False
  Else
     vbDigDNI = ValDNI(13)
     If Not vbDigDNI Then Cancel = True
  End If
'Else
'    If Not vbDigApeMat Then Cancel = True
'End If

End Sub

Private Sub txtRuc_LostFocus()
txtRuc.BackColor = &H80000005
End Sub

Private Sub txtTelefAfi_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 300, 320
End Sub

Private Sub txtTelefEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1819, 936
Else
    Resaltado 0, 0, 1200, 150, 1656, 1245
End If
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
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 446
Else
    Resaltado 0, 0, 1200, 150, 666, 633
End If
txtTipDoc.BackColor = &HC0FFFF

End Sub

Private Sub txtTipDoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtTipDoc_LostFocus()
txtTipDoc.BackColor = &H80000005
End Sub

Private Sub txtTipFD_GotFocus()
ImgEdit1.Zoom = 60
Resaltado 0, 0, 1200, 150, 0, 732
End Sub

Private Sub txtTipFD_KeyPress(KeyAscii As Integer)
Dim vbtemkey As Integer

vbtemkey = KeyAscii

If Not (KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or KeyAscii = vbKey3 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyF9 Or KeyAscii = 13) Then
    KeyAscii = 0
Else
    cboPrefDirAfi.SetFocus
End If

'TabSig vbtemkey
End Sub

Private Sub txtTipFD_LostFocus()
'If txtTipFD <> "1" And txtTipFD <> "2" Or txtTipFD <> "3" Then
'    txtTipFD = ""
'End If
End Sub

Private Sub txtTipFD_Validate(Cancel As Boolean)
If txtTipFD = "1" Or txtTipFD = "2" Or txtTipFD = "3" Or txtTipFD = "" Then
    Cancel = False
  Else
    Cancel = True
End If
End Sub

Private Sub txtTipFO_GotFocus()
ImgEdit1.Zoom = 60
Resaltado 0, 0, 1200, 150, 0, 732
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

Private Sub txtUrba_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
'Resaltado 0, 0, 1200, 150, 1786, 586
Resaltado 0, 0, 1200, 150, 650, 360
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

Private Sub txtUrbaEmp_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 1769, 846
Else
    Resaltado 0, 0, 1200, 150, 1695, 1161
End If
txtUrbaEmp.BackColor = &HC0FFFF
End Sub

Private Sub txtUrbaEmp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtUrbaEmp_LostFocus()
txtUrbaEmp.BackColor = &H80000005
End Sub

Private Sub txtUrbaEmp_Validate(Cancel As Boolean)
If MsgBox("¿Verificó la Urbanización contra la imagen?", vbQuestion + vbYesNo, Caption) = vbYes Then
    Cancel = False
Else
    Cancel = True
End If
End Sub


Function validar() As Boolean
validar = True
If Trim(txtNroSolicitud) = "" Then
    MsgBox "Debe Ingresar el Nro de Solicitud", vbCritical, Caption
    txtNroSolicitud.SetFocus
    validar = False
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
    
If frmAfiliado.Enabled Then

    If Trim(Trim(txtcorreo.Text) & Trim(cboDominios.Text)) <> "" Then
        If Not Validar_Email(Trim(txtcorreo.Text) & "@" & Trim(cboDominios.Text)) Then
            validar = False
            Exit Function
        End If
    End If
    
    
    If Not optVacio.Value And Trim(txtcorreo.Text) = "" Then
        MsgBox " Ingrese correo electrónico ", 64, Caption
        validar = False
        Exit Function
    End If
    
    If optVacio.Value And Trim(txtcorreo.Text) <> "" Then
        MsgBox " Coloque indicador de correo electrónico (Particular o Trabajo)", 64, Caption
        validar = False
        Exit Function
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
    
End If
'
'If txtcorreo.Text = "" Then
'        If MsgBox("No ingresó correo electrónico ¿Desea continuar? ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
'            validar = False
'            txtcorreo.SetFocus
'            Exit Function
'        End If
'End If
    
If (cboPrefDirEmp.ListIndex <> -1 And Trim(txtDireEmp.Text) = "") Or (cboPrefDirEmp.ListIndex = -1 And Trim(txtDireEmp.Text) <> "") Then
        MsgBox " Revise la dirección del empleador y el prefijo", 64, Caption
        validar = False
        Exit Function
End If


If (cboPrefLocEmp.ListIndex <> -1 And Trim(txtUrbaEmp.Text) = "") Or (cboPrefLocEmp.ListIndex = -1 And Trim(txtUrbaEmp.Text) <> "") Then
        MsgBox " Revise la Urba del empleador y el prefijo", 64, Caption
        validar = False
        Exit Function
End If



End Function

Private Sub txtAv_GotFocus()
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
'Resaltado 0, 0, 1200, 150, 0, 586
Resaltado 0, 0, 1200, 150, 200, 360

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

Private Sub ucUbigeoAfi_GotFocus()
 ImgEdit1.Zoom = 50
 Resaltado 0, 15, 1200, 50, 38, 320
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
'KeyAscii = ValidarDigitacion(txtTelefAfi, KeyAscii, "QWERTYUIOPASDFGHJKLZXCVBNMÑ- " & vbCaracteresNumero)

If KeyAscii = 13 Then
    If LLenarDatosUbigeo Then TabSig KeyAscii
Else
    If KeyAscii = vbKeyF9 Then TabSig 13
End If
End Sub

Function LLenarDatosUbigeo() As Boolean
'Ubigeo de Afiliados
On Error Resume Next
Dim rsUbigeoAfi As New ADODB.Recordset
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

Function ValidarDigitacion(Cajatexto As TextBox, KeyCode As Integer, CadenaVal As String) As Integer

KeyCode = Asc(Chr(KeyCode))

If Not ValidarCaracteres(KeyCode, CadenaVal) Then
    ValidarDigitacion = 0
Else
    ValidarDigitacion = KeyCode
End If

End Function

Private Sub ucUbigeoEmp_GotFocus()
 ImgEdit1.Zoom = 50
 If IndDoc = "T" Then
    Resaltado 0, 15, 1200, 50, 38, 437
 Else
     Resaltado 0, 15, 1200, 50, 38, 609
 End If
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

Private Sub txtRuc_GotFocus()
'Set controlf = ActiveControl
If ImgEdit1.Zoom = 50 Then ImgEdit1.Zoom = 100
If IndDoc = "T" Then
    Resaltado 0, 0, 1200, 150, 0, 777
Else
    Resaltado 0, 0, 1200, 150, 0, 1089
End If
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

Sub CargarDatosSolicitud()
Dim Strsqltexto As String

Strsqltexto = "Select a.CodAfp ,Secu, a.FecSal , a.Lote , a.NumSol , " & _
" eDesgeo = f.desgeo, a.posicion, imagen, d.Ruta, LoteFisico " & _
" from solicitud a join _Imagenes c on c.idImagen = a.idImagen and  a.posicion = c.posicion " & _
" join _Volumenes d on d.idvolumen = c.idvolumen " & _
" left join Empleadores e on e.codafp = a.codafp and e.numsol = a.numsol and a.fecsal = e.fecsal and   e.Secuencial = a.secu " & _
" left join Ubigeo f on E.ubigeo = F.codgeo " & _
" where a.CodAfp = '" & vbGAFP & "' and a.Lote = '" & vbGLte & "' and  isnull(indD2,0) = 0 and (TipImag = 'S' or  TipImag = 'E') and  a.inddoc = '" & IndDoc & "' order by a.numsol,a.secu  "


If rsSolicitud.State = 1 Then rsSolicitud.Close
rsSolicitud.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
 
With rsSolicitud


If (.EOF And .BOF) Then
     'Call CargarDatosEmpleadores
     cmdGrabar.Enabled = False
     MsgBox "No existen Registros para este Lote", 64, Caption
End If

lblTotal.Caption = .RecordCount

End With
End Sub

Sub RefrescarSol(vbResImag As Boolean)

With rsSolicitud
If Not (.EOF And .BOF) Then
        txtFecsal = IIf(IsNull(.Fields("Fecsal")), "", .Fields("Fecsal"))
        TxtLote = IIf(IsNull(.Fields("Lote")), "", .Fields("Lote"))
        txtNroSolicitud = IIf(IsNull(.Fields("Numsol")), "", .Fields("Numsol"))
                         
        '----------------------
        'Visualizar Imagen correspondiente
        If vbResImag Then
            ImgEdit1.Image = Rutaimag(.Fields("Ruta"), .Fields("LoteFisico"), .Fields("Imagen"))
            ImgEdit1.Display
        End If
        
        
        ImgEdit2.Image = ImgEdit1.Image
        ImgEdit2.FitTo 0
        ImgEdit2.Display
        
        lblcont.Caption = .AbsolutePosition
        Set RsImag = CargarImagenes(IIf(IsNull(.Fields("numsol")), "", Trim(.Fields("numsol"))), .Fields("LoteFisico"))
        
Else
        txtApePat.Text = ""
        txtApeMat.Text = ""
        txtNombreuno = ""
        txtNombredos = ""
        txtTipDoc = ""
        txtNumDoc = ""
        txtCUSPP = ""
        txtFecNac = ""
                                                        
        txtRuc = ""
        lblRazonSoc.Text = ""
        txtDireEmp.Text = ""
        txtNroEmp.Text = ""
        txtUrbaEmp.Text = ""
        ucUbigeoEmp.cTexto = ""
        txtTelefEmp.Text = ""
        
        '------------------------------------------------------------------------
        ' Se agregaron nuevos controles a solicitud
        ' Danes Vilela
        ' 18/Abril/2013
        cboPrefDirAfi.ListIndex = -1
        cboPrefLocAfi.ListIndex = -1
        txtAv = ""
        txtNroDir = ""
        txtUrba = ""
        ucUbigeoAfi.cTexto = ""
        txtTelefAfi = ""
        '------------------------------------------------------------------------
    
        
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
ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY + txtFactor
End Sub

Private Sub UpDown1_UpClick()
ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY - txtFactor
End Sub

Sub ProcesarValidos(Lote As String)

Dim sqltexto As String
 sqltexto = " Update Solicitud set indVerD2 = 1, UsuVerD2 = 'SIST'" & _
            " from Solicitud a left join Empleadores e on e.codafp = a.codafp and e.numsol = a.numsol and e.Secuencial = a.secu " & _
            " where indD2 = 1 and a.secu = 0  " & _
            " and (a.ApePat = a.ApePatD2 and a.ApeMat = a.ApeMatD2 and a.Nombre1 = a.Nombre1D2 and " & _
            " a.Nombre2 = a.Nombre2D2 and a.tipodoc = a.tipodocd2 and a.numdoc = numdocd2 and isnull(a.CUSPP,'') = isnull(a.CUSPPD2,'') and isnull(a.fecnac,'') = isnull(a.fecnacD2,'') and e.NumRuc = e.NumRucD2 and e.Direccion = e.DireccionD2 and " & _
            " e.Numero = e.NumeroD2 and e.Urb = e.UrbD2 and e.Telefono = e.TelefonoD2 and e.PrefDirEmp = e.PrefDirEmpD2 and e.PrefLocEmp = e.PrefLocEmpD2 ) and a.lote = '" & Lote & "'"

Dim CmdProceso As Command

Set CmdProceso = New Command
With CmdProceso
    .ActiveConnection = Con.Cnx
    .CommandText = sqltexto
    .Execute , , adExecuteNoRecords
End With

End Sub


Sub MovImg(ByRef KeyCode As Integer)
On Error Resume Next

If KeyCode = vbKeyF12 Then
    If Not (RsImag.EOF And RsImag.BOF) Then
      RsImag.MoveNext
      If RsImag.EOF Then RsImag.MoveLast
      'ImgEdit1.Image = Rutaimag(rsImag!Ruta, vbGLteAfp, rsImag!imagen)
      ImgEdit2.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      txtnumpag = RsImag!Posicion
      ImgEdit2.Display
            
      ImgEdit2.FitTo 1
      ImgEdit2.Refresh
      KeyCode = 0
 
     ' Call VerPlanillaScala
      'Call MovLinea
    End If
End If

If KeyCode = vbKeyF11 Then
    If Not (RsImag.EOF And RsImag.BOF) Then
      RsImag.MovePrevious
      If RsImag.BOF Then RsImag.MoveFirst
      'ImgEdit1.Image = Rutaimag(rsImag!Ruta, vbGLteAfp, rsImag!imagen)
      ImgEdit2.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      txtnumpag = RsImag!Posicion
      ImgEdit2.Display
      
      ImgEdit2.FitTo 1
      ImgEdit2.Refresh
      KeyCode = 0
  
      'Call MovLinea
    End If
End If


End Sub

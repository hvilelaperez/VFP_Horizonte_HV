VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmValidaciones 
   Caption         =   "Validaciones"
   ClientHeight    =   11010
   ClientLeft      =   255
   ClientTop       =   255
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkBoleta 
      Caption         =   "Sin Boleta de Pago y Otros / Copia Simple"
      Height          =   255
      Left            =   4320
      TabIndex        =   51
      Top             =   480
      Width           =   3975
   End
   Begin VB.CheckBox chk50 
      Caption         =   "Visualizar mayores de 50 años"
      Height          =   255
      Left            =   10080
      TabIndex        =   45
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdTerminar 
      Caption         =   "&Terminar"
      Height          =   495
      Left            =   10080
      TabIndex        =   44
      Top             =   10320
      Width           =   1455
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   11640
      TabIndex        =   8
      Top             =   10320
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00000000&
      Caption         =   "&Salir"
      Height          =   495
      Left            =   13200
      TabIndex        =   2
      Top             =   10320
      Width           =   1455
   End
   Begin VB.CheckBox chkOmitir 
      Caption         =   "Visualizar  todos los lotes"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.CheckBox chkFedOblig 
      Caption         =   "Firma de Fedatario Obigatorio"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10095
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   17806
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Firmas y Sellos"
      TabPicture(0)   =   "frmValidaciones.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImgEdit1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Validación Física"
      TabPicture(1)   =   "frmValidaciones.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmDatDif"
      Tab(1).Control(1)=   "frmExiEnm"
      Tab(1).Control(2)=   "frmVerSol"
      Tab(1).Control(3)=   "frmIAEmp"
      Tab(1).Control(4)=   "frmIAfi"
      Tab(1).Control(5)=   "TDBGrid2"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Validación SBS"
      TabPicture(2)   =   "frmValidaciones.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtNumdocCus"
      Tab(2).Control(1)=   "txtFecNacCus"
      Tab(2).Control(2)=   "txtApePatCus"
      Tab(2).Control(3)=   "chCambios"
      Tab(2).Control(4)=   "SSTab2"
      Tab(2).Control(5)=   "txtCusSbs"
      Tab(2).Control(6)=   "txtAfpOri"
      Tab(2).Control(7)=   "chkOtraAfp"
      Tab(2).Control(8)=   "chkExistSpp"
      Tab(2).Control(9)=   "txtTipDoc"
      Tab(2).Control(10)=   "txtUbiNacRENIEC"
      Tab(2).Control(11)=   "txtUbiNacINEI"
      Tab(2).Control(12)=   "txtFecNac"
      Tab(2).Control(13)=   "txtSexo"
      Tab(2).Control(14)=   "txtNumdoc"
      Tab(2).Control(15)=   "txtNom2"
      Tab(2).Control(16)=   "txtNom1"
      Tab(2).Control(17)=   "txtApeMat"
      Tab(2).Control(18)=   "txtApepat"
      Tab(2).Control(19)=   "txtCuspp"
      Tab(2).Control(20)=   "txtNumsol"
      Tab(2).Control(21)=   "TDBGrid3"
      Tab(2).Control(22)=   "Command1"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "cmdToolBar(3)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "cmdToolBar(2)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "cmdToolBar(1)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "cmdToolBar(0)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "ImgEdit2"
      Tab(2).Control(28)=   "Label26"
      Tab(2).Control(29)=   "Label25"
      Tab(2).Control(30)=   "Label21"
      Tab(2).Control(31)=   "Label15"
      Tab(2).Control(32)=   "Label14"
      Tab(2).Control(33)=   "Label13"
      Tab(2).Control(34)=   "Label12"
      Tab(2).Control(35)=   "Label11"
      Tab(2).Control(36)=   "Label10"
      Tab(2).Control(37)=   "Label9"
      Tab(2).Control(38)=   "Label8"
      Tab(2).Control(39)=   "Label7"
      Tab(2).Control(40)=   "Label6"
      Tab(2).Control(41)=   "Label5"
      Tab(2).Control(42)=   "Label3"
      Tab(2).Control(43)=   "Label2"
      Tab(2).Control(44)=   "Label1"
      Tab(2).ControlCount=   45
      Begin VB.Frame frmDatDif 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   -63840
         TabIndex        =   109
         Top             =   6240
         Width           =   2055
         Begin VB.CheckBox chkDDNroDoc 
            Caption         =   "Nro Documento"
            Height          =   375
            Left            =   120
            TabIndex        =   115
            Top             =   2520
            Width           =   1815
         End
         Begin VB.CheckBox chkDDTipDoc 
            Caption         =   "Tipo Documento"
            Height          =   375
            Left            =   120
            TabIndex        =   113
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CheckBox chkDDFecNac 
            Caption         =   "Fec. Nacimiento"
            Height          =   375
            Left            =   120
            TabIndex        =   112
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CheckBox chkDDNomAfi 
            Caption         =   "Nombre Afiliado"
            Height          =   375
            Left            =   120
            TabIndex        =   111
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox chkDDApeAfi 
            Caption         =   "Ape. Afiliado"
            Height          =   375
            Left            =   120
            TabIndex        =   110
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label24 
            Caption         =   "Diferentes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame frmExiEnm 
         Caption         =   "Existen Enmendaduras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   -66120
         TabIndex        =   103
         Top             =   6240
         Width           =   2175
         Begin VB.CheckBox chkEEDatAfi 
            Caption         =   "Datos Afiliado"
            Height          =   375
            Left            =   120
            TabIndex        =   108
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkEEDatEmp 
            Caption         =   "Datos Empleador"
            Height          =   375
            Left            =   120
            TabIndex        =   107
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox chkEEDatPro 
            Caption         =   "Datos Promotor"
            Height          =   375
            Left            =   120
            TabIndex        =   106
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox chkEEAntPrev 
            Caption         =   "Antecedente Previsionales"
            Height          =   495
            Left            =   120
            TabIndex        =   105
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CheckBox chkEEDecSol 
            Caption         =   "Verificación o Declaración de la Solicitud"
            Height          =   615
            Left            =   120
            TabIndex        =   104
            Top             =   2160
            Width           =   1815
         End
      End
      Begin VB.Frame frmVerSol 
         Caption         =   "Llenado  Incorrecto Datos "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   -68760
         TabIndex        =   97
         Top             =   6240
         Width           =   2535
         Begin VB.CheckBox chkVSFecSus 
            Caption         =   "Fecha Suscripción"
            Height          =   375
            Left            =   120
            TabIndex        =   101
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chkVSTipFon 
            Caption         =   "Elección Tipo Fondo"
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CheckBox chkVSDecTra 
            Caption         =   "Declaración de Traspaso"
            Height          =   375
            Left            =   120
            TabIndex        =   99
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox chkVSVS 
            Caption         =   "Verificación de Solicitud"
            Height          =   375
            Left            =   120
            TabIndex        =   98
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label23 
            Caption         =   "Verificación Solicitud"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame frmIAEmp 
         Caption         =   "Llenado Incorrecto "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   -70680
         TabIndex        =   90
         Top             =   6240
         Width           =   1815
         Begin VB.CheckBox chkDEUbi 
            Caption         =   "Ubigeo "
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox chkDEDir 
            Caption         =   "Dirección"
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox chkDERazSoc 
            Caption         =   "Razón Social"
            Height          =   375
            Left            =   120
            TabIndex        =   92
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox chkDENroRuc 
            Caption         =   "Nro Ruc"
            Height          =   375
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkDENroEmp 
            Caption         =   "Nro de Empleadores"
            Height          =   375
            Left            =   120
            TabIndex        =   93
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label22 
            Caption         =   "Datos Empleador"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame frmIAfi 
         Caption         =   "Llenado Incorrecto Afiliado"
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
         Height          =   3015
         Left            =   -74760
         TabIndex        =   76
         Top             =   6240
         Width           =   3975
         Begin VB.CheckBox chkIncAfiDir 
            Caption         =   "Dirección del Afiliado"
            Height          =   375
            Left            =   2040
            TabIndex        =   82
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkIncAfiFalDat 
            Caption         =   "Faltan Datos en formatos copias"
            Height          =   375
            Left            =   2040
            TabIndex        =   87
            Top             =   2520
            Width           =   1815
         End
         Begin VB.CheckBox chkIncAfiUbi 
            Caption         =   "Ubigeo"
            Height          =   375
            Left            =   2040
            TabIndex        =   89
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CheckBox chkIncAfiTipDoc 
            Caption         =   "Tipo documento"
            Height          =   375
            Left            =   240
            TabIndex        =   88
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CheckBox chkIncAfiNac 
            Caption         =   "Nacionalidad"
            Height          =   375
            Left            =   2040
            TabIndex        =   86
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkIncAfiIpss 
            Caption         =   "Carnet IPSS"
            Height          =   375
            Left            =   2040
            TabIndex        =   85
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkIncAfiSex 
            Caption         =   "Sexo"
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Top             =   2520
            Width           =   2775
         End
         Begin VB.CheckBox chkIncAfiCuspp 
            Caption         =   "Cuspp"
            Height          =   375
            Left            =   240
            TabIndex        =   83
            Top             =   2160
            Width           =   2775
         End
         Begin VB.CheckBox chkIncAfiFirRev 
            Caption         =   "Firmas al reverso"
            Height          =   375
            Left            =   2040
            TabIndex        =   81
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chkIncAfiFecNac 
            Caption         =   "Fecha Nac."
            Height          =   375
            Left            =   240
            TabIndex        =   80
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox chkIncAfiNroDoc 
            Caption         =   "Nro Documento"
            Height          =   375
            Left            =   240
            TabIndex        =   79
            Top             =   1800
            Width           =   2775
         End
         Begin VB.CheckBox chkIncAfiNom 
            Caption         =   "Nombres"
            Height          =   375
            Left            =   240
            TabIndex        =   78
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox chkIncAfiApe 
            Caption         =   "Apellidos"
            Height          =   375
            Left            =   240
            TabIndex        =   77
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.TextBox txtNumdocCus 
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
         Left            =   -74760
         MaxLength       =   12
         TabIndex        =   74
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox txtFecNacCus 
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
         Left            =   -72960
         MaxLength       =   10
         TabIndex        =   72
         Top             =   6480
         Width           =   1095
      End
      Begin VB.TextBox txtApePatCus 
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
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   70
         Top             =   5880
         Width           =   6135
      End
      Begin VB.CheckBox chCambios 
         Caption         =   "Hacer Cambios"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   3600
         Width           =   1935
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5175
         Left            =   -67680
         TabIndex        =   52
         Top             =   4200
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9128
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Internet"
         TabPicture(0)   =   "frmValidaciones.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "WebBrowser1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Base Local"
         TabPicture(1)   =   "frmValidaciones.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(1)=   "dtgRuc"
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame1 
            Caption         =   "Datos"
            ForeColor       =   &H00C00000&
            Height          =   1815
            Left            =   -74880
            TabIndex        =   55
            Top             =   360
            Width           =   7455
            Begin VB.TextBox txtNumdocLc 
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
               Left            =   4920
               TabIndex        =   68
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txtApematLc 
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
               Left            =   1680
               TabIndex        =   61
               Top             =   960
               Width           =   1815
            End
            Begin VB.CommandButton cmdBuscar 
               Caption         =   "&Buscar"
               Height          =   375
               Left            =   5880
               TabIndex        =   60
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox txtApepatLc 
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
               Left            =   1680
               TabIndex        =   59
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtCuspplc 
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
               Left            =   1680
               TabIndex        =   58
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtNombre2Lc 
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
               Left            =   4920
               TabIndex        =   57
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtNombre1Lc 
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
               Left            =   4920
               TabIndex        =   56
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "DNI :"
               Height          =   255
               Left            =   3480
               TabIndex        =   69
               Top             =   1060
               Width           =   1335
            End
            Begin VB.Label lblRazSoc 
               Caption         =   "Apellido Materno :"
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label20 
               Caption         =   "Apellido Paterno :"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "CUSPP :"
               Height          =   255
               Left            =   240
               TabIndex        =   64
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "Nombre 2 :"
               Height          =   255
               Left            =   3480
               TabIndex        =   63
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "Nombre 1 :"
               Height          =   255
               Left            =   3480
               TabIndex        =   62
               Top             =   360
               Width           =   1335
            End
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   4815
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   7455
            ExtentX         =   13150
            ExtentY         =   8493
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin MSDataGridLib.DataGrid dtgRuc 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   54
            Top             =   2280
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4895
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
            ColumnCount     =   18
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "Apepat"
               Caption         =   "Ape Paterno"
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
               DataField       =   "apemat"
               Caption         =   "Ape. Materno"
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
            BeginProperty Column04 
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
            BeginProperty Column05 
               DataField       =   "numruc"
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
            BeginProperty Column06 
               DataField       =   "razsoc"
               Caption         =   "Razón Social"
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
               DataField       =   "Direcc"
               Caption         =   "Dirección"
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
               DataField       =   "CodPos"
               Caption         =   "Cod Pos"
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
               DataField       =   "codgeo"
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
            BeginProperty Column10 
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
            BeginProperty Column11 
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
            BeginProperty Column12 
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
            BeginProperty Column13 
               DataField       =   "Jubilado"
               Caption         =   "Jubilado"
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
               DataField       =   "tipdoc"
               Caption         =   "Tipo Doc."
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
            BeginProperty Column16 
               DataField       =   "ipss"
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
            BeginProperty Column17 
               DataField       =   "fecnac"
               Caption         =   "Fecha Nac."
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
                  ColumnWidth     =   1244.976
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
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2894.74
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   2264.882
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   1379.906
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   659.906
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   810.142
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column16 
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column17 
                  ColumnWidth     =   975.118
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtCusSbs 
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
         Left            =   -69600
         MaxLength       =   12
         TabIndex        =   26
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox txtAfpOri 
         Height          =   330
         Left            =   -70200
         MaxLength       =   2
         TabIndex        =   46
         Top             =   5280
         Width           =   495
      End
      Begin VB.CheckBox chkOtraAfp 
         Caption         =   "Otra AFP"
         Height          =   375
         Left            =   -69240
         TabIndex        =   43
         Top             =   5160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkExistSpp 
         Caption         =   "Existe en SPP"
         Height          =   375
         Left            =   -69240
         TabIndex        =   42
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTipDoc 
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
         Left            =   -72000
         MaxLength       =   23
         TabIndex        =   23
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtUbiNacRENIEC 
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
         Left            =   -73200
         MaxLength       =   6
         TabIndex        =   29
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtUbiNacINEI 
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
         Left            =   -71520
         MaxLength       =   6
         TabIndex        =   28
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtFecNac 
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
         Left            =   -74760
         MaxLength       =   10
         TabIndex        =   27
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtSexo 
         Alignment       =   2  'Center
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
         Left            =   -70080
         MaxLength       =   1
         TabIndex        =   25
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtNumdoc 
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
         Left            =   -71280
         MaxLength       =   12
         TabIndex        =   24
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox txtNom2 
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
         Left            =   -73320
         MaxLength       =   10
         TabIndex        =   22
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtNom1 
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
         Left            =   -74760
         MaxLength       =   10
         TabIndex        =   21
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox txtApeMat 
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
         Left            =   -70080
         MaxLength       =   20
         TabIndex        =   20
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtApepat 
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
         Left            =   -72120
         MaxLength       =   20
         TabIndex        =   19
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtCuspp 
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
         Left            =   -73560
         MaxLength       =   12
         TabIndex        =   18
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtNumsol 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   17
         Top             =   4080
         Width           =   1095
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid3 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   16
         Top             =   6960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Solicitud"
         Columns(0).DataField=   "Numsol"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Doc. Iden."
         Columns(1).DataField=   "Numdoc"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Cuspp"
         Columns(2).DataField=   "CUSPP"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Ape. Paterno"
         Columns(3).DataField=   "ApePat"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Ape. Materno"
         Columns(4).DataField=   "ApeMat"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nombre1"
         Columns(5).DataField=   "Nombre1"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Nombre2"
         Columns(6).DataField=   "Nombre2"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).ShowCollapseExpandIcons=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1535"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1508"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1429"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2223"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2143"
         Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=260"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=1799"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1720"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=260"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=1931"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1852"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=260"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(25)=   "Column(5).Width=1879"
         Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1799"
         Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=260"
         Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(30)=   "Column(6).Width=1640"
         Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1561"
         Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=260"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   1289.764
         ViewColumnWidth =   5745.26
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Arial"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Arial"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=25,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=44,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=27,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=30,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=29,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=31,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=32,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=43,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=45,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=46,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=25"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=26,.alignment=0"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=27"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=29"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=16,.parent=25"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=26"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=27"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=29"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=25"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=26,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=27"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=29"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=25"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=26,.alignment=0"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=27"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=29"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=110,.parent=25"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=26,.alignment=0"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=27"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=29"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=114,.parent=25"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=111,.parent=26,.alignment=0"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=112,.parent=27"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=113,.parent=29"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=118,.parent=25"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=115,.parent=26,.alignment=0"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=116,.parent=27"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=117,.parent=29"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Credenciales  &Web"
         Height          =   375
         Left            =   -67800
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Actualizar"
         Height          =   375
         Index           =   3
         Left            =   -62280
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Detener"
         Height          =   375
         Index           =   2
         Left            =   -63360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   ">"
         Height          =   375
         Index           =   1
         Left            =   -64440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "<"
         Height          =   375
         Index           =   0
         Left            =   -65520
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3720
         Width           =   975
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit1 
         Height          =   5175
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   14655
         _Version        =   131074
         _ExtentX        =   25850
         _ExtentY        =   9128
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit2"
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   438978816
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
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Height          =   3735
         Left            =   360
         TabIndex        =   6
         Top             =   5640
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   6588
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nro Solicitud"
         Columns(0).DataField=   "Numsol"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Doc. Iden."
         Columns(1).DataField=   "Numdoc"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Firma Representante AFP"
         Columns(2).DataField=   "IndFirRep"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Sello de Repres.AFP"
         Columns(3).DataField=   "IndSellRep"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Lugar Suscripción"
         Columns(4).DataField=   "IndLugSus"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Firma Afiliado"
         Columns(5).DataField=   "IndFirAfi"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   4
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Firma Promotor"
         Columns(6).DataField=   "IndFirPro"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Sello de Recepción"
         Columns(7).DataField=   "IndSelRec"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   4
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Huella Digital"
         Columns(8).DataField=   "IndHueDig"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   4
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Firma Fedatario"
         Columns(9).DataField=   "IndFirFed"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   4
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "EECC Via Mail"
         Columns(10).DataField=   "IndEnvMail"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   18
         Columns(11)._MaxComboItems=   3
         Columns(11).ValueItems(0)._DefaultItem=   -1
         Columns(11).ValueItems(0).Value=   "0"
         Columns(11).ValueItems(0).Value.vt=   8
         Columns(11).ValueItems(0).DisplayValue=   "S"
         Columns(11).ValueItems(0).DisplayValue.vt=   8
         Columns(11).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(11).ValueItems(1)._DefaultItem=   0
         Columns(11).ValueItems(1).Value=   "1"
         Columns(11).ValueItems(1).Value.vt=   8
         Columns(11).ValueItems(1).DisplayValue=   "P"
         Columns(11).ValueItems(1).DisplayValue.vt=   8
         Columns(11).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(11).ValueItems(2)._DefaultItem=   0
         Columns(11).ValueItems(2).Value=   "2"
         Columns(11).ValueItems(2).Value.vt=   8
         Columns(11).ValueItems(2).DisplayValue=   "T"
         Columns(11).ValueItems(2).DisplayValue.vt=   8
         Columns(11).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(11).ValueItems.Count=   3
         Columns(11).Caption=   "Email Particular"
         Columns(11).DataField=   "IndMail"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Pos"
         Columns(12).DataField=   "Posicion"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   50
         Columns(13)._MaxComboItems=   5
         Columns(13).ValueItems(0)._DefaultItem=   0
         Columns(13).ValueItems(0).Value=   "S"
         Columns(13).ValueItems(0).Value.vt=   8
         Columns(13).ValueItems(0).DisplayValue=   "S"
         Columns(13).ValueItems(0).DisplayValue.vt=   8
         Columns(13).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(13).ValueItems(1)._DefaultItem=   0
         Columns(13).ValueItems(1).Value=   "N"
         Columns(13).ValueItems(1).Value.vt=   8
         Columns(13).ValueItems(1).DisplayValue=   "N"
         Columns(13).ValueItems(1).DisplayValue.vt=   8
         Columns(13).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(13).ValueItems.Count=   2
         Columns(13).Caption=   "Comprobacion Boleta de Pago"
         Columns(13).DataField=   "BolPago"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   18
         Columns(14)._MaxComboItems=   5
         Columns(14).ValueItems(0)._DefaultItem=   0
         Columns(14).ValueItems(0).Value=   "S"
         Columns(14).ValueItems(0).Value.vt=   8
         Columns(14).ValueItems(0).DisplayValue=   "S"
         Columns(14).ValueItems(0).DisplayValue.vt=   8
         Columns(14).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(14).ValueItems(1)._DefaultItem=   0
         Columns(14).ValueItems(1).Value=   "N"
         Columns(14).ValueItems(1).Value.vt=   8
         Columns(14).ValueItems(1).DisplayValue=   "N"
         Columns(14).ValueItems(1).DisplayValue.vt=   8
         Columns(14).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(14).ValueItems.Count=   2
         Columns(14).Caption=   "Copia Simple Doc. Iden."
         Columns(14).DataField=   "CopiaSim"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerStyle=   1
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131585"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1905"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1826"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131585"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2408"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2328"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131585"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2011"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1931"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=131585"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2064"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1984"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=131585"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1720"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1640"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=131585"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2275"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2196"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=131585"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2196"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2117"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=131585"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=1693"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1614"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=131585"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=1931"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1852"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=131585"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(51)=   "Column(10).Width=1746"
         Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1667"
         Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=131585"
         Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(56)=   "Column(11).Width=2355"
         Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2275"
         Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=131585"
         Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(61)=   "Column(12).Width=2381"
         Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=2302"
         Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=131585"
         Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(66)=   "Column(13).Width=2487"
         Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2408"
         Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=131585"
         Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(71)=   "Column(14).Width=2196"
         Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2117"
         Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=131585"
         Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   2
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Arial"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Arial"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Arial"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7,.alignment=2,.valignment=2"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=74,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(98)  =   "Named:id=33:Normal"
         _StyleDefs(99)  =   ":id=33,.parent=0,.alignment=2"
         _StyleDefs(100) =   "Named:id=34:Heading"
         _StyleDefs(101) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(102) =   ":id=34,.wraptext=-1"
         _StyleDefs(103) =   "Named:id=35:Footing"
         _StyleDefs(104) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(105) =   "Named:id=36:Selected"
         _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(107) =   "Named:id=37:Caption"
         _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(109) =   "Named:id=38:HighlightRow"
         _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(111) =   "Named:id=39:EvenRow"
         _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(113) =   "Named:id=40:OddRow"
         _StyleDefs(114) =   ":id=40,.parent=33"
         _StyleDefs(115) =   "Named:id=41:RecordSelector"
         _StyleDefs(116) =   ":id=41,.parent=34"
         _StyleDefs(117) =   "Named:id=42:FilterBar"
         _StyleDefs(118) =   ":id=42,.parent=33"
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit2 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   14775
         _Version        =   131074
         _ExtentX        =   26061
         _ExtentY        =   5318
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit2"
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   438972928
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
      Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
         Height          =   5595
         Left            =   -74760
         TabIndex        =   49
         Top             =   600
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   9869
         _LayoutType     =   4
         _RowHeight      =   20
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nro Solicitud"
         Columns(0).DataField=   "Numsol"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Doc. Iden."
         Columns(1).DataField=   "NumDoc"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   16
         Columns(2)._MaxComboItems=   5
         Columns(2).ValueItems(0)._DefaultItem=   0
         Columns(2).ValueItems(0).Value=   "0"
         Columns(2).ValueItems(0).Value.vt=   8
         Columns(2).ValueItems(0).DisplayValue=   "1"
         Columns(2).ValueItems(0).DisplayValue.vt=   8
         Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(1)._DefaultItem=   0
         Columns(2).ValueItems(1).Value=   "1"
         Columns(2).ValueItems(1).Value.vt=   8
         Columns(2).ValueItems(1).DisplayValue=   "2"
         Columns(2).ValueItems(1).DisplayValue.vt=   8
         Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(2)._DefaultItem=   0
         Columns(2).ValueItems(2).Value=   "2"
         Columns(2).ValueItems(2).Value.vt=   8
         Columns(2).ValueItems(2).DisplayValue=   "3"
         Columns(2).ValueItems(2).DisplayValue.vt=   8
         Columns(2).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(3)._DefaultItem=   0
         Columns(2).ValueItems(3).Value=   "3"
         Columns(2).ValueItems(3).Value.vt=   8
         Columns(2).ValueItems(3).DisplayValue=   "4"
         Columns(2).ValueItems(3).DisplayValue.vt=   8
         Columns(2).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(4)._DefaultItem=   0
         Columns(2).ValueItems(4).Value=   "4"
         Columns(2).ValueItems(4).Value.vt=   8
         Columns(2).ValueItems(4).DisplayValue=   "5"
         Columns(2).ValueItems(4).DisplayValue.vt=   8
         Columns(2).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(5)._DefaultItem=   0
         Columns(2).ValueItems(5).Value=   "5"
         Columns(2).ValueItems(5).Value.vt=   8
         Columns(2).ValueItems(5).DisplayValue=   "6"
         Columns(2).ValueItems(5).DisplayValue.vt=   8
         Columns(2).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(6)._DefaultItem=   0
         Columns(2).ValueItems(6).Value=   "6"
         Columns(2).ValueItems(6).Value.vt=   8
         Columns(2).ValueItems(6).DisplayValue=   "7"
         Columns(2).ValueItems(6).DisplayValue.vt=   8
         Columns(2).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(7)._DefaultItem=   0
         Columns(2).ValueItems(7).Value=   "7"
         Columns(2).ValueItems(7).Value.vt=   8
         Columns(2).ValueItems(7).DisplayValue=   "8"
         Columns(2).ValueItems(7).DisplayValue.vt=   8
         Columns(2).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(8)._DefaultItem=   0
         Columns(2).ValueItems(8).Value=   "8"
         Columns(2).ValueItems(8).Value.vt=   8
         Columns(2).ValueItems(8).DisplayValue=   "9"
         Columns(2).ValueItems(8).DisplayValue.vt=   8
         Columns(2).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(9)._DefaultItem=   0
         Columns(2).ValueItems(9).Value=   "9"
         Columns(2).ValueItems(9).Value.vt=   8
         Columns(2).ValueItems(9).DisplayValue=   "10"
         Columns(2).ValueItems(9).DisplayValue.vt=   8
         Columns(2).ValueItems(9)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(10)._DefaultItem=   0
         Columns(2).ValueItems(10).Value=   "10"
         Columns(2).ValueItems(10).Value.vt=   8
         Columns(2).ValueItems(10).DisplayValue=   "11"
         Columns(2).ValueItems(10).DisplayValue.vt=   8
         Columns(2).ValueItems(10)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(11)._DefaultItem=   0
         Columns(2).ValueItems(11).Value=   "11"
         Columns(2).ValueItems(11).Value.vt=   8
         Columns(2).ValueItems(11).DisplayValue=   "12"
         Columns(2).ValueItems(11).DisplayValue.vt=   8
         Columns(2).ValueItems(11)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems.Count=   12
         Columns(2).Caption=   "Sec. Emp"
         Columns(2).DataField=   "Secu"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Falta DNI"
         Columns(3).DataField=   "IndFalDni"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DNI Caduco"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DNI Ilegible"
         Columns(5).DataField=   "IndDniLeg"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   4
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DNI Indicios de Fasificacion"
         Columns(6).DataField=   "IndDNIFal"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Huella Digital Ilegible"
         Columns(7).DataField=   "IndHueIle"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   4
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Huella Digital Diferente"
         Columns(8).DataField=   "IndHueDif"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   4
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Firma Afiliado Diferente"
         Columns(9).DataField=   "IndFirAfiDif"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   4
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Sin Boleta o Dec. Jurada"
         Columns(10).DataField=   "IndBolFal"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   4
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Falta Formato varios Empleadores"
         Columns(11).DataField=   "IndForAdi"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   4
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Formatos Incompletos"
         Columns(12).DataField=   "IndForInc"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   4
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Falta Hoja de Vida"
         Columns(13).DataField=   "IndHojVid"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   4
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Falta Declaracion + 50 años"
         Columns(14).DataField=   "IndFalDec"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   4
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Existen Enmendaduras"
         Columns(15).DataField=   "IndForEmn"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   4
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Datos Generales Diferentes"
         Columns(16).DataField=   "IndDGenDif"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   4
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Llenado Incorrecto Afiliado"
         Columns(17).DataField=   "IndLlenInc"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   4
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "Llenado Incorrecto Empleador "
         Columns(18).DataField=   "IndLlenIncEmp"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   4
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "Llenado Incorrecto Promotor"
         Columns(19).DataField=   "IndLlenIncPro"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   4
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "Llenado Incorrecto Ant. Previsio."
         Columns(20).DataField=   "IndLlenIncAnt"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   4
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "Llenado Incorrecto Verif. Sol."
         Columns(21).DataField=   "IndLlenIncVrS"
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   4
         Columns(22)._MaxComboItems=   5
         Columns(22).Caption=   "Fecha Suscripción reverso dif."
         Columns(22).DataField=   "IndFecRevr"
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   23
         Splits(0)._UserFlags=   0
         Splits(0).ShowCollapseExpandIcons=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=23"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1746"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1085"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1005"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1217"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1138"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1191"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1111"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1217"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1138"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1561"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1482"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=1535"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1455"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=1535"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1455"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=513"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=1561"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1482"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=513"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(51)=   "Column(10).Width=1508"
         Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1429"
         Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=513"
         Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(56)=   "Column(11).Width=1561"
         Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=1482"
         Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=513"
         Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(61)=   "Column(12).Width=1667"
         Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=1588"
         Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=513"
         Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(66)=   "Column(13).Width=1296"
         Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=1217"
         Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=513"
         Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(71)=   "Column(14).Width=1640"
         Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=1561"
         Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=513"
         Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(76)=   "Column(15).Width=1535"
         Splits(0)._ColumnProps(77)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(78)=   "Column(15)._WidthInPix=1455"
         Splits(0)._ColumnProps(79)=   "Column(15)._ColStyle=513"
         Splits(0)._ColumnProps(80)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(81)=   "Column(16).Width=1535"
         Splits(0)._ColumnProps(82)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(16)._WidthInPix=1455"
         Splits(0)._ColumnProps(84)=   "Column(16)._ColStyle=513"
         Splits(0)._ColumnProps(85)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(86)=   "Column(17).Width=1746"
         Splits(0)._ColumnProps(87)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(88)=   "Column(17)._WidthInPix=1667"
         Splits(0)._ColumnProps(89)=   "Column(17)._ColStyle=513"
         Splits(0)._ColumnProps(90)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(91)=   "Column(18).Width=1773"
         Splits(0)._ColumnProps(92)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(93)=   "Column(18)._WidthInPix=1693"
         Splits(0)._ColumnProps(94)=   "Column(18)._ColStyle=513"
         Splits(0)._ColumnProps(95)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(96)=   "Column(19).Width=1640"
         Splits(0)._ColumnProps(97)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(98)=   "Column(19)._WidthInPix=1561"
         Splits(0)._ColumnProps(99)=   "Column(19)._ColStyle=513"
         Splits(0)._ColumnProps(100)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(101)=   "Column(20).Width=1852"
         Splits(0)._ColumnProps(102)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(103)=   "Column(20)._WidthInPix=1773"
         Splits(0)._ColumnProps(104)=   "Column(20)._ColStyle=513"
         Splits(0)._ColumnProps(105)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(106)=   "Column(21).Width=1799"
         Splits(0)._ColumnProps(107)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(108)=   "Column(21)._WidthInPix=1720"
         Splits(0)._ColumnProps(109)=   "Column(21)._ColStyle=513"
         Splits(0)._ColumnProps(110)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(111)=   "Column(22).Width=1931"
         Splits(0)._ColumnProps(112)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(113)=   "Column(22)._WidthInPix=1852"
         Splits(0)._ColumnProps(114)=   "Column(22)._ColStyle=513"
         Splits(0)._ColumnProps(115)=   "Column(22).Order=23"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   3
         FootLines       =   1
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgpicMode=1,.bold=0,.fontsize=825"
         _StyleDefs(10)  =   ":id=4,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Arial"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=900"
         _StyleDefs(13)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Arial"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Arial"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=99,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=108,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=100,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=101,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=102,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=104,.parent=6,.alignment=2"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=103,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=105,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=106,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=107,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=109,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=110,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=114,.parent=99,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=111,.parent=100"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=112,.parent=101"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=113,.parent=103"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=16,.parent=99"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=100"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=101"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=103"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=54,.parent=99,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=100"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=101"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=103"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=118,.parent=99,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=100"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=101"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=103"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=20,.parent=99,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=100"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=101"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=103"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=122,.parent=99,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=119,.parent=100"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=120,.parent=101"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=121,.parent=103"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=126,.parent=99,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=123,.parent=100"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=124,.parent=101"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=125,.parent=103"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=134,.parent=99,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=131,.parent=100"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=132,.parent=101"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=133,.parent=103"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=138,.parent=99,.alignment=2"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=135,.parent=100"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=136,.parent=101"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=137,.parent=103"
         _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=142,.parent=99,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=139,.parent=100"
         _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=140,.parent=101"
         _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=141,.parent=103"
         _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=146,.parent=99,.alignment=2"
         _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=143,.parent=100"
         _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=144,.parent=101"
         _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=145,.parent=103"
         _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=150,.parent=99,.alignment=2"
         _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=147,.parent=100"
         _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=148,.parent=101"
         _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=149,.parent=103"
         _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=154,.parent=99,.alignment=2"
         _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=151,.parent=100"
         _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=152,.parent=101"
         _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=153,.parent=103"
         _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=158,.parent=99,.alignment=2"
         _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=155,.parent=100"
         _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=156,.parent=101"
         _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=157,.parent=103"
         _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=162,.parent=99,.alignment=2"
         _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=159,.parent=100"
         _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=160,.parent=101"
         _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=161,.parent=103"
         _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=166,.parent=99,.alignment=2"
         _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=163,.parent=100"
         _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=164,.parent=101"
         _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=165,.parent=103"
         _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=170,.parent=99,.alignment=2"
         _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=167,.parent=100"
         _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=168,.parent=101"
         _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=169,.parent=103"
         _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=174,.parent=99,.alignment=2"
         _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=171,.parent=100"
         _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=172,.parent=101"
         _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=173,.parent=103"
         _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=24,.parent=99,.alignment=2"
         _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=21,.parent=100"
         _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=22,.parent=101"
         _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=23,.parent=103"
         _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=28,.parent=99,.alignment=2"
         _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=25,.parent=100"
         _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=26,.parent=101"
         _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=27,.parent=103"
         _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=32,.parent=99,.alignment=2"
         _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=29,.parent=100"
         _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=30,.parent=101"
         _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=31,.parent=103"
         _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=46,.parent=99,.alignment=2"
         _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=43,.parent=100"
         _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=44,.parent=101"
         _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=45,.parent=103"
         _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=50,.parent=99,.alignment=2"
         _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=47,.parent=100"
         _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=48,.parent=101"
         _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=49,.parent=103"
         _StyleDefs(130) =   "Named:id=33:Normal"
         _StyleDefs(131) =   ":id=33,.parent=0"
         _StyleDefs(132) =   "Named:id=34:Heading"
         _StyleDefs(133) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(134) =   ":id=34,.wraptext=-1"
         _StyleDefs(135) =   "Named:id=35:Footing"
         _StyleDefs(136) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(137) =   "Named:id=36:Selected"
         _StyleDefs(138) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(139) =   "Named:id=37:Caption"
         _StyleDefs(140) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(141) =   "Named:id=38:HighlightRow"
         _StyleDefs(142) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(143) =   "Named:id=39:EvenRow"
         _StyleDefs(144) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(145) =   "Named:id=40:OddRow"
         _StyleDefs(146) =   ":id=40,.parent=33"
         _StyleDefs(147) =   "Named:id=41:RecordSelector"
         _StyleDefs(148) =   ":id=41,.parent=34"
         _StyleDefs(149) =   "Named:id=42:FilterBar"
         _StyleDefs(150) =   ":id=42,.parent=33"
      End
      Begin VB.Label Label26 
         Caption         =   "Num. Doc. CUS"
         Height          =   255
         Left            =   -74640
         TabIndex        =   75
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "Fec. Nacimiento CUS"
         Height          =   255
         Left            =   -72960
         TabIndex        =   73
         Top             =   6240
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "Ape. Paterno CUS"
         Height          =   255
         Left            =   -74640
         TabIndex        =   71
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "CUSPP SBS"
         Height          =   255
         Left            =   -69480
         TabIndex        =   50
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Afp Ori"
         Height          =   255
         Left            =   -70200
         TabIndex        =   47
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Tip Docu."
         Height          =   255
         Left            =   -72120
         TabIndex        =   41
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Ubi. Nac. RENIEC"
         Height          =   255
         Left            =   -73200
         TabIndex        =   40
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Ubi. Nac. INEI"
         Height          =   255
         Left            =   -71520
         TabIndex        =   39
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Fec. Nacimiento"
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   -70080
         TabIndex        =   37
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Num. Doc."
         Height          =   255
         Left            =   -71160
         TabIndex        =   36
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Nombre 2"
         Height          =   255
         Left            =   -73320
         TabIndex        =   35
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre 1"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Ape. Materno"
         Height          =   255
         Left            =   -69960
         TabIndex        =   33
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Ape. Paterno"
         Height          =   255
         Left            =   -72000
         TabIndex        =   32
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "CUSPP"
         Height          =   255
         Left            =   -73440
         TabIndex        =   31
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Solicitud"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   3840
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBuscarVal 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtLote 
      Alignment       =   2  'Center
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblIndDoc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   8040
      TabIndex        =   48
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Lote"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmValidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsImag As ADODB.Recordset
Private RsImagSBS As ADODB.Recordset
Private rsValidacionImag As ADODB.Recordset
Private rsValidacionFisica As ADODB.Recordset
Private rsValidacionSBS As ADODB.Recordset
Private TipoDocLote As Boolean  '0 Traspaso 1 Contrato
Private PosicionImag As Integer

Private Sub chCambios_Click()
If chCambios.Value = 1 Then
    ActDesCampos False
Else
    ActDesCampos True
End If
End Sub

Private Sub Check5_Click()

End Sub

Private Sub Check11_Click()

End Sub

Private Sub cmdActualizar_Click()
Select Case SSTab1.Tab
Case 0
    Call GrabarImag
Case 1
    Call GrabarFis
Case 2
    Call GrabarSBS
End Select


Call ActualizarLoteVal(TxtLote)



End Sub

Private Sub cmdBuscar_Click()
BuscarCusppLc
End Sub

Private Sub cmdBuscarVal_Click()

If Trim(TxtLote) = "" Then MsgBox "Debe ingresar un lote", 64: Exit Sub

If chkOmitir.Value = 0 Then
    If Not ValidarLote(TxtLote) Then Exit Sub
End If


'Select Case SSTab1.Tab
'Case 0
    ActualizaIndDig Trim(TxtLote)
    BuscarValIma
    CargarImagen
'Case 1
    BuscarFisicas
'Case 2
    BuscarSBS
    CargarImagenSBS
    RelacionarDatosSBS
'End Select

End Sub

Sub CargarImagen()
On Error Resume Next

Dim strsql As String

strsql = " Select imagen, c.Ruta, LoteFisico, fecsal,posicion " & _
         " from _Imagenes b join _Volumenes c on c.idvolumen = b.idvolumen " & _
         " join CtrlDoc d on d.Lote = b.Lotefisico " & _
         " where barcode = 0 and b.Lotefisico = '" & Trim(TxtLote.Text) & "' " & _
         " and idimagen = '01" & Trim(TDBGrid1.Columns(0)) & "' and posicion = " & rsValidacionImag.Fields("posicion") & " "
                  
         'and  b.posicion = " & Trim(TDBGrid1.Columns(12)) & "

    If RsImag.State = 1 Then RsImag.Close
    RsImag.Open strsql, Con.Cnx, adOpenStatic, adLockReadOnly
     
    If Not (RsImag.EOF And RsImag.BOF) Then
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, vbGLteAfp, rsImag!imagen)
        ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
        ImgEdit1.Enabled = True
        
        'If vbGTippag = "04" Or vbGTippag = "06" Then
            ImgEdit1.FitTo 1
           ' ImgEdit1.Zoom = 20
            
            ImgEdit1.Display
        'End If
                                  
        If Err.Number > 0 Then ImgEdit1.Enabled = False
        
        'ImgEdit1.Refresh
        ImgEdit1.ScrollPositionY = Int((ImgEdit1.ImageScaleHeight / 3)) * 1.8
        
     Else
        ImgEdit1.Enabled = False
     End If
End Sub

Sub CargarImagenSBS()
On Error Resume Next

Dim strsql As String

strsql = " Select imagen, c.Ruta, LoteFisico, fecsal,posicion " & _
         " from _Imagenes b join _Volumenes c on c.idvolumen = b.idvolumen " & _
         " join CtrlDoc d on d.Lote = b.Lotefisico " & _
         " where barcode = 0 and b.Lotefisico = '" & Trim(TxtLote.Text) & "' " & _
         " and idimagen = '01" & Trim(TDBGrid3.Columns(0)) & "' order by  posicion "

    If RsImagSBS.State = 1 Then RsImagSBS.Close
    RsImagSBS.Open strsql, Con.Cnx, adOpenStatic, adLockReadOnly
     
    If Not (RsImagSBS.EOF And RsImagSBS.BOF) Then
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, vbGLteAfp, rsImag!imagen)
        ImgEdit2.Image = Rutaimag(RsImagSBS!Ruta, RsImagSBS!lotefisico, RsImagSBS!Imagen)
        ImgEdit2.Enabled = True
        
        'If vbGTippag = "04" Or vbGTippag = "06" Then
            ImgEdit2.FitTo 1
            ImgEdit2.Display
        'End If
                                  
        If Err.Number > 0 Then ImgEdit2.Enabled = False
        
        'ImgEdit1.Refresh
        ImgEdit2.ScrollPositionY = Int((ImgEdit2.ImageScaleHeight / 3)) * 0.2
        
     Else
        ImgEdit2.Enabled = False
     End If
End Sub
Sub BuscarFisicas()
Dim sqltexto As String

sqltexto = " select NumSol, Numdoc,Secu, posicion, IndFalDni, IndDniLeg, IndDniCad, IndDNIFal, IndSellONP,  IndFirAfi," & _
           " IndFirAfiDif, IndFirPro, IndSelRec, IndHueDig, IndHueIle,IndHueDif, " & _
           " IndFirFed, IndFecVer, IndFecDec, IndFirRep, IndSellRep, IndBolFal, IndForAdi, IndForInc, " & _
           " IndHojVid, IndFalDec, IndForEmn, IndDGenDif, IndLlenInc,IndLlenIncEmp,IndLlenIncPro,IndLlenIncAnt,IndLlenIncVrS,IndFecRevr, indIAApel," & _
           " indIANomb, indIAFecNac,indIATipDoc,indIANroDoc,indIACuspp,indIASexo,indIAIpss, " & _
           " indIANacio,indIADirAfi,indIAFirRev,indIAUbig,indIAFalDat,indIERuc,indIERazSoc, " & _
           " indIENroEmp,indIEDir,indIEUbigeo,indIVFecSus,indIVTipFon,indIVDecTra,indIVVerSol, " & _
           " indEEDatAfi, indEEDatEmp,indEEDatPro,indEEAntPrev,indEEDecSol,indDDApeAfi,indDDNomAfi, " & _
           " indDDFecNac, indDDTipDoc,indDDNumDoc  from solicitud " & _
           " where Lote = '" & Trim(TxtLote.Text) & "'  "

If chkFedOblig.Value = 1 Then
    sqltexto = sqltexto & " and  CodPromo in (Select CodPromo from TPromotor where isnull(TipSanc,'') <> '' ) "
End If


If chk50.Value = 1 Then
    sqltexto = sqltexto & " and  year(getdate()) - year(Fecnac) >= 50 "
End If


'Relacionar datos check


If rsValidacionFisica.State = 1 Then rsValidacionFisica.Close
rsValidacionFisica.Open sqltexto, Con.Cnx, adOpenStatic, adLockBatchOptimistic

RelacionarDatosIndicadores

Set TDBGrid2.DataSource = rsValidacionFisica
'LblNumReg.Caption = rsValidacionFisica.RecordCount

End Sub

Sub BuscarValIma()
Dim sqltexto As String

sqltexto = " select numsol , inddoc , numdoc, IndFirRep, IndSellRep, IndSellRep, IndLugSus, IndFirAfi, IndFirPro, " & _
           " IndSelRec, IndHueDig, IndFirFed, IndEnvMail , IndMail, posicion, BolPago, CopiaSim from solicitud " & _
           " where Lote = '" & Trim(TxtLote.Text) & "' "


If chkFedOblig.Value = 1 Then
    sqltexto = sqltexto & " and  CodPromo in (Select CodPromo from TPromotor where isnull(TipSanc,'') <> '' ) "
End If
'chkBoleta


If chkBoleta.Value = 1 Then
    sqltexto = sqltexto & " and indDoc = 'T' and ( (BolPago = 'N' and OtroDoc = 'N') or  (CopiaSim = 'N') )"
End If


If rsValidacionImag.State = 1 Then rsValidacionImag.Close
rsValidacionImag.Open sqltexto, Con.Cnx, adOpenStatic, adLockBatchOptimistic

Set TDBGrid1.DataSource = rsValidacionImag

PosicionImag = 0

If rsValidacionImag.RecordCount > 0 Then
        
        If rsValidacionImag.Fields("Inddoc") = "C" Then
            TDBGrid1.Columns(2).Visible = False
            TDBGrid1.Columns(3).Visible = False
            TDBGrid1.Columns(7).Visible = False
            TDBGrid1.Columns(13).Visible = False
            TDBGrid1.Columns(14).Visible = False
            
            lblIndDoc.Caption = "AFILIACIONES"
            RevCamposSbs False
            
            TDBGrid3.Top = 5880
            TDBGrid3.Height = 4095
        Else
            TDBGrid1.Columns(2).Visible = True
            TDBGrid1.Columns(3).Visible = True
            TDBGrid1.Columns(7).Visible = True
            
            TDBGrid1.Columns(13).Visible = True
            TDBGrid1.Columns(14).Visible = True
                                    
            lblIndDoc.Caption = "TRASPASOS"
            RevCamposSbs True
            
            TDBGrid3.Top = 6960
            TDBGrid3.Height = 3015
            
        End If
        
        PosicionImag = rsValidacionImag.Fields("posicion")

End If

 
'LblNumReg.Caption = rsValidacionImag.RecordCount


End Sub

Sub RevCamposSbs(vbhabilitar As Boolean)
txtApePatCus.Visible = vbhabilitar
txtNumdocCus.Visible = vbhabilitar
txtFecNacCus.Visible = vbhabilitar
Label21.Visible = vbhabilitar
Label26.Visible = vbhabilitar
Label25.Visible = vbhabilitar
End Sub
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTerminar_Click()
Call TerminarValidacion
End Sub

Private Sub cmdToolBar_Click(Index As Integer)
 On Error GoTo mensajeError

Select Case Index
 
  Case 0
    WebBrowser1.GoBack 'ir a la página anterior
  Case 1
    WebBrowser1.GoForward 'Siguiente página
  Case 2
    WebBrowser1.Stop 'Detener la carga de la página
  Case 3
    WebBrowser1.Refresh ' Actualizar la página
  Case 4
    WebBrowser1.GoHome 'Ir a la página de inicio
End Select


Exit Sub
'rutina de Error
mensajeError:
If Err.Number = -2147467259 Then
   Resume Next
MsgBox Err.Description
End If
End Sub

'Private Sub Command1_Click()
'
'End Sub

Private Sub Command1_Click()
 WebBrowser1.SetFocus
'TabSig 13

 'keybd_event VK_TAB, 0, 0, 0
    SendKeys "{TAB}"
    SendKeys "{TAB}"
     SendKeys "20100100399"
      SendKeys "{TAB}"
     SendKeys "8492"
     SendKeys "{ENTER}"
'     For i = 0 To 50000
'     Next
    Command1.Enabled = False
End Sub

Private Sub Form_Activate()
'Call Command1_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TxtLote.Text <> "" Then
    Select Case SSTab1.Tab
    Case 0
        MovImg KeyCode
    Case 2
        MovImgSBS KeyCode
    End Select
End If
End Sub

Private Sub Form_Load()
Set rsValidacionImag = New ADODB.Recordset
Set rsValidacionFisica = New ADODB.Recordset
Set rsValidacionSBS = New ADODB.Recordset
Set RsImag = New ADODB.Recordset
Set RsImagSBS = New ADODB.Recordset

rsValidacionImag.CursorLocation = adUseClient
rsValidacionFisica.CursorLocation = adUseClient

'WebBrowser1.ShowBrowserBar 1
WebBrowser1.AddressBar = True
WebBrowser1.Navigate ("http://www.sbs.gob.pe/0/modulos/JER/JER_Interna.aspx?ARE=0&PFL=1&JER=272")
chk50.Enabled = False
ActDesCampos True

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 If SSTab1.Tab = 2 Then
    txtNumsol.SetFocus
 End If
 
 If SSTab1.Tab = 1 Then
    chk50.Enabled = True
 Else
    chk50.Enabled = False
 End If
 
 If SSTab1.Tab = 0 Then
    chkBoleta.Enabled = True
 Else
    chkBoleta.Enabled = False
 End If
 
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 11 Then
'    If IIf(IsNull(rsValidacionImag.Fields("IndEnvMail")), 0, rsValidacionImag.Fields("IndEnvMail")) = False Then
'        rsValidacionImag.Fields("IndMail") = False
'        TDBGrid1.Columns(11).Value = False
'    End If
End If
End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'If ColIndex = 11 Then
'    If IIf(IsNull(rsValidacionImag.Fields("IndEnvMail")), 0, rsValidacionImag.Fields("IndEnvMail")) <> 1 Then
'        rsValidacionImag.Fields("IndMail") = 0
'    End If
'End If
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'If ColIndex = 11 Then
'    If IIf(IsNull(rsValidacionImag.Fields("IndEnvMail")), 0, rsValidacionImag.Fields("IndEnvMail")) <> 1 Then
'        rsValidacionImag.Fields("IndMail") = 0
'    End If
'End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If (TDBGrid1.Row + 1) <> LastRow Then
    If TxtLote.Text <> "" Then CargarImagen
End If

End Sub

Private Sub TDBGrid3_Click()
chCambios.Value = 0
Call CargarImagenSBS
End Sub

Private Sub TDBGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'ActDesCampos True
End Sub

Private Sub TxtApeMat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtCuspp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFecNac_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
FVFecha txtFecNac, KeyAscii
End Sub

Private Sub txtFecNac_LostFocus()
ValidaFecha txtFecNac
End Sub

Private Sub TxtLote_LostFocus()
TxtLote = Format(TxtLote, "0000000#")
End Sub

Sub MovImg(KeyCode As Integer)
 On Error Resume Next
If KeyCode = vbKeyPageDown Then
   ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY + 400
End If
If KeyCode = vbKeyPageUp Then
   ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY - 400
End If
If KeyCode = vbKeyRight Then
   ImgEdit1.ScrollPositionX = ImgEdit1.ScrollPositionX + 400
End If
If KeyCode = vbKeyLeft Then
   ImgEdit1.ScrollPositionX = ImgEdit1.ScrollPositionX - 400
End If
If KeyCode = vbKeyDown Then
   ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY + 100
End If
If KeyCode = vbKeyUp Then
   ImgEdit1.ScrollPositionY = ImgEdit1.ScrollPositionY - 100
End If

If KeyCode = vbKeyF12 Then
    If Not (RsImag.EOF And RsImag.BOF) Then
      RsImag.MoveNext
      If RsImag.EOF Then RsImag.MoveLast
      'ImgEdit1.Image = Rutaimag(rsImag!Ruta, vbGLteAfp, rsImag!imagen)
      ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      txtnumpag = RsImag!Posicion
      ImgEdit1.Display
            
      ImgEdit1.FitTo 1
      ImgEdit1.Refresh
      
 
     ' Call VerPlanillaScala
      'Call MovLinea
    End If
End If

If KeyCode = vbKeyF11 Then
    If Not (RsImag.EOF And RsImag.BOF) Then
      RsImag.MovePrevious
      If RsImag.BOF Then RsImag.MoveFirst
      'ImgEdit1.Image = Rutaimag(rsImag!Ruta, vbGLteAfp, rsImag!imagen)
      ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      txtnumpag = RsImag!Posicion
      ImgEdit1.Display
      
      ImgEdit1.FitTo 1
      ImgEdit1.Refresh
      
  
      'Call MovLinea
    End If
End If


End Sub

Sub MovImgSBS(KeyCode As Integer)
 On Error Resume Next
If KeyCode = vbKeyPageDown Then
   ImgEdit2.ScrollPositionY = ImgEdit2.ScrollPositionY + 400
End If
If KeyCode = vbKeyPageUp Then
   ImgEdit2.ScrollPositionY = ImgEdit2.ScrollPositionY - 400
End If
If KeyCode = vbKeyRight Then
   ImgEdit2.ScrollPositionX = ImgEdit2.ScrollPositionX + 400
End If
If KeyCode = vbKeyLeft Then
   ImgEdit2.ScrollPositionX = ImgEdit2.ScrollPositionX - 400
End If
If KeyCode = vbKeyDown Then
   ImgEdit2.ScrollPositionY = ImgEdit2.ScrollPositionY + 100
End If
If KeyCode = vbKeyUp Then
   ImgEdit2.ScrollPositionY = ImgEdit2.ScrollPositionY - 100
End If

If KeyCode = vbKeyF12 Then
    If Not (RsImagSBS.EOF And RsImagSBS.BOF) Then
      RsImagSBS.MoveNext
      If RsImagSBS.EOF Then RsImagSBS.MoveLast
      'ImgEdit2.Image = Rutaimag(RsImagSBS!Ruta, vbGLteAfp, RsImagSBS!imagen)
      ImgEdit2.Image = Rutaimag(RsImagSBS!Ruta, RsImagSBS!lotefisico, RsImagSBS!Imagen)
      txtnumpag = RsImagSBS!Posicion
      ImgEdit2.Display
            
      ImgEdit2.FitTo 1
      ImgEdit2.Refresh
      
 
     ' Call VerPlanillaScala
      'Call MovLinea
    End If
End If

If KeyCode = vbKeyF11 Then
    If Not (RsImagSBS.EOF And RsImagSBS.BOF) Then
      RsImagSBS.MovePrevious
      If RsImagSBS.BOF Then RsImagSBS.MoveFirst
      'ImgEdit2.Image = Rutaimag(RsImagSBS!Ruta, vbGLteAfp, RsImagSBS!imagen)
      ImgEdit2.Image = Rutaimag(RsImagSBS!Ruta, RsImagSBS!lotefisico, RsImagSBS!Imagen)
      txtnumpag = RsImagSBS!Posicion
      ImgEdit2.Display
      
      ImgEdit2.FitTo 1
      ImgEdit2.Refresh
       
      'Call MovLinea
    End If
End If


End Sub

Sub GrabarImag()
'rsValidacionImag.Supports adUpdateBatch
On Error Resume Next
If rsValidacionImag.State = 1 Then rsValidacionImag.UpdateBatch
CambiarEstadoLote TxtLote, "8"

'29587

End Sub
Sub GrabarFis()
'Se validará
On Error Resume Next
With rsValidacionFisica
    If .State = 1 Then
      
    .MoveFirst
        
    Do While Not .EOF
         
        If .Fields("indIANomb") Or .Fields("indIAFecNac") Or .Fields("indIATipDoc") Or _
           .Fields("indIANroDoc") Or .Fields("indIACuspp") Or .Fields("indIASexo") Or _
           .Fields("indIAIpss") Or .Fields("indIANacio") Or .Fields("indIADirAfi") Or _
           .Fields("indIAFirRev") Or .Fields("indIAUbig") Or .Fields("indIAFalDat") Then
               .Fields("IndLlenInc") = 1
        End If
         
         
        If .Fields("indIERuc") Or .Fields("indIERazSoc") Or .Fields("indIENroEmp") Or _
           .Fields("indIEDir") Or .Fields("indIEUbigeo") Then
                .Fields("IndLlenIncEmp") = 1
        End If
         
        If .Fields("indIVFecSus") Or .Fields("indIVTipFon") Or .Fields("indIVDecTra") Or _
           .Fields("indIVVerSol") Then
                .Fields("IndLlenIncVrS") = 1
        End If
    
        If .Fields("indEEDatAfi") Or .Fields("indEEDatEmp") Or .Fields("indEEDatPro") Or _
           .Fields("indEEAntPrev") Or .Fields("indEEDecSol") Then
                      
           .Fields("IndForEmn") = 1
        End If
        
        
        If .Fields("indDDApeAfi") Or .Fields("indDDNomAfi") Or .Fields("indDDFecNac") Or _
           .Fields("indDDTipDoc") Or .Fields("indDDNumDoc") Then
           .Fields("IndDGenDif") = 1
        End If
        
        
        'Validacion Mensajes
        
        If .Fields("indIANomb") = 0 And .Fields("indIAFecNac") = 0 And .Fields("indIATipDoc") = 0 And _
           .Fields("indIANroDoc") = 0 And .Fields("indIACuspp") = 0 And .Fields("indIASexo") = 0 And _
           .Fields("indIAIpss") = 0 And .Fields("indIANacio") = 0 And .Fields("indIADirAfi") = 0 And _
           .Fields("indIAFirRev") = 0 And .Fields("indIAUbig") = 0 And .Fields("indIAFalDat") = 0 And .Fields("IndLlenInc") Then
                 MsgBox "Debe seleccionar el detalle para Llenado Incorrecto Afiliado", vbCritical, "Validación Física"
                 Exit Sub
        End If
        

        If .Fields("indIERuc") = 0 And .Fields("indIERazSoc") = 0 And .Fields("indIENroEmp") = 0 And _
           .Fields("indIEDir") = 0 And .Fields("indIEUbigeo") = 0 And .Fields("IndLlenIncEmp") Then
                MsgBox "Debe seleccionar el detalle para Llenado Incorrecto Datos del Empleador", vbCritical, "Validación Física"
                Exit Sub
        End If
        
        If .Fields("indIVFecSus") = 0 And .Fields("indIVTipFon") = 0 And .Fields("indIVDecTra") = 0 And _
           .Fields("indIVVerSol") = 0 And .Fields("IndLlenIncVrS") Then
                MsgBox "Debe seleccionar el detalle para Llenado Incorrecto Datos Verificación de Solicitud", vbCritical, "Validación Física"
                Exit Sub
        End If
        
        If .Fields("indEEDatAfi") = 0 And .Fields("indEEDatEmp") = 0 And .Fields("indEEDatPro") = 0 And _
           .Fields("indEEAntPrev") = 0 And .Fields("indEEDecSol") = 0 And .Fields("IndForEmn") Then
                MsgBox "Debe seleccionar el detalle para Existen Enmendaduras", vbCritical, "Validación Física"
                Exit Sub
        End If
        
        
        If .Fields("indDDApeAfi") = 0 And .Fields("indDDNomAfi") = 0 And .Fields("indDDFecNac") = 0 And _
           .Fields("indDDTipDoc") = 0 And .Fields("indDDNumDoc") = 0 And .Fields("IndDGenDif") Then
            MsgBox "Debe seleccionar el detalle para Datos Generales Diferentes", vbCritical, "Validación Física"
            Exit Sub
        End If
                        
        .MoveNext
    Loop
        
    .UpdateBatch
    
    End If

End With
End Sub
Sub GrabarSBS()
On Error Resume Next
If rsValidacionSBS.State = 1 Then rsValidacionSBS.UpdateBatch
End Sub

Function ValidarLote(Lote As String) As Boolean
Dim vbEstado As String
Dim vbdesEstado As String
vbEstado = fgDatosql("Select Estado from _scan where Lotefisico = '" & Lote & "' ")
If vbEstado = "7" Then
    ValidarLote = True
Else
    ValidarLote = False
    vbdesEstado = fgDatosql("Select desEstado from Estado where Estado = '" & vbEstado & "' ")
    MsgBox "El Lote " & Lote & " se encuentra en " & vbdesEstado & "", 64, Caption
End If

End Function

Sub BuscarSBS()
Dim sqltexto As String

sqltexto = " select NumSol, CodAfpOri, CUSPP, ApePat, ApeMat, Nombre1, Nombre2,TipoDoc, NumDoc, Sexo, FecNac, " & _
           " UbigeoNac, UbigeoRen, IndAfiSpp, IndAfiOtraAfp, CUSSBS , ApePatCUS, ApeMatCUS, Nombre1CUS, Nombre2CUS, " & _
           " FecNacCUS, NumDocCUS " & _
           " from Solicitud " & _
           " where Lote = '" & Trim(TxtLote.Text) & "'  "

'0 Traspaso 1 Contrato
Dim vbTipo As String

vbTipo = fgDatosql("Select indTipDoc from ctrldoc where Lote = '" & Trim(TxtLote.Text) & "'")

If vbTipo = "" Then
    Exit Sub
Else
    TipoDocLote = vbTipo
End If


If TipoDocLote Then
    'Contrato
    chkExistSpp.Visible = True
    chkOtraAfp.Visible = False
Else
    'Traspaso
    chkExistSpp.Visible = False
    chkOtraAfp.Visible = True
End If



If chkFedOblig.Value = 1 Then
sqltexto = sqltexto & " and  CodPromo in (Select CodPromo from TPromotor where isnull(TipSanc,'') <> '' ) "
End If

If rsValidacionSBS.State = 1 Then rsValidacionSBS.Close
rsValidacionSBS.Open sqltexto, Con.Cnx, adOpenStatic, adLockBatchOptimistic

Set TDBGrid3.DataSource = rsValidacionSBS
 
End Sub

Sub RelacionarDatosSBS()

Set txtNumsol.DataSource = rsValidacionSBS
txtNumsol.DataField = "Numsol"

Set txtCUSPP.DataSource = rsValidacionSBS
txtCUSPP.DataField = "cuspp"

Set txtApePat.DataSource = rsValidacionSBS
txtApePat.DataField = "ApePat"

Set txtApeMat.DataSource = rsValidacionSBS
txtApeMat.DataField = "Apemat"

Set txtNom1.DataSource = rsValidacionSBS
txtNom1.DataField = "Nombre1"

Set txtNom2.DataSource = rsValidacionSBS
txtNom2.DataField = "Nombre2"

Set txtTipDoc.DataSource = rsValidacionSBS
txtTipDoc.DataField = "TipoDoc"

Set txtNumDoc.DataSource = rsValidacionSBS
txtNumDoc.DataField = "Numdoc"

Set txtSexo.DataSource = rsValidacionSBS
txtSexo.DataField = "Sexo"

Set txtFecNac.DataSource = rsValidacionSBS
txtFecNac.DataField = "FecNac"

Set txtUbiNacINEI.DataSource = rsValidacionSBS
txtUbiNacINEI.DataField = "UbigeoNac"

Set txtUbiNacRENIEC.DataSource = rsValidacionSBS
txtUbiNacRENIEC.DataField = "UbigeoRen"

Set chkOtraAfp.DataSource = rsValidacionSBS
chkOtraAfp.DataField = "IndAfiOtraAfp"

Set chkExistSpp.DataSource = rsValidacionSBS
chkExistSpp.DataField = "IndAfiSpp"

Set txtAfpOri.DataSource = rsValidacionSBS
txtAfpOri.DataField = "CodAfpOri"

Set txtCusSbs.DataSource = rsValidacionSBS
txtCusSbs.DataField = "CUSSBS"

Set txtApePatCus.DataSource = rsValidacionSBS
txtApePatCus.DataField = "ApePatCUS"

Set txtNumdocCus.DataSource = rsValidacionSBS
txtNumdocCus.DataField = "NumdocCus"

Set txtFecNacCus.DataSource = rsValidacionSBS
txtFecNacCus.DataField = "FecnacCus"

End Sub

Private Sub TxtNom1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNom2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNumSol_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSexo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSexo_Validate(Cancel As Boolean)
If txtSexo = "F" Or txtSexo = "M" Then
 Cancel = False
 Else
 Cancel = True
End If
End Sub

Private Sub txtTipDoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUbiNacINEI_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'TabSig KeyAscii
If KeyAscii = 13 Then Call ValLugarNac(KeyAscii)
End Sub

Private Sub txtUbiNacRENIEC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
 Call ValLugarNac(KeyCode)
End If
End Sub

Private Sub txtUbiNacRENIEC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Function ValLugarNac(KeyAscii) As Boolean
Dim vCodGeo As String

If KeyAscii = 13 Then

If Trim(txtUbiNacRe) = "" Then
         MsgBox "Debe Ingresar Código de Reniec", 64, Caption
         Beep
         ValLugarNac = False
         txtUbiNacRENIEC.SetFocus
         Exit Function
End If

        vCodGeo = fgDatosql("select codgeo from ubigeo where codubic = '" & Trim(txtUbiNacINEI.Text) & "'")
                If vCodGeo <> "" Then
                'Consultar no existe razon en truc
                'txtRazSoc.Text = fgDatosql("Select RAZSOC from TRUC where CODAFP = '" & vbGAFP & "'and  NUMRUC = '" & txtRuc & "'")
                   ValLugarNac = True
                   lblMensaje.Caption = ""
                   txtUbiNacINEI = vCodGeo
                   txtUbiNacRENIEC.SetFocus
                   
                Else
                   MsgBox "Código de Lugar Reniec no se encuentra en la Tabla Ubigeo", vbCritical, Caption
                   ValLugarNac = False
                   txtUbiNacRENIEC.SetFocus
                   Beep
                End If
                              
End If
 If KeyAscii = vbKeyF9 Then
   'Fuerza el pase al sgte campo
    ValLugarNac = True
    txtNumsol.SetFocus
 End If

End Function

Sub TerminarValidacion()

If MsgBox(" ¿ Desea Concluir la valiciones para el Lote " & TxtLote & " ?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
    CambiarEstadoLote TxtLote, "4"
    
    Set RsImag.DataSource = Nothing
    Set RsImagSBS.DataSource = Nothing
    Set rsValidacionFisica.DataSource = Nothing
    Set rsValidacionSBS.DataSource = Nothing
    Set rsValidacionImag.DataSource = Nothing
    
    Set TDBGrid1.DataSource = Nothing
    Set TDBGrid2.DataSource = Nothing
    Set TDBGrid3.DataSource = Nothing
    
    
    Set txtNumsol.DataSource = Nothing
    Set txtCUSPP.DataSource = Nothing
    Set txtApePat.DataSource = Nothing
    Set txtApeMat.DataSource = Nothing
    Set txtNom1.DataSource = Nothing
    Set txtNom2.DataSource = Nothing
    Set txtTipDoc.DataSource = Nothing
    Set txtNumDoc.DataSource = Nothing
    Set txtSexo.DataSource = Nothing
    Set txtFecNac.DataSource = Nothing
    Set txtUbiNacINEI.DataSource = Nothing
    Set txtUbiNacRENIEC.DataSource = Nothing
    Set chkOtraAfp.DataSource = Nothing
    Set chkExistSpp.DataSource = Nothing
    Set txtCusSbs.DataSource = Nothing

    ImgEdit1.ClearDisplay
    ImgEdit2.ClearDisplay
    
    chkExistSpp.Visible = False
    chkOtraAfp.Visible = False
    
    Me.TxtLote.Text = ""
    
End If


End Sub

Sub ValidarAfpOri()

End Sub

Sub BuscarCusppLc()
On Error GoTo TRATA_ERROR
Dim rsResBan As ADODB.Recordset
Set rsResBan = New ADODB.Recordset

Dim lcCadena As String
Dim Condicion As String
Screen.MousePointer = MousePointerConstants.vbHourglass
DoEvents
Strsqltexto = " select codafp, numcus, jubilado, apepat,  apemat,  nombre1, nombre2, numruc, " & _
              " razsoc , direcc, codpos, a.codgeo, " & _
              " tipdoc , numdoc, ipss, fecnac, departamento , " & _
              " provincia , distrito " & _
              " from tcuspp a left join ubigeo b on a.codgeo = b.codgeo "
              
     lcCadena = Trim(Strsqltexto)
     If rsResBan.State = 1 Then
       rsResBan.Close
     End If
     Condicion = " Where "
    
     'Condicion = Condicion & " codafp = '" & vbGAFP & "' and "
      
        If Trim(txtCuspplc.Text) <> "" Then Condicion = Condicion & " numcus   Like '%" & Trim(txtCuspplc.Text) & "%' And "
        If Trim(txtApepatLc.Text) <> "" Then Condicion = Condicion & " apepat like '%" & Trim(txtApepatLc.Text) & "%' And "
        If Trim(txtApematLc.Text) <> "" Then Condicion = Condicion & " apemat Like '%" & Trim(txtApematLc.Text) & "%' And "
        If Trim(txtNombre1Lc.Text) <> "" Then Condicion = Condicion & " nombre1 Like '%" & Trim(txtNombre1Lc.Text) & "%' And "
        If Trim(txtNombre2Lc.Text) <> "" Then Condicion = Condicion & " nombre2 Like '%" & Trim(txtNombre2Lc.Text) & "%' And "
        If Trim(txtNumdocLc.Text) <> "" Then Condicion = Condicion & " numdoc = '" & Trim(txtNumdocLc.Text) & "' And "
        
         
    If Condicion = " Where " Then
       Condicion = ""
    Else
       Condicion = Left(Condicion, Len(Condicion) - 4)
    End If
    lcCadena = lcCadena & Condicion & "order by numcus"
    With rsResBan
         .CursorType = adOpenStatic
         .CursorLocation = adUseClient
         .LockType = adLockReadOnly
    End With
    
    If rsResBan.State = 1 Then rsResBan.Close
    'Set rsResBan = comandob.Execute
    rsResBan.Open lcCadena, Con.Cnx
    If Not (rsResBan.EOF And rsResBan.BOF) Then
     dtgRuc.Enabled = True
     Set dtgRuc.DataSource = rsResBan
     Else
     dtgRuc.Enabled = False
     Set dtgRuc.DataSource = Nothing
     MsgBox "No encontrado", 64, Caption
    End If
     dtgRuc.Refresh
     Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
TRATA_ERROR:
    MsgBox Err.Description
End Sub

Sub ActDesCampos(Value As Boolean)
    txtCUSPP.Locked = Value
    txtApePat.Locked = Value
    txtApeMat.Locked = Value
    txtNom1.Locked = Value
    txtNom2.Locked = Value
    txtTipDoc.Locked = Value
    txtNumDoc.Locked = Value
    txtSexo.Locked = Value
    txtFecNac.Locked = Value
    txtUbiNacRENIEC.Locked = Value
    txtUbiNacINEI.Locked = Value
    txtAfpOri.Locked = Value
End Sub


'
'lslTablaDE.Left = TbdEstruc.Left + TbdEstruc.Columns(1).Left
'   lslTablaDE.Top = TbdEstruc.Top + TbdEstruc.RowTop(TbdEstruc.Row) + TbdEstruc.RowHeight
'   lslTablaDE.Visible = True
Sub RelacionarDatosIndicadores()

    Set chkIncAfiApe.DataSource = rsValidacionFisica
    chkIncAfiApe.DataField = "indIAApel"
    
    Set chkIncAfiNom.DataSource = rsValidacionFisica
    chkIncAfiNom.DataField = "indIANomb"
    
    Set chkIncAfiFecNac.DataSource = rsValidacionFisica
    chkIncAfiFecNac.DataField = "indIAFecNac"
    
    Set chkIncAfiTipDoc.DataSource = rsValidacionFisica
    chkIncAfiTipDoc.DataField = "indIATipDoc"
    
    Set chkIncAfiNroDoc.DataSource = rsValidacionFisica
    chkIncAfiNroDoc.DataField = "indIANroDoc"
    
    
    Set chkIncAfiCuspp.DataSource = rsValidacionFisica
    chkIncAfiCuspp.DataField = "indIACuspp"
    
    Set chkIncAfiSex.DataSource = rsValidacionFisica
    chkIncAfiSex.DataField = "indIASexo"
    
    Set chkIncAfiIpss.DataSource = rsValidacionFisica
    chkIncAfiIpss.DataField = "indIAIpss"
    
    Set chkIncAfiNac.DataSource = rsValidacionFisica
    chkIncAfiNac.DataField = "indIANacio"
    
    Set chkIncAfiDir.DataSource = rsValidacionFisica
    chkIncAfiDir.DataField = "indIADirAfi"
    
    Set chkIncAfiFirRev.DataSource = rsValidacionFisica
    chkIncAfiFirRev.DataField = "indIAFirRev"
    
    Set chkIncAfiUbi.DataSource = rsValidacionFisica
    chkIncAfiUbi.DataField = "indIAUbig"
    
    Set chkIncAfiFalDat.DataSource = rsValidacionFisica
    chkIncAfiFalDat.DataField = "indIAFalDat"
    
    '--Empleadores
    
    Set chkDENroRuc.DataSource = rsValidacionFisica
    chkDENroRuc.DataField = "indIERuc"
    
    Set chkDERazSoc.DataSource = rsValidacionFisica
    chkDERazSoc.DataField = "indIERazSoc"
    
    Set chkDENroEmp.DataSource = rsValidacionFisica
    chkDENroEmp.DataField = "indIENroEmp"
    
    Set chkDEDir.DataSource = rsValidacionFisica
    chkDEDir.DataField = "indIEDir"
    
    Set chkDEUbi.DataSource = rsValidacionFisica
    chkDEUbi.DataField = "indIEUbigeo"
    
    '-- Veri Soli
    
    Set chkVSFecSus.DataSource = rsValidacionFisica
    chkVSFecSus.DataField = "indIVFecSus"
    
    Set chkVSTipFon.DataSource = rsValidacionFisica
    chkVSTipFon.DataField = "indIVTipFon"
    
    Set chkVSDecTra.DataSource = rsValidacionFisica
    chkVSDecTra.DataField = "indIVDecTra"
    
    Set chkVSVS.DataSource = rsValidacionFisica
    chkVSVS.DataField = "indIVVerSol"
    
    
    Set chkEEDatAfi.DataSource = rsValidacionFisica
    chkEEDatAfi.DataField = "indEEDatAfi"
    
    Set chkEEDatEmp.DataSource = rsValidacionFisica
    chkEEDatEmp.DataField = "indEEDatEmp"
    
    Set chkEEDatPro.DataSource = rsValidacionFisica
    chkEEDatPro.DataField = "indEEDatPro"
    
    Set chkEEAntPrev.DataSource = rsValidacionFisica
    chkEEAntPrev.DataField = "indEEAntPrev"
    
    Set chkEEDecSol.DataSource = rsValidacionFisica
    chkEEDecSol.DataField = "indEEDecSol"
    
    Set chkDDApeAfi.DataSource = rsValidacionFisica
    chkDDApeAfi.DataField = "indDDApeAfi"
    
    Set chkDDNomAfi.DataSource = rsValidacionFisica
    chkDDNomAfi.DataField = "indDDNomAfi"
    
    Set chkDDFecNac.DataSource = rsValidacionFisica
    chkDDFecNac.DataField = "indDDFecNac"
    
    Set chkDDTipDoc.DataSource = rsValidacionFisica
    chkDDTipDoc.DataField = "indDDTipDoc"
    
    Set chkDDNroDoc.DataSource = rsValidacionFisica
    chkDDNroDoc.DataField = "indDDNumDoc"


End Sub
Sub ActualizarLoteVal(Lote As String)
On Error GoTo ERRORES
Dim cmdAct As Command
Set cmdAct = New Command

With cmdAct
    .ActiveConnection = Con.Cnx
    .CommandText = " Update Ctrldoc set CodUsuVal = '" & gNomUsu & "', FecVal = getdate()  where Lote = '" & Lote & "'"
    .Execute
End With

Set cmdAct = Nothing

Exit Sub
ERRORES:
MsgBox Err.Description
End Sub

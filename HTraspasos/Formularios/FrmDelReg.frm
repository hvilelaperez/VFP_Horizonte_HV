VERSION 5.00
Begin VB.Form FrmDelReg 
   Caption         =   "Elimina Registros"
   ClientHeight    =   8325
   ClientLeft      =   1860
   ClientTop       =   2130
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10995
   Begin VB.Frame Frame2 
      Caption         =   "Elimina Lote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   52
      Top             =   6240
      Width           =   10695
      Begin VB.Frame Frame3 
         Caption         =   "Opciones de Eliminación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         TabIndex        =   56
         Top             =   240
         Width           =   3975
         Begin VB.OptionButton optcomp 
            Caption         =   "Completa"
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
            Left            =   1920
            TabIndex        =   58
            ToolTipText     =   "Imagenes, Registro y Lote"
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optdig 
            Caption         =   "Solo Digitación"
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
            ToolTipText     =   "Registro"
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.TextBox txtLoteElim 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   54
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminarLote 
         BackColor       =   &H00FFFFFF&
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
         Height          =   495
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   55
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.Frame FraDelImag 
      Caption         =   "Elimina Registro de Imagenes"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   27
      Top             =   3960
      Width           =   10695
      Begin VB.CommandButton CmdEliminaImg 
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
         Height          =   375
         Left            =   9240
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton CmdBusca2 
         Caption         =   "Buscar Reg &Imagen"
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
         Left            =   6600
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxtImagen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4920
         TabIndex        =   7
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox TxtLote2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Imagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   46
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label LblImagen 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   360
         TabIndex        =   45
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label LblTipImag 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   8160
         TabIndex        =   41
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "TipImag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8160
         TabIndex        =   40
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label LblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7320
         TabIndex        =   39
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label LblLotFisico 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6000
         TabIndex        =   38
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label LblPosicion 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5040
         TabIndex        =   37
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label LblBarcode 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   36
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label LblIDImagen 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2280
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7320
         TabIndex        =   34
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Lote_Fisico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6000
         TabIndex        =   33
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "IDImagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   32
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Posicion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   31
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4080
         TabIndex        =   30
         Top             =   1320
         Width           =   795
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   10680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Num Imagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   28
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.Frame FraDelSol 
      Caption         =   "Elimina Registro de Solicitud  y Empleadores"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10695
      Begin VB.TextBox txtFecSal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         TabIndex        =   0
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   3840
         TabIndex        =   42
         Top             =   2400
         Width           =   4455
         Begin VB.OptionButton OptSoloEmp 
            Caption         =   "Solo Reg Empleadores"
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
            Left            =   1680
            TabIndex        =   44
            Top             =   360
            Width           =   2295
         End
         Begin VB.OptionButton OptAmbos 
            Caption         =   "Ambos"
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
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmdEliminaSol 
         Caption         =   "&Deletea Regs"
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
         Left            =   8760
         TabIndex        =   5
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton CmdBusca1 
         Caption         =   "Busca &Solicitud"
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
         Left            =   8640
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxtSecu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtNumSol 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtLote1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4320
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   51
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label20 
         Caption         =   "IndDoc"
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
         Left            =   3000
         TabIndex        =   50
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblIndDoc 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3120
         TabIndex        =   49
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "Num Reg"
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
         Left            =   2040
         TabIndex        =   48
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label LblNumReg 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2040
         TabIndex        =   47
         Top             =   2760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   10680
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Secu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5760
         TabIndex        =   26
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Solicitud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   25
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4560
         TabIndex        =   24
         Top             =   480
         Width           =   420
      End
      Begin VB.Label LblFecSal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   360
         TabIndex        =   23
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   22
         Top             =   2520
         Width           =   1560
      End
      Begin VB.Label LblNom2 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   21
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label LblNom1 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6960
         TabIndex        =   20
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label LblApeMat 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4560
         TabIndex        =   19
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label LblApePat 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2040
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label LblCuspp 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8880
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nombre1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6960
         TabIndex        =   15
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cuspp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ap Materno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4560
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ap Paterno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   12
         Top             =   1680
         Width           =   1065
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
      Height          =   495
      Left            =   9240
      TabIndex        =   10
      Top             =   7680
      Width           =   1095
   End
End
Attribute VB_Name = "FrmDelReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsSol As New Recordset
Public RsImag As New Recordset
Public ximagen As String
Public xlote As String
Dim xcondicion As String
Dim xcondicion1 As String
Dim xcondicion2 As String
Dim xcondi1 As String
Dim xcondi2 As String
Dim xcondi3 As String
Dim xcondi3e As String
Dim xcondi4 As String
Dim xcondi5 As String
Dim SQLstr As String

Private Sub Command1_Click()

End Sub

Private Sub cmdEliminarLote_Click()
On Error GoTo errores

If val(txtLoteElim.Text) = 0 Then
    MsgBox "Debe ingresar un número de lote válido", vbCritical, "Eliminación de Lote"
    txtLoteElim.SetFocus
    Exit Sub
End If

If MsgBox("¿Esta seguro de eliminar el lote (" & IIf(optdig.Value, "Solo Digitación", " Completo") & ") ", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminación de Lote") = vbYes Then
    If optdig.Value Then
        EliminarLoteDigitacion txtLoteElim.Text
    Else
        EliminarLoteCompleto txtLoteElim.Text
    End If
End If


MsgBox "Proceso de Eliminación Completado ", 64, Caption
txtLoteElim.Text = ""
Exit Sub
errores:
MsgBox Err.Description, vbCritical, Caption

End Sub


Private Sub Form_Load()
    
'    CmdBusca1.Enabled = False
'    CmdBusca2.Enabled = False
End Sub

Private Sub CmdBusca1_Click()
    lblCUSPP = " "
    lblApePat = " "
    lblApeMat = " "
    LblNom1 = " "
    LblNom2 = " "
    LblFecSal = " "
    lblnumreg = " "
    
    xcondi1 = ""
    xcondi2 = ""
    xcondi3 = ""
    xcondi3e = ""
    xcondi4 = ""
    xcondi4e = ""
    
    
    
    If txtNumsol <> "" Then
        xcondi1 = " numsol = '" & txtNumsol & "' "
    Else
        MsgBox "Debe Ingresar Solicitud", 64, "Busqueda"
        txtNumsol.SetFocus
    End If
    
    If txtFecsal <> "" Then
        xcondi4 = " and fecsal = '" & txtFecsal & "' "
        xcondi4e = " and fecsal = '" & txtFecsal & "' "
    Else
        MsgBox "Debe Ingresar la fecha de salida", 64, "Busqueda"
        txtFecsal.SetFocus
    End If
    
    If TxtLote1 <> "" Then
        xcondi2 = " and Lote = '" & TxtLote1 & "' "
    End If
    
    If TxtSecu <> "" Then
        xcondi3 = " and Secu = '" & TxtSecu & "' "
        xcondi3e = " and Secuencial = '" & TxtSecu & "' "
    End If
    
    
    xcondicion1 = RTrim(LTrim(xcondi1 + xcondi2 + xcondi3 + xcondi4))
    xcondicion2 = RTrim(LTrim(xcondi1 + xcondi2 + xcondi3e + xcondi4e))
  
    
    If RsSol.State = 1 Then RsSol.Close

    SQLstr = " SELECT fecsal,lote,numsol,secu,cuspp,apepat,apemat,nombre1,nombre2,codusu,inddoc " & _
             " FROM solicitud " & _
             " Where " & xcondicion1 & " "

    RsSol.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    If RsSol.EOF Then
        MsgBox "La Solicitud ingresada No Existe"
        txtNumsol.SetFocus
        Exit Sub
    Else
        LlenaControles1
    End If
End Sub

Private Sub LlenaControles1()
    With RsSol
        lblCUSPP = IIf(IsNull(!Cuspp), "", !Cuspp)
        lblApePat = IIf(IsNull(!ApePat), "", !ApePat)
        lblApeMat = IIf(IsNull(!ApeMat), "", !ApeMat)
        LblNom1 = IIf(IsNull(!Nombre1), "", !Nombre1)
        LblNom2 = IIf(IsNull(!Nombre2), "", !Nombre2)
        LblFecSal = IIf(IsNull(!Fecsal), "", !Fecsal)
        lblnumreg = RsSol.RecordCount
        lblIndDoc = IIf(IsNull(!IndDoc), "", !IndDoc)
        If TxtSecu = "" Or TxtSecu = Null Then
            TxtSecu = IIf(IsNull(!Secu), "", !Secu)
        End If
    End With
End Sub

Private Sub CmdEliminaSol_Click()
    Dim Resp As Integer
    Resp = MsgBox("¿Esta seguro que desea eliminar el registro?", 4 + 32, "Confirma Eliminación de Registro ")
    If Resp = vbYes Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "update empleadores " & _
            "Set empleadores.lote = solicitud.lote " & _
            "FROM empleadores, solicitud WHERE empleadores.numsol=solicitud.numsol and empleadores.secuencial=solicitud.secu and empleadores.lote = null "
            .Execute
            
            .CommandText = "delete familiar where  " & _
            " fecsal = '" & txtFecsal & "' and  numsol = '" & Me.txtNumsol & "'" & IIf(Trim(TxtSecu) <> "", " and SecHv = '" & TxtSecu & "' ", "")
            .Execute
            
            .CommandText = "delete hojavida where  " & _
            " fecsal = '" & txtFecsal & "' and  numsol = '" & Me.txtNumsol & "'" & IIf(Trim(TxtSecu) <> "", " and Secu = '" & TxtSecu & "' ", "")
            .Execute
            
            .CommandText = "delete empleadores " & _
            "Where " & xcondicion2 & " "
            .Execute , , adExecuteNoRecords
            
            .CommandText = "Update solicitud set indD2 = 0, codmodd2 = null, UsuVerD2 = null, IndVerD2 = 0 " & _
            "Where " & xcondicion1 & " "
            .Execute , , adExecuteNoRecords
            
            .CommandText = "Update _scan set Estado = '2' " & _
            "Where lotefisico = '" & TxtLote1.Text & "' "
            .Execute , , adExecuteNoRecords
            
'            .CommandText = "Update _imagenes set idImagen = null, Tipimag = null, Posicion = null, Barcode = '0' , status = 1  " & _
'            "Where lotefisico = '" & TxtLote1.Text & "' and idimagen = '01" & txtNumsol & "'"
'            .Execute , , adExecuteNoRecords
            
        End With
        
        If OptAmbos.Value Then
            RsSol.Delete
            RsSol.UpdateBatch
        End If
        
        MsgBox "Registro Eliminado", vbInformation, "Proceso de Eliminación de Registros"
    End If
    lblCUSPP = " "
    lblApePat = " "
    lblApeMat = " "
    LblNom1 = " "
    LblNom2 = " "
    LblFecSal = " "
    lblnumreg = " "
    lblIndDoc = " "
End Sub

Private Sub CmdBusca2_Click()
    LblImagen = " "
    LblIDImagen = " "
    LblBarcode = " "
    LblPosicion = " "
    LblLotFisico = " "
    LblStatus = " "
    LblTipImag = " "
    
    xcondi4 = ""
    xcondi5 = ""
      
    If TxtImagen <> "" Then
        xlote = Format(TxtLote2, "000#")
        ximagen = Format(TxtImagen, "0000000#")
        ximagen = xlote + "_" + ximagen
        xcondi4 = " Imagen = '" & ximagen & "' "
    Else
        MsgBox "Debe Ingresar Imagen", vbCritical, "Error en Busqueda"
        TxtImagen.SetFocus
    End If
    If TxtLote2 <> "" Then
        xcondi5 = " and Lote = " & TxtLote2 & " "
    End If
    
    xcondicion = RTrim(LTrim(xcondi4 + xcondi5))
    
    If RsImag.State = 1 Then RsImag.Close

    SQLstr = "SELECT idimagen,imagen,barcode,posicion,lotefisico,status,tipimag " & _
                "FROM _imagenes " & _
                "Where " & xcondicion & "  "

    RsImag.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    If RsImag.EOF Then
        MsgBox "La Imagen ingresada No Existe"
        TxtImagen.SetFocus
        Exit Sub
    Else
        LlenaControles2
    End If
End Sub

Private Sub LlenaControles2()
    With RsImag
        LblImagen = IIf(IsNull(!Imagen), "", !Imagen)
        LblIDImagen = IIf(IsNull(!idImagen), "", !idImagen)
        LblBarcode = IIf(IsNull(!Barcode), "", !Barcode)
        LblPosicion = IIf(IsNull(!Posicion), "", !Posicion)
        LblLotFisico = IIf(IsNull(!lotefisico), "", !lotefisico)
        LblStatus = IIf(IsNull(!Status), "", !Status)
        LblTipImag = IIf(IsNull(!TipImag), "", !TipImag)
    End With
End Sub

Private Sub CmdEliminaImg_Click()
    Dim Resp As Integer
    Resp = MsgBox("¿Esta seguro que desea eliminar el registro?", 4 + 32, "Confirma Eliminación de Registro ")
    If Resp = vbYes Then
        RsImag.Delete
        RsImag.UpdateBatch
        MsgBox "Registro Eliminado", vbInformation, "Proceso de Eliminacion de Registros"
    End If
    LblImagen = " "
    LblIDImagen = " "
    LblBarcode = " "
    LblPosicion = " "
    LblLotFisico = " "
    LblStatus = " "
    LblTipImag = " "
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub txtFecsal_KeyPress(KeyAscii As Integer)
FVFecha txtFecsal, KeyAscii
End Sub

Private Sub txtFecsal_LostFocus()
ValidaFecha txtFecsal
End Sub

Private Sub TxtLote1_GotFocus()
    Marcado TxtLote1
End Sub

Private Sub TxtLote1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            txtNumsol.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLote1_LostFocus()
TxtLote1 = Format(TxtLote1, "0000000#")
End Sub

Private Sub txtLoteElim_LostFocus()
txtLoteElim = Format(Me.txtLoteElim.Text, "0000000#")
End Sub

Private Sub TxtNumSol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtSecu.SetFocus
        Case Else
           ' KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol_LostFocus()
    txtNumsol = Format(txtNumsol, "000000#")
End Sub

Private Sub TxtNumSol_Change()
    'CmdBusca1.Enabled = val(TxtLote1) > 0 And val(txtNumsol) > 0
End Sub

Private Sub TxtSecu_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdBusca1.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLote2_GotFocus()
    Marcado TxtLote2
End Sub

Private Sub TxtLote2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtImagen.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtImagen_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdBusca2.SetFocus
        Case Else
            KeyAscii = 0
    End Select
    'Case 95
End Sub

Private Sub TxtImagen_Change()
    CmdBusca2.Enabled = val(TxtLote2) > 0 And TxtImagen <> ""
End Sub


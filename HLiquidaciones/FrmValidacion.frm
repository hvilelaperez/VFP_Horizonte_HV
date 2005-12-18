VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmValidacion 
   Caption         =   "Control de Calidad"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   14415
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraprincipal 
      Height          =   11775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      Begin VB.Frame fracontrol 
         Height          =   8175
         Left            =   12840
         TabIndex        =   18
         Top             =   240
         Width           =   1335
         Begin VB.CommandButton Cmd1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd2 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton Cmd3 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1680
            Width           =   495
         End
         Begin VB.CommandButton Cmdvalidacion 
            Caption         =   "&Validacion"
            Height          =   735
            Left            =   120
            TabIndex        =   19
            Top             =   2400
            Width           =   1095
         End
      End
      Begin VB.Frame fradiseño 
         Height          =   8175
         Left            =   2760
         TabIndex        =   16
         Top             =   120
         Width           =   9975
         Begin ImgeditLibCtl.ImgEdit Imgintegra 
            Height          =   7815
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   9735
            _Version        =   131076
            _ExtentX        =   17171
            _ExtentY        =   13785
            _StockProps     =   96
            BorderStyle     =   1
            ImageControl    =   "ImgEdit1"
            UndoBufferSize  =   413076480
            ForceFileLinking1x=   -1  'True
            sReserved1      =   -3656
            sReserved2      =   -3656
            lReserved1      =   2010226912
            lReserved2      =   2010226912
            bReserved1      =   -1  'True
            bReserved2      =   -1  'True
         End
      End
      Begin VB.Frame farimagen 
         Caption         =   "Imagenes"
         Height          =   8175
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2535
         Begin VB.FileListBox Flimagen 
            Height          =   7500
            Left            =   120
            Pattern         =   "*.tif"
            TabIndex        =   15
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame fravaidar 
         Caption         =   "Corregir"
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   8640
         Width           =   12735
         Begin VB.TextBox txtfechafin 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6648
            TabIndex        =   7
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtfechaini 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4512
            TabIndex        =   6
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtcodigo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2376
            TabIndex        =   5
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Txtcuspp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtremu 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8784
            TabIndex        =   3
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtnropla 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10920
            TabIndex        =   2
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "FECHA INICO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4800
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "CODIGO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblcuspp 
            Caption         =   "CUSPP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "FECHA FINAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6960
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "REMUNERACION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8520
            TabIndex        =   9
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "NRO DE PLANILLA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   10965
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "FrmValidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdvalidacion_Click()
Dim Archdao As New Recordset
Dim Xcodigo As String
Dim Xcomentario As String
'''validar informacion levantada con la captura de datos dependiendo del archivo de estrucutra que ellos  quieren
    If Abrir_conexion Then
    
      Set Archdao = Sel_captura
      If Not Archdao.EOF Then
             While Not Archdao.EOF
               Xcodigo = Archdao.Fields("cobcodigo1")
               Select Case Xcodigo
               Case "01"
                  If Len(Trim$(Archdao.Fields("cobfechaini1"))) Then
                  Xcomentario = "INICIO DE RELACION LABORAL"
                  End If
                  
                  If Len(Trim$(Archdao.Fields("cobcodigo1"))) Then
                  Xcomentario = "INICIO DE RELACION LABORAL con fecha en blanco"
                  End If
                  
                  If Len(Trim$(Archdao.Fields("cobcodigo1"))) Then
                  Xcomentario = "INICIO DE RELACION LABORAL con fecha en blanco"
                  End If
                  
                  If Len(Trim$(Archdao.Fields("COBFECHAINI1"))) And Len(Trim$(Archdao.Fields("COBFECHAFIN1"))) Then
                  Xcomentario = "INICIO DE RELACION LABORAL se ingresa fecha"
                  End If
               Case "02"
               Case "03"
               Case "04"
               Case "05"
               Case "18"
               Case "20"
                 'MsgBox "se encontro el archivo"
               Case "21"
               Case "22"
               Case "99"
                            
               End Select
                    If Archdao.Fields("cobcodigo1") Then
                    
                    End If
               Archdao.MoveNext
             Wend
      End If
    End If
'''''
End Sub
Private Sub Flimagen_Click()
Imgintegra.Image = App.Path & "\imagenes\" & Flimagen
             lngZomm = 40 '100 '350
             lngScrollX = 1539
             lngScrollY = 3494

 Imgintegra.Zoom = lngZomm
 Imgintegra.Display
 Imgintegra.ScrollPositionY = lngScrollY
 Imgintegra.ScrollPositionX = lngScrollX
 Imgintegra.ScrollBars = True
'Imgintegra
End Sub

Private Sub Form_Load()
Flimagen.Path = App.Path & "\imagenes"
End Sub

Private Sub Imgintegra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Imgintegra.Zoom = 40 'lngZomm
Imgintegra.Display
End If
End Sub

Private Sub Imgintegra_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
Imgintegra.ZoomToSelection
End Sub

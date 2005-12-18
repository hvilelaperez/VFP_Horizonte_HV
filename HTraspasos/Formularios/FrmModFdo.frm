VERSION 5.00
Begin VB.Form FrmModFdo 
   Caption         =   "Busca y Modifica registros de Tipo de Fondo"
   ClientHeight    =   7605
   ClientLeft      =   975
   ClientTop       =   1920
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   12885
   Begin VB.Frame Frame3 
      Caption         =   "Mensajes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Top             =   3840
      Width           =   12615
      Begin VB.Label lblMensaje 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   840
         TabIndex        =   40
         Top             =   240
         Width           =   11415
      End
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
      Height          =   615
      Left            =   8520
      TabIndex        =   38
      Top             =   240
      Width           =   1335
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
      Height          =   615
      Left            =   10920
      TabIndex        =   37
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Cargados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   12615
      Begin VB.Label LblApePat 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   10335
      End
      Begin VB.Label LblLote 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10680
         TabIndex        =   35
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LblAnno 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   34
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LblMes 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.Label LblCUSPP 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   32
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label LblAFP 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2640
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label LblFdoDes 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1440
         TabIndex        =   30
         Top             =   840
         Width           =   615
      End
      Begin VB.Label LblFdoOri 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos y Nombres"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1950
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSPP"
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
         Left            =   4200
         TabIndex        =   27
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AFP Origen"
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
         Left            =   2520
         TabIndex        =   26
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Fdo Orig"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Fdo Dest"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   24
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   10680
         TabIndex        =   23
         Top             =   525
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         TabIndex        =   22
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
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
         Left            =   9240
         TabIndex        =   21
         Top             =   525
         Width           =   375
      End
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TxtNumSol 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3960
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Ingresados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   12615
      Begin VB.TextBox TxtAnno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   9120
         MaxLength       =   4
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtMes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   8160
         MaxLength       =   2
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   10680
         MaxLength       =   6
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtAFP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TxtFdoDes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TxtFdoOri 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   480
         MaxLength       =   1
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TxtCUSPP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4200
         MaxLength       =   12
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox TxtApePat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaxLength       =   60
         TabIndex        =   9
         Top             =   1920
         Width           =   10455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
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
         Left            =   9240
         TabIndex        =   18
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         TabIndex        =   17
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   10680
         TabIndex        =   16
         Top             =   525
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Fdo Dest"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   15
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fdo Orig"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AFP Origen"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSPP"
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
         Left            =   4200
         TabIndex        =   12
         Top             =   525
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos y Nombres"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1950
      End
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese Solicitud a modificar :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   19
      Top             =   360
      Width           =   3390
   End
End
Attribute VB_Name = "FrmModFdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsModFdo As New Recordset
Public RsCarFdo As New Recordset
Dim SQLstr As String
Dim xnumreg As Integer

Private Sub Form_Load()
    TxtNumSol = ""
    Inicializar

    CmdBuscar.Enabled = False
End Sub

Private Sub CmdBuscar_Click()
    Inicializar
   
    If RsModFdo.State = 1 Then RsModFdo.Close
    SQLstr = "SELECT * FROM RevFdodig Where NumSol = '" & TxtNumSol & "' "
    RsModFdo.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    If RsModFdo.EOF Then
        MsgBox "La Solicitud ingresada No Existe"
        TxtNumSol.SetFocus
        Exit Sub
    Else
        If RsCarFdo.State = 1 Then RsCarFdo.Close
        SQLstr2 = "SELECT * FROM RevFdoCarga Where N_solici = '" & TxtNumSol & "' "
        RsCarFdo.Open SQLstr2, Con.Cnx, adOpenStatic, adLockReadOnly
        xnumreg = RsCarFdo.RecordCount
    
        LlenaControles
    End If
End Sub

Private Sub LlenaControles()
    With RsModFdo
        TxtCUSPP = IIf(IsNull(!CUSPP), "", !CUSPP)
        TxtApePat = IIf(IsNull(!Ape_Pat), "", !Ape_Pat)
       
        TxtAFP = IIf(IsNull(!codafp), "", !codafp)
        TxtFdoOri = IIf(IsNull(!TipFdoOri), "", !TipFdoOri)
        TxtFdoDes = IIf(IsNull(!TipFdoDes), "", !TipFdoDes)
        TxtMes = IIf(IsNull(!Mes), "", !Mes)
        TxtAnno = IIf(IsNull(!Anno), "", !Anno)
        TxtLote = IIf(IsNull(!Lote), "", !Lote)
    End With

    If xnumreg > 0 Then
        With RsCarFdo
            lblCUSPP = IIf(IsNull(!CUSPP), "", !CUSPP)
            LblApePat = IIf(IsNull(!Ape_Pat), "", !Ape_Pat)
           
            LblAFP = IIf(IsNull(!AFP), "", !AFP)
            LblFdoOri = IIf(IsNull(!ORI), "", !ORI)
            LblFdoDes = IIf(IsNull(!DES), "", !DES)
            LblMes = IIf(IsNull(!Mes), "", !Mes)
            LblAnno = IIf(IsNull(!Anno), "", !Anno)
            LblLote = IIf(IsNull(!Lote), "", !Lote)
            
            If TxtFdoOri <> LblFdoOri Then
                lblMensaje.Caption = "Error en Fondo Origen "
            Else
                If TxtFdoDes <> LblFdoDes Then
                    lblMensaje.Caption = "Error en Fondo Destino "
                Else
                    If TxtAFP <> LblAFP Then
                        lblMensaje.Caption = "Error en AFP "
                    Else
                        If TxtCUSPP <> lblCUSPP Then
                            lblMensaje.Caption = "Error en CUSPP "
                        Else
                            If TxtApePat <> LblApePat Then
                                lblMensaje.Caption = "Error en Apellidos y Nombres "
                                
                            End If
                        End If
                    End If
                End If
            End If
            
        End With
    Else
        LblApePat = "NO EXISTE REGISTRO"
    End If
        
End Sub

Sub Inicializar()
    TxtCUSPP = ""
    TxtApePat = ""
   
    TxtFdoOri = ""
    TxtFdoDes = ""
    TxtAFP = ""
    TxtMes = ""
    TxtAnno = ""
    TxtLote = ""
    
    lblCUSPP = " "
    LblApePat = " "
    LblFdoOri = " "
    LblFdoDes = " "
    LblAFP = " "
    LblMes = " "
    LblAnno = " "
    LblLote = " "
    
    lblMensaje = ""
End Sub



Private Sub Txtnumsol_GotFocus()
    Marcado TxtNumSol
End Sub

Private Sub TxtNumSol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdBuscar.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol_LostFocus()
    TxtNumSol = Format(TxtNumSol, "000000#")
End Sub

Private Sub TxtNumSol_Change()
    CmdBuscar.Enabled = Val(TxtNumSol) > 0
End Sub

Private Sub TxtFdoOri_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 49 To 51
        Case 8
        Case 13
            TxtFdoDes.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtFdoDes_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 49 To 51
        Case 8
        Case 13
            TxtAFP.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtAFP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case 72 To 73
        Case 78 To 80
        Case 82
        Case 8
        Case 13
            TxtCUSPP.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCuspp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'vbDigCusPP = ValCuss(KeyAscii)
    Select Case KeyAscii
        Case 49 To 51
        Case 65 To 90
        Case 8
        Case 13
            If IsNumeric(Left(Trim(TxtCUSPP), 6)) And IsNumeric(Right(Trim(TxtCUSPP), 1)) And Not IsNumeric(Mid(Trim(TxtCUSPP), 7, 5)) Then
                TxtMes.SetFocus
            Else
                lblMensaje.Caption = "CUSPP No cumple estructura establecida ######AAAAA#"
                TxtCUSPP.SetFocus
            End If
        Case Else
            KeyAscii = 0
    End Select
    
End Sub

Private Sub TxtMes_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            If Val(TxtMes) > 12 Then
                lblMensaje.Caption = "Error en mes ingresado "
                TxtMes.SetFocus
            Else
                TxtAnno.SetFocus
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtAnno_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            If Val(TxtAnno) < 2006 Or Val(TxtAnno) > 2100 Then
                lblMensaje.Caption = "Error en año ingresado "
                TxtAnno.SetFocus
            Else
                TxtLote.SetFocus
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLote_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            If Val(TxtLote) = 0 Then
                lblMensaje.Caption = "Debe ingresar el Lote "
                TxtLote.SetFocus
            Else
                TxtApePat.SetFocus
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case KeyAscii
        Case 65 To 90
        Case 8
        Case 165    'Ñ
        Case 209    '
        Case 13
            CmdGrabar.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub




Private Sub CmdGrabar_Click()
    With RsModFdo
        !TipFdoOri = Trim(TxtFdoOri)
        !TipFdoDes = Trim(TxtFdoDes)
        !codafp = Trim(TxtAFP)
        
        !CUSPP = Trim(TxtCUSPP)
        !Ape_Pat = Trim(TxtApePat)
          
        !Mes = Trim(TxtMes)
        !Anno = Trim(TxtAnno)
        !Lote = Trim(TxtLote)
                
    End With
    RsModFdo.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


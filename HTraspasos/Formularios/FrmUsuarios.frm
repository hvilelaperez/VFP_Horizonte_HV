VERSION 5.00
Begin VB.Form FrmUsuarios 
   Caption         =   "Mantenimiento de Usuarios"
   ClientHeight    =   3900
   ClientLeft      =   4395
   ClientTop       =   3630
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   6330
   Begin VB.Frame frameMantenimiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4560
      TabIndex        =   15
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
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
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
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
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1095
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
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "E&liminar"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.Frame FraIr 
      Caption         =   "Ir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   10
      Top             =   2880
      Width           =   3135
      Begin VB.CommandButton CmdIr 
         Caption         =   "<<"
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
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdIr 
         Caption         =   "<"
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
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdIr 
         Caption         =   ">"
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
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdIr 
         Caption         =   ">>"
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
         Index           =   3
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox TxtClave 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox TxtLogin 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TxtApe 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox TxtNom 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Top             =   675
      Width           =   2655
   End
   Begin VB.TextBox TxtCodUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Clave"
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
      TabIndex        =   4
      Top             =   2400
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Login"
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
      TabIndex        =   3
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombres"
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
      TabIndex        =   1
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cod User"
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
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsU As New Recordset        'Usuarios
Dim SW As Boolean
Dim xMarcador As Variant

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "2326", "si-006", "Traspasos"
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    DatosServidor
'***
    If RsU.State = 1 Then RsU.Close
    RsU.CursorLocation = adUseClient
    Call Con.AbrirTabla("select * from Usuarios", RsU, adOpenStatic, adLockBatchOptimistic)
'para que aparezcan datos cuando abra pantalla
LlenaControles
'para que no se puedan modificar
Desabilitatextos
End Sub

Private Sub CmdIr_Click(Index As Integer)
    Select Case Index
        Case 0      ' Inicio
            RsU.MoveFirst
            LlenaControles
            CmdIr(0).Enabled = False
            CmdIr(1).Enabled = False
            CmdIr(2).Enabled = True
            CmdIr(3).Enabled = True
        Case 1      'Anterior
            RsU.MovePrevious
            LlenaControles
            If RsU.BOF Then
                CmdIr(0).Enabled = False
                CmdIr(1).Enabled = False
                CmdIr(2).Enabled = True
                CmdIr(3).Enabled = True
                RsU.MoveFirst
            Else
                CmdIr(0).Enabled = True
                CmdIr(1).Enabled = True
                CmdIr(2).Enabled = True
                CmdIr(3).Enabled = True
            End If
        Case 2      'Siguiente
            RsU.MoveNext
            LlenaControles
            If RsU.EOF Then
                CmdIr(0).Enabled = True
                CmdIr(1).Enabled = True
                CmdIr(2).Enabled = False
                CmdIr(3).Enabled = False
                RsU.MoveLast
            Else
                CmdIr(0).Enabled = True
                CmdIr(1).Enabled = True
                CmdIr(2).Enabled = True
                CmdIr(3).Enabled = True
            End If
        Case 3      'Final
            RsU.MoveLast
            LlenaControles
            CmdIr(0).Enabled = True
            CmdIr(1).Enabled = True
            CmdIr(2).Enabled = False
            CmdIr(3).Enabled = False
    End Select
End Sub

Private Sub LlenaControles()
    If RsU.EOF Then Exit Sub
    If RsU.BOF Then Exit Sub
    With RsU
        TxtCodUser = IIf(IsNull(!CodUser), "", !CodUser)
        TxtNom = IIf(IsNull(!Nombres), "", !Nombres)
        TxtApe = IIf(IsNull(!Apellidos), "", !Apellidos)
        TxtLogin = IIf(IsNull(!Login), "", !Login)
        TxtClave = IIf(IsNull(!clave), "", !clave)
    End With
End Sub

Private Sub CmdSalir_Click()
    Unload Me
  'Me.Hide
End Sub

Private Sub CmdNuevo_Click()
    If CmdNuevo.Caption = "&Nuevo" Then
        Habilitatextos
        Limpia_Controles
        With RsU
            'captura el ultimo numero de envio
            If .BOF And .EOF Then   'pregunta si tabla esta vacia
                TxtCodUser = 1
            Else
                .MoveLast           'desplaza hasta la ultima registro y suma 1
                TxtCodUser = Val(!CodUser) + 1
            End If
        End With
        RsU.AddNew
        TxtCodUser.Enabled = False
        TxtNom.SetFocus
        FraIr.Enabled = False   'deshabilita el frame desplazamiento
        DesabilitaBotones
        SW = True
        CmdNuevo.Caption = "&Grabar"        'cambia nombre de botones
        CmdEditar.Caption = "&Cancelar"
    Else
        On Error GoTo Err
        Grabar_Controles
        RsU.UpdateBatch
        If SW = True Then
            RsU.MoveLast
        End If
        Desabilitatextos
        FraIr.Enabled = True
        HabilitaBotones
        CmdNuevo.Caption = "&Nuevo"
        CmdEditar.Caption = "&Editar"
        Exit Sub
Err:
    MsgBox Err.Description
    End If
End Sub

Private Sub CmdEditar_Click()
    If CmdEditar.Caption = "&Editar" Then
        Habilitatextos
        TxtCodUser.Enabled = False
        TxtNom.SetFocus
        FraIr.Enabled = False
        DesabilitaBotones
        SW = False
        xMarcador = RsU.Bookmark         'marca registro
        CmdNuevo.Caption = "&Grabar"
        CmdEditar.Caption = "&Cancelar"
    Else
        RsU.CancelUpdate        'cancela modificacion
        If SW = True Then       'si intento ingresar nuevo
            RsU.MoveFirst
        Else                    'si intento modificar
            RsU.Bookmark = xMarcador
        End If
        LlenaControles
        Desabilitatextos
        FraIr.Enabled = True        'habilita frame de desplazamiento
        HabilitaBotones
        CmdNuevo.Caption = "&Nuevo"
        CmdEditar.Caption = "&Editar"
    End If
End Sub

Private Sub CmdBuscar_Click()
    Dim xcod As String          'variable de busqueda
    str1 = RsU.Bookmark         'marca registro
    xcod = InputBox("Ingrese código a Buscar")
    RsU.MoveFirst
    RsU.Find "Login ='" + Trim(xcod) + "'"
    'xcod = Format(xcod, "0000#")
    If RsU.EOF Then
        MsgBox "El codigo ingresado no existe"
        RsU.Bookmark = str1
    End If
    LlenaControles
End Sub

Private Sub CmdEliminar_Click()
    If MsgBox("Desea eliminar el Registro con código: " + Str(RsU(0)), vbOKCancel + vbQuestion, "CONFIRMA ELIMINACIÓN") <> vbCancel Then
        RsU.Delete
        RsU.UpdateBatch
        RsU.MoveNext
        If RsU.EOF Then
            RsU.MoveLast
        End If
    End If
    LlenaControles
End Sub

Sub Habilitatextos()
    ' Permite habilitar las cajas de Textos
    TxtNom.Enabled = True
    TxtApe.Enabled = True
    TxtLogin.Enabled = True
    TxtClave.Enabled = True
End Sub

Sub Desabilitatextos()
    ' Permite Desabilitar las cajas de Textos
    TxtNom.Enabled = False
    TxtApe.Enabled = False
    TxtLogin.Enabled = False
    TxtClave.Enabled = False
End Sub

Private Sub Limpia_Controles()
    TxtCodUser = ""
    TxtNom = ""
    TxtApe = ""
    TxtLogin = ""
    TxtClave = ""
End Sub

Sub DesabilitaBotones()
    CmdBuscar.Enabled = False
    CmdEliminar.Enabled = False
End Sub

Sub HabilitaBotones()
    CmdBuscar.Enabled = True
    CmdEliminar.Enabled = True
End Sub

Private Sub Grabar_Controles()
    With RsU
        !CodUser = Val(TxtCodUser)
        !Nombres = Trim(TxtNom)
        !Apellidos = Trim(TxtApe)
        !Login = Trim(TxtLogin)
        !clave = Trim(TxtClave)
    End With
End Sub

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


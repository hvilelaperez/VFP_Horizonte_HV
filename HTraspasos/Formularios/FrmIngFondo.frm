VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmIngFondo 
   Caption         =   "Ingreso para Verificacion de Tipos de Fondo"
   ClientHeight    =   9105
   ClientLeft      =   1890
   ClientTop       =   1245
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   11370
   Begin VB.Frame Frame2 
      Caption         =   "Mensajes"
      Height          =   735
      Left            =   240
      TabIndex        =   26
      Top             =   8280
      Width           =   9495
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
         TabIndex        =   27
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   11055
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9720
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtFdoDes 
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
         Height          =   400
         Left            =   8760
         MaxLength       =   1
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TxtFdoOri 
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
         Height          =   400
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   7
         Top             =   840
         Width           =   615
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
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox TxtCUSPP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   5
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox TxtNumSol 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   240
         MaxLength       =   7
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Fondo Destino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8640
         TabIndex        =   21
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fondo Origen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7680
         TabIndex        =   20
         Top             =   240
         Width           =   795
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
         Left            =   4920
         TabIndex        =   19
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label lblCUSPP 
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
         Left            =   2040
         TabIndex        =   18
         Top             =   525
         Width           =   675
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Solicitud"
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
         TabIndex        =   17
         Top             =   525
         Width           =   1230
      End
   End
   Begin VB.Frame Frame9 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7815
      Begin VB.TextBox TxtAnno 
         Alignment       =   2  'Center
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
         Height          =   400
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   1
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox TxtMes 
         Alignment       =   2  'Center
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
         Height          =   400
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   0
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
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
         Height          =   400
         Left            =   5760
         MaxLength       =   8
         TabIndex        =   2
         Top             =   315
         Width           =   1575
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
         Left            =   2280
         TabIndex        =   22
         Top             =   360
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
         Left            =   360
         TabIndex        =   15
         Top             =   360
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
         Left            =   4920
         TabIndex        =   14
         Top             =   360
         Width           =   420
      End
   End
   Begin MSDataGridLib.DataGrid DtGFondos 
      Height          =   5535
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9763
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "anno"
         Caption         =   "Año"
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
         DataField       =   "Mes"
         Caption         =   "Mes"
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
      BeginProperty Column03 
         DataField       =   "numsol"
         Caption         =   "N Solicitud"
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
         DataField       =   "cuspp"
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
      BeginProperty Column05 
         DataField       =   "codafp"
         Caption         =   "AFP"
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
         DataField       =   "tipfdoori"
         Caption         =   "Fdo Origen"
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
         DataField       =   "tipfdodes"
         Caption         =   "Fdo Destino"
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
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label LblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total Docs:"
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
      Left            =   9960
      TabIndex        =   24
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label LblTotDoc 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   9960
      TabIndex        =   23
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmIngFondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsFondos As New Recordset
Dim Activado As Boolean
Dim xcodafp As String
Dim xcodusu As String
Dim xcruzo As String
Dim xhorini As Date

'Variable Obligatorio
Private vbNroSol As Boolean
Private vbNroSol2 As Boolean
Private vtmpNroSol As String
Private vbDigNroSol As Boolean

Private vbCuss As Boolean
Private vbCuss2 As Boolean
Private vbnoCuss As Boolean
Private vbDigCuss As Boolean
Private vtmpCuss As String

Private vbtmpFTipoO As Integer  'Origen
Private vbtmpFTipoD As Integer  'Destino
Private vbdigTipFondo As Boolean ' Indica si ya se digito El Tipo de Fondo
Private vbdigF9FO As Boolean 'Indica q presiono f9 Fondo Origen
Private vbdigF9FD As Boolean 'Indica q presiono f9 Fondo destino

Private Sub DtGFondos_SelChange(Cancel As Integer)
    Activado = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    vbGAFP = "01"
    xcodusu = gNomUsu
    LblTotDoc = "0"
    TxtMes = ""
    TxtAnno = ""
    TxtLote = ""
    
    CargarComboSql cboAfp, "select codafp + ' - ' + descri, codafp from tafp"
    Call Con.AbrirTabla("select * from RevFdoDig", RsFondos, adOpenStatic, adLockBatchOptimistic)
    Nuevo

End Sub

Sub Nuevo()
    Inicializar
    
    With RsFondos
        'captura el ultimo numero de Lote
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            TxtMes = 1
            TxtAnno = 2006
            TxtMes = Format(TxtMes, "0#")
            TxtAnno = Format(TxtAnno, "000#")
            TxtLote = ""
        Else
            .MoveLast           'desplaza hasta la ultima registro y suma 1
            TxtMes = !Mes
            TxtAnno = !Anno
            TxtLote = !Lote
        End If
    
        LblTotDoc = "0"
    
        CmdBuscar.Enabled = True
        CmdAgregar.Enabled = True
        CmdGrabar.Enabled = True
        CmdSalir.Enabled = True
    End With
End Sub

Sub Inicializar()
    xhorini = fgDatosql("select convert(datetime,getdate(),114)")
    
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
        
    vbtmpFTipoO = 0  'Origen
    vbtmpFTipoD = 0 'Destino
    vbdigTipFondo = False

    vbdigF9FD = False
    vbdigF9FO = False
        
    cboAfp.ListIndex = -1
    TxtNumSol = ""
    TxtCUSPP = ""
    TxtFdoOri = ""
    TxtFdoDes = ""

End Sub

Private Sub TxtMes_GotFocus()
    Marcado TxtMes
    TxtMes.BackColor = &HC0FFFF
End Sub

Private Sub TxtMes_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            If Val(TxtMes) = 0 Then
                lblMensaje.Caption = "Debe ingresar el Mes "
                TxtMes.SetFocus
            Else
                If Val(TxtMes) > 12 Then
                    lblMensaje.Caption = "Error en mes ingresado "
                    TxtMes.SetFocus
                Else
                    TabSig KeyAscii
                End If
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtMes_LostFocus()
    TxtMes.BackColor = &H80000005
    lblMensaje.Caption = ""
    TxtMes = Format(TxtMes, "0#")
End Sub

Private Sub TxtAnno_GotFocus()
    Marcado TxtAnno
    TxtAnno.BackColor = &HC0FFFF
End Sub

Private Sub TxtAnno_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            If Val(TxtAnno) = 0 Then
                lblMensaje.Caption = "Debe ingresar el Año "
                TxtAnno.SetFocus
            Else
                If Val(TxtAnno) < 2006 Or Val(TxtAnno) > 2100 Then
                    lblMensaje.Caption = "Error en año ingresado "
                    TxtAnno.SetFocus
                Else
                    TabSig KeyAscii
                End If
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtAnno_LostFocus()
    TxtAnno.BackColor = &H80000005
    lblMensaje.Caption = ""
    TxtAnno = Format(TxtAnno, "000#")
End Sub

Private Sub Txtlote_GotFocus()
    Marcado TxtLote
    TxtLote.BackColor = &HC0FFFF
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
                TabSig KeyAscii
                'CmdBuscar.SetFocus
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CmdBuscar_Click()
    Dim strsql As String
    TxtLote = Format(TxtLote, "0000000#")
    LblTotDoc = "0"
    strsql = "Select * From RevFdoDig Where Lote = '" & Trim(TxtLote) & "' Order By Lote, NumSol"
    If RsFondos.State = 1 Then RsFondos.Close
    RsFondos.Open strsql, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    Set DtGFondos.DataSource = RsFondos
    
    DtGFondos.Refresh
    
    'Formato_Grid
    If DtGFondos.ApproxCount <> 0 Then
        TxtLote = RsFondos.Fields("Lote")
        RsFondos.MoveFirst
        For i = 1 To DtGFondos.ApproxCount
            DtGFondos.Col = 3
            LblTotDoc = Val(LblTotDoc) + 1
            RsFondos.MoveNext
        Next i
    Else
        MsgBox " No hay registros de ese Lote "
    End If
    TxtNumSol.SetFocus
    
End Sub

Private Sub CmdAgregar_Click()
    If xcodafp = "" Then
        MsgBox "Debe ingresar AFP ", vbCritical, "Error en Ingreso de AFP"
        cboAfp.SetFocus
        Exit Sub
    End If
    
    If vbNroSol2 Or vbCuss2 Or vbdigTipFondo Then
        With RsFondos
            .AddNew
            .Fields("Mes") = TxtMes
            .Fields("Anno") = TxtAnno
            .Fields("Lote") = TxtLote
            .Fields("NumSol") = TxtNumSol
            .Fields("CUSPP") = TxtCUSPP
            .Fields("CodAfp") = xcodafp
        
            .Fields("TipFdoOri") = Val(TxtFdoOri)
            .Fields("TipFdoDes") = Val(TxtFdoDes)
            .Fields("cruzo") = "N"
            .Fields("UsuDig") = xcodusu
            .Fields("FecDig") = gDate
            .Fields("F_Horini") = xhorini
            .Fields("F_Horfin") = fgDatosql("select convert(datetime,getdate(),114)")
            '.Update
        End With
    
        LblTotDoc = Val(LblTotDoc) + 1
        Inicializar
        CmdGrabar.Enabled = True
        TxtNumSol.SetFocus
    
    Else
        lblMensaje.Caption = "Realizar 2da Digitacion"
        Beep
        vtmpNroSol = TxtNumSol.Text
        TxtNumSol.Text = ""
        vbNroSol2 = True
        
        vtmpCuss = TxtCUSPP.Text
        TxtCUSPP.Text = ""
        vbCuss2 = True  '2da digitacion
        
        vbtmpFTipoO = TxtFdoOri.Text
        TxtFdoOri.Text = ""

        vbtmpFTipoD = TxtFdoDes.Text
        TxtFdoDes.Text = ""

        vbdigTipFondo = True
        TxtNumSol.SetFocus
                
        vbDig2 = True
        TxtNumSol.SetFocus
        Exit Sub
    End If

End Sub

Private Sub CmdGrabar_Click()
If RsFondos.State = 1 And RsFondos.RecordCount > 0 Then
    RsFondos.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End If
End Sub

Private Sub CmdEliminar_Click()
    If Activado Then
        'On Error GoTo Salir
        With RsFondos
            xcruzo = .Fields("cruzo")
            If xcruzo <> "S" Then
                DtGFondos.Col = 0
                DtGFondos.Col = 3
                LblTotDoc = Val(LblTotDoc) - 1
                .Delete
            Else
                MsgBox "No se puede eliminar registro ya cruzado", vbCritical, "Error en Eliminación "
            End If
        End With
    Else
        MsgBox "Selecciona el registro a eliminar", vbCritical, "Error en Eliminación "
    End If
Salir:
    Exit Sub
End Sub

Private Sub cboAfp_GotFocus()
    cboAfp.BackColor = &HC0FFFF
End Sub

Private Sub cboAfp_KeyPress(KeyAscii As Integer)
    TabSig KeyAscii
    'pgBuscarEnCombo cboAfp, KeyAscii, 2
    xcodafp = Left(Trim(cboAfp.Text), 2)
End Sub

Private Sub cboAfp_LostFocus()
    cboAfp.BackColor = &H80000005
    lblMensaje.Caption = ""
End Sub

Private Sub Txtnumsol_GotFocus()
'On Local Error Resume Next
    Marcado TxtNumSol
    TxtNumSol.BackColor = &HC0FFFF
End Sub

Private Sub TxtNumSol_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtNumSol.Text = Format(TxtNumSol.Text, "000000#")
            If Val(TxtNumSol) = 0 Then
                lblMensaje.Caption = "Debe ingresar el Numero de Solicitud "
                TxtNumSol.SetFocus
            Else
                vbDigNroSol = ValNroSol(KeyAscii)
                'Txtnumsol.Text = Format(Txtnumsol.Text, "000000#")
                If vbDigNroSol Then
                
                   ' TxtNumSol.Text = Format(TxtNumSol.Text, "000000#")
                    Call BuscarSolicitud
                Else
                    lblMensaje.Caption = "Valores no coinciden "
                    TxtNumSol.SetFocus
                
                End If
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol_LostFocus()
    TxtNumSol.Text = Format(TxtNumSol.Text, "000000#")
    TxtNumSol.BackColor = &H80000005
    lblMensaje.Caption = ""
End Sub

Sub BuscarSolicitud()
    'Existe el Nro de Solicitud
    If fgDatosql("select NUMSOL from RevFdoDig where NUMSOL = '" & TxtNumSol & "'") <> "" Then
        MsgBox "La solicitud ya existe. Debe ingresar otro registro", vbCritical, "Error en Ingreso de Solicitud "
        Inicializar
        TxtNumSol.SetFocus
    Else
        TxtCUSPP.SetFocus
    End If
End Sub

Private Sub TxtNumSol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        vbDigNroSol = ValNroSol(13)
        Exit Sub
    End If
    If KeyCode = vbKeyF9 Then
        vbDigNroSol = ValNroSol(KeyCode)
        TxtNumSol.Text = Format(TxtNumSol.Text, "000000#")
        Call BuscarSolicitud
        
        
    End If
End Sub

Function ValNroSol(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto ingreso Obligatorio

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Not vbNroSol Then
        If Trim(TxtNumSol) = "" Then
            lblMensaje.Caption = "Debe ingresar Nro de Solicitud "
            ValNroSol = False
            TxtNumSol.SetFocus  'TabSig 13
            Exit Function
        End If
    End If

    If vbNroSol2 Then
        If Comparar(vtmpNroSol, TxtNumSol.Text, True) Then
            'Pasa al sgte campo
            lblMensaje.Caption = ""
            ValNroSol = True
            TabSig KeyAscii
        Else
            lblMensaje.Caption = " Valores no coinciden"
            Beep
            ValNroSol = False
            TxtNumSol.SetFocus
        End If
    Else
        vbNroSol = True
        lblMensaje.Caption = ""
        ValNroSol = True
        TabSig KeyAscii
    End If
End If

If KeyAscii = vbKeyF9 Then
    If vbNroSol2 Then '2da Digitacion
        'Fuerza el pase al sgte campo
        ValNroSol = True
        lblMensaje.Caption = ""
        'TabSig 13
    Else
        ValNroSol = False
        lblMensaje = "Debe digitar 2 veces el campo"
        Beep
    End If
End If

End Function

Private Sub TxtCUSPP_GotFocus()
    Marcado TxtCUSPP
    TxtCUSPP.BackColor = &HC0FFFF
End Sub

Private Sub TxtCUSPP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        vbDigCuss = ValCuss(13)
        Exit Sub
    End If
    If KeyCode = vbKeyF9 Or KeyCode = vbKeyF8 Then
        vbDigCuss = ValCuss(KeyCode)
    End If
End Sub

Private Sub TxtCuspp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    vbDigCusPP = ValCuss(KeyAscii)
End Sub

Private Sub TxtCUSPP_LostFocus()
    TxtCUSPP.BackColor = &H80000005
    lblMensaje.Caption = ""
End Sub

Private Sub TxtCUSPP_Validate(Cancel As Boolean)
'On Error Resume Next
    If (GetKeyState(vbKeyTab) < 0) Then
        If Not ValCuss(13) Then
            Cancel = True
        End If
        Exit Sub
    End If

    If vbDigCuss Then Cancel = False: Exit Sub
  
    vbDigCuss = ValCuss(13)
    If Not vbDigCuss Then Cancel = True
End Sub

Function ValCuss(KeyAscii As Integer) As Boolean
'Funcion que valida el correcto ingreso del CUSPP

If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
    If Trim(TxtCUSPP) = "" Then
        lblMensaje.Caption = "Debe Ingresar Valor"
        Beep
        ValCuss = False
        TxtCUSPP.SetFocus
        Exit Function
    End If

    If Not vbCuss Then
        If Not ValCusp2(TxtCUSPP) Then
            ValCuss = False
            lblMensaje.Caption = "CUSPP No cumple estructura establecida ######AAAAA#"
            TxtCUSPP.SetFocus
            Exit Function
        End If
        vbCuss = True
        TxtCUSPP.SetFocus
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

    If Not ValCusp2(TxtCUSPP) Then
        ValCuss = False
        lblMensaje.Caption = "CUSPP No cumple estructura establecida ######AAAAA#"
        TxtCUSPP.SetFocus
        'Exit Function
    End If

End If

If vbCuss Then
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 165 Or KeyAscii = 209) Then
        If Not vbnoCuss Then  'si el falso

            If Not CompararTexto(vtmpCuss, TxtCUSPP.Text, KeyAscii, Len(TxtCUSPP) + 1) Then
                KeyAscii = 0
                Beep
                lblMensaje.Caption = " Caracter no Corresponde"
                TxtCUSPP.SetFocus
                ValCuss = False
            Else
                lblMensaje.Caption = ""
            End If
        Else
            vbnoCuss = False
        End If
    Else
       If KeyAscii = 13 And (Len(vtmpCuss) <= Len(TxtCUSPP.Text)) Then vbCuss = True: lblMensaje = "": cboAfp.SetFocus
    End If
 End If

End Function

Function ValCusp2(texto As String) As Boolean
    ValCusp2 = True
    If Len(Trim(texto)) < 11 Then ValCusp2 = False: Exit Function

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

Private Sub TxtFdoOri_GotFocus()
    Marcado TxtFdoOri
    TxtFdoOri.BackColor = &HC0FFFF
End Sub

Private Sub TxtFdoOri_LostFocus()
    TxtFdoOri.BackColor = &H80000005
    lblMensaje.Caption = ""
End Sub

Private Sub TxtFdoOri_KeyPress(KeyAscii As Integer)
    Dim vbtemkey As Integer

    vbtemkey = KeyAscii
    If Not (KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or KeyAscii = vbKey3 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyF9) Then
        KeyAscii = 0
    End If
    TabSig vbtemkey
End Sub

Private Sub TxtFdoOri_KeyUp(KeyCode As Integer, Shift As Integer)
    '-------Validar--------
    If vbdigTipFondo Then
        If KeyCode = vbKeyF9 Then
            vbdigF9FO = True
            TxtFdoDes.SetFocus
        End If
    End If
End Sub

Private Sub txtFTipoO_Validate(Cancel As Boolean)
    If TxtFdoOri = "1" Or TxtFdoOri = "2" Or TxtFdoOri = "3" Then
        Cancel = False
    Else
        Cancel = True
    End If

    If vbdigTipFondo Then
        If vbdigF9FO Then
            lblMensaje.Caption = ""
            TxtFdoDes.SetFocus
            Exit Sub
        End If

        If Trim(vbtmpFTipoO) <> TxtFdoOri.Text Then
            lblMensaje.Caption = "Valores no coinciden"
            'txtFTipoO.SetFocus
            Cancel = True
            Exit Sub
        Else
            lblMensaje.Caption = ""
            'TabSig KeyAscii
            Exit Sub
        End If
    End If
End Sub

Private Sub TxtFdoDes_GotFocus()
    Marcado TxtFdoDes
    TxtFdoDes.BackColor = &HC0FFFF
End Sub

Private Sub TxtFdoDes_LostFocus()
    TxtFdoDes.BackColor = &H80000005
    lblMensaje.Caption = ""
End Sub

Private Sub TxtFdoDes_KeyPress(KeyAscii As Integer)
    Dim vbtempKey As Integer
    vbtempKey = KeyAscii

    If Not (KeyAscii = vbKey1 Or KeyAscii = vbKey2 Or KeyAscii = vbKey3 Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If

    If vbtempKey = 13 Then Call TxtFdoDes_Validate(False)
End Sub

Private Sub TxtFdoDes_KeyUp(KeyCode As Integer, Shift As Integer)
If vbdigTipFondo Then
    If KeyCode = vbKeyF9 Then
        vbdigF9FD = True
        CmdAgregar.SetFocus
    End If
End If

End Sub

Private Sub TxtFdoDes_Validate(Cancel As Boolean)
    If TxtFdoDes = "1" Or TxtFdoDes = "2" Or TxtFdoDes = "3" Then
        Cancel = False
    Else
        Cancel = True
    End If

    If vbdigTipFondo Then
        If vbdigF9FD Then
            lblMensaje.Caption = ""
            CmdAgregar.SetFocus
            Exit Sub
        End If

        If Trim(vbtmpFTipoD) <> TxtFdoDes.Text Then
            lblMensaje.Caption = " Valores no coinciden"
            Cancel = True
            Exit Sub
        Else
            lblMensaje.Caption = ""
            CmdAgregar.SetFocus
            'Exit Sub
        End If
    Else
        CmdAgregar.SetFocus
    End If
End Sub

Private Sub TxtNumSol_Validate(Cancel As Boolean)

 If (GetKeyState(vbKeyTab) < 0) Then
  TxtNumSol.Text = Format(TxtNumSol.Text, "000000#")
  vbDigNroSol = ValNroSol(13)
    If Not vbDigNroSol Then
       Cancel = True
    Else
       
           BuscarSolicitud
    End If
    'Exit Sub
  End If

  If vbDigNroSol Then
    Cancel = False ' Exit Sub
  Else
   If Not vbNroSol Then
    Cancel = False
   Else
    Cancel = True
  'vbNroSol2
   End If
  End If
 
'  If TxtApePat = rsDetalle.Fields("apepat") Then
'     Cancel = False
'  Else
'     vbDigApepat = ApePat(13)
'     If Not vbDigApepat Then Cancel = True
'  End If
End Sub

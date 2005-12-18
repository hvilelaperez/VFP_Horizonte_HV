VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCargos 
   Caption         =   "Consulta Cargos"
   ClientHeight    =   7320
   ClientLeft      =   2205
   ClientTop       =   1845
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10275
   Begin VB.Frame Frame2 
      Caption         =   "Agregar Cargos"
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
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   8775
      Begin VB.TextBox TxtCCargo 
         Alignment       =   2  'Center
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
         MaxLength       =   5
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtDesCargo 
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
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   3
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod Cargo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion Cargo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
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
      Left            =   9000
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
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
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CmdEliminar 
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
      Left            =   9000
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame FraCriterio 
      Caption         =   "Criterio de Busqueda"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Codigo "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox TxtBusqueda 
      Height          =   405
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton CmdVisualizar 
      Caption         =   "Visualizar Toda la tabla"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox txtCaptura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   6480
      Width           =   5655
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
      Left            =   8880
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DtGCargo 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "CodCargo"
         Caption         =   "Cod Cargo"
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
         DataField       =   "descargo"
         Caption         =   "Descripcion"
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7994.835
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCriterio 
      AutoSize        =   -1  'True
      Caption         =   "Campo de Busqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "FrmCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsCargo As New Recordset
Public RsAgrega As New Recordset
Dim Activado As Boolean
Dim xcodusu As String
Dim valor As String

Private Sub Form_Load()
    CmdVisualizar.Enabled = False
    valor = "Descargo"
    xcodusu = gNomUsu
    CmdAgregar.Enabled = False
    CmdEliminar.Enabled = False
    TxtCCargo = ""
    TxtDesCargo = ""
    TxtBusqueda = ""
End Sub

Private Sub DtGCargo_HeadClick(ByVal ColIndex As Integer)
    RsCargo.Sort = DtGCargo.Columns(ColIndex).DataField
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub DtGCargo_SelChange(Cancel As Integer)
    Activado = True
    CmdEliminar.Enabled = True
End Sub

Private Sub TxtBusqueda_Change()
    Dim strsql As String
    
    If valor = "Descargo" Then
        SQLstr = "SELECT * FROM Cargo WHERE " + valor + " like '" + Trim(TxtBusqueda) & "%' ORDER BY descargo"
    Else
        SQLstr = "SELECT * FROM Cargo WHERE " + valor + " like '" + Trim(TxtBusqueda) & "%' ORDER BY codcargo"
    End If
    
    If RsCargo.State = 1 Then RsCargo.Close
    RsCargo.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    Set DtGCargo.DataSource = RsCargo
    DtGCargo.Refresh
    
    If DtGCargo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGCargo.Columns(1).CellText(DtGCargo.FirstRow))
        Final = Trim(DtGCargo.Columns(1).CellText(DtGCargo.ApproxCount))
        numReg = DtGCargo.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron los Distrito desde { " & Comienzo & " }  a  { " & Final & " },  con un total de : " & numReg & " Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
    Select Case Index
        Case 0
            valor = "Descargo"
        Case 1
            valor = "Codcargo"
    End Select
    TxtBusqueda.SetFocus
End Sub

Private Sub CmdVisualizar_Click()
'    With RsCargo
'    .MoveFirst
'    If .BOF And .EOF Then   'pregunta si tabla esta vacia
'        MsgBox "No existen registros ", vbCritical, "Error en Consulta "
'        Exit Sub
'    End If
'    If DtGCargo.ApproxCount <> 0 Then
'        .Find "Codgeo = '" + Trim(DtGCargo.Columns(0).Text) + "'"
'    End If
'    txtCaptura = ""
'    Set DtGCargo.DataSource = Nothing
'    End With
    Dim SQLstr As String
    TxtBusqueda = ""
    
    If RsCargo.State = 1 Then RsCargo.Close
    SQLstr = "SELECT * FROM Cargo ORDER BY descargo"
    
    RsCargo.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    Set DtGCargo.DataSource = RsCargo
    DtGCargo.Refresh
    
    If DtGCargo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGCargo.Columns(1).CellText(DtGCargo.FirstRow))
        Final = Trim(DtGCargo.Columns(1).CellText(DtGCargo.ApproxCount))
        numReg = DtGCargo.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron los Cargos desde { " & Comienzo & " }  a  { " & Final & " },  con un total de : " & numReg & " Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
    CmdVisualizar.Enabled = False
End Sub

Private Sub CmdAgregar_Click()
    TxtCCargo = Format(TxtCCargo, "0000#")
    If TxtCCargo = "" Or TxtCCargo = "00000" Then
        MsgBox "Debe ingresar Codigo de Cargo ", vbCritical, "Error en Ingreso de Cargo"
        TxtCCargo.SetFocus
        Exit Sub
    End If
    If TxtDesCargo = "" Then
        MsgBox "Debe ingresar Descripcion de Cargo ", vbCritical, "Error en Ingreso de Cargo"
        TxtDesCargo.SetFocus
        Exit Sub
    End If
    Dim SQLstrA As String
    
    If RsAgrega.State = 1 Then RsAgrega.Close
    SQLstrA = "SELECT * FROM Cargo where CodCargo = '" & TxtCCargo & "'  ORDER BY CodCargo "
    
    RsAgrega.Open SQLstrA, Con.Cnx, adOpenStatic, adLockReadOnly

    If RsAgrega.EOF Then
    Else
        MsgBox "El Cargo YA EXISTE"
        TxtCCargo = ""
        TxtDesCargo = ""
        TxtBusqueda.SetFocus
        Exit Sub
    End If

    With RsCargo
        .AddNew
        .Fields("CodCargo") = Format(TxtCCargo, "0000#")
        .Fields("DesCARGO") = TxtDesCargo
'        .Fields("CodUsu") = xcodusu
'        .Fields("Fecdig") = gDate
'        .Fields("Horini") = ObtenerHora
'        .Fields("Horfin") = ObtenerHora
        .UpdateBatch
    End With
    TxtCCargo = ""
    TxtDesCargo = ""
    
    CmdGrabar.Enabled = True
    CmdGrabar.SetFocus
End Sub

Private Sub CmdGrabar_Click()
    RsCargo.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub cmdEliminar_Click()
    If Activado Then
        On Error GoTo Salir
        Dim Resp As Integer
        Resp = MsgBox("¿Esta seguro que desea eliminar el registro?", 4 + 32, "Confirma Eliminación de Registro ")
        If Resp = vbYes Then
            RsCargo.Delete
            RsCargo.UpdateBatch
            MsgBox "Registro Eliminado", vbInformation, "Proceso de Eliminacion de Registros"
        End If
    Else
        MsgBox "Selecciona el registro a eliminar", vbCritical, "Error en Eliminación "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub TxtCCargo_GotFocus()
    CmdAgregar.Enabled = True
    Marcado TxtCCargo
    TxtCCargo.BackColor = &HC0FFFF
End Sub

Private Sub TxtCCargo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtDesCargo.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCCargo_LostFocus()
    If TxtCCargo = "" Then
        TxtCCargo = "00000"
    End If
    TxtCCargo.BackColor = &H80000005
    TxtCCargo = Format(TxtCCargo, "0000#")
End Sub

Private Sub TxtDesCargo_GotFocus()
    Marcado TxtDesCargo
    TxtDesCargo.BackColor = &HC0FFFF
End Sub

Private Sub TxtDesCargo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case KeyAscii
        Case 65 To 90
        Case 8
        Case 165    'Ñ
        Case 209    '
        Case 13
            If Len(TxtDesCargo) = 0 Then
                txtCaptura.Text = "Debe ingresar el nombre del Cargo "
                TxtDesCargo.SetFocus
            Else
                CmdAgregar.SetFocus
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtDesCargo_LostFocus()
    TxtDesCargo.BackColor = &H80000005
    txtCaptura.Text = ""
End Sub


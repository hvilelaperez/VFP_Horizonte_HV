VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmConsPromo1 
   Caption         =   "Consulta Promotor"
   ClientHeight    =   7365
   ClientLeft      =   1845
   ClientTop       =   2610
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11340
   Begin VB.TextBox TxtBusqueda 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4335
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
      Height          =   1935
      Left            =   8280
      TabIndex        =   7
      Top             =   240
      Width           =   2895
      Begin VB.OptionButton OptOpcion 
         BackColor       =   &H80000013&
         Caption         =   "Codigo Promotor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton OptOpcion 
         BackColor       =   &H80000013&
         Caption         =   "Apellidos , Nombres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   2175
      End
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
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox txtCaptura 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   960
      Width           =   5775
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
      Left            =   6600
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DtGPromo 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8070
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Codpromo"
         Caption         =   "Cod Promo"
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
         DataField       =   "Ape_Pat"
         Caption         =   "Apellido Paterno"
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
         DataField       =   "Ape_Mat"
         Caption         =   "Apellido Materno"
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
         DataField       =   "Nombres"
         Caption         =   "Nombres"
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
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCriterio 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "FrmConsPromo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsPromo As New Recordset      'TPromotor
Dim valor As String

Private Sub Form_Load()
    CmdVisualizar.Enabled = False
    valor = "Despromo"
End Sub

Private Sub DtGPromo_HeadClick(ByVal ColIndex As Integer)
    RsPromo.Sort = DtGPromo.Columns(ColIndex).DataField
End Sub

Private Sub TxtBusqueda_Change()
    Dim strsql As String
    If RsPromo.State = 1 Then RsPromo.Close
    If valor = "Despromo" Then
        SQLstr = "SELECT * FROM TPromotor WHERE " + valor + " like '%" + Trim(TxtBusqueda) & "%' ORDER BY Ape_Pat,Ape_Mat,Nombres"
    Else
        SQLstr = "SELECT * FROM TPromotor WHERE " + valor + " like '%" + Trim(TxtBusqueda) & "%' ORDER BY codpromo"
    End If
    RsPromo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly

    Set DtGPromo.DataSource = RsPromo
    DtGPromo.Refresh
    If DtGPromo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGPromo.Columns(1).CellText(DtGPromo.FirstRow))
        Final = Trim(DtGPromo.Columns(1).CellText(DtGPromo.ApproxCount))
        numReg = DtGPromo.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron los Promotores desde { " & Comienzo & " }  a  { " & Final & " }, con un total de : " & numReg & "  Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
    Select Case Index
        Case 0
            valor = "Despromo"
        Case 1
            valor = "Codpromo"
    End Select
    TxtBusqueda.SetFocus
End Sub

Private Sub CmdVisualizar_Click()
    Dim strsql As String
    TxtBusqueda = ""
    If RsPromo.State = 1 Then RsPromo.Close
    SQLstr = "SELECT * FROM TPromotor ORDER BY codpromo"
    
    RsPromo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtGPromo.DataSource = RsPromo
    DtGPromo.Refresh
    
    If DtGPromo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGPromo.Columns(1).CellText(DtGPromo.FirstRow))
        Final = Trim(DtGPromo.Columns(1).CellText(DtGPromo.ApproxCount))
        numReg = DtGPromo.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron los Promotores desde { " & Comienzo & " }  a  { " & Final & " }, con un total de : " & numReg & "  Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
    CmdVisualizar.Enabled = False
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


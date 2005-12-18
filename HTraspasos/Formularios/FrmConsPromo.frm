VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmConsPromo 
   Caption         =   "Consulta Promotor"
   ClientHeight    =   6735
   ClientLeft      =   765
   ClientTop       =   2370
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   13500
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
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   6135
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
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Codigo Interno"
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
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Codigo Promotor SBS"
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
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton OptOpcion 
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
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
      Left            =   10200
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtCaptura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   615
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   960
      Width           =   6735
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
      Left            =   11040
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DtGPromo 
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   8493
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
      ColumnCount     =   8
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
         DataField       =   "codproad"
         Caption         =   "Cd Interno"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "tipsanc"
         Caption         =   "Sancion"
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
         DataField       =   "codage"
         Caption         =   "Agencia"
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
         DataField       =   "codsup"
         Caption         =   "Dir Negoc"
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2594.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2594.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1395.213
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
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "FrmConsPromo"
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
        If valor = "Codpromo" Then
            SQLstr = "SELECT * FROM TPromotor WHERE " + valor + " like '%" + Trim(TxtBusqueda) & "%' ORDER BY codpromo"
        Else
            SQLstr = "SELECT * FROM TPromotor WHERE " + valor + " like '%" + Trim(TxtBusqueda) & "%' ORDER BY codproad"
        End If
    End If
    RsPromo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly

    Set DtGPromo.DataSource = RsPromo
    DtGPromo.Refresh
    If DtGPromo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGPromo.Columns(2).CellText(DtGPromo.FirstRow))
        Final = Trim(DtGPromo.Columns(2).CellText(DtGPromo.ApproxCount))
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
        Case 2
            valor = "Codproad"
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
        Comienzo = Trim(DtGPromo.Columns(2).CellText(DtGPromo.FirstRow))
        Final = Trim(DtGPromo.Columns(2).CellText(DtGPromo.ApproxCount))
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


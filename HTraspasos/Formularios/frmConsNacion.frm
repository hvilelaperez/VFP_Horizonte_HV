VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConsNacion 
   Caption         =   "Consulta de Nacionalidades"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
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
      Left            =   10920
      TabIndex        =   6
      Top             =   840
      Width           =   1095
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
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   840
      Width           =   6735
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
      Left            =   10080
      TabIndex        =   4
      Top             =   120
      Width           =   2655
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Descripcion Larga"
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
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Descripción Corta"
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
   End
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
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin MSDataGridLib.DataGrid DtGPromo 
      Height          =   4815
      Left            =   0
      TabIndex        =   7
      Top             =   1560
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
         DataField       =   "descri_larga"
         Caption         =   "Des. Larga"
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
         DataField       =   "Descri_corta"
         Caption         =   "Des. Corta"
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
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   5130.142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3825.071
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsNacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsPromo As New Recordset      'TPromotor
Dim valor As String

Private Sub Form_Load()
    CmdVisualizar.Enabled = False
    valor = "Descri_Larga"
End Sub

Private Sub DtGPromo_HeadClick(ByVal ColIndex As Integer)
    RsPromo.Sort = DtGPromo.Columns(ColIndex).DataField
End Sub

Private Sub TxtBusqueda_Change()
    Dim strsql As String
    If RsPromo.State = 1 Then RsPromo.Close
    If valor = "Descri_Larga" Then
        SQLstr = "SELECT * FROM nacionalidad WHERE " + valor + " like '%" + Trim(TxtBusqueda) & "%' ORDER BY Descri_Larga"
    Else
        If valor = "Descri_Corta" Then
            SQLstr = "SELECT * FROM nacionalidad WHERE " + valor + " like '%" + Trim(TxtBusqueda) & "%' ORDER BY Descri_Corta"
        Else
            SQLstr = "SELECT * FROM nacionalidad WHERE " + valor + " like '%" + Trim(TxtBusqueda) & "%' ORDER BY codigo"
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
        txtCaptura.Text = "Se Filtraron los Nacionalidades desde { " & Comienzo & " }  a  { " & Final & " }, con un total de : " & numReg & "  Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
    Select Case Index
        Case 0
            valor = "Descri_Larga"
        Case 1
            valor = "Descri_Corta"
        Case 2
            valor = "Codigo"
    End Select
    TxtBusqueda.SetFocus
End Sub

Private Sub CmdVisualizar_Click()
    Dim strsql As String
    TxtBusqueda = ""
    If RsPromo.State = 1 Then RsPromo.Close
    SQLstr = "SELECT * FROM Nacionalidad ORDER BY codigo"
    
    RsPromo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtGPromo.DataSource = RsPromo
    DtGPromo.Refresh
    
    If DtGPromo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGPromo.Columns(2).CellText(DtGPromo.FirstRow))
        Final = Trim(DtGPromo.Columns(2).CellText(DtGPromo.ApproxCount))
        numReg = DtGPromo.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron las Nacionalidades desde { " & Comienzo & " }  a  { " & Final & " }, con un total de : " & numReg & "  Registros."
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



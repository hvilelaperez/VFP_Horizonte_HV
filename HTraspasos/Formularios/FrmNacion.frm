VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmNacion 
   Caption         =   "Consulta Nacionalidades"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdVisualizar 
      Caption         =   "Visualizar Toda la tabla"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Left            =   7920
      TabIndex        =   1
      Top             =   240
      Width           =   2175
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
      Left            =   9000
      TabIndex        =   2
      Top             =   960
      Width           =   1095
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
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   840
      Width           =   6495
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
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   5415
   End
   Begin VB.Frame FraCriterio 
      Caption         =   "Criterio "
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
      TabIndex        =   4
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Descripcion PAIS"
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
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Codigo"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DtGNacion 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9975
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
         DataField       =   "Descri_Corta"
         Caption         =   "Descripcion Corta"
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
         DataField       =   "Descri_Larga"
         Caption         =   "Descripcion Larga"
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
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4500.284
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
      Left            =   2160
      TabIndex        =   7
      Top             =   60
      Width           =   1980
   End
End
Attribute VB_Name = "FrmNacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsNacion As New Recordset      'Nacionalidad
Dim valor As String

Private Sub Form_Load()
    CmdVisualizar.Enabled = False
    valor = "Descri"
End Sub

Private Sub DtGNacion_HeadClick(ByVal ColIndex As Integer)
    RsNacion.Sort = DtGNacion.Columns(ColIndex).DataField
End Sub

Private Sub TxtBusqueda_Change()
    Dim strsql As String
    If RsNacion.State = 1 Then RsNacion.Close
    If valor = "Descri" Then
        SQLstr = "SELECT * FROM Nacionalidad WHERE Descri_Corta like '" + Trim(TxtBusqueda) & "%' or Descri_Larga like '" + Trim(TxtBusqueda) & "%' ORDER BY Descri_Larga "
    Else
        SQLstr = "SELECT * FROM Nacionalidad WHERE " + valor + " like '" + Trim(TxtBusqueda) & "%' ORDER BY codigo"
    End If
    RsNacion.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly

   'strSQL = "Select * From Clientes Where " + Trim(valor) + " like '" + Trim(TxtBusqueda) & "%' Order By Codigo"
    Set DtGNacion.DataSource = RsNacion
    DtGNacion.Refresh
    If DtGNacion.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGNacion.Columns(1).CellText(DtGNacion.FirstRow))
        Final = Trim(DtGNacion.Columns(1).CellText(DtGNacion.ApproxCount))
        numReg = DtGNacion.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron los Paises desde { " & Comienzo & " }  a  { " & Final & " },  con un total de : " & numReg & " Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
    Select Case Index
        Case 0
            valor = "Descri"
        Case 1
            valor = "Codigo"
    End Select
    TxtBusqueda.SetFocus
End Sub

Private Sub CmdVisualizar_Click()
    Dim strsql As String
    TxtBusqueda = ""
    If RsNacion.State = 1 Then RsNacion.Close
    
    SQLstr = "SELECT * FROM Nacionalidad ORDER BY codigo"
    RsNacion.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtGNacion.DataSource = RsNacion
    DtGNacion.Refresh
    
    If DtGNacion.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGNacion.Columns(1).CellText(DtGNacion.FirstRow))
        Final = Trim(DtGNacion.Columns(1).CellText(DtGNacion.ApproxCount))
        numReg = DtGNacion.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron los Distrito desde { " & Comienzo & " }  a  { " & Final & " },  con un total de : " & numReg & " Registros."
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


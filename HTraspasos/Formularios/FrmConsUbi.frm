VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmConsUbi 
   Caption         =   "Consulta Ubigeo"
   ClientHeight    =   6300
   ClientLeft      =   1230
   ClientTop       =   2595
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   11895
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
      Height          =   1575
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Codigo Ubigeo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Distrito, Provincia, Depart"
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
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Ubigeo RENIEC"
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
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
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
      Left            =   6120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DtGUbigeo 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7646
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Codgeo"
         Caption         =   "Ubigeo"
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
         DataField       =   "Distrito"
         Caption         =   "Distrito"
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
         DataField       =   "Provincia"
         Caption         =   "Provincia"
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
         DataField       =   "Departamento"
         Caption         =   "Departamento"
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
         DataField       =   "codubic"
         Caption         =   "Ubg Reniec"
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
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
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
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   5175
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
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox TxtBusqueda 
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4455
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "FrmConsUbi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsUbi As New Recordset      'Ubigeo
Dim valor As String

Private Sub Form_Load()
    CmdVisualizar.Enabled = False
    valor = "Desgeo"
End Sub

Private Sub DtGUbigeo_HeadClick(ByVal ColIndex As Integer)
    RsUbi.Sort = DtGUbigeo.Columns(ColIndex).DataField
End Sub

Private Sub TxtBusqueda_Change()
    Dim strsql As String
    If RsUbi.State = 1 Then RsUbi.Close
    If valor = "Desgeo" Then
        SQLstr = "SELECT * FROM Ubigeo WHERE " + valor + " like '" + Trim(TxtBusqueda) & "%' ORDER BY distrito,provincia,departamento"
    Else
        SQLstr = "SELECT * FROM Ubigeo WHERE " + valor + " like '" + Trim(TxtBusqueda) & "%' ORDER BY codgeo"
    End If
    RsUbi.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly

   'strSQL = "Select * From Clientes Where " + Trim(valor) + " like '" + Trim(TxtBusqueda) & "%' Order By Codigo"
    Set DtGUbigeo.DataSource = RsUbi
    DtGUbigeo.Refresh
    If DtGUbigeo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGUbigeo.Columns(1).CellText(DtGUbigeo.FirstRow))
        Final = Trim(DtGUbigeo.Columns(1).CellText(DtGUbigeo.ApproxCount))
        numReg = DtGUbigeo.ApproxCount
        CmdVisualizar.Enabled = True
        txtCaptura.Text = "Se Filtraron los Distrito desde { " & Comienzo & " }  a  { " & Final & " },  con un total de : " & numReg & " Registros."
    Else
        txtCaptura.Text = " No hay registros con esa letra o codigo "
    End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
    Select Case Index
        Case 0
            valor = "Desgeo"
        Case 1
            valor = "Codgeo"
        Case 2
            valor = "CodUBIC"
    End Select
    TxtBusqueda.SetFocus
End Sub

Private Sub CmdVisualizar_Click()
'    With RsUbi
'    .MoveFirst
'    If .BOF And .EOF Then   'pregunta si tabla esta vacia
'        MsgBox "No existen registros ", vbCritical, "Error en Consulta "
'        Exit Sub
'    End If
'    If DtGUbigeo.ApproxCount <> 0 Then
'        .Find "Codgeo = '" + Trim(DtGUbigeo.Columns(0).Text) + "'"
'    End If
'    txtCaptura = ""
'    Set DtGUbigeo.DataSource = Nothing
'    End With
    Dim strsql As String
    TxtBusqueda = ""
    If RsUbi.State = 1 Then RsUbi.Close
    SQLstr = "SELECT * FROM Ubigeo ORDER BY codgeo"
    RsUbi.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtGUbigeo.DataSource = RsUbi
    DtGUbigeo.Refresh
    If DtGUbigeo.ApproxCount <> 0 Then
        Dim Comienzo, Final, numReg
        Comienzo = Trim(DtGUbigeo.Columns(1).CellText(DtGUbigeo.FirstRow))
        Final = Trim(DtGUbigeo.Columns(1).CellText(DtGUbigeo.ApproxCount))
        numReg = DtGUbigeo.ApproxCount
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

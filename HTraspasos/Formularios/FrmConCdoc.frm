VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmConCdoc 
   Caption         =   "Consulta por Envios"
   ClientHeight    =   7590
   ClientLeft      =   2850
   ClientTop       =   1935
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9600
   Begin VB.Frame FraIr 
      Caption         =   "Ir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   10
      Top             =   0
      Width           =   2295
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
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   495
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
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   495
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
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
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
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox TxtTotImag 
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
      Height          =   405
      Left            =   4320
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtTotdoc 
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
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
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
      Left            =   8400
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtEnvio 
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
      Height          =   405
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtFecha 
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
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DtGAgencia 
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
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
      ColumnCount     =   6
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "tdoc"
         Caption         =   "Tipo"
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
         DataField       =   "Agencia"
         Caption         =   "Cod Agen"
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
         DataField       =   "DesAgen"
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
      BeginProperty Column04 
         DataField       =   "CantSol"
         Caption         =   "Cantidad"
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
         DataField       =   "CantImag"
         Caption         =   "Imagenes"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2954.835
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin VB.Label LblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total Docs"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Imag"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label LblEnvio 
      AutoSize        =   -1  'True
      Caption         =   "Envio"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   525
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      TabIndex        =   4
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "FrmConCdoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsCtrlD As New Recordset      'CtrlDoc
Public RsConsulCdoc As New Recordset        'TempConsulta
Dim xprimer As Integer

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "DIGITACION\DIGITACION", "BDTraspasos"
'    Con.ConexionSQL "sa", "", "DG-63\LASERFICHE", "BDTraspasosDig2"
'    DatosServidor
'***
    If RsCtrlD.State = 1 Then RsCtrlD.Close
    SQLstr = "SELECT distinct envio, fecsal FROM ctrldoc "
    RsCtrlD.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
'***
'    Call Con.AbrirTabla("select distinct envio from CtrlDoc", RsCtrlD, adOpenStatic, adLockReadOnly)
    xprimer = 0
    txtEnvio.Enabled = False
    TxtFecha.Enabled = False
    TxtTotdoc.Enabled = False
    TxtTotImag.Enabled = False
    LlenaControles
End Sub

Private Sub DtGAgencia_HeadClick(ByVal ColIndex As Integer)
    RsConsulCdoc.Sort = DtGAgencia.Columns(ColIndex).DataField
End Sub

Private Sub LlenaControles()
    With RsCtrlD
    If .BOF And .EOF Then   'pregunta si tabla esta vacia
        MsgBox "No existen registros ", vbCritical, "Error en Consulta por Envio"
        Exit Sub
    Else
        If xprimer = 0 Then
            .MoveFirst
            xprimer = 1
        Else
        End If
        txtEnvio = !Envio
    End If
    End With
    If RsConsulCdoc.State = 1 Then RsConsulCdoc.Close
    
    'strSQL = "Select * From CtrlDoc Where Envio = '" + Trim(TxtEnvio) & "' Order By Lote"
    strsql = "SELECT Lote, Envio, Fecsal,indtipdoc,tdoc=case when indtipdoc = 1 then 'C' else 'T' end, Agencia, CantSol, CantImag, DesAgen FROM CtrlDoc, TAgenAfp WHERE Agencia = Codigo and  Envio = '" + Trim(txtEnvio) & "' Order By Lote"
    RsConsulCdoc.Open strsql, Con.Cnx, adOpenStatic, adLockReadOnly
'***
    If RsConsulCdoc.EOF Then Exit Sub
    If RsConsulCdoc.BOF Then Exit Sub
    'tambien puede ser: txtcod = rsc!codigo
    'TxtEnvio.Text = RsConsulCdoc.Fields("Envio")
    TxtFecha.Text = RsConsulCdoc.Fields("FecSal")
    Set DtGAgencia.DataSource = RsConsulCdoc
    DtGAgencia.Refresh
    TxtTotdoc = 0
    TxtTotImag = 0
    If DtGAgencia.ApproxCount <> 0 Then
        RsConsulCdoc.MoveFirst
        For i = 1 To DtGAgencia.ApproxCount
            DtGAgencia.Col = 4
            TxtTotdoc = val(TxtTotdoc) + val(DtGAgencia.Text)
            DtGAgencia.Col = 5
            TxtTotImag = val(TxtTotImag) + val(DtGAgencia.Text)
            RsConsulCdoc.MoveNext
        Next i
    End If
End Sub

Private Sub CmdIr_Click(Index As Integer)
    With RsCtrlD
    Select Case Index
        Case 0      ' Inicio
            .MoveFirst
            CmdIr(0).Enabled = False
            CmdIr(1).Enabled = False
            CmdIr(2).Enabled = True
            CmdIr(3).Enabled = True
            LlenaControles
        Case 1      'Anterior
            .MovePrevious
            If .BOF Then
                CmdIr(0).Enabled = False
                CmdIr(1).Enabled = False
                CmdIr(2).Enabled = True
                CmdIr(3).Enabled = True
                .MoveFirst
            Else
                CmdIr(0).Enabled = True
                CmdIr(1).Enabled = True
                CmdIr(2).Enabled = True
                CmdIr(3).Enabled = True
            End If
            LlenaControles
        Case 2      'Siguiente
            .MoveNext
            If .EOF Then
                CmdIr(0).Enabled = True
                CmdIr(1).Enabled = True
                CmdIr(2).Enabled = False
                CmdIr(3).Enabled = False
                .MoveLast
            Else
                CmdIr(0).Enabled = True
                CmdIr(1).Enabled = True
                CmdIr(2).Enabled = True
                CmdIr(3).Enabled = True
            End If
            LlenaControles
        Case 3      'Final
            .MoveLast
            CmdIr(0).Enabled = True
            CmdIr(1).Enabled = True
            CmdIr(2).Enabled = False
            CmdIr(3).Enabled = False
            LlenaControles
    End Select
    End With
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
End Sub

Private Sub TxtFecha_LostFocus()
    ValidaFecha TxtFecha
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub


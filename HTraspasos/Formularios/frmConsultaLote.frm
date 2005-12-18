VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConsultaLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Estado de Lotes"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   13335
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3720
      TabIndex        =   10
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid dtgLote 
      Height          =   6015
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Lotefisico"
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
         DataField       =   "Envio"
         Caption         =   "Envio"
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
         DataField       =   "Fecsal"
         Caption         =   "Fecha salida"
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
         DataField       =   "Tipdoc"
         Caption         =   "Tip. Doc."
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
         DataField       =   "desEstado"
         Caption         =   "Estado"
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
         DataField       =   "Codmod"
         Caption         =   "Usuario Dig."
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
         DataField       =   "codmodd2"
         Caption         =   "Usuario Verif."
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
         DataField       =   "codusuval"
         Caption         =   "Usuario Val."
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
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3270.047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Criterio"
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
      TabIndex        =   8
      Top             =   0
      Width           =   3495
      Begin VB.OptionButton optEnvio 
         Caption         =   "Por Envio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   285
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtEnvio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptLote 
         Caption         =   "Por Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   1050
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Por Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   645
         Width           =   1215
      End
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtFecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   840
      Width           =   9375
      Begin VB.Label lblnumreg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3120
         TabIndex        =   13
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Reg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmConsultaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
Buscar
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub
Sub Buscar()
Dim sqltexto As String
Dim rsLote As Recordset
Set rsLote = New ADODB.Recordset

sqltexto = " select Lotefisico , b.Fecsal, Envio, TipDoc = case when indtipdoc = 0 then 'T' else 'C' end, " & _
           " a.Estado, DesEstado, codmod, codmodd2, codusuval from _scan a " & _
           " left join ctrldoc b on a.lotefisico = b.lote " & _
           " join Estado c on a.Estado = c.Estado "
           
Condicion = " where "
If optEnvio Then
    Condicion = Condicion & " Envio = '" & txtEnvio & "' "
End If

If OptFecha Then
    Condicion = Condicion & " b.Fecsal = '" & txtFecsal & "'  "
End If
 
If OptLote Then
    Condicion = Condicion & " a.lotefisico = '" & TxtLote & "' "
End If

sqltexto = sqltexto & Condicion & " order by Lotefisico "

rsLote.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

Set dtgLote.DataSource = rsLote

lblnumreg.Caption = rsLote.RecordCount

If rsLote.RecordCount = 0 Then MsgBox "No se encontraron registros", 64, Caption

End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub txtEnvio_LostFocus()
txtEnvio.Text = Format(txtEnvio.Text, "000000#")
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
 FVFecha TxtFecha, KeyAscii
End Sub
Private Sub TxtLote_KeyPress(KeyAscii As Integer)
'  TxtLote = Format(TxtLote, "0000000#")
End Sub

Private Sub TxtLote_LostFocus()
    TxtLote = Format(TxtLote, "0000000#")
End Sub

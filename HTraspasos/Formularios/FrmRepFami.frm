VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "Imgedit.ocx"
Begin VB.Form FrmRepFami 
   Caption         =   "Revision de Datos de Familiares"
   ClientHeight    =   10575
   ClientLeft      =   315
   ClientTop       =   330
   ClientWidth     =   14400
   LinkTopic       =   "Form2"
   ScaleHeight     =   10575
   ScaleWidth      =   14400
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   4695
      Left            =   120
      TabIndex        =   24
      Top             =   5760
      Width           =   14055
      _Version        =   131076
      _ExtentX        =   24791
      _ExtentY        =   8281
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   441921024
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   2117707439
      lReserved2      =   2117707439
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Frame Frame4 
      Caption         =   "Modo de Consulta"
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
      TabIndex        =   21
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton optEnvio 
         Caption         =   "Por Envio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   960
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
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   25
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
         Left            =   1680
         TabIndex        =   0
         Top             =   240
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
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Por Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptLote 
         Caption         =   "Por Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
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
      Height          =   1455
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton OptSoloTra 
         Caption         =   "Solo Traspasos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton OptSoloAfi 
         Caption         =   "Solo Afiliaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos los Registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
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
      Height          =   375
      Left            =   13080
      TabIndex        =   4
      Top             =   1080
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
      Height          =   375
      Left            =   13080
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CmdUbigeo 
      Caption         =   "Cons Ubigeo"
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
      Left            =   11400
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame FraRep 
      Caption         =   "Criterios"
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
      Height          =   1455
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton CmdCons 
         Caption         =   "&Consulta "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4080
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptTodosF 
         Caption         =   "Todos"
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptAPateB 
         Caption         =   "Ap Paterno = blanco"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OptParenB 
         Caption         =   "Parentesco = b"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptFNacB 
         Caption         =   "Fec Nacim = b"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton OptNom1B 
         Caption         =   "Nombre 1 = b"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton OptAMateB 
         Caption         =   "Ap Materno = blanco"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton OptSexB 
         Caption         =   "Sexo = b"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DtGFami 
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   7011
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
      ColumnCount     =   11
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
         DataField       =   "Numsol"
         Caption         =   "Solicitud"
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
         DataField       =   "inddoc"
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
      BeginProperty Column03 
         DataField       =   "Secuencia"
         Caption         =   "Sec"
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
         DataField       =   "apepat"
         Caption         =   "Ap Paterno"
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
         DataField       =   "apemat"
         Caption         =   "Ap Materno"
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
         DataField       =   "NomUno"
         Caption         =   "Nombre 1"
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
         DataField       =   "NomDos"
         Caption         =   "Nombre 2"
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
      BeginProperty Column08 
         DataField       =   "codpare"
         Caption         =   "Parentesco"
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
      BeginProperty Column09 
         DataField       =   "sexo"
         Caption         =   "Sexo"
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
      BeginProperty Column10 
         DataField       =   "fecnac"
         Caption         =   "Fec Nacim"
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
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Registros :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   11400
      TabIndex        =   16
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label LblTotReg 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   11520
      TabIndex        =   15
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "FrmRepFami"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsFami As New Recordset
Dim RsImag As New Recordset
Public xnumsol As String
Public xidimagen As String
Public xlote As String
Dim xlote2 As String
Public xfrmrep As String
Dim xtipimag As String
Dim xrutavol As String
Dim ximagen As String
Dim xnumpag As Integer
'Public xsecuencial As Integer
Dim xprimer As String
Dim Activado As Boolean
Dim B1 As String

Private Sub DtGFami_HeadClick(ByVal ColIndex As Integer)
    RsFami.Sort = DtGFami.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
    cmdGrabar.Enabled = False
    xfrmrep = "V"
    TxtFecha = gDate
    B1 = ""
End Sub

Private Sub CmdCons_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""

    If OptTodos.Value Then
        B1_1 = "SELECT a.fecsal,b.lote,a.numsol,a.inddoc,a.secuencia,a.apepat,a.apemat,a.nomuno,a.nomdos,a.codpare,a.sexo,a.fecnac " & _
        "FROM FAMILIAR a join SOLICITUD b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.secu  join  Ctrldoc c on b.lote = c.lote and b.fecsal = c.fecsal WHERE a.indDoc is not NULL "
    End If
    If OptSoloAfi.Value Then
        B1_1 = "SELECT a.fecsal,b.lote,a.numsol,a.inddoc,a.secuencia,a.apepat,a.apemat,a.nomuno,a.nomdos,a.codpare,a.sexo,a.fecnac " & _
        "FROM FAMILIAR a join SOLICITUD b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.secu join Ctrldoc c on b.lote = c.lote and b.fecsal = c.fecsal WHERE a.indDoc = 'C' "
    End If
    If OptSoloTra.Value Then
        B1_1 = "SELECT a.fecsal,b.lote,a.numsol,a.inddoc,a.secuencia,a.apepat,a.apemat,a.nomuno,a.nomdos,a.codpare,a.sexo,a.fecnac " & _
        "FROM FAMILIAR a join SOLICITUD b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.secu join Ctrldoc c on b.lote = c.lote and b.fecsal = c.fecsal WHERE a.indDoc = 'T' "
    End If
'*****
    If OptFecha.Value Then
        B1_2 = " and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
    Else
        If OptLote Then
            If Trim(TxtLote) <> "" Then
                B1_2 = " and b.Lote = '" & TxtLote & "' "
            Else
                MsgBox "Debe ingresar el Lote a buscar", vbCritical, "Búsqueda"
                TxtLote.SetFocus
                Exit Sub
            End If
        Else
            If txtEnvio <> "" Then
                B1_2 = " and c.Envio = '" & txtEnvio & "' "
            Else
                MsgBox "Debe ingresar el Envio a buscar", vbCritical, "Búsqueda"
                txtEnvio.SetFocus
                Exit Sub
            End If
        End If
    End If
'*****
    If OptTodosF.Value Then
        B1_3 = ""
    End If
    If OptAPateB.Value Then
        B1_3 = " and a.apepat = '' "
    End If
    If OptAMateB.Value Then
        B1_3 = " and a.apemat = '' "
    End If
    If OptNom1B.Value Then
        B1_3 = " and a.nomuno = '' "
    End If
    If OptParenB.Value Then
        B1_3 = " and a.codpare = '' "
    End If
    If OptSexB.Value Then
        B1_3 = " and a.sexo = '' "
    End If
    If OptFNacB.Value Then
        B1_3 = " and (a.fecnac is NULL or a.fecnac >= CONVERT(datetime,'" & TxtFecha & "',103) ) "
    End If
'*****
    B1_4 = " order by a.numsol,a.secuencia "

    B1 = B1_1 + B1_2 + B1_3 + B1_4
    
    Call CargarFami

End Sub

Sub CargarFami()
    If RsFami.State = 1 Then RsFami.Close
    
    RsFami.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    If RsFami.EOF And RsFami.BOF Then
        Set DtGFami.DataSource = Nothing
        DtGFami.Refresh
        cmdGrabar.Enabled = False
'        CmdElimina.Enabled = False
        MsgBox "No se encontraron registros", vbInformation, "Revision de Datos de Familiares "
    Else
        Set DtGFami.DataSource = RsFami
        DtGFami.Refresh
        cmdGrabar.Enabled = True
'        CmdElimina.Enabled = True
        LblTotReg.Caption = RsFami.RecordCount
        CargaImag
    End If
       
End Sub

Private Sub txtEnvio_LostFocus()
    txtEnvio.Text = Format(txtEnvio.Text, "000000#")
End Sub

Private Sub TxtFecha_GotFocus()
    Marcado TxtFecha
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
    If KeyAscii = 13 Then CmdCons.SetFocus
End Sub

Private Sub TxtFecha_LostFocus()
    ValidaFecha TxtFecha
End Sub

Private Sub TxtLote_GotFocus()
    Marcado TxtLote
End Sub

Private Sub TxtLote_LostFocus()
    TxtLote = Format(TxtLote, "0000000#")
End Sub

Private Sub TxtLote_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdCons.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CmdGrabar_Click()
    RsFami.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub CmdSalir_Click()
    xfrmrep = ""
    Unload Me
End Sub

Private Sub CmdUbigeo_Click()
    FrmConsUbi.Show
End Sub

Private Sub DtGFami_SelChange(Cancel As Integer)
    Activado = True
    CargaImag
End Sub

Private Sub CargaImag()
    With RsFami
        xlote = Format(!Lote, "0000000#")
        xlote2 = Format(!Lote, "0000000#")
        xnumsol = Format(!Numsol, "000000#")
        xtipimag = "V"
        xidimagen = "01" + xnumsol

    End With
    xnumpag = 1
      
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient
    
    SQLImag = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and TipImag = '" & xtipimag & "' and substring(idimagen,3,9)= '" & xnumsol & "' order by posicion "
    
    RsImag.Open SQLImag, Con.Cnx, adOpenStatic, adLockReadOnly
    
    With RsImag          'coneccion a _Imagenes
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            .MoveFirst
            MuestraImg
        End If
    End With
               
End Sub

Private Sub MuestraImg()
    
    With RsImag  'coneccion a _Imagenes
        .Find "[posicion]='" & xnumpag & "'"
'        .Find "[idimagen]='" & xidimagen & "'"
        
        If .EOF Then .MovePrevious
        ximagen = IIf(IsNull(!Imagen), "", !Imagen)
        xrutavol = .Fields("ruta")
        xrutavol = Replace(xrutavol, "\\", "\")
        xidvolumen = IIf(IsNull(!idvolumen), "", !idvolumen)
    End With
      
    ImgEdit1.Image = xrutavol & "\" & xlote2 & "\" & ximagen & ".TIF"
    ImgEdit1.Enabled = True

    ImgEdit1.FitTo 1
'    ImgEdit1.Zoom = 50
    ImgEdit1.Display
    
    ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 3)
    ImgEdit1.ScrollPositionX = 0
   
'    ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    ImgEdit1.Refresh
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
' alto
    ImgEdit1.Height = FrmRepFami.Height - 7800
' ancho
    ImgEdit1.Width = FrmRepFami.Width - 300
' alto
'    DtGFami.Height = 1200 - ImgEdit1.Height
    ImgEdit1.Height = FrmRepFami.Height - 6500
End Sub

Sub MovImg(KeyCode As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF12 Then
        If Not (RsImag.EOF And RsImag.BOF) Then
            RsImag.MoveNext
            If RsImag.EOF Then RsImag.MoveLast
            'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
            ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If

    If KeyCode = vbKeyF11 Then
        If Not (RsImag.EOF And RsImag.BOF) Then
            RsImag.MovePrevious
            If RsImag.BOF Then RsImag.MoveFirst
            'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
            ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!Imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If
End Sub

Private Sub ImgEdit1_DblClick()
    On Error Resume Next
    ImgEdit1.FitTo 1
End Sub

Private Sub ImgEdit1_KeyDown(KeyCode As Integer, Shift As Integer)
    MovImg KeyCode
End Sub

Private Sub ImgEdit1_MarkEnd(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MarkType As Integer, ByVal GroupName As String)
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

Private Sub ImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

'********************
Private Sub ImgEdit1_Click()
    On Error Resume Next
End Sub

Private Sub UpDown1_DownClick()
    On Error Resume Next
    If Not (RsImag.EOF And RsImag.BOF) Then
        RsImag.MoveNext
        If RsImag.EOF Then RsImag.MoveLast
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
        ImgEdit1.Image = Rutaimag(xrutavol, RsImag!lotefisico, RsImag!Imagen)
        ImgEdit1.Display
        ImgEdit1.Refresh
    End If
End Sub

Private Sub UpDown1_UpClick()
    On Error Resume Next
    If Not (RsImag.EOF And RsImag.BOF) Then
        RsImag.MovePrevious
        If RsImag.BOF Then RsImag.MoveFirst
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
        ImgEdit1.Image = Rutaimag(xrutavol, RsImag!lotefisico, RsImag!Imagen)
        ImgEdit1.Display
        ImgEdit1.Refresh
    End If
End Sub


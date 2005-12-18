VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmConsFondo 
   Caption         =   "Consulta Tipo de Fondo"
   ClientHeight    =   10560
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   14610
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
      Height          =   1335
      Left            =   3120
      TabIndex        =   16
      Top             =   120
      Width           =   2295
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
         Left            =   120
         TabIndex        =   19
         Top             =   280
         Value           =   -1  'True
         Width           =   2055
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
         Left            =   120
         TabIndex        =   18
         Top             =   620
         Width           =   1815
      End
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
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
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
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2895
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
         Left            =   120
         TabIndex        =   25
         Top             =   1000
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
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   24
         Top             =   960
         Width           =   1215
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
         Left            =   120
         TabIndex        =   15
         Top             =   650
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
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1335
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
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   13
         Top             =   600
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
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
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
      Height          =   1335
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   6135
      Begin VB.OptionButton OptFNacFdo 
         Caption         =   "F Nac vs T Fondo"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   620
         Width           =   1815
      End
      Begin VB.OptionButton OptFecFirma 
         Caption         =   "Fec Veri vs Fec Firma (T)"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   300
         Width           =   2175
      End
      Begin VB.CommandButton CmdCons 
         Caption         =   "&Consulta"
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
         Left            =   4800
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton OptTFbAFPb 
         Caption         =   "T Fondo = b / AFP = b"
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
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton OptSolo2 
         Caption         =   "Solo Tipo Mixto (2)"
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
         Left            =   120
         TabIndex        =   8
         Top             =   620
         Width           =   2055
      End
      Begin VB.OptionButton OptTodosF 
         Caption         =   "Todos"
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
         Left            =   3000
         TabIndex        =   7
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptSolo13 
         Caption         =   "Solo Tipos 1 y 3 <>  Mixto (2)"
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
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   2775
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
      Left            =   13320
      TabIndex        =   2
      Top             =   960
      Width           =   975
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
      Left            =   11880
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CmdVerImag 
      Caption         =   "Ver &Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DtGTFon 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "lote"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "secu"
         Caption         =   "sec"
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
         DataField       =   "CodAfpOri"
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
      BeginProperty Column05 
         DataField       =   "TipoFonOri"
         Caption         =   "Fon Ori"
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
         DataField       =   "TipoFonDes"
         Caption         =   "Fon Des"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "fecregpro"
         Caption         =   "F Reg Promo"
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
         DataField       =   "fecfirmsol"
         Caption         =   "F Firma Solic"
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
         DataField       =   "indenvmail"
         Caption         =   "EECC"
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
         DataField       =   "indmail"
         Caption         =   "Cas/Trab"
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
      BeginProperty Column11 
         DataField       =   "fecnaci"
         Caption         =   "Fec Nac HV"
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
      BeginProperty Column12 
         DataField       =   "fecnac"
         Caption         =   "Fec Nac T"
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
      BeginProperty Column13 
         DataField       =   "codusu"
         Caption         =   "Cod Usu"
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
      BeginProperty Column14 
         DataField       =   "UsuRepTF"
         Caption         =   "Usu. Mod."
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
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column14 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   4815
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   14295
      _Version        =   131074
      _ExtentX        =   25215
      _ExtentY        =   8493
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   375235072
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   45
      lReserved2      =   45
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
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
      Left            =   11880
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11760
      TabIndex        =   3
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "FrmConsFondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsTipFon As New Recordset
Dim RsImag As New Recordset
Public xfrmrep As String
Public xnumsol As String
Public xidimagen As String
Public xlote As String
Dim xlote2 As String
Dim xtipimag As String
Dim xrutavol As String
Dim ximagen As String
Dim xnumpag As Integer
Dim B1 As String
Dim SQLstr As String
Dim Activado As Boolean

Private Sub DtGTFon_HeadClick(ByVal ColIndex As Integer)
    RsTipFon.Sort = DtGTFon.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
    xfrmrep = "F"
    TxtFecha = gDate
    TxtloteI = ""
    TxtLoteF = ""
    B1 = ""
    cmdGrabar.Enabled = False
    CmdVerImag.Enabled = False

End Sub

Private Sub CmdCons_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""
    B1_5 = ""
    B1_6 = ""

    B1_1 = " SELECT a.fecsal,a.lote,a.secu,a.numsol,a.inddoc,a.CodPromo,a.codafpori,a.TipoFonOri,a.TipoFonDes,a.IndEnvMail,a.IndMail,a.Fecregpro,a.Fecfirmsol,a.fecnac,b.fecnaci,a.status,a.idimagen,a.codusu, UsuRepTF "
    B1_2 = " FROM SOLICITUD a join HOJAVIDA b on a.numsol=b.numsol and a.fecsal=b.fecsal and a.secu=b.secu join Ctrldoc c on a.lote = c.lote and a.fecsal = c.fecsal "
'*****
    If OptTodos.Value Then
        B1_3 = "WHERE a.indDoc is not NULL and a.secu=0 "
    End If
    If OptSoloAfi.Value Then
        B1_3 = "WHERE a.indDoc = 'C' and a.secu=0 "
    End If
    If OptSoloTra.Value Then
        B1_3 = "WHERE a.indDoc = 'T' and a.secu=0 "
    End If
'*****
    If OptFecha.Value Then
        If TxtFecha <> "  /  /   " And TxtFecha <> "" Then
            B1_4 = " and a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
        Else
            MsgBox "Ingresar Fecha a buscar", vbCritical, "Error en la busqueda "
            TxtFecha.SetFocus
            Exit Sub
        End If
    Else
        If OptLote Then
            If val(TxtLote) > 0 Then
                B1_4 = " and a.Lote = '" & TxtLote & "' "
            Else
                MsgBox "Ingresar Lote a buscar", vbCritical, "Error en la busqueda "
                TxtLote.SetFocus
                Exit Sub
            End If
        Else
            If val(txtEnvio) > 0 Then
                B1_4 = " and c.envio = '" & txtEnvio & "' "
            Else
                MsgBox "Ingresar Envio a buscar", vbCritical, "Error en la busqueda "
                TxtLote.SetFocus
                Exit Sub
            End If
        End If
    End If
'*****
    If OptTodosF.Value Then
        B1_5 = ""
    End If
    If OptSolo13.Value Then
        B1_5 = " and ( (a.TipoFonDes<>'' and a.TipoFonDes<>'2') ) "
    End If
    If OptSolo2.Value Then
        B1_5 = " and ( a.TipoFonDes=2 ) "
    End If
    If OptTFbAFPb.Value Then
        B1_5 = " and ( a.CodAfpOri='' or a.TipoFonOri='' or a.TipoFonDes='' ) "
    End If
    If OptFecFirma.Value Then
        B1_5 = " and ( a.Fecregpro>a.fecsal or a.Fecregpro<(a.fecsal-5) or a.Fecfirmsol>a.fecsal or a.Fecfirmsol<(a.fecsal-5) or a.Fecregpro<>a.Fecfirmsol or a.Fecregpro is null or a.Fecfirmsol is null ) "
    End If
    If OptFNacFdo.Value Then
        B1_5 = " and ( year(a.Fecnac)<=year(getdate())-64 or year(b.Fecnaci)<=year(getdate())-64  ) and a.tipofondes<>1 "
    End If

'*****
    B1_6 = " order by a.lote,a.numsol,a.secu "

    B1 = B1_1 + B1_2 + B1_3 + B1_4 + B1_5 + B1_6
    
    LlenaControles
End Sub

Private Sub LlenaControles()
    LblTotReg = 0
    
    If RsTipFon.State = 1 Then RsTipFon.Close
        
    RsTipFon.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    With RsTipFon
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            Set DtGTFon.DataSource = Nothing
            DtGTFon.Refresh
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Verificacion"
            Exit Sub
        Else
            Set DtGTFon.DataSource = RsTipFon
            DtGTFon.Refresh
            cmdGrabar.Enabled = True
            LblTotReg.Caption = RsTipFon.RecordCount
            CargaImag
        End If
    End With
End Sub

Private Sub DtGTFon_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "F"
    If Activado Then
        On Error GoTo Salir
        With RsTipFon
            xlote = .Fields("lote")
            xnumsol = .Fields("numsol")
            xidimagen = "01" + xnumsol
            FrmMuestraImg.Show
        End With
    Else
        MsgBox "Selecciona la solicitud a buscar", vbCritical, "Error en la busqueda "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo errores
    RsTipFon.Fields("UsuRepTF") = gNomUsu
    RsTipFon.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
    Exit Sub
errores:
    If Err.Number = -2147217864 Then
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
    Else
    MsgBox Err.Description, vbCritical, Caption
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

Private Sub CmdSalir_Click()
    xfrmrep = ""
    Unload Me
End Sub

Private Sub CargaImag()
    With RsTipFon
        xlote = Format(!Lote, "0000000#")
        xlote2 = Format(!Lote, "0000000#")
        xnumsol = Format(!Numsol, "000000#")
        xtipimag = "S"
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
    ImgEdit1.Height = FrmConsFondo2.Height - 7800
' ancho
    ImgEdit1.Width = FrmConsFondo2.Width - 300
' alto
'    DtGHV.Height = 1200 - ImgEdit1.Height
    ImgEdit1.Height = FrmConsFondo2.Height - 6500
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


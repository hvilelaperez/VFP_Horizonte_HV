VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmRepImag 
   Caption         =   "Consulta de Archivo Imagenes"
   ClientHeight    =   8520
   ClientLeft      =   180
   ClientTop       =   1080
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   14610
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
      Left            =   8880
      TabIndex        =   8
      Top             =   240
      Width           =   2895
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
         TabIndex        =   12
         Top             =   880
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
         TabIndex        =   11
         Top             =   400
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtLote 
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
         Height          =   360
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   10
         Top             =   840
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
         Height          =   380
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   4
      Top             =   1800
      Width           =   975
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
      Left            =   8880
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DtGImag 
      Height          =   8295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   14631
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Lotefisico"
         Caption         =   "Lote Fisico"
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
         DataField       =   "IDImagen"
         Caption         =   "ID Imagen"
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
         DataField       =   "Imagen"
         Caption         =   "Imagen"
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
         DataField       =   "Barcode"
         Caption         =   "Barcode"
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
         DataField       =   "Posicion"
         Caption         =   "Posicion"
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
         DataField       =   "Status"
         Caption         =   "Status"
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
         DataField       =   "Tipimag"
         Caption         =   "Tipo Imagen"
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
            Alignment       =   2
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1035.213
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdRepo 
      Caption         =   "&Reporte Vista Prev"
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
      Left            =   10320
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
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
      Height          =   495
      Left            =   12000
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   6015
      Left            =   8760
      TabIndex        =   13
      Top             =   2400
      Width           =   5775
      _Version        =   131074
      _ExtentX        =   10186
      _ExtentY        =   10610
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   372015360
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   37
      lReserved2      =   37
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
      Left            =   13320
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Left            =   13200
      TabIndex        =   6
      Top             =   480
      Width           =   1305
   End
End
Attribute VB_Name = "FrmRepImag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsConsImag As New Recordset        'TempConsulta
Public RsRepoImag As New Recordset        'TempConsulta
Dim RsImag As New Recordset
Public ximagen As String
Public xidimagen As String
Public xlote As String
Dim xlote2 As String
Public xfrmrep As String
'Dim xtipimag As String
Dim xrutavol As String
Dim xnumpag As Integer
'Dim xprimer As String
Dim SQLstr As String
Dim B1 As String
Dim Activado As Boolean

Private Sub DtGImag_HeadClick(ByVal ColIndex As Integer)
    RsConsImag.Sort = DtGImag.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "2326", "si-006", "BDTraspasosdig2"
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    DatosServidor
'***
    cmdGrabar.Enabled = False
    CmdVerImag.Enabled = False
    xfrmrep = "I"
    TxtFecha = gDate
    B1 = ""
End Sub

Private Sub CmdCons_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""

    If TxtFecha = "" And TxtLote = "" Then
        MsgBox "Debe ingresar algun dato de busqueda", vbInformation, "Error en Busqueda"
        TxtFecha.SetFocus
        Exit Sub
    End If
    B1_1 = "SELECT a.lotefisico,a.idimagen,a.imagen,a.barcode,a.posicion,a.lote,a.status,a.tipimag " & _
    "FROM _IMAGENES a left join CTRLDOC b on a.lotefisico = b.lote "

    B1_2 = "WHERE a.imagen is not null  "

'*****
    If OptFecha.Value Then
        B1_3 = " and b.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
    Else
        B1_3 = " and a.Lotefisico = '" & TxtLote & "' "
    End If
'*****
    B1_4 = " ORDER BY a.lotefisico, a.imagen "

    B1 = B1_1 + B1_2 + B1_3 + B1_4
   
    LlenaControles
End Sub

Private Sub LlenaControles()
    LblTotReg = 0
    
    If RsConsImag.State = 1 Then RsConsImag.Close
    RsConsImag.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    With RsConsImag
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Solicitudes"
            DtGImag.Refresh
            cmdGrabar.Enabled = False
            Exit Sub
        Else
            Set DtGImag.DataSource = RsConsImag
            DtGImag.Refresh
            cmdGrabar.Enabled = True
            LblTotReg.Caption = RsConsImag.RecordCount
            CargaImag
        End If
    End With
End Sub

Private Sub CmdRepo_Click()
    TxtLote = Format(TxtLote, "0000000#")
    If RsRepoImag.State = 1 Then RsRepoImag.Close
    
    If TxtFecha <> "" And TxtLote = "" Then
        SQLstr = "SELECT * FROM _imagenes, ctrldoc WHERE _imagenes.lotefisico = ctrldoc.lote and ctrldoc.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) ORDER BY _imagenes.lotefisico, imagen"
    End If
    If TxtLote <> "" Then
        SQLstr = "SELECT * FROM _imagenes WHERE lotefisico = '" & TxtLote & "' ORDER BY _imagenes.lotefisico, imagen"
    End If
    
    If TxtFecha = "" And TxtLote = "" Then
        MsgBox "Debe ingresar algun dato de busqueda", vbInformation, "Error en Busqueda"
        TxtFecha.SetFocus
    End If
    
    RsRepoImag.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtRImagLote.DataSource = RsRepoImag
    DtRImagLote.Show
End Sub

Private Sub TxtFecha_GotFocus()
    Marcado TxtFecha
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
    If KeyAscii = 13 Then TxtLote.SetFocus
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
    RsConsImag.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub CmdSalir_Click()
    xfrmrep = ""
    Unload Me
End Sub

Private Sub DtGImag_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "I"
    If Activado Then
        On Error GoTo Salir
        With RsConsImag
            xlote = .Fields("lotefisico")
            xidimagen = IIf(IsNull(!idImagen), "", !idImagen)  '.Fields("idimagen")
            ximagen = .Fields("imagen")
            FrmMuestraImg.Show
        End With
    Else
        MsgBox "Seleccione la Imagen a buscar", vbCritical, "Error en la busqueda "
    End If
Salir:
    Exit Sub
End Sub

Private Sub CargaImag()
    With RsConsImag
        xlote2 = Format(!lotefisico, "0000000#")
        xlote = .Fields("lotefisico")
        xidimagen = IIf(IsNull(!idImagen), "", !idImagen)
        ximagen = .Fields("imagen")

    End With
    xnumpag = 1
      
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient
    
    SQLImag = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico = '" & xlote & "' AND imagen = '" & ximagen & "' and Tipo = 'B' order by posicion "
    
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
    
'    ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 3)
'    ImgEdit1.ScrollPositionX = 0
   
'    ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    ImgEdit1.Refresh
    
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'' alto
'    ImgEdit1.Height = FrmRepSolic2.Height - 7800
'' ancho
'    ImgEdit1.Width = FrmRepSolic2.Width - 300
'' alto
''    DtGHV.Height = 1200 - ImgEdit1.Height
'    ImgEdit1.Height = FrmRepSolic2.Height - 6500
'End Sub

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



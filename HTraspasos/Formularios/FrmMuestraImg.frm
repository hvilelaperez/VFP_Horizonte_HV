VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "Imgedit.ocx"
Begin VB.Form FrmMuestraImg 
   Caption         =   "Muestra Imagen"
   ClientHeight    =   9945
   ClientLeft      =   1395
   ClientTop       =   1005
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   12615
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   8655
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   12375
      _Version        =   131076
      _ExtentX        =   21828
      _ExtentY        =   15266
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   373602816
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
   Begin VB.CommandButton CmdRelee 
      Caption         =   "&Refresh"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox CboTipo 
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
      Left            =   5520
      TabIndex        =   15
      Text            =   "Tipo"
      Top             =   240
      Width           =   855
   End
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
      Left            =   7920
      TabIndex        =   10
      Top             =   0
      Width           =   2895
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
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   615
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
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   615
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
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   615
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
         Width           =   615
      End
   End
   Begin VB.TextBox TxtloteI 
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
      Left            =   3360
      TabIndex        =   8
      Top             =   120
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
      Left            =   11280
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtImagen 
      Enabled         =   0   'False
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
      Left            =   3360
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Txtidimagen 
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
      Left            =   1200
      TabIndex        =   1
      Top             =   570
      Width           =   1335
   End
   Begin VB.TextBox Txtnumsol 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   90
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Lote"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TIF"
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
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ID Imagen"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Solicitud"
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
      TabIndex        =   2
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "FrmMuestraImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVol As New Recordset           '_Volumenes
Public RsImag As New Recordset           '_Imagenes
Public xrutavol As String

Private Sub Form_Load()
'***
    If FrmRepSolic.xfrmrep = "S" Then
        txtNumsol = FrmRepSolic.xnumsol
        Txtidimagen = Trim(FrmRepSolic.xidimagen)
        TxtloteI = Trim(FrmRepSolic.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = ""
    End If
    If FrmRepEmple.xfrmrep = "E" Then
        txtNumsol = FrmRepEmple.xnumsol
        Txtidimagen = Trim(FrmRepEmple.xidimagen)
        TxtloteI = Trim(FrmRepEmple.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = ""
    End If
    If FrmVeriRiaFec2.xfrmrep = "V" Then
        txtNumsol = FrmVeriRiaFec2.xnumsol
        Txtidimagen = Trim(FrmVeriRiaFec2.xidimagen)
        TxtloteI = Trim(FrmVeriRiaFec2.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = ""
    End If
    If FrmVeriCuspp.xfrmrep = "C" Then
        txtNumsol = FrmVeriCuspp.xnumsol
        Txtidimagen = Trim(FrmVeriCuspp.xidimagen)
        TxtloteI = Trim(FrmVeriCuspp.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = ""
    End If
    If FrmBuscaSol.xfrmrep = "B" Then
        txtNumsol = FrmBuscaSol.xnumsol
        Txtidimagen = Trim(FrmBuscaSol.xidimagen)
        TxtloteI = Trim(FrmBuscaSol.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = ""
    End If
    If FrmRepImag.xfrmrep = "I" Then
        txtNumsol = Mid(FrmRepImag.xidimagen, 3, 7)
        Txtidimagen = Trim(FrmRepImag.xidimagen)
        TxtloteI = Trim(FrmRepImag.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = Trim(FrmRepImag.ximagen)
    End If
    If FrmConsFondo.xfrmrep = "F" Then
        txtNumsol = FrmConsFondo.xnumsol
        Txtidimagen = Trim(FrmConsFondo.xidimagen)
        TxtloteI = Trim(FrmConsFondo.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = ""
    End If
    If FrmRepHV.xfrmrep = "D" Then
        txtNumsol = FrmRepHV.xnumsol
        Txtidimagen = Trim(FrmRepHV.xidimagen)
        TxtloteI = Trim(FrmRepHV.xlote)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = ""
    End If
                            
    txtNumsol.Enabled = False
    Txtidimagen.Enabled = False
    TxtloteI.Enabled = False
    TxtImagen.Enabled = False
    With cboTipo
      .AddItem "S"
      .AddItem "E"
      .AddItem "V"
      .AddItem "D"
      .AddItem "B"
      .AddItem "M"
      .AddItem "H"
      .AddItem "O"
      .AddItem "Todos"
      .ListIndex = 0
    End With

    If FrmRepImag.xfrmrep = "I" Then
        cboTipo = "Todos"
    End If
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient

    'Call Con.AbrirTabla("select * from _Volumenes", RsVol, adOpenStatic, adLockReadOnly)
    If cboTipo = "Todos" Then
        SQLstr = " select a.*, b.Ruta from _Imagenes a join _volumenes b on " & _
                 " a.idvolumen = b.idvolumen  where Lotefisico = '" & TxtloteI & "' order by imagen "
    Else
        SQLstr = "select a.*, b.Ruta from _Imagenes a join _volumenes b on " & _
                 " a.idvolumen = b.idvolumen where Tipimag = '" & cboTipo & "' order by imagen "
    End If
    RsImag.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
'    If xrutavol = "" Then
'        With RsVol
'            .MoveFirst
'            xrutavol = .Fields("ruta")
'        End With
'    End If

    
    With RsImag  'coneccion a _Imagenes
        .MoveFirst
        If FrmRepImag.xfrmrep <> "I" Then
            .Find "[idimagen]='" & Txtidimagen & "'"
        Else
            .Find "[imagen]='" & TxtImagen & "'"
        End If

        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            If .EOF Then .MovePrevious
            TxtImagen = IIf(IsNull(!Imagen), "", !Imagen)
            TxtloteI = Format(!lotefisico, "0000000#")
        End If
        xrutavol = .Fields("Ruta")
    End With
    
    ImgEdit1.Image = xrutavol & "\" & TxtloteI & "\" & TxtImagen & ".TIF"
    ImgEdit1.Display
    ImgEdit1.FitTo 0

End Sub

Private Sub CmdRelee_Click()
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient
    If cboTipo = "Todos" Then
        SQLstr = "select * from _Imagenes where Lotefisico = '" & TxtloteI & "' order by imagen "
    Else
        SQLstr = "select * from _Imagenes where Tipimag = '" & cboTipo & "' order by imagen "
    End If
    RsImag.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    With RsImag  'coneccion a _Imagenes
        .MoveFirst
        If FrmRepImag.xfrmrep <> "I" Then
            .Find "[idimagen]='" & Txtidimagen & "'"
        Else
            .Find "[imagen]='" & TxtImagen & "'"
        End If

        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            If .EOF Then .MovePrevious
            TxtImagen = IIf(IsNull(!Imagen), "", !Imagen)
            TxtloteI = Format(!lotefisico, "0000000#")
        End If
    End With
    
    If Dir(xrutavol & "\" & TxtloteI & "\" & TxtImagen & ".TIF", vbArchive) <> "" Then
        ImgEdit1.Image = xrutavol & "\" & TxtloteI & "\" & TxtImagen & ".TIF"
        ImgEdit1.Display
        ImgEdit1.FitTo 0
    Else
        ImgEdit1.ClearDisplay
    End If
    
End Sub

Private Sub CmdSalir_Click()
    FrmRepSolic.xfrmrep = ""
    FrmRepEmple.xfrmrep = ""
    FrmVeriRiaFec2.xfrmrep = ""
    FrmVeriCuspp.xfrmrep = ""
    FrmBuscaSol.xfrmrep = ""
    FrmRepImag.xfrmrep = ""
    FrmRepHV.xfrmrep = ""
    txtNumsol = ""
    Txtidimagen = ""
    TxtloteI = ""
    Unload Me
End Sub

Private Sub CmdIr_Click(Index As Integer)
    With RsImag
    Select Case Index
        Case 0      ' Inicio
            .MoveFirst
            LlenaControles
            CmdIr(0).Enabled = False
            CmdIr(1).Enabled = False
            CmdIr(2).Enabled = True
            CmdIr(3).Enabled = True
        Case 1      'Anterior
            .MovePrevious
            LlenaControles
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
        Case 2      'Siguiente
            .MoveNext
            LlenaControles
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
        Case 3      'Final
            .MoveLast
            LlenaControles
            CmdIr(0).Enabled = True
            CmdIr(1).Enabled = True
            CmdIr(2).Enabled = False
            CmdIr(3).Enabled = False
    End Select
    End With
End Sub

Private Sub LlenaControles()
    If RsImag.EOF Then Exit Sub
    If RsImag.BOF Then Exit Sub
    With RsImag
        Txtidimagen = IIf(IsNull(!idImagen), "", !idImagen)
        txtNumsol = Mid(Txtidimagen, 3, 7)
        TxtloteI = IIf(IsNull(!lotefisico), "", !lotefisico)
        TxtloteI = Format(TxtloteI, "0000000#")
        TxtImagen = IIf(IsNull(!Imagen), "", !Imagen)
        If Dir(xrutavol & "\" & TxtloteI & "\" & TxtImagen & ".TIF", vbArchive) <> "" Then
            ImgEdit1.Image = xrutavol & "\" & TxtloteI & "\" & TxtImagen & ".TIF"
            ImgEdit1.Display
            ImgEdit1.FitTo 0
        Else
            ImgEdit1.ClearDisplay
        End If
    End With
End Sub

Private Sub ImgEdit1_DblClick()
    ImgEdit1.FitTo 0
End Sub

Private Sub ImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Local Error Resume Next
    ImgEdit1.ZoomToSelection
End Sub


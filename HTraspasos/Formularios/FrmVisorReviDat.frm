VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmVisorReviDat 
   Caption         =   "Visor Revisa Datos"
   ClientHeight    =   10440
   ClientLeft      =   1395
   ClientTop       =   675
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   12975
   Begin VB.TextBox TxtNom2 
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
      Left            =   3960
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox TxtNom1 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox TxtCuspp 
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
      Left            =   6600
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox TxtApeMat 
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
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox TxtApePat 
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
      Left            =   840
      TabIndex        =   4
      Text            =   "AAAA"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton CmdBusca 
      Caption         =   "&Buscar"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   1095
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
      Left            =   8640
      TabIndex        =   16
      Top             =   600
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
      Height          =   360
      Left            =   840
      TabIndex        =   0
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
      Height          =   375
      Left            =   11760
      TabIndex        =   10
      Top             =   720
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
      Height          =   360
      Left            =   11040
      TabIndex        =   9
      Top             =   120
      Width           =   1695
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
      Height          =   360
      Left            =   9120
      TabIndex        =   3
      Top             =   90
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
      Height          =   360
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin ImgeditLibCtl.ImgEdit ImgEditVista 
      Height          =   8775
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   12735
      _Version        =   131074
      _ExtentX        =   22463
      _ExtentY        =   15478
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   63236608
      OcrZoneVisibility=   -3516
      AnnotationOcrType=   25774
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25774
      sReserved1      =   -3516
      sReserved2      =   -3516
      lReserved1      =   1241728
      lReserved2      =   1241728
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Nombre2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3120
      TabIndex        =   25
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Nombre1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Cuspp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6000
      TabIndex        =   23
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Ap Mat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3120
      TabIndex        =   22
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ap Pat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Lote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TIF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10680
      TabIndex        =   14
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ID Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8160
      TabIndex        =   12
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Solicitud"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmVisorReviDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVol As New Recordset           '_Volumenes
Public RsImag As New Recordset           '_Imagenes
Public RsSoli As New Recordset          'Solicitud
Dim xrutavol As String


Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "DIGITACION\DIGITACION", "BDTraspasos2"
'    Con.ConexionSQL "sa", "2326", "si-006", "TraspasosPiloto"
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    DatosServidor
'***
    TxtLoteI = ""
    TxtNumSol = ""
    Txtidimagen = ""
    TxtImagen = ""
    TxtApePat = ""
    TxtApeMat = ""
    TxtNom1 = ""
    TxtNom2 = ""
    TxtCUSPP = ""
    TxtDNI = ""
    TxtBoleta = ""
    TxtFNac = ""
    TxtFSuscrip = ""
    Txtidimagen.Enabled = False
    TxtImagen.Enabled = False
    TxtApePat.Enabled = False
    TxtApeMat.Enabled = False
    TxtNom1.Enabled = False
    TxtNom2.Enabled = False
    TxtCUSPP.Enabled = False
'    TxtDNI.Enabled = False
'    TxtBoleta.Enabled = False
'    TxtFNac.Enabled = False
'    TxtFSuscrip.Enabled = False
    CmdBusca.Enabled = False
    FraIr.Enabled = False
    
    Call Con.AbrirTabla("select * from _Volumenes", RsVol, adOpenStatic, adLockReadOnly)

    With RsVol
        .MoveFirst
        xrutavol = .Fields("ruta")
    End With
    
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient

    SQLstr = "select * from _Imagenes where Tipimag = 'S'"
    RsImag.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic

End Sub

Private Sub TxtLoteI_Change()
    CmdBusca.Enabled = Val(TxtLoteI) > 0 And Val(TxtNumSol) > 0
    FraIr.Enabled = Val(TxtLoteI) > 0 And Val(TxtNumSol) > 0
End Sub

Private Sub CmdBusca_Click()
    TxtLoteI = Format(TxtLoteI, "0000000#")
    TxtNumSol = Format(TxtNumSol, "000000#")
    
    If RsSoli.State = 1 Then RsSoli.Close
    RsSoli.CursorLocation = adUseClient

    SQLstr = "select * from Solicitud where secu = 0 and lote = '" & TxtLoteI & "' order by lote,numsol"
    RsSoli.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    With RsSoli
        .MoveFirst
        '.Find "[lote]='" & TxtloteI & "'"
        .Find "[numsol]='" & TxtNumSol & "'"
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            If .EOF Then .MovePrevious
            LlenaControles
            MuestraImg
        End If
    End With
    
End Sub

Private Sub MuestraImg()
    
    With RsImag  'coneccion a _Imagenes
        .MoveFirst
        .Find "[idimagen]='" & Txtidimagen & "'"
        '.Find "[Imagen]='" & TxtArchivo & "'"
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            'If .EOF Then .MovePrevious
            TxtImagen = IIf(IsNull(!imagen), "", !imagen)
        End If
    End With
    ImgEditVista.Image = xrutavol & "\" & TxtLoteI & "\" & TxtImagen & ".TIF"
                
    ImgEditVista.Zoom = 50
    ImgEditVista.Display
    'ImgEdit1.Draw Izq, Top, Width, Height
    'ImgEditVista.Draw 120, 1440, 12735, 8895
    ImgEditVista.ScrollPositionY = (Int(ImgEditVista.ImageScaleHeight / 4) * 3)
    ImgEditVista.ScrollPositionX = 0
    
    ImgEditVista.Refresh
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdIr_Click(Index As Integer)
    With RsSoli
    Select Case Index
        Case 0      ' Inicio
            .MoveFirst
            CmdIr(0).Enabled = False
            CmdIr(1).Enabled = False
            CmdIr(2).Enabled = True
            CmdIr(3).Enabled = True
            LlenaControles
            MuestraImg
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
            MuestraImg
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
            MuestraImg
        Case 3      'Final
            .MoveLast
            CmdIr(0).Enabled = True
            CmdIr(1).Enabled = True
            CmdIr(2).Enabled = False
            CmdIr(3).Enabled = False
            LlenaControles
            MuestraImg
    End Select
    End With
End Sub

Private Sub LlenaControles()
    If RsSoli.EOF Then Exit Sub
    If RsSoli.BOF Then Exit Sub
    With RsSoli
        'TxtloteI = !lote
        TxtLoteI = Format(!Lote, "0000000#")
        TxtNumSol = Format(!Numsol, "000000#")
        Txtidimagen = !idimagen
        TxtApePat = IIf(IsNull(!ApePat), "", !ApePat)
        TxtApeMat = IIf(IsNull(!ApeMat), "", !ApeMat)
        TxtNom1 = IIf(IsNull(!Nombre1), "", !Nombre1)
        TxtNom2 = IIf(IsNull(!Nombre2), "", !Nombre2)
        TxtCUSPP = IIf(IsNull(!CUSPP), "", !CUSPP)
'        TxtDNI = IIf(IsNull(!CopiaSim), "", !CopiaSim)
'        TxtBoleta = IIf(IsNull(!BolPago), "", !BolPago)
'        TxtFNac = IIf(IsNull(!FecNac), "", !FecNac)
'        TxtFSuscrip = IIf(IsNull(!FecFirmSol), "", !FecFirmSol)
        Txtidimagen = IIf(IsNull(!idimagen), "", !idimagen)
    End With
End Sub

Private Sub ImgEditVista_DblClick()
    ImgEditVista.FitTo 0
End Sub

Private Sub ImgEditVista_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Local Error Resume Next
    ImgEditVista.ZoomToSelection
End Sub

Private Sub TxtloteI_GotFocus()
    Marcado TxtLoteI
End Sub

Private Sub Txtnumsol_GotFocus()
    Marcado TxtNumSol
End Sub

Private Sub TxtloteI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtNumSol.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdBusca.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

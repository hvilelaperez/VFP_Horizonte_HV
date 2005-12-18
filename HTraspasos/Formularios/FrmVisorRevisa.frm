VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmVisorRevisa 
   Caption         =   "Visor Revisa Datos"
   ClientHeight    =   10440
   ClientLeft      =   1395
   ClientTop       =   675
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   12975
   Begin VB.TextBox TxtFSuscrip 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6960
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TxtFNac 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TxtBoleta 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox TxtDNI 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox TxtCuspp 
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox TxtApeMat 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox TxtApePat 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton CmdBusca 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame FraIr 
      Caption         =   "Ir"
      Height          =   855
      Left            =   8520
      TabIndex        =   18
      Top             =   480
      Width           =   3135
      Begin VB.CommandButton CmdIr 
         Caption         =   ">>"
         Height          =   495
         Index           =   3
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdIr 
         Caption         =   ">"
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdIr 
         Caption         =   "<"
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdIr 
         Caption         =   "<<"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox TxtloteI 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   11880
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TxtImagen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11280
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Txtidimagen 
      Height          =   285
      Left            =   9120
      TabIndex        =   3
      Top             =   90
      Width           =   1215
   End
   Begin VB.TextBox Txtnumsol 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin ImgeditLibCtl.ImgEdit ImgEditVista 
      Height          =   8895
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   12735
      _Version        =   131076
      _ExtentX        =   22463
      _ExtentY        =   15690
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   84609792
      ForceFileLinking1x=   -1  'True
      sReserved1      =   -3516
      sReserved2      =   -3516
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Label Label11 
      Caption         =   "Fec Suscrip"
      Height          =   255
      Left            =   6000
      TabIndex        =   29
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fec Nac"
      Height          =   195
      Left            =   3600
      TabIndex        =   28
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Boleta"
      Height          =   195
      Left            =   2040
      TabIndex        =   27
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "LE / DNI"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Cuspp"
      Height          =   195
      Left            =   6000
      TabIndex        =   25
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Ap Mat"
      Height          =   195
      Left            =   3120
      TabIndex        =   24
      Top             =   600
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ap Pat"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Lote"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TIF"
      Height          =   195
      Left            =   10680
      TabIndex        =   16
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ID Imagen"
      Height          =   195
      Left            =   8160
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Solicitud"
      Height          =   195
      Left            =   2160
      TabIndex        =   13
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "FrmVisorRevisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVol As New Recordset           '_Volumenes
Public RsImag As New Recordset           '_Imagenes
Public RsSoli As New Recordset          'Solicitud
Dim xrutavol As String


Private Sub Form_Load()
    Con.ConexionSQL "sa", "", "DIGITACION\DIGITACION", "BDTraspasos2"
'    Con.ConexionSQL "sa", "2326", "si-006", "TraspasosPiloto"
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
    DatosServidor
'***
    TxtloteI = ""
    Txtnumsol = ""
    Txtidimagen = ""
    TxtImagen = ""
    TxtApePat = ""
    TxtApeMat = ""
    TxtCuspp = ""
    TxtDNI = ""
    TxtBoleta = ""
    TxtFNac = ""
    TxtFSuscrip = ""
    Txtidimagen.Enabled = False
    TxtImagen.Enabled = False
    TxtApePat.Enabled = False
    TxtApeMat.Enabled = False
    TxtCuspp.Enabled = False
    TxtDNI.Enabled = False
    TxtBoleta.Enabled = False
    TxtFNac.Enabled = False
    TxtFSuscrip.Enabled = False
    
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

Private Sub CmdBusca_Click()
    TxtloteI = Format(TxtloteI, "0000000#")
    Txtnumsol = Format(Txtnumsol, "000000#")
    
    If RsSoli.State = 1 Then RsSoli.Close
    RsSoli.CursorLocation = adUseClient

    SQLstr = "select * from Solicitud where secu = 0 and lote = '" & TxtloteI & "' order by lote,numsol"
    RsSoli.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    With RsSoli
        .MoveFirst
        '.Find "[lote]='" & TxtloteI & "'"
        .Find "[numsol]='" & Txtnumsol & "'"
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
    ImgEditVista.Image = xrutavol & "\" & TxtloteI & "\" & TxtImagen & ".TIF"
                
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
        TxtloteI = Format(!lote, "0000000#")
        Txtnumsol = Format(!Numsol, "000000#")
        Txtidimagen = !idimagen
        TxtApePat = IIf(IsNull(!ApePat), "", !ApePat)
        TxtApeMat = IIf(IsNull(!ApeMat), "", !ApeMat)
        TxtCuspp = IIf(IsNull(!CUSPP), "", !CUSPP)
        TxtDNI = IIf(IsNull(!CopiaSim), "", !CopiaSim)
        TxtBoleta = IIf(IsNull(!BolPago), "", !BolPago)
        TxtFNac = IIf(IsNull(!FecNac), "", !FecNac)
        TxtFSuscrip = IIf(IsNull(!FecFirmSol), "", !FecFirmSol)
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
    Marcado TxtloteI
End Sub

Private Sub Txtnumsol_GotFocus()
    Marcado Txtnumsol
End Sub

Private Sub TxtloteI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            Txtnumsol.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtnumsol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdBusca.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

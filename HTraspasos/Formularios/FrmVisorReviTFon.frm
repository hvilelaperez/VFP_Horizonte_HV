VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "Imgedit.ocx"
Begin VB.Form FrmVisorReviTFon 
   Caption         =   "Visor Revisa Datos"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   14895
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   14655
      Begin VB.CheckBox ChkCasa 
         Caption         =   "Mail de casa"
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
         Left            =   9120
         TabIndex        =   30
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox ChkEECC 
         Caption         =   "EECC x mail"
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
         Left            =   6960
         TabIndex        =   29
         Top             =   480
         Width           =   1695
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
         Height          =   350
         Left            =   1320
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtLote 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtAFPOri 
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
         Height          =   350
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtTFonOri 
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
         Height          =   350
         Left            =   4200
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtTFonDes 
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
         Height          =   350
         Left            =   5760
         TabIndex        =   6
         Top             =   480
         Width           =   495
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
         Left            =   12840
         TabIndex        =   7
         Top             =   360
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
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "AFP Orig"
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
         Height          =   240
         Left            =   2760
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tip Fdo Orig"
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
         Height          =   240
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tip Fdo Dest"
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
         Height          =   240
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingresar Lotes de Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6135
      Begin VB.TextBox TxtLoteI 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   1095
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
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtLoteF 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lote Inicial"
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
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lote Final"
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
         Height          =   240
         Left            =   2640
         TabIndex        =   19
         Top             =   360
         Width           =   960
      End
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   7935
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   14655
      _Version        =   131074
      _ExtentX        =   25850
      _ExtentY        =   13996
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   450759168
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   38
      lReserved2      =   38
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
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
      Left            =   6480
      TabIndex        =   12
      Top             =   120
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   240
         Width           =   615
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
      Left            =   13440
      TabIndex        =   3
      Top             =   360
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
      Height          =   350
      Left            =   11280
      TabIndex        =   9
      Top             =   480
      Width           =   1575
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
      Height          =   350
      Left            =   9720
      TabIndex        =   8
      Top             =   480
      Width           =   1335
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
      Left            =   11400
      TabIndex        =   11
      Top             =   240
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
      Left            =   9720
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmVisorReviTFon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVol As New Recordset           '_Volumenes
Public RsImag As New Recordset           '_Imagenes
Public RsSoli As New Recordset          'Solicitud
Dim xtipimag As String
Dim xrutavol As String
Dim ximagen As String
Dim xnumpag As Integer
Dim B1 As String

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    DatosServidor
'***
    TxtloteI = ""
    TxtLoteF = ""
    txtLote = ""
    txtNumsol = ""
    Txtidimagen = ""
    TxtImagen = ""
    TxtTFonOri = ""
    TxtTFonDes = ""
    txtAfpOri = ""
    Txtidimagen.Enabled = False
    TxtImagen.Enabled = False
    txtLote.Enabled = False
    txtNumsol.Enabled = False
    CmdBusca.Enabled = False
    FraIr.Enabled = False
    
End Sub

Private Sub CmdBusca_Click()
    xtipimag = "S"
    
    If RsSoli.State = 1 Then RsSoli.Close
    RsSoli.CursorLocation = adUseClient

    B1 = "SELECT a.fecsal,a.lote,a.secu,a.numsol,a.inddoc,a.codafpori,a.TipoFonOri,a.TipoFonDes,a.IndEnvMail,a.IndMail,a.status,a.idimagen,b.imagen FROM solicitud a join _Imagenes b on a.idimagen=b.idimagen WHERE a.Lote >= '" & TxtloteI & "' AND a.Lote <= '" & TxtLoteF & "' AND (b.BarCode = 0) and b.tipimag='S' and a.inddoc='T' order by a.lote,a.numsol"
    RsSoli.Open B1, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    
    With RsSoli
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            FraIr.Enabled = False
            Exit Sub
        Else
            .MoveFirst
            If .EOF Then .MovePrevious
            CmdIr(0).Enabled = False
            CmdIr(1).Enabled = False
            CmdIr(2).Enabled = True
            CmdIr(3).Enabled = True
            LlenaControles
            MuestraImg
        End If
    End With
    
End Sub

Private Sub LlenaControles()
    If RsSoli.EOF Then Exit Sub
    If RsSoli.BOF Then Exit Sub
    With RsSoli
        txtLote = Format(!Lote, "0000000#")
        txtNumsol = Format(!Numsol, "000000#")
        Txtidimagen = IIf(IsNull(!idImagen), "", !idImagen)
        TxtTFonOri = IIf(IsNull(!TipoFonOri), "", !TipoFonOri)
        TxtTFonDes = IIf(IsNull(!TipoFonDes), "", !TipoFonDes)
        txtAfpOri = IIf(IsNull(!CodAFPOri), "", !CodAFPOri)
        If !IndEnvMail = Null Then
            ChkEECC = 0
        Else
            If !IndEnvMail = False Then
                ChkEECC = 0
            Else
                ChkEECC = 1
            End If
        End If
        If !IndMail = Null Then
            ChkCasa = 0
        Else
            If !IndMail = False Then
                ChkCasa = 0
            Else
                ChkCasa = 1
            End If
        End If
    End With
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdGrabar_Click()
    With RsSoli
'        !TipoFonOri = Val(TxtTFonOri)
        !TipoFonOri = Trim(TxtTFonOri)
        !TipoFonDes = Trim(TxtTFonDes)
        !CodAFPOri = Trim(txtAfpOri)
        If ChkEECC.Value = 0 Then
            !IndEnvMail = False
        Else
            !IndEnvMail = True
        End If
        If ChkCasa.Value = 0 Then
            !IndMail = False
        Else
            !IndMail = True
        End If
    End With
    RsSoli.UpdateBatch
    MsgBox "Se grabaron registros", vbInformation, "Transaccion Grabada"
End Sub

Private Sub TxtLoteI_Change()
    CmdBusca.Enabled = val(TxtloteI) > 0 And val(TxtLoteF) > 0
    FraIr.Enabled = val(TxtloteI) > 0 And val(TxtLoteF) > 0
End Sub

Private Sub TxtLoteF_Change()
    CmdBusca.Enabled = val(TxtloteI) > 0 And val(TxtLoteF) > 0
    FraIr.Enabled = val(TxtloteI) > 0 And val(TxtLoteF) > 0
End Sub

Private Sub TxtloteI_GotFocus()
    Marcado TxtloteI
End Sub

Private Sub TxtLoteF_GotFocus()
    Marcado TxtLoteF
End Sub

Private Sub TxtLoteI_LostFocus()
    TxtloteI = Format(TxtloteI, "0000000#")
End Sub

Private Sub TxtLoteF_LostFocus()
    TxtLoteF = Format(TxtLoteF, "0000000#")
End Sub

Private Sub TxtloteI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtLoteF.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLoteF_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdBusca.SetFocus
        Case Else
            KeyAscii = 0
    End Select
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

Private Sub MuestraImg()
    
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient

    SQLImag = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico >= '" & TxtloteI & "' AND Lotefisico <= '" & TxtLoteF & "' AND BarCode = 0 and idimagen = '" & Txtidimagen & "' and Tipo = 'B' and TipImag = '" & xtipimag & "' order by posicion "
    
    RsImag.Open SQLImag, Con.Cnx, adOpenStatic, adLockReadOnly
    
    With RsImag  'coneccion a _Imagenes
        .MoveFirst
        .Find "[idimagen]='" & Txtidimagen & "'"
        '.Find "[Imagen]='" & TxtArchivo & "'"
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            'If .EOF Then .MovePrevious
            TxtImagen = IIf(IsNull(!Imagen), "", !Imagen)
            ximagen = IIf(IsNull(!Imagen), "", !Imagen)
            xrutavol = .Fields("ruta")
            xrutavol = Replace(xrutavol, "\\", "\")
            xidvolumen = IIf(IsNull(!idvolumen), "", !idvolumen)
        End If
    End With
    ImgEdit1.Image = xrutavol & "\" & txtLote & "\" & TxtImagen & ".TIF"
                
    ImgEdit1.Zoom = 50
    ImgEdit1.Display
    'ImgEdit1.Draw Izq, Top, Width, Height
    'ImgEdit1.Draw 120, 1440, 12735, 8895
    ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 3)
    ImgEdit1.ScrollPositionX = 0
    
    ImgEdit1.Refresh
End Sub

Private Sub ImgEdit1_DblClick()
    ImgEdit1.FitTo 0
End Sub

Private Sub ImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Local Error Resume Next
    ImgEdit1.ZoomToSelection
End Sub


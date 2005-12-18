VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form FrmVisorArchivo 
   Caption         =   "Visor de Archivos"
   ClientHeight    =   8670
   ClientLeft      =   390
   ClientTop       =   1665
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   14340
   Begin VB.CommandButton cmdGImag 
      Caption         =   "Generar &Imágenes"
      Height          =   495
      Left            =   4080
      TabIndex        =   37
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenerarRegistros 
      Caption         =   "&Generar Registros"
      Height          =   495
      Left            =   2160
      TabIndex        =   36
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdPosicion 
      Caption         =   "&Actualizar Posicion"
      Height          =   495
      Left            =   240
      TabIndex        =   35
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnlazar 
      Caption         =   "&De"
      Height          =   495
      Left            =   3360
      TabIndex        =   33
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton CmdO 
      Caption         =   "&O"
      Height          =   495
      Left            =   2520
      TabIndex        =   32
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton CmdM 
      Caption         =   "&M"
      Height          =   495
      Left            =   2040
      TabIndex        =   31
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton CmdB 
      Caption         =   "&B"
      Height          =   495
      Left            =   1560
      TabIndex        =   30
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "&D"
      Height          =   495
      Left            =   1080
      TabIndex        =   29
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton CmdS 
      Caption         =   "&S"
      Height          =   495
      Left            =   600
      TabIndex        =   28
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton CmdH 
      Caption         =   "&H"
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   6960
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   5175
      Begin VB.TextBox txtposicion 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4440
         MaxLength       =   1
         TabIndex        =   34
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TxtImagen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TxtLoteI 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtArchivo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtTipImag 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         MaxLength       =   1
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtLote 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Archivo Imag"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Visor"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Imag"
         Height          =   195
         Left            =   4320
         TabIndex        =   22
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   3000
         TabIndex        =   21
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   6960
      Width           =   1215
   End
   Begin ImgeditLibCtl.ImgEdit ImgEditVista 
      Height          =   7695
      Left            =   5640
      TabIndex        =   11
      Top             =   360
      Width           =   8535
      _Version        =   131074
      _ExtentX        =   15055
      _ExtentY        =   13573
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   390810880
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
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox TxtRuta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox TxtUnidad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Lblmsg 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   7560
      Width           =   5055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ruta"
      Height          =   195
      Left            =   4320
      TabIndex        =   14
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Unidad"
      Height          =   195
      Left            =   4320
      TabIndex        =   13
      Top             =   4320
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Imagen:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label LblVista 
      AutoSize        =   -1  'True
      Caption         =   "Vista previa"
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
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label LblArchivo 
      AutoSize        =   -1  'True
      Caption         =   "Archivo"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
   Begin VB.Label LblCarpeta 
      AutoSize        =   -1  'True
      Caption         =   "Carpeta"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   675
   End
   Begin VB.Label LblUnidad 
      AutoSize        =   -1  'True
      Caption         =   "Unidad"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmVisorArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVol As New Recordset           '_Volumenes
Public RsImag As New Recordset           '_Imagenes
'Public RsT As New Recordset             'TMPTipodoc
Dim xrutavol As String
Dim xidvol As String

Private Sub cmdEnlazar_Click()
  With RsImag
        .Fields("TipImag") = Null
        .Fields("posicion") = Null
        .Fields("idimagen") = Null
        .Fields("status") = "1"
        .Fields("barcode") = "0"
        .UpdateBatch
    End With
    Lblmsg = ""
    Lblmsg = "Se quitó enlace a imagen " & TxtImagen & ""
End Sub

Private Sub cmdGenerarRegistros_Click()
On Error GoTo ERRORES
'    File1.Path = Dir1.Path
'    TxtRuta = Dir1.Path
'    TxtLote = Right(TxtRuta, 8)


If val(Me.TxtLote) = 0 Then MsgBox "Debe selecciona un Lote", vbCritical, Caption: Exit Sub

Set comando = New Command

'si existe imagenes del lote
If fgDatosql("select lotefisico from _imagenes where lotefisico = " & val(TxtLote) & " ") <> "" Then
    If MsgBox("Existe Registro de Imagenes para el lotefisico  " & TxtLote & " Debe Eliminar los registros existente ", vbYesNo, Caption) = vbYes Then
        comando.ActiveConnection = Con.Cnx
        comando.CommandText = "Delete from _imagenes where lotefisico = " & val(Me.TxtLote) & " "
        comando.Execute
        MsgBox "Registros Eliminados", 64, Caption
     Else
     Exit Sub
    End If
    
End If


'Verifica si existe lote : sino lo crea

If fgDatosql("Select LoteFisico from _scan where LoteFisico = '" & Format(TxtLote.Text, "0000000#") & "'") = "" Then
    'Crear el Lote
 With comando
    .ActiveConnection = Con.Cnx
    .CommandText = " sp_InsertarPreScan ?,?,? "
    .Parameters("@idVolumen") = xidvol
    .Parameters("@Lote") = val(TxtLote)
    .Parameters("@LoteFisico") = Format(val(TxtLote), "0000000#")
    .Execute
    
    .ActiveConnection = Con.Cnx
    .CommandText = " sp_InsertarScan ?,?,? "
    .Parameters("@idVolumen") = xidvol
    .Parameters("@Lote") = val(TxtLote)
    .Parameters("@LoteFisico") = Format(val(TxtLote), "0000000#")
    .Execute
    
    .ActiveConnection = Con.Cnx
    .CommandText = "Update _Scan set Estado = 1 where idvolumen = " & xidvol & "  and lote = " & val(TxtLote) & ""
    .Execute
    
 End With
End If


'va al primer registro
File1.ListIndex = 0


With comando

    For i = 0 To File1.ListCount - 1
        .ActiveConnection = Con.Cnx
        .CommandText = "exec sp_RegistrarImagen ?,?,?,?,?,?"
        .Parameters("@idVolumen") = xidvol
        .Parameters("@Lote") = val(TxtLote)
        .Parameters("@LoteFisico") = Format(val(TxtLote), "0000000#")
        .Parameters("@Imagen") = Trim(Mid(Trim(File1.FileName), 1, 13))
        .Parameters("@Tipo") = "B"
        .Parameters("@Replicar") = 0
        .Execute
        If (File1.ListIndex + 1) <= (File1.ListCount - 1) Then File1.ListIndex = File1.ListIndex + 1
    Next i
    
End With
Exit Sub
ERRORES:
MsgBox Err.Description, vbCritical, Caption
End Sub

Private Sub cmdGImag_Click()
frmGenImagenes.Show 1
End Sub

Private Sub cmdPosicion_Click()
  With RsImag
        .Fields("posicion") = txtposicion
        .UpdateBatch
    End With
    Lblmsg = ""
    Lblmsg = "Se Actualizó posición a imagen " & TxtImagen & ""

End Sub

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "DIGITACION\DIGITACION", "BDTraspasos"
'    Con.ConexionSQL "sa", "2326", "si-006", "TraspasosPiloto"
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    DatosServidor
'***
'    If RsT.State = 1 Then RsT.Close
'    RsT.CursorLocation = adUseClient

    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient

    Call Con.AbrirTabla("select * from _Volumenes", RsVol, adOpenStatic, adLockReadOnly)
    Call Con.AbrirTabla("select * from _Imagenes", RsImag, adOpenStatic, adLockBatchOptimistic)
    'Call Con.AbrirTabla("select * from TMPTipoDoc", RsT, adOpenStatic, adLockBatchOptimistic)
    With RsVol
        .MoveFirst
        
        xrutavol = .Fields("ruta")
        xidvol = .Fields("idVolumen")
    End With
    Drive1.Drive = Mid(xrutavol, 1, 2)
    Dir1.Path = Replace(xrutavol, "\\", "\")

    With cboTipo
        .AddItem "*.tif"
        .AddItem "*.bmp"
        .AddItem "*.gif"
        .AddItem "*.jpg"
        .AddItem "*.tif;*.bmp;*.gif;*.jpg"
        .ListIndex = 0
    End With
    TxtUnidad = Drive1.Drive
    txtruta = Dir1.Path
    txtArchivo = File1.FileName
    TxtLote = Right(txtruta, 8)
    Lblmsg = ""
End Sub

Private Sub Drive1_Change()
    TxtUnidad = Drive1.Drive
    On Error GoTo DRIVE_ERROR
        Dir1.Path = Drive1.Drive
    Exit Sub
DRIVE_ERROR:
    MsgBox "ERROR: Unidad no preparada", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    txtruta = Dir1.Path
    TxtLote = Right(txtruta, 8)
    dd = Split(Dir1.Path, "\")
End Sub

Private Sub cboTipo_Click()
   File1.Pattern = cboTipo
End Sub

Private Sub File1_Click()
    Dim ARCHIVO As String
    TxtTipImag = ""
    TxtImagen = ""
    TxtloteI = ""
    Lblmsg = ""
    ARCHIVO = File1.Path & "\" & File1.FileName
    'ARCHIVO = File1.Path & "\" & File1.List(File1.ListIndex)
    ImgEditVista.Image = ARCHIVO
    ImgEditVista.Display
    ImgEditVista.FitTo 0
'   imgEditvista.Picture = file1.Path & "\" & file1.FileName
    txtArchivo = File1.FileName
    txtArchivo = Mid(txtArchivo, 1, InStr(txtArchivo, ".") - 1)
    
    With RsImag  'coneccion a _Imagenes
        .MoveFirst
        .Find "[Imagen]='" + txtArchivo + "'"
        If Not .EOF Then
            TxtImagen = IIf(IsNull(!Imagen), "", !Imagen)
            TxtTipImag = IIf(IsNull(!TipImag), "", !TipImag)
            TxtloteI = IIf(IsNull(!Lote), "", !Lote)
            Me.txtposicion = IIf(IsNull(!Posicion), "", !Posicion)
        End If
    End With

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

'Private Sub TxtTipo_Change()
'    CmdGrabar.Enabled = True
'End Sub

Private Sub CmdH_Click()
    TxtTipImag = "H"
    GrabaTipo
End Sub

Private Sub CmdS_Click()
    TxtTipImag = "S"
    GrabaTipo
End Sub

Private Sub CmdD_Click()
    TxtTipImag = "D"
    GrabaTipo
End Sub

Private Sub CmdB_Click()
    TxtTipImag = "B"
    GrabaTipo
End Sub

Private Sub CmdM_Click()
    TxtTipImag = "M"
    GrabaTipo
End Sub

Private Sub CmdO_Click()
    TxtTipImag = "O"
    GrabaTipo
End Sub

Private Sub GrabaTipo()
    With RsImag
        .Fields("TipImag") = TxtTipImag
        .UpdateBatch
    End With
    Lblmsg = ""
    Lblmsg = "Se grabo Tipo " & TxtTipImag
End Sub

Private Sub ImgEditVista_DblClick()
ImgEditVista.FitTo 0
End Sub

'Private Sub CmdGrabar_Click()
'    If TxtTipo = "" Then
'        Lblmsg = ""
'        Lblmsg = "Debe ingresar Tipo "
'        'MsgBox "Debe ingresar Tipo ", vbCritical, "Error en Ingreso de Tipo"
'        TxtTipo.SetFocus
'        Exit Sub
'    End If
'    With RsT
'        .AddNew
'        .Fields("ImagenTif") = TxtArchivo
'        .Fields("TipoDoc") = TxtTipo
'        .Fields("Loteimag") = TxtLote
'        .UpdateBatch
'    End With
'    Lblmsg = ""
'    Lblmsg = "Se grabo Tipo " & TxtTipo
'    'MsgBox "Se grabo Tipo ", vbInformation, "Graba Tipo de Documento"
'    TxtTipo.SetFocus
'End Sub

'Private Sub TxtTipo_GotFocus()
'    Marcado TxtTipo
'End Sub

'Private Sub TxtTipo_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    Select Case KeyAscii
'        Case 66
'        Case 68
'        Case 72
'        Case 77
'        Case 83
'        Case 8
'        Case 13
'            CmdGrabar.SetFocus
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub
Private Sub ImgEditVista_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub ImgEditVista_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
On Local Error Resume Next
ImgEditVista.ZoomToSelection
End Sub

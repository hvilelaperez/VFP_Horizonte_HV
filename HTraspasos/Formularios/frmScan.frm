VERSION 5.00
Object = "{00120421-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltbar12n.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltocx12n.ocx"
Object = "{84BCAF03-5BB5-11CF-96D7-00AA004A32FB}#3.0#0"; "kscn3032.ocx"
Object = "{9E0EC5A3-23FD-11CF-AFBE-00AA006B77CB}#3.0#0"; "kviw3032.ocx"
Begin VB.Form frmScan 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PROFUTURO [ Scan ]"
   ClientHeight    =   10755
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   717
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Terminar digitalización ..."
      Top             =   4200
      Width           =   2655
   End
   Begin KScanLib.KScan KScan1 
      Left            =   3240
      Top             =   10200
      _Version        =   196608
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DeviceEngineMaster=   "Software VRS -- ISIS scanning"
      ScanContinuous  =   0   'False
   End
   Begin KViewLib.KView KView1 
      Height          =   9735
      Left            =   3000
      TabIndex        =   28
      Top             =   120
      Width           =   8655
      _Version        =   196608
      _ExtentX        =   15266
      _ExtentY        =   17171
      _StockProps     =   101
      BackColor       =   4210752
   End
   Begin LEADLib.LEAD LEAD1 
      Height          =   375
      Left            =   10680
      TabIndex        =   26
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
      _Version        =   65539
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   229
      ScaleHeight     =   25
      ScaleWidth      =   33
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
   Begin MSMask.MaskEdBox mskReplicar 
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   4560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdInicioLote 
      Caption         =   "Iniciar nuevo lote "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton cmdRescan 
      Caption         =   "Redigitalizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
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
      TabStop         =   0   'False
      ToolTipText     =   "Redigitalizar"
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton cmdPararScan 
      Caption         =   "Terminar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
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
      TabStop         =   0   'False
      ToolTipText     =   "Terminar digitalización ..."
      Top             =   3855
      Width           =   2655
   End
   Begin VB.CommandButton cmdSeleccionScanner 
      Caption         =   "Scanner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Seleccionar scanner ..."
      Top             =   2415
      Width           =   2655
   End
   Begin VB.CommandButton cmdPropiedades 
      Caption         =   "Propiedades"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Propiedades del scanner ..."
      Top             =   2775
      Width           =   2655
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Digitalizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Digitalizar ..."
      Top             =   3135
      Width           =   2655
   End
   Begin VB.CommandButton cmdNuevoLote 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Nuevo lote ..."
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdVolumen 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Volumenes ..."
      Top             =   375
      Width           =   375
   End
   Begin VB.CommandButton cmdCerrarLote 
      Caption         =   "Teminar lote actual"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9360
      Width           =   2655
   End
   Begin VB.CommandButton cmdCerrarVolumen 
      Caption         =   "Terminar volumen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Terminar digitalización ..."
      Top             =   11160
      Width           =   2655
   End
   Begin VB.ListBox lstImagenes 
      Appearance      =   0  'Flat
      Height          =   1290
      ItemData        =   "frmScan.frx":030A
      Left            =   120
      List            =   "frmScan.frx":030C
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdActualizarCapacidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2400
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Actualizar  ..."
      Top             =   7095
      Width           =   375
   End
   Begin VB.CheckBox chkReplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Replicar imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   1620
   End
   Begin LTBARLibCtl.LEADBarCode LEADBarCode1 
      Left            =   4560
      OleObjectBlob   =   "frmScan.frx":030E
      Top             =   9240
   End
   Begin VB.Label lblDirectorio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   24
      Top             =   1335
      Width           =   810
   End
   Begin VB.Label lblDirectorioVolumen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   1575
      Width           =   2655
   End
   Begin VB.Label lblTituloLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   735
      Width           =   375
   End
   Begin VB.Label lblVolumen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblTituloVolumen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Volumen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   750
   End
   Begin VB.Label lblLote2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblTotalLote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label lblTotalLotes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total lotes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   870
   End
   Begin VB.Label lblLote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Label lblTotalImagenes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6525
      Width           =   2655
   End
   Begin VB.Label lblCapacidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño del volumen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   1725
   End
   Begin VB.Label lblCapacidadVolumen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label lblImagenes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imagenes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   4845
      Width           =   825
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Eliminar"
      Visible         =   0   'False
      Begin VB.Menu mnuEliminar 
         Caption         =   "Eliminar imagenes"
      End
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strDirectorioImagenes As String
Public lngVolumenActual As String
Dim bolTransaccion As Boolean

Dim strImagenScan As String
Dim strNuevaImagen As String
Public lngImagenActual As Long
Public lngTamañoVolumen As Long
Public strPrefijoRuta As String
Public lngTotalImagenes As Long
Dim bytPaso As Byte
Dim bolRescan As Boolean
Public bolIniciarLote As Boolean
Dim bytPesoIndex  As Integer
Private Sub chkReplicar_Click()
    If chkReplicar.Value = 1 Then
        mskReplicar.SetFocus
    End If
End Sub

Private Sub cmdActualizarCapacidad_Click()
    Me.MousePointer = 11
    If lngIdVolumenActual > 0 Then TamañoDirectorio strDirectorioImagenes
    Me.MousePointer = 0
End Sub
Private Sub cmdCerrarLote_Click()
    TerminarLote
End Sub
Sub TerminarLote()
If Not ValidaLote Then Exit Sub
    If lngIdVolumenActual > 0 Then
        If vbYes = MsgBox("Desea terminar el lote " & Trim(lblLote2.Caption) & " ?", vbQuestion + vbYesNo, "Cerrar lote " & Trim(lblLote2.Caption)) Then
            clsDataScan.Conexion.Execute ("EXECUTE sp_FinalizaScan " & lngIdVolumenActual & "," & lngLoteActual)
            lstImagenes.Clear
            lblLote2.Caption = ""
            
            lblTotalImagenes.Caption = ""
            lblCapacidadVolumen.Caption = ""
            lblTotalImagenes.Caption = ""
        End If
    End If

End Sub
Function ValidaLote() As Boolean
ValidaLote = True
If lblLote2.Caption = "Lote nuevo" Then
    ValidaLote = False
    MsgBox "El Lote no tiene Hoja de Control", vbCritical, Caption
    Exit Function
End If

End Function
Private Sub cmdCerrarVolumen_Click()
    If lngIdVolumenActual > 0 Then
        If vbYes = MsgBox("Desea terminar el volumen actual ?", vbQuestion + vbYesNo, "Cerrar volumen") Then
            KScan1.DeviceMethod = KSDEVICEMETHODSINGLE
            bytPaso = 4
        End If
    End If
End Sub

Private Sub cmdImportar_Click()
'frmImporImag.FormOrig = Me
If lblLote2.Caption = "" Or UCase(lblLote2.Caption) = UCase("Lote Nuevo") Then ' MsgBox "Debe crear o tener un lote activo", 64, Caption
    frmImporImag.vbindLoteCreado = False
Else
    frmImporImag.lblLote.Caption = lblLote2.Caption
    frmImporImag.vbindLoteCreado = True
End If

frmImporImag.Show 1

End Sub

Private Sub cmdInicioLote_Click()
On Local Error GoTo Err:
    If lngIdVolumenActual = 0 Or bolVolumenSeleccionado = False Then
        MsgBox "Seleccione un Volumen", vbCritical, "Lote"
    Else
            DatosVolumen
            If vbYes = MsgBox("Desea iniciar un nuevo lote para el volumen seleccionado ?", vbQuestion + vbYesNo, App.EXEName) Then
             'cambió aqui lngLoteActual = lngLoteActual + 1
                    bolTransaccion = True
                    
                    lstImagenes.Clear
                    lblTotalImagenes.Caption = ""
                    bolIniciarLote = True
                    
            End If
            lblLote2.Caption = "Lote nuevo"
    End If
    Exit Sub
Err:
   ' clsDataScan.Conexion.RollbackTrans
End Sub

Private Sub cmdNuevoLote_Click()
Dim strDirScan As String
    If lngIdVolumenActual = 0 Or bolVolumenSeleccionado = False Then
        MsgBox "Seleccione un Volumen", vbCritical, "Lote"
    Else

        strLoteActual = lblLote.Caption
        frmLotes.Show vbModal
        
        strDirScan = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "00000000")
        If Trim(Dir(strDirScan, vbDirectory)) = "" Then MkDir strDirScan
        lstImagenes.Clear
        If lblLote2.Caption <> "" Then CargarImagenes
        
           
    End If
End Sub
Private Sub cmdPararScan_Click()
    If KScan1.DeviceReserved = True Then
        KScan1.Action = KSACTIONSTOPIMMEDIATE
        KScan1.Action = KSACTIONUNRESERVE
        cmdRescan.Enabled = False
        cmdPropiedades.Enabled = False
        cmdScan.Enabled = False
        cmdPararScan.Enabled = False
    Else
        MsgBox "Seleccione un scanner", vbCritical, "Digitalizar"
    End If
End Sub

Private Sub cmdPropiedades_Click()
    If KScan1.DeviceReserved = True Then
        KScan1.ScanSource = KSSOURCEADF
        KScan1.ScanSize = KGSIZEMAX 'KSIZEA4 '
        KScan1.ScanDpi = 16
        KScan1.ActiveDialog = 2
        KScan1.Action = 7
        If KScan1.ScanDpi < 16 Then
            MsgBox "La resolución usada para la digitalización debe ser mayor igual a 240 dpi", vbCritical, "Scan"
        Else
            cmdScan.Enabled = True
            cmdRescan.Enabled = True
            cmdPararScan.Enabled = True
        End If
    Else
        MsgBox "Seleccione un scanner", vbCritical, "Digitalizar"
    End If
End Sub

Private Sub cmdRescan_Click()
    If Me.lstImagenes.ListCount > 0 Then
        If Me.lstImagenes.ListIndex > -1 Then
                bolRescan = True
                KScan1.DeviceMethod = KSDEVICEMETHODSINGLE
                KScan1.Action = KSACTIONSTART
        End If
    End If
End Sub

Private Sub cmdScan_Click()
    bolRescan = False
'    If LoteActivo = True Then
'        MsgBox "El lote actual esta terminado, genere otro lote", vbCritical, "Digitalizar"
'        Exit Sub
'    End If
    
    If KScan1.DeviceReserved = True Then
        cmdRescan.Enabled = True
        cmdCerrarVolumen.Enabled = True
        cmdCerrarLote.Enabled = True
        KView1.BackColor = vbBlack
        KView1.FitWidth = True
        KView1.FitHeight = True
        KView1.Gray = True
        KScan1.DisplayCtl1 = 0
        KScan1.DisplayCtl1 = KView1
        KScan1.Display = True
    

            Dim rsLoteActual As New Recordset
            Set rsLoteActual = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Scan,[Contador],'where idvolumen=" & lngIdVolumenActual & " and lote = " & lngLoteActual & "',''")
            If rsLoteActual.RecordCount > 0 Then lngImagenActual = rsLoteActual!contador
            If rsLoteActual.State = 1 Then rsLoteActual.Close
            Set rsLoteActual = Nothing

            strPrefijoRuta = Format(Trim(lblLote2.Caption), "0000") & "_"  '& Format(Day(Date), "00") & Format(Month(Date), "00") & Year(Date) & "_"
            If (KScan1.ActiveDevice = KSACTIVEDEVICEDISK) Then KScan1.ScanDuplex = False
            KScan1.DeviceMethod = KSDEVICEMETHODBATCH
            KScan1.Action = KSACTIONSTART
            DoEvents
       ' End If
    Else
        MsgBox "Seleccione un scanner", vbCritical, "Digitalizar"
    End If
End Sub
Private Sub EmpezarScan()
   If KScan1.ActiveDevice = KSACTIVEDEVICESCANNER Then
        If KScan1.ScanDuplex Then
        End If
        If KScan1.DeviceMethod = KSDEVICEMETHODBATCH Then
            KView1.Action = KVACTIONCLEAR
        Else
            KView1.Action = KVACTIONCLEAR
        End If
    Else
        If KScan1.DeviceMethod = KSDEVICEMETHODBATCH Then
            KView1.Action = KVACTIONCLEAR
        End If
    End If
End Sub

Private Sub cmdSeleccionScanner_Click()
On Local Error GoTo Err:
    If lngIdVolumenActual = 0 Then 'Or lngLoteActual = 0 Then
        MsgBox "Seleccione un Volumen y un Lote", vbCritical, "Scanner"
    Else
        If KScan1.DeviceReserved = False Then
            KScan1.ActiveDevice = KSACTIVEDEVICESCANNER
            KScan1.ActiveDialog = KSDIALOGSELECTDEVICE
            KScan1.Action = KSACTIONOPENDIALOG
            KScan1.Action = KSACTIONRESERVE
            cmdPropiedades.Enabled = True
        End If
    End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Scanner"
    Exit Sub
End Sub
Private Sub cmdVolumen_Click()
Dim strDirScan As String
    frmVolumen.Show vbModal
    If bolVolumenSeleccionado = True Then
      ' If bolTransaccion = True Then clsDataScan.Conexion.RollbackTrans
       bolTransaccion = False
       
       Dim rsVolumenActual As New Recordset
       Set rsVolumenActual = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Volumenes,[*],'where idvolumen=" & lngIdVolumenActual & "',''")
       If rsVolumenActual.RecordCount > 0 Then
            lblVolumen.Caption = " " & rsVolumenActual!nombre
            lblDirectorioVolumen.Caption = " " & UCase(rsVolumenActual!Ruta)
            lblTotalLote.Caption = " " & rsVolumenActual!total_lotes
            lngTotalImagenes = rsVolumenActual!total_imagenes
            lngLoteActual = IIf(IsNull(rsVolumenActual!Lote_Actual), "0", rsVolumenActual!Lote_Actual)
            lngVolumenActual = rsVolumenActual!idvolumen
            
            'If LoteActivo = False Then CargarImagenes
'            CargarImagenes
'            If lngLoteActual > 0 Then
'                lblLote.Caption = " " & rsVolumenActual!Lote_Actual
'                lblLote2.Caption = " " & rsVolumenActual!Lote_Fisico
'                cmdCerrarLote.Enabled = True
'            Else
'                lblLote.Caption = " "
'                lblLote2.Caption = " "
'            End If
            
            If ExisteDirVolumen(rsVolumenActual!Ruta) Then
                strDirectorioImagenes = rsVolumenActual!Ruta
                Me.MousePointer = 11
                TamañoDirectorio strDirectorioImagenes
                Me.MousePointer = 0
                strDirScan = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "00000000")
                If Trim(Dir(strDirScan, vbDirectory)) = "" Then MkDir strDirScan
            Else
                MsgBox "No existe el directorio " & UCase(rsVolumenActual!Ruta) & " , verifique", vbCritical, "Error"
            End If
            cmdCerrarVolumen.Enabled = True
       End If
       If rsVolumenActual.State = 1 Then rsVolumenActual.Close
       Set rsVolumenActual = Nothing
       KView1.FitWidth = True
       KView1.FitHeight = True
    End If
End Sub

Sub CargarImagenes()
    If clsDataScan.ConexionActiva = True Then
        Dim rsimagenes As New Recordset
        If rsimagenes.State = 1 Then rsimagenes.Close
        Set rsimagenes = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Imagenes,'imagen','where idVolumen=" & lngIdVolumenActual & " and lote = " & lngLoteActual & "','order by imagen'")
        lstImagenes.Clear
        If rsimagenes.RecordCount > 0 Then
            lstImagenes.Clear
            Do While Not rsimagenes.EOF
                lstImagenes.AddItem rsimagenes!Imagen
                rsimagenes.MoveNext
            Loop
            cmdCerrarLote.Enabled = True
            lstImagenes.ListIndex = lstImagenes.ListCount - 1
        End If
        lblTotalImagenes.Caption = lstImagenes.ListCount & " imagenes"
        If rsimagenes.State = 1 Then rsimagenes.Close
        Set rsimagenes = Nothing
    End If

End Sub

Public Sub InicializaKofax()
On Local Error Resume Next

    If KScan1.DeviceReserved Then
        KScan1.Action = KSACTIONUNRESERVE
    End If
             
    Me!KScan1.DeviceAlias = "(Default)"
    Me!KScan1.ActiveDevice = KSACTIVEDEVICESCANNER
    Me!KScan1.Action = KSACTIONRESET     'Este reiniciliza todo el driver utilizar en caso que se cuelgue el scanner
                
    Me!KScan1.Action = KSACTIONRESERVE
    'Seteando el alimentador del scanner ADF
                
    Me!KScan1.ScanSource = KSSOURCEADF

    If strTipoScan Like "*VRS*" Then
        KScan1.ScanSize = KGSIZEA4
    Else
        KScan1.ScanSize = KGSIZEA4
        If KScan1.ScanContinuousCap = True Then
            KScan1.ScanContinuous = True
            'PENDIENTEMsgBox KScan1.ScanContinuous
        End If
        KScan1.ScanContinuous = True
    End If
    'Seteando el scanner en simplex o duplex
    Me!KScan1.DeviceMethod = KSDEVICEMETHODSINGLE
    'KScan1.ScanWithoutSetSettings = False
    'KScan1.ScanContrast = KGCONTRAST2
    'KScan1.ScanDensity = KGCONTRASTPANEL
    'KScan1.PEVerticalDpi = 240
    'KScan1.PEHorizontalDpi = 240
    KScan1.ScanDpi = KGDPI200
    'KScan1.ScanContinuous = True
    'KScan1.Action = KSACTIONSENDCUSTOMCMD
    KScan1.Action = KSACTIONSETSETTINGS
    'Seteando el scanner a duplex
    'If Me!KScan1.ScanDuplexCap = True Then
    '    Me!KScan1.ScanDuplex = True
    'End If
    'Conectando Kview1 hacia el control KScan1
    Me!KScan1.DisplayCtl1 = Me!KView1
    'If Me!KScan1.ScanDuplex Then
    '    Me!KScan1.DisplayCtl2 = KView2
    'End If
    'Mostrando las imagenes que han sido escaneadas
    Me!KScan1.Display = True
End Sub
Private Sub Form_Load()
Dim strServerSQL As String
Dim strDataBaseSQL As String
Dim strUserSQL As String
Dim strPasswordUserSql As String
Dim strParametros() As String

'    bytPesoIndex = 0
'    bytPesoIndex = LeerIni(App.Path & "\digitacion.ini", "seteos", "pesoindex")
    
    mskReplicar.Text = 0
    bolIniciarLote = False
    
    LEAD1.UnlockSupport L_SUPPORT_TIFLZW, "gZWEhj9ZX2j"
    
    
'    If Command <> "" Then
'        strParametros = Split(Command$, ",")
'
'        strServerSQL = strParametros(0)
'        strDataBaseSQL = strParametros(1)
'        strUserSQL = strParametros(2)
'        strPasswordUserSql = strParametros(3)
        
        Set clsDataScan = New clsDataSQL
'        clsDataScan.SQLServer = strServerSQL
'        clsDataScan.SQLDataBase = strDataBaseSQL
'        clsDataScan.SQLUser = strUserSQL
'        clsDataScan.SQLUserPassword = strPasswordUserSql
'        clsDataScan.Conectar
        lngTamañoVolumen = 0
        clsDataScan.Conexion = Con.Cnx
        If clsDataScan.ConexionActiva = False Then
            MsgBox clsDataScan.MsgError, vbCritical, "Inicio de SQL Server"
            Set clsDataScan = Nothing
            Unload Me
        Else
            bolVolumenSeleccionado = False
            InicializarControles
            IniciarValores
            intMenu = 1
        End If
'    Else
'        MsgBox "Faltan los parámetros de inicio", vbCritical, "Scan"
''        Unload Me
'    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    KView1.Height = ScaleHeight - KView1.Top
    KView1.Width = ScaleWidth - KView1.Left
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If KScan1.DeviceReserved = True Then
        KScan1.Action = KSACTIONSTOPIMMEDIATE
        KScan1.DisplayCtl1 = 0
        KScan1.Action = KSACTIONUNRESERVE
    End If
'    If Not (clsDataScan Is Nothing) Then
'        If clsDataScan.ConexionActiva = True Then clsDataScan.Desconectar
'        Set clsDataScan = Nothing
'    End If
    If Dir(App.Path & "\tempo.tif", vbArchive) <> "" Then
        Kill App.Path & "\tempo.tif"
    End If
'    End
End Sub
Private Sub KScan1_Error(nError As Integer, strError As String)
    If strError = "Scanner Out Of Paper" Then
        strError = "El alimentador del scanner se encuentra disponible"
    Else
        If strError = "Scanner Paper Jam" Then
            strError = "Papel atorado en el scanner, revise el Alimentador del Equipo"
        Else
            strError = Err.Description
        End If
    End If
    MsgBox strError, vbInformation, "Scan"
    Me.KScan1.Action = KSACTIONNONE
    
End Sub
Private Sub KScan1_PageDone(nStatus As Integer, strStatus As String)
On Local Error GoTo error
Dim bolOk As Boolean
Dim strMensaje As String
Dim bolDobleLote As Boolean
 
DoEvents
bolOk = True
    
    bolDobleLote = False
    
    If bolRescan = False Then
        If bolIniciarLote = True And Me.lstImagenes.ListCount = 0 Then
            Dim strLoteFisico As String
            strLoteFisico = ""
            strLoteFisico = ObtenerLoteFisico(App.Path & "\tempo.tif")
            If strLoteFisico = "" Then
                strLoteCaptura = ""
                bolOk = False
                strMensaje = "Debe generar un numero de lote válido"
            Else
                If lblLote2.Caption <> "Lote nuevo" Then
                    If ExisteLote(strLoteFisico) = True Then
                        strMensaje = "Ya existe un lote con el mismo numero de lote a crear"
                        bolOk = False
                    End If
                End If
            End If
            
            If bolOk = False Then
                bolIniciarLote = False
                lngImagenActual = lngImagenActual - 1
               ' If bolTransaccion = True Then clsDataScan.Conexion.RollbackTrans
                bolTransaccion = False
                MsgBox strMensaje, vbCritical, App.EXEName
                cmdPararScan_Click
                
                Exit Sub
            Else
                'If bolTransaccion = True Then clsDataScan.Conexion.CommitTrans
                'clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarDatosLote " & lngIdVolumenActual & "," & lngLoteActual & ",'" & strLoteFisico & "'")
                If bolIniciarLote Then
                    lblLote2.Caption = strLoteFisico
                    strPrefijoRuta = Format(Trim(lblLote2.Caption), "000#") & "_"
                    lngImagenActual = 1
                    strNuevaImagen = strPrefijoRuta & Format(IIf(lngImagenActual = 0, 1, lngImagenActual), "0000000#")
                    strImagenScan = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "0000000#") & "\" & strNuevaImagen & "." & cExtensionImagen
                    If Dir(strDirectorioImagenes & "\" & Format(lblLote2.Caption, "0000000#"), vbDirectory) = "" Then
                        MkDir strDirectorioImagenes & "\" & Format(lblLote2.Caption, "0000000#")
                    End If
                    
                    If Dir(strImagenScan, vbArchive) <> "" Then Kill strImagenScan
                    FileCopy App.Path & "\tempo.tif", strImagenScan
                    
                    lngLoteActual = fgDatosql("Select isnull(max(Lote)+1,1) from _Scan where idvolumen = " & lngIdVolumenActual & " ")
                    
                    strCampos = lngIdVolumenActual & "," & lngLoteActual & ",'" & Trim(Format(strLoteFisico, "0000000#")) & "'"
                    clsDataScan.Conexion.Execute ("EXECUTE sp_InsertarPreScan " & strCampos)
                    clsDataScan.Conexion.Execute ("EXECUTE sp_InsertarScan " & strCampos)
                    clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTabla _Volumenes ,'set lote_actual=" & lngLoteActual & ", Lote_Fisico = ''" & Trim(Format(strLoteFisico, "0000000#")) & "''','idVolumen= " & lngIdVolumenActual & "'")
    
                    lblLote.Caption = " " & lngLoteActual
                    'lblLote2.Caption = " "
                End If
            End If
            bolIniciarLote = False
        End If
        
        'lstImagenes.AddItem strNuevaImagen
        If lngImagenActual = 1 Then clsDataScan.Conexion.Execute ("EXECUTE sp_IniciaScan " & lngIdVolumenActual & "," & lngLoteActual)
        clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTabla _Scan,'set Contador=" & lngImagenActual & "','idVolumen=" & lngIdVolumenActual & " and lote = " & lngLoteActual & "'")
        If FileLen(strImagenScan) > 0 Then
            If lstImagenes.ListIndex > 0 Then
                If Round(FileLen(strImagenScan) / 1024) < bytPesoIndex Then
                        Dim strNuevoLote As String
                        strNuevoLote = ""
                        strNuevoLote = ObtenerLoteFisico(strImagenScan)
                        If strNuevoLote <> "" Then
                           strMensaje = "Se ha detectado una hoja de control, cada lote solo debe tener una hoja de control" & Chr(13) & Chr(13) & "En caso de que el documento reconocido no sea una hoja de control por favor vuelva a escanear la última hoja, el documento que acaba de escanear no se registro"
                           bolDobleLote = True
                           lngImagenActual = lngImagenActual - 1
                           MsgBox strMensaje, vbCritical, "Scan"
                           Kill strImagenScan
                           lstImagenes_DblClick
                        End If
                End If
            End If
            If bolDobleLote = False Then
                lstImagenes.AddItem strNuevaImagen
                lblTotalImagenes.Caption = lstImagenes.ListCount & " imagenes"
                lstImagenes.ListIndex = lstImagenes.ListCount - 1
                
                If chkReplicar.Value = 0 Then
                    clsDataScan.Conexion.Execute ("EXECUTE sp_RegistrarImagen " & lngIdVolumenActual & "," & lngLoteActual & ",'" & IIf(Trim(lblLote2.Caption) <> "", Format(Trim(lblLote2.Caption), "0000000#"), Trim(lblLote2.Caption)) & "',[" & strNuevaImagen & "],'B',0")
                Else
                    If vbYes = MsgBox("Desea replicar la imagen " & mskReplicar.Text & " veces", vbQuestion + vbYesNo, "Replicar") Then
                        clsDataScan.Conexion.Execute ("EXECUTE sp_RegistrarImagen " & lngIdVolumenActual & "," & lngLoteActual & ",'" & IIf(Trim(lblLote2.Caption) <> "", Format(Trim(lblLote2.Caption), "0000000#"), Trim(lblLote2.Caption)) & "',[" & strNuevaImagen & "],'B'," & mskReplicar.Text)
                    Else
                        clsDataScan.Conexion.Execute ("EXECUTE sp_RegistrarImagen " & lngIdVolumenActual & "," & lngLoteActual & ",'" & IIf(Trim(lblLote2.Caption) <> "", Format(Trim(lblLote2.Caption), "0000000#"), Trim(lblLote2.Caption)) & "',[" & strNuevaImagen & "],'B',0")
                    End If
                End If
            End If
        End If
 
    Else
        lstImagenes_DblClick
    End If
 
    If bolDobleLote = False Then lngTamañoVolumen = lngTamañoVolumen + FileLen(strImagenScan)
    DoEvents
    Exit Sub
error:
    MsgBox Err.Description, vbCritical, "Scan"
    If bolIniciarLote = True Then
        'clsDataScan.Conexion.RollbackTrans
        bolIniciarLote = False
    End If
End Sub


Function ObtenerLoteFisico(strImagenScan As String) As String
Dim strBarCode As String
strBarCode = ""
On Local Error Resume Next
    If Dir(strImagenScan, vbArchive) <> "" Then
        LEADBarCode1.Units = 0
        LEADBarCode1.Linear.EnableErrorCheck = False
        LEADBarCode1.EnableUseRegion = 0
        LEADBarCode1.BarColor = 0
        LEADBarCode1.SpaceColor = 0
    
        LEAD1.PaintSizeMode = PAINTSIZEMODE_FIT
        LEAD1.BitonalScaling = 2
        strBarCode = ""
    
        LEADBarCode1.Bitmap = 0
        LEADBarCode1.Linear.EnableErrorCheck = 0
        LEADBarCode1.Linear.Granularity = 9
        LEADBarCode1.Linear.MinLength = 3
        LEADBarCode1.Linear.ReadDirection = 15
        LEAD1.Load strImagenScan, 0, 1, 1
    
        LEAD1.BorderRemove 0, 15, 20, 5, 5
        LEAD1.Despeckle
        LEAD1.ForceRepaint
        LEADBarCode1.Bitmap = LEAD1.Bitmap
        intRun = LEADBarCode1.Read(16, 0, 0, 0, 0, 0)
        If intRun = 0 Then
            If LEADBarCode1.TotalCount = 0 Then
                intRotar = 0
                Do While True
                    If LEADBarCode1.TotalCount = 0 Then
                        intRotar = intRotar + 90
                        LEAD1.FastRotate intRotar
                        LEADBarCode1.Linear.ReadDirection = 15
                    Else
                        Exit Do
                    End If
                    intRun = LEADBarCode1.Read(16, 0, 0, 0, 0, 0)
                    If intRotar = 360 Then Exit Do
                Loop
            End If
            If LEADBarCode1.TotalCount > 0 Then
                strBarCode = Trim(Me.LEADBarCode1.BarCodeData(0).data)
                If IsNumeric(strBarCode) And Len(strBarCode) = 7 Then
                    
                Else
                    strBarCode = ""
                End If
            End If
        End If
    End If
    LEADBarCode1.Bitmap = 0
    
    ObtenerLoteFisico = strBarCode
End Function
Private Sub KScan1_PageEnd()
On Local Error GoTo error
    DoEvents
        
       
        If bolIniciarLote = True And Me.lstImagenes.ListCount = 0 Then
            strImagenScan = App.Path & "\tempo.tif"
             lngImagenActual = lngImagenActual + 1
            If Dir(strImagenScan, vbArchive) <> "" Then
                Kill strImagenScan
            End If
        Else
            If bolRescan = False Then
            lngImagenActual = lngImagenActual + 1
            strNuevaImagen = strPrefijoRuta & Format(lngImagenActual, "00000000")
            strImagenScan = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "00000000") & "\" & strNuevaImagen & "." & cExtensionImagen
            Else
                strImagenScan = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "00000000") & "\" & lstImagenes.list(lstImagenes.ListIndex) & "." & cExtensionImagen
                lngTamañoVolumen = lngTamañoVolumen - FileLen(strImagenScan)
            End If

        End If
        
        KScan1.PEFileName = strImagenScan
    DoEvents
    Exit Sub
error:
    MsgBox Err.Description, vbCritical, "Scan"
End Sub
Private Sub KView1_DblClick()
    If Dir(KView1.FileName, vbArchive) <> "" Then
        KView1.FitHeight = True
        KView1.FitWidth = True
        KView1.Refresh
    End If
End Sub

Private Sub lstImagenes_DblClick()
On Local Error GoTo Err:

    If lstImagenes.ListCount > 0 Then
        KScan1.DisplayCtl1 = 0
        KView1.BackColor = vbBlack
        KView1.Gray = True
        KScan1.Display = False
        KView1.FileName = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "00000000") & "\" & lstImagenes.list(lstImagenes.ListIndex) & "." & cExtensionImagen
        If FileLen(KView1.FileName) > 0 Then KView1.Refresh
    End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Error"
End Sub
Sub TamañoDirectorio(strDirectorio As String)
End Sub
Function CalcularMB(lngBytes As Variant)
Dim strMB As String
    CalcularMB = "0 KB"
    If lngBytes < 1024 Then
        strMB = lngBytes & " bytes"
    Else
        If lngBytes >= 1024 And lngBytes < 1048476 Then
            strMB = Round((lngBytes / 1024), 2) & " KB"
        Else
            strMB = Round(((lngBytes / 1024) / 1024), 2) & " MB"
        End If
    End If
    clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTamañoVolumen " & lngIdVolumenActual & ",'" & strMB & "'")
    CalcularMB = strMB
End Function

Sub CerrarVolumen()
    clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTabla _Volumenes,'set Estado=1','idVolumen=" & lngIdVolumenActual & "'")
    clsDataScan.Conexion.Execute ("EXECUTE sp_CerrarVolumen " & lngIdVolumenActual)
    IniciarValores
    InicializarControles
End Sub
Sub IniciarValores()
    lngNuevaimagen = "00000000"
    lngImagenActual = 0
    lngLoteActual = 0
    lngIdVolumenActual = 0
    bolVolumenSeleccionado = False
    
    strDirectorioImagenes = ""
    strLoteActual = ""
    lblCapacidadVolumen.Caption = "0 KB"
    bolRescan = False
End Sub
Sub InicializarControles()
    lblVolumen.Caption = ""
    lblDirectorioVolumen.Caption = ""
    lblLote.Caption = ""
    cmdCerrarVolumen.Enabled = False
    cmdCerrarLote.Enabled = False
    cmdPararScan.Enabled = False
    cmdPropiedades.Enabled = False
    cmdRescan.Enabled = False
    cmdScan.Enabled = False
    lstImagenes.Clear
    lblTotalImagenes.Caption = ""
    lblCapacidadVolumen.Caption = ""
    lblTotalImagenes.Caption = ""
    cmdNuevoLote.Enabled = True
    cmdVolumen.Enabled = True
End Sub

Private Sub lstImagenes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuOpciones
    End If
End Sub

Private Sub mnuEliminar_Click()
Dim strImagenBorrar As String
    If lstImagenes.SelCount > 0 Then
        If vbYes = MsgBox("Desea eliminar los " & lstImagenes.SelCount & " elementos seleccionados ?", vbQuestion + vbYesNo, "Eliminar") Then
            Dim intIndice As Integer
            intIndice = 0
            MousePointer = 11
            Do While lstImagenes.SelCount <> 0
                lstImagenes.ListIndex = intIndice
                If lstImagenes.Selected(intIndice) = True Then
                    Debug.Print lstImagenes.list(intIndice)
                    clsDataScan.Conexion.Execute ("EXECUTE sp_EliminarImagen '" & lstImagenes.list(intIndice) & "'," & lngIdVolumenActual & "," & lngLoteActual)
                    strImagenBorrar = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "00000000") & "\" & lstImagenes.list(intIndice) & "." & cExtensionImagen
                    If Dir(strImagenBorrar, vbArchive) <> "" Then Kill strImagenBorrar
                    lstImagenes.RemoveItem (intIndice)
                Else
                    intIndice = intIndice + 1
                End If
                If intIndice > lstImagenes.ListCount - 1 Then Exit Do
            Loop
            MsgBox "Se elimino los elementos seleccionados", vbInformation, "Eliminar"
        End If
    End If
    lblTotalImagenes.Caption = lstImagenes.ListCount & " imagenes"
    If lstImagenes.ListCount = 0 Then bolIniciarLote = True
    
    MousePointer = 0
End Sub
Private Sub mskReplicar_GotFocus()
    mskReplicar.SelStart = 0
    mskReplicar.SelLength = Len(mskReplicar)
End Sub

'Function ExisteLote(strLoteFisico As String) As Boolean
'Dim rsTempo As New Recordset
'Dim bolOk As Boolean
'    bolOk = False
'    If rsTempo.State = 1 Then rsTempo.Close
'    Set rsTempo = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Scan,[lote],'where idvolumen=" & lngIdVolumenActual & " and lotefisico = " & strLoteFisico & "',''")
'    If rsTempo.RecordCount > 0 Then
'        bolOk = True
'    End If
'    If rsTempo.State = 1 Then rsTempo.Close
'    Set rsTempo = Nothing
'
'
'    ExisteLote = bolOk
'End Function

Sub DatosVolumen()
Dim rsVolumenActual As New Recordset
       Set rsVolumenActual = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Volumenes,[*],'where idvolumen=" & lngIdVolumenActual & "',''")
       If rsVolumenActual.RecordCount > 0 Then
            lblVolumen.Caption = " " & rsVolumenActual!nombre
            lblDirectorioVolumen.Caption = " " & UCase(rsVolumenActual!Ruta)
            lblTotalLote.Caption = " " & rsVolumenActual!total_lotes
            lngTotalImagenes = rsVolumenActual!total_imagenes
            lngLoteActual = IIf(IsNull(rsVolumenActual!Lote_Actual), "0", rsVolumenActual!Lote_Actual)
            lngVolumenActual = rsVolumenActual!idvolumen
            
            'If LoteActivo = False Then CargarImagenes
            CargarImagenes
            If lngLoteActual > 0 Then
                lblLote.Caption = " " & rsVolumenActual!Lote_Actual
                lblLote2.Caption = " " & rsVolumenActual!Lote_Fisico
                cmdCerrarLote.Enabled = True
            Else
                lblLote.Caption = " "
                lblLote2.Caption = " "
            End If
            
            If ExisteDirVolumen(rsVolumenActual!Ruta) Then
                strDirectorioImagenes = rsVolumenActual!Ruta
                Me.MousePointer = 11
                TamañoDirectorio strDirectorioImagenes
                Me.MousePointer = 0
                strDirScan = strDirectorioImagenes & "\" & Format(lblLote2.Caption, "00000000")
                If Trim(Dir(strDirScan, vbDirectory)) = "" Then MkDir strDirScan
            Else
                MsgBox "No existe el directorio " & UCase(rsVolumenActual!Ruta) & " , verifique", vbCritical, "Error"
            End If
            cmdCerrarVolumen.Enabled = True
       End If
       If rsVolumenActual.State = 1 Then rsVolumenActual.Close
       Set rsVolumenActual = Nothing
End Sub

Sub ImportarImagen(Pos_Ima As Integer, rutaImagen As String)
'Copiar la imagen
'Crear el registro
On Local Error GoTo ERRORES
Dim vbtImagen As String

Dim comm As Command
Set comm = New ADODB.Command

'strDirectorioImagenes
'lngVolumenActual
'lblLote

With comm
            
           '-----Insertar archivo --------
        
          
           vbtImagen = fgDatosql(" pNomImagen  " & CInt(lngVolumenActual) & ", " & CInt(lblLote) & ", '" & lblLote2 & "' ")
           
           If vbtImagen = "" Then
            MsgBox "detener"
            Exit Sub
           End If
           
                                 
           lngImagenActual = val(fgDatosql("Select contador from _scan where idvolumen = " & lngVolumenActual & " and Lote = " & lblLote & "  "))
           lngImagenActual = lngImagenActual + 1
           
           If Not CopiarFile(rutaImagen, strDirectorioImagenes & "\" & Format(Trim(lblLote2), "0000000#") & "\" & Trim(vbtImagen) & "." & cExtensionImagen) Then
                    MsgBox "Error al copiar el archivo ", vbCritical, Caption
           Else
                             
                 strImagenScan = strDirectorioImagenes & "\" & Format(Trim(lblLote2), cFormatoLote) & "\" & Trim(vbtImagen) & "." & cExtensionImagen
                                     
                 'Call InsertarImagen(CInt(Vol.idVolumen), CInt(Lot.idlote), Lot.Lote, vbtImagen, "B", val(mskReplicar.Text), lngImagenActual, TamImag, FecImag, KScan1.ScanDpi, 0, 0)
                 
                 clsDataScan.Conexion.Execute ("EXECUTE sp_RegistrarImagen " & lngVolumenActual & "," & lblLote & ",'" & IIf(Trim(lblLote2.Caption) <> "", Format(Trim(lblLote2.Caption), "0000000#"), Trim(lblLote2.Caption)) & "',[" & vbtImagen & "],'B',0")
                                                   
                 indice = lstImagenes.ListIndex + 1
                
                'Ingresa el lote fisico
                'CargarImagenes
                
                'lstImagenes.ListIndex = indice
                'lstImagenes2.ListIndex = indice
                'lngImagenActual = lngImagenActual + 1
                Call ActualizaContador(lngImagenActual)
                
                'lblTotalImagenes.Caption = lstImagenes.ListCount & " imágenes"
           
'                If Trim(vbtImagen) <> "" Then
'                   'vbtImagen
'                 '  Call GuarLog(vbgCliente, vbgProyecto, CInt(Vol.idVolumen), CInt(Lot.idlote), Trim(vbtImagen), "1", gNomUsu, "Digitalizacion", Trim(Lot.Lote))
'                 '  Call ActImagInser(vbgCliente, vbgProyecto, CInt(Vol.idVolumen), CInt(Lot.idlote))
'                Else
'                   If ExisteArchivo(Vol.RutaVolumen & "\" & Format(Trim(Lot.Lote), cFormatoLote) & "\" & Trim(vbtImagen) & "." & cExtensionImagen) Then
'                      Kill Vol.RutaVolumen & "\" & Format(Trim(Lot.Lote), cFormatoLote) & "\" & Trim(vbtImagen) & "." & cExtensionImagen
'                   End If
'                End If
            
           End If
End With
Exit Sub
ERRORES:
    With comm
        If ExisteArchivo(Vol.RutaVolumen & "\" & Format(Trim(Lot.Lote), cFormatoLote) & "\" & Trim(vbtImagen) & "." & cExtensionImagen) Then
                      Kill Vol.RutaVolumen & "\" & Format(Trim(Lot.Lote), cFormatoLote) & "\" & Trim(vbtImagen) & "." & cExtensionImagen
        End If
    End With
        MsgBox "Error en la Inserción de la Imagen  " & Err.Description, vbCritical, Caption

End Sub

Sub InsertarImagen(idvolumen As Integer, idlote As Integer, Lote As String, _
                    Imagen As String, Tipo As String, Replicar As Integer, Posicion As Long, _
                    Tamano As Variant, FecImag As String, Dpi As Integer, Alto As Integer, Ancho As Integer)
                    
On Local Error GoTo ERRORES

Dim comm As New ADODB.Command

Con.Cnx.BeginTrans


With comm

    .ActiveConnection = Con.Cnx
    .CommandText = " Update Imagenes set Pos_Ima = Pos_Ima + 1  " & _
                   " where  idProyecto = " & vbgProyecto & "  and  idcliente = " & vbgCliente & "  " & _
                   " and idVolumen = " & Vol.idvolumen & " and idLote = " & Lot.idlote & " and Pos_Ima >= " & Posicion & " "
                 
    .Execute

    .ActiveConnection = Con.Cnx
    .CommandText = " INSERT INTO Imagenes (idProyecto, idCliente,  idVolumen, idLote,   Lote, Imagen, Tipo, Pos_Ima, Tamano, FecImag, Dpi, Alto, Ancho  ) values " & _
                   " (" & vbgProyecto & ", " & vbgCliente & ", " & Vol.idvolumen & ", " & Lot.idlote & ", '" & Lote & "','" & Imagen & "','" & Tipo & "' , " & Posicion & ", " & Tamano & ", '" & FecImag & "', " & Dpi & ", " & Alto & ", " & Ancho & "  ) "
    .Execute
    

     'Call ActImagInser(vbgCliente, vbgProyecto, Vol.idVolumen, Lot.idlote)
     
End With

Con.Cnx.CommitTrans
Set comm = Nothing
Exit Sub
ERRORES:
Con.Cnx.RollbackTrans
Set comm = Nothing
MsgBox Err.Description

End Sub

Sub ActualizaContador(contador As Long)

On Local Error GoTo ERRORES

Dim comm As New ADODB.Command

'Con.Cnx.BeginTrans

With comm
    .ActiveConnection = Con.Cnx
    .CommandText = "Update _Scan set contador = " & contador & " where  idVolumen = " & lngIdVolumenActual & " and Lote= " & Me.lblLote & ""
    .Execute , , adExecuteNoRecords

End With

'Con.Cnx.CommitTrans
Set comm = Nothing
Exit Sub
ERRORES:
'Con.Cnx.RollbackTrans
Set comm = Nothing
MsgBox Err.Description
End Sub


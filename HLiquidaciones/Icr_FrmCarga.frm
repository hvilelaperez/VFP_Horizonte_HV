VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form Icr_FrmCarga 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Carga de Imagenes"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   Icon            =   "Icr_FrmCarga.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6120
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   8655
         Begin VB.ListBox lstMensajes 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1710
            ItemData        =   "Icr_FrmCarga.frx":030A
            Left            =   120
            List            =   "Icr_FrmCarga.frx":030C
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   8415
         End
         Begin ImgeditLibCtl.ImgEdit Imgvisor 
            Height          =   1215
            Left            =   360
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
            _Version        =   131076
            _ExtentX        =   3836
            _ExtentY        =   2143
            _StockProps     =   96
            BorderStyle     =   1
            ImageControl    =   "ImgEdit1"
            UndoBufferSize  =   412727040
            ForceFileLinking1x=   -1  'True
            sReserved1      =   -3656
            sReserved2      =   -3656
            lReserved1      =   2010226912
            lReserved2      =   2010226912
            bReserved1      =   -1  'True
            bReserved2      =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   4560
         Width           =   8655
         Begin VB.CommandButton cmdverificar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Procesar Archivo Excell/"
            Height          =   735
            Left            =   360
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   7935
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Generar Archivo Lst"
            Height          =   735
            Left            =   6120
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmdFirma 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Firma Virtual Imagen"
            Height          =   735
            Left            =   3360
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
      End
      Begin VB.Frame FraUbicacion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Ubicación Archivos Origén Integra"
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8655
         Begin VB.DriveListBox Drvxls 
            Height          =   315
            Left            =   480
            TabIndex        =   4
            Top             =   360
            Width           =   8055
         End
         Begin VB.DirListBox Dirxls 
            Height          =   1440
            Left            =   480
            TabIndex        =   3
            Top             =   720
            Width           =   2655
         End
         Begin VB.FileListBox flxls 
            Height          =   1455
            Left            =   3120
            TabIndex        =   2
            Top             =   720
            Width           =   5415
         End
      End
   End
End
Attribute VB_Name = "Icr_FrmCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdFirma_Click()
'''Realizar la anotacion en los archivos de imagenes colocar el numero de cargo''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Cmdverificar_Click()
''''Archivo en excell para la Carga de Datos'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Trim$(flxls.FileName) <> "" Then
        xhoja = UCase(InputBox("Ingresar el nombre de la Hoja del Archivo en excel :", "Liquidacion Horizonte"))
        RutaArchivoExcel = Dirxls.Path & "\" & flxls.FileName
        RutaDirExcel = Dirxls.Path
        Generar_Archivo
Else
        MsgBox "Debe Seleccionar el Archivo", vbExclamation, "Liquidacion Horizonte"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Public Sub Generar_Archivo()
Dim XCAMPO1 As String
Dim XCAMPO2 As String
Dim XCAMPO3 As String
Dim XCAMPO4 As String
Dim XCAMPO5 As String
Dim XCAMPO6 As String
Dim XCAMPO7 As String
Dim XCAMPO8 As String
Dim XCAMPO9 As String
Dim XCAMPO10 As String
Dim XCAMPO11 As String
Dim XCAMPO12 As String
Dim XCAMPO13 As String
Dim XCAMPO14 As String
Dim XCAMPO15 As String
Dim XCAMPO16 As String
Dim XBATCH As String
Dim XNum As Double
Dim intArch As Integer
Dim xRuta As String
Dim Xreclamo As String
Dim XNumre As Integer
Dim xMin As String
Dim xMax As String
Dim xExcellUltimo As String
Dim xxRuta As String
Dim intCantLotes As Integer
Dim intLotes As Integer
Dim IntContador As Integer
Dim StrReclamo As String


Dim RutaArchivo As String

'Dim xlotes(1 To 30)

Dim strFolderPadre As String
                        Dim strFolderActual As String
                        Dim strPeriodo As String
                        Dim strAnno As String

'Dim XNum As String
Const ALL_MARKS = 0
Const WINDOWS_DEFAULT = 2
Dim RsExcell As New Recordset
On Error GoTo Error_Handling_Generar_Archivo

With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RutaArchivoExcel & _
                ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "select distinct lote from [" & xhoja & "$]"

End With

If RsExcell.RecordCount > 0 Then
 ReDim xlotes(1 To RsExcell.RecordCount)
 intCantLotes = RsExcell.RecordCount
 For intLotes = 1 To RsExcell.RecordCount
 xlotes(intLotes) = "" & RsExcell.Fields("lote")
 RsExcell.MoveNext
 Next intLotes
End If

RsExcell.Close


With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RutaArchivoExcel & _
                ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "select * from [" & xhoja & "$]where ruta is not null"

End With

Icr_FrmCarga.MousePointer = vbHourglass
If RsExcell.RecordCount > 0 Then
  For X = 1 To RsExcell.RecordCount
    DoEvents
    'PrbBarra.Min = 1
    'PrbBarra.Max = RsExcell.RecordCount + 1
    'PrbBarra.Value = X + 1
    'XNum = Val(Txtinico) + X - 1
    XCAMPO1 = Trim$(IIf(IsNull(RsExcell.Fields("Ruta")), "", RsExcell.Fields("Ruta")))
    XCAMPO2 = Trim$(IIf(IsNull(RsExcell.Fields("nroreclamo")), "", RsExcell.Fields("nroreclamo")))
    XBATCH = Trim$(IIf(IsNull(RsExcell.Fields("LOTE")), "", RsExcell.Fields("LOTE")))
    
    'XCAMPO3 = Trim$(IIf(IsNull(RsExcell.Fields(2)), "", RsExcell.Fields(2)))
    'XCAMPO4 = Trim$(IIf(IsNull(RsExcell.Fields(3)), "", RsExcell.Fields(3)))
    'XCAMPO5 = Trim$(IIf(IsNull(RsExcell.Fields(4)), "", RsExcell.Fields(4)))
    'Grabar_TemporalEX XCAMPO1, XCAMPO2, XCAMPO3, XCAMPO4, XCAMPO5, XCAMPO6, XCAMPO7, XCAMPO8, XCAMPO9, XCAMPO10, _
    'XCAMPO11, XCAMPO12, XCAMPO13, XCAMPO14, XCAMPO15, XCAMPO16
    
    If Dir(XCAMPO1, vbArchive) <> "" Then
    Else
    MsgBox "No existe la ruta para este archivo: " & XCAMPO1, vbCritical, "Liquidacion"
    End If
    XNumre = XNumre + 1
    RsExcell.MoveNext
  Next X
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   RsExcell.MoveLast
   RsExcell.MoveFirst
  If RsExcell.RecordCount > 0 Then
     For intArch = 1 To RsExcell.RecordCount
      XCAMPO2 = Trim$(IIf(IsNull(RsExcell.Fields("ruta")), "", RsExcell.Fields("ruta")))
      xRuta = XCAMPO2
      Xreclamo = Trim$(IIf(IsNull(RsExcell.Fields("nroreclamo")), "", RsExcell.Fields("nroreclamo")))
      imgVisor.Image = xRuta
      imgVisor.Zoom = 80
      imgVisor.Display
      imgVisor.AnnotationFont = "Courier"
      imgVisor.AnnotationFont.Bold = True
      imgVisor.AnnotationFont.Size = 8
      imgVisor.AnnotationFontColor = vbRed
      imgVisor.AnnotationStampText = "Numero de Reclamo :" & Xreclamo & "-" & Format(Now, "dd/mm/yyyy")
      imgVisor.AnnotationType = wiTextStamp  '8
      imgVisor.Draw 380, 15
      imgVisor.ConvertPageType wiPageTypeBW, True
      imgVisor.BurnInAnnotations ALL_MARKS, WINDOWS_DEFAULT
      imgVisor.SaveAs xRuta, wiFileTypeTIFF, wiPageTypeBW
     
     RsExcell.MoveNext
     Next intArch
     lstMensajes.AddItem "Se procesaron la cantidad de: " & XNumre & "  Imagenes con Nro de Reclamo"
     
  End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''ARCHIVO LST
StrArchivo = strDirectorioLST & "\Lst_" & Format(Now, "YYYY") & "_General.lst"
'lstMensajes.AddItem "Se genero el archivo LST: " & StrArchivo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   RsExcell.MoveLast
   RsExcell.MoveFirst
'   xMin = "0000" 'Select_Imagen_Min(XBATCH)
'   xMax = Select_Imagen_Max(XBATCH)
If RsExcell.RecordCount > 0 Then
                intPedidosRechazados = RsExcell.RecordCount
                        
                If rsSQL.State = 1 Then rsSQL.Close
                    rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & ("" & RsExcell!LOTE) & "' order by ds_perioDO,ds_año,ds_batch"
                    rsSQL.Open
                    If rsSQL.RecordCount > 0 Then
                         strPeriodo = rsSQL!ds_periodo
                         strAnno = rsSQL!ds_año
                    End If
                StrArchivo = strDirectorioLST & "\Lst_" & Format(Now, "YYYY") & "_" & "General.lst"
                If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
                'lstMensajes.AddItem "Lst_" & ("" & strAnno) & "_" & ("" & RsExcell!LOTE) & ".lst"
                lstMensajes.AddItem "Se genero el archivo LST: " & StrArchivo
                RsExcell.MoveFirst
                lngFile = FreeFile
                Open StrArchivo For Output As #lngFile

                    Print #lngFile, "LASERFICHE IMPORT LIST"
                    Print #lngFile, "FOLDERS(RECLAMOS)"
                    Print #lngFile, "FOLDERS(" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ")"
                    strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & RsExcell!LOTE)
                    strFolderActual = strFolderPadre
                    'Print #lngFile, "FOLDERS(" & strFolderActual & ")"
                    
          For intLotes = 1 To intCantLotes
                RsExcell.Close

                With RsExcell
                .ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                "Data Source=" & RutaArchivoExcel & _
                                ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                                
                .CursorType = adOpenStatic
                .CursorLocation = adUseClient
                .Open "select * from [" & xhoja & "$]where lote= '" & xlotes(intLotes) & "' and ruta is not null"
                
                End With
          
                     xMin = "0000" 'Select_Imagen_Min(XBATCH)
                     xMax = Select_Imagen_Max(RsExcell.Fields("lote"))
                     RsExcell.MoveLast
                     xExcellUltimo = Mid(RsExcell.Fields("ruta"), Len(RsExcell!RUTA) - 7, 4)
                     RsExcell.MoveFirst
                    Do While Not RsExcell.EOF
                    xMin = xMin + 1
                      
                                    If Format(xMin, "0000") = Mid(RsExcell!RUTA, Len(RsExcell!RUTA) - 7, 4) Then
                                        strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & RsExcell!LOTE)
                                        Print #lngFile, "DOCUMENT(" & "" & RsExcell!nroreclamo & ")"
                                        StrReclamo = "" & RsExcell!nroreclamo
                                        Print #lngFile, "STARTFIELDS"
                                        Print #lngFile, ("RECLAMOS")
                                        Print #lngFile, ("")
                                        Print #lngFile, ("" & RsExcell!nroreclamo)
                                        Print #lngFile, "ENDFIELDS"
                                        Print #lngFile, "STARTLIST"
                                        Print #lngFile, ("" & RsExcell!RUTA)
                                        RutaArchivo = Mid(("" & RsExcell!RUTA), 1, Len(RsExcell!RUTA) - 25)
                                        Print #lngFile, "ENDLIST"
                                        
                                        If Mid(RsExcell.Fields("ruta"), Len(RsExcell!RUTA) - 7, 4) = xExcellUltimo Then
                                        
                                         If xMax - xExcellUltimo > 0 Then
                                             For IntContador = 1 To xMax - xExcellUltimo
                                                xExcellUltimo = xExcellUltimo + 1
                                                xxRuta = Select_Imagen_Existe("" & RsExcell!LOTE, Format(xExcellUltimo, "0000"))
                                                Print #lngFile, "DOCUMENT(" & "" & StrReclamo & ")"
                                                Print #lngFile, "STARTFIELDS"
                                                Print #lngFile, ("RECLAMOS")
                                                Print #lngFile, ("")
                                                Print #lngFile, ("" & StrReclamo)
                                                Print #lngFile, "ENDFIELDS"
                                                Print #lngFile, "STARTLIST"
                                                Print #lngFile, ("" & RutaArchivo & xxRuta)
                                                Print #lngFile, "ENDLIST"
                                             Next IntContador
                                         End If
                                        End If
                                        RsExcell.MoveNext
                                    Else
                                                xxRuta = Select_Imagen_Existe("" & RsExcell!LOTE, Format(xMin, "0000"))
                                                Print #lngFile, "DOCUMENT(" & "" & StrReclamo & ")"
                                                Print #lngFile, "STARTFIELDS"
                                                Print #lngFile, ("RECLAMOS")
                                                Print #lngFile, ("")
                                                Print #lngFile, ("" & StrReclamo)
                                                Print #lngFile, "ENDFIELDS"
                                                Print #lngFile, "STARTLIST"
                                                Print #lngFile, ("" & RutaArchivo & xxRuta)
                                                Print #lngFile, "ENDLIST"
                                    End If
                        
                    Loop
                
            Next intLotes
                
                Close #lngFile
                
MsgBox "Se Procesaron con exito las imagenes con nro de Reclamo y generado el LST ", vbInformation, "Liquidacion"

End If
If RsExcell.State = 1 Then RsExcell.Close
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Icr_FrmCarga.MousePointer = vbNormal
End If
Error_Handling_Generar_Archivo:
        If Err.Number = "-2147467259" Then
        strError = strError & Err.Number & " " & Err.Description & Chr$(13)
        MsgBox "Debe cambiar el nombre a hoja1 tiene otro nombre no valido", vbExclamation
        End If
        If Err.Number = "3265" Then
        strError = strError & Err.Number & " " & Err.Description & Chr$(13)
        MsgBox "Debe verificar la estructura del archivo excel tiene un nombre invalido " & strError & "  ", vbExclamation
        End If
        If Err.Number = "-2147217904 " Then
        strError = strError & Err.Number & " " & Err.Description & Chr$(13)
        MsgBox "Debe verificar la estructura del archivo excel tiene un nombre invalido " & strError & "  ", vbExclamation
        End If
        If Err.Number = -2147217887 Then
        MsgBox "Error de campo demasiado largo verificar archivo excell", vbCritical, ""
        End If
        If Err.Number = -2147352571 Then
        MsgBox "Error de tipo de dato numerico se esta enviando nullo", vbCritical, ""
        End If
End Sub


Private Sub cmverificar_Click()

End Sub

Private Sub Dirxls_Change()
flxls.Path = Dirxls.Path
End Sub

Private Sub Drvxls_Change()
Dirxls.Path = Drvxls.Drive
End Sub


Public Sub Generar_Lst()
If rsSQL.State = 1 Then rsSQL.Close
                    rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch"

                    rsSQL.Open
                    If rsSQL.RecordCount > 0 Then
                        intPedidosRechazados = rsSQL.RecordCount

                        StrArchivo = strDirectorioLST & "\Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
                        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo

                        lstMensajes.AddItem "Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
                        rsSQL.MoveFirst

                        lngFile = FreeFile


                        Dim strFolderPadre As String
                        Dim strFolderActual As String

                        Open StrArchivo For Output As #lngFile

                            Print #lngFile, "LASERFICHE IMPORT LIST"
                            Print #lngFile, "FOLDERS(RECLAMOS)"
                            Print #lngFile, "FOLDERS(" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ")"
                            strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
                            strFolderActual = strFolderPadre
                            'Print #lngFile, "FOLDERS(" & strFolderActual & ")"
                            Do While Not rsSQL.EOF
                                strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
                                If strFolderPadre <> strFolderActual Then
                                    Print #lngFile, "FOLDERS()"
                                    strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
                                    strFolderActual = strFolderPadre
                                    Print #lngFile, "FOLDERS(" & strFolderActual & ")"
                                End If
                                Print #lngFile, "DOCUMENT(" & "LP_" & Format(rsSQL!id_liq, "0000000000") & "_" & ("" & rsSQL!ds_nroliq) & ")"
                                Print #lngFile, "STARTFIELDS"
                                'Print #lngFile, strPlantillaLF
                                Print #lngFile, ("RECLAMOS")
                                Print #lngFile, ("")
                                Print #lngFile, ("" & rsSQL!id_liq)
                                'Print #lngFile, ("" & rsSQL!ds_nroliq)
                                'Print #lngFile, Format(Year(Date), "0000")

                                'Print #lngFile, (rsSQL!ds_periodo)

                                'Print #lngFile, (rsSQL!priapell) & " " & (rsSQL!segapell) & " " & (rsSQL!nombres)   '''APELLIDOS Y NOMBRE  80 CARACTERES
                                'Print #lngFile, ("Vigente")
                                Print #lngFile, "ENDFIELDS"

'                                Print #lngFile, "DOCUMENT(" & "LP_" & Format(rsSQL!ID_liq, "0000000000") & "_" & ("" & rsSQL!ds_nroliq) & ")"
                                Print #lngFile, "STARTLIST"
                                Print #lngFile, rsSQL!Orig_file
                                Print #lngFile, "ENDLIST"
                                rsSQL.MoveNext
                            Loop
                        Close #lngFile

                    End If
                    If rsSQL.State = 1 Then rsSQL.Close

End Sub

Private Sub Form_Load()
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
End Sub

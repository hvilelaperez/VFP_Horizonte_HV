VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Icr_Frmenvio 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Archivos de Entrega"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   Icon            =   "Icr_Frmenvio.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Caption         =   "Tipo Archivo"
         Height          =   855
         Left            =   2400
         TabIndex        =   8
         Top             =   1560
         Width           =   3015
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000009&
            Caption         =   "Excell"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000009&
            Caption         =   "Texto"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox CHKTODOS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TODOS"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   4695
      End
      Begin Crystal.CrystalReport Crtxls 
         Left            =   120
         Top             =   4080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cmbBatch 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
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
         Height          =   2760
         ItemData        =   "Icr_Frmenvio.frx":030A
         Left            =   600
         List            =   "Icr_Frmenvio.frx":030C
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4815
      End
      Begin VB.PictureBox picAvance 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   7  'Invert
         FillColor       =   &H00400000&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   600
         ScaleHeight     =   270
         ScaleWidth      =   4785
         TabIndex        =   2
         Top             =   1200
         Width           =   4815
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   1680
         Width           =   1455
      End
      Begin Crystal.CrystalReport crCargo 
         Left            =   120
         Top             =   3360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport crar1 
         Left            =   120
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport crdupli 
         Left            =   120
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label lblEnvio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nùmero de batch"
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
         Left            =   600
         TabIndex        =   5
         Top             =   285
         Width           =   1410
      End
   End
End
Attribute VB_Name = "Icr_Frmenvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSQL As New Recordset
Dim StrArchivo As String
Dim strDirectorio As String
Dim strEjecutar As String

Dim strZona As String * 4

Dim strBatch As String * 7 '4
Dim strAsesora As String * 14
Dim strCampaña As String * 2
Dim strCodigoImportacion As String * 2
Dim strCliente As String * 3
Dim strFolletoSiguiente As String * 2
Dim strFolletoPosterior As String * 2
Dim strFiller As String * 4
Dim strCL As String * 4
Dim strCantidad As String * 2
Private Sub cmdprocesar_Click()
    'VerificarProductosNull
    GenerarEntregableTXT
End Sub

Private Sub Command1_Click()
lst_otros
End Sub

Private Sub Form_Load()
On Error GoTo Err:
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    ListarBatch
Err:
 If Len(Err.Description) > 0 Then
  MsgBox "Se a encontrado el siguiente error: " & Err.Description, vbCritical, "Liquidación Horizonte"
 End If
End Sub
Sub ListarBatch()
On Error GoTo Err:
    cmbBatch.Clear
    If rsSQL.State = 1 Then rsSQL.Close
    'rsSQL.Source = "Select distinct ds_batch from afiliacion order by 1"
    rsSQL.Source = "Select distinct ds_batch from TB_liquidacion order by 1"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        rsSQL.MoveFirst
        Do While Not rsSQL.EOF
            cmbBatch.AddItem rsSQL!DS_batch
            rsSQL.MoveNext
        Loop
    End If
    If rsSQL.State = 1 Then rsSQL.Close

Err:
 If Len(Err.Description) > 0 Then
  MsgBox "Se a encontrado el siguiente error: " & Err.Description, vbCritical, "Liquidación Horizonte"
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsSQL Is Nothing) Then
        If rsSQL.State = 1 Then rsSQL.Close
        Set rsSQL = Nothing
    End If
End Sub

Private Sub mskEnvio_GotFocus()
    mskEnvio.SelStart = 0
    mskEnvio.SelLength = Len(mskEnvio.Text)
End Sub
Private Sub GeneraArchivo(StrArchivo As String, ByVal strLinea)
    Dim myFile As Integer
    myFile = FreeFile()
    Open StrArchivo For Append Access Write As myFile
    Print #myFile, strLinea
    Close #myFile
End Sub
Sub VerificarProductosNull()
On Local Error GoTo Err:
Dim intProductos As Integer
Dim intIndice As Integer

Dim strProd As String
Dim strCant As String
Dim intTotProd As Integer

Screen.MousePointer = 11

    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from tb_Parametros"
    rsSQL.Open
    intProductos = 0 & rsSQL!nu_cl
    If rsSQL.State = 1 Then rsSQL.Close

    
    If rsSQL.State = 1 Then rsSQL.Close
'    rsSQL.Source = "Select * from vw_afiliacionAxis where ds_batch='" & cmbBatch.Text & "'"
     rsSQL.Source = "Select * from vw_liquidacionentrega where ds_batch='" & cmbBatch.Text & "'"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        rsSQL.MoveFirst
        Do While Not rsSQL.EOF
            intTotProd = 0
'            For intIndice = 1 To intProductos
'                strProd = "cl" & Format(intIndice, "00")
'                strCant = "can" & Format(intIndice, "00")
'
'                If Trim("" & rsSQL.Fields(strProd)) <> "" Then
'                    If (ExisteCL("" & rsSQL.Fields(strProd)) = True) And _
'                       ((0 & rsSQL.Fields(strCant)) > 0) Then
'                        If DigitoOK(Trim("" & rsSQL.Fields(strProd))) = True Then
'                            intTotProd = intTotProd + 1
'                        End If
'                    End If
'                End If
'            Next
'
'            If intTotProd = 0 Then
'                strCadenaSQL = "EXEC sp_InsertarRechazo '" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & rsSQL!id_hojapedido & ",3,'" & strProd & "','',0" & ",'" & ("" & rsSQL!ASESORA) & "'"
'                cnServer.Execute strCadenaSQL
'            End If
            rsSQL.MoveNext
        Loop
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    Screen.MousePointer = 0
    Exit Sub
Err:
    Resume Next
End Sub

Sub GenerarEntregableTXT()
On Local Error GoTo Err:
'On Error Resume Next
Dim lngFile As Long
Dim strLinea As String
Dim intOk As Integer

Dim strErrZona As String * 4
Dim strErrIDHojaPedido As String * 8
Dim strErrAsesora As String * 14
Dim strErrRechazo As String * 30
Dim strErrProducto As String * 10
Dim strErrValor As String * 20

Dim bolZip As Boolean
Dim strTXT As String
Dim strZip As String

Dim intProductos As Integer
Dim intIndice As Integer

Dim strProd As String
Dim strCant As String
Dim intTotProd As Integer

Dim intPedidosProcesados As Integer
Dim intPedidosRechazados As Integer
Dim intPedidosValidados As Integer
Dim intRegistroDuplicados As Integer
Dim intRegistroValidos As Integer
Dim rsZonas As New Recordset

Dim rsFormulario As New Recordset
Dim rsDuplica As New Recordset
Dim StrCampo As String
Dim StrDato1 As String
Dim StrDato2 As String
Dim StrDato As String

Dim StrRucArchivo As String
Dim StrRaSocArchivo As String
Dim StrComentario As String
Dim StrCodigo As String
Dim Strcuspp As String
Dim strPeriodo As String
Dim StrFechainicio As String
Dim StrFechaCese As String
'JN 16/04/2013
Dim strPath As String

Dim StrNrolircargo As String
Dim StrLoteCargo As String
Dim StrPagCargo As String
Dim StrRutaCargo As String
Dim StrArchivotif As String
Dim StrRemuneracion As String
Dim StrRequerimiento As String
Dim StrObservacion As String
Dim StrArchivo_Texto As String

Dim StrTramiteR As String

Dim Z As Integer

Dim StrDireccion As String
Dim StrEmail As String
Dim StrTelefono As String
Dim StrDocumento As String
Dim StrNom_Ape As String
Dim StrCambio_RS As String

Dim Strorden As String




Dim StrFirma_Bo As Boolean
Dim StrFecha_Bo As Boolean
Dim Strperiodo_HO As String

Dim StrRazonAd As String
Dim StrRucAd As String
Dim StrPeriodoAd As String

    rsZonas.CursorType = adOpenDynamic
    rsZonas.LockType = adLockBatchOptimistic
    rsZonas.ActiveConnection = cnServer
    
    rsFormulario.CursorType = adOpenDynamic
    rsFormulario.LockType = adLockBatchOptimistic
    rsFormulario.ActiveConnection = cnServer
    
    rsDuplica.CursorType = adOpenDynamic
    rsDuplica.LockType = adLockBatchOptimistic
    rsDuplica.ActiveConnection = cnServer
    
    Screen.MousePointer = 11
    
    lstMensajes.Clear
    
    If Trim(cmbBatch.Text) = "" Then
        MsgBox "Seleccione un número de batch", vbCritical, "Generaciòn de archivos de retorno"
       Screen.MousePointer = 0
        Exit Sub
    End If
    
    'cmbbatch.Text = Padln(Trim(mskEnvio.Text), 4, "0")
    
    cmdProcesar.Enabled = False
    cmbBatch.Enabled = False
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "delete from tb_Entrega"
    rsSQL.Open
    
    
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from tb_Parametros"
    rsSQL.Open
    
    strCampañaActual = "" & rsSQL!ds_periodo
    strZonaActual = "" & rsSQL!ds_año
    strDirectorio = "" & rsSQL!ds_ftpaxis
    intProductos = 0 & rsSQL!nu_cl
    
    
    Dim strsqlp As String
    Dim HojaExcel As String
    'ExcelFile = App.Path & "\Libro1.xls"
    'HojaExcel = "[Hoja1]"
    'cnServer.Execute "SELECT * INTO [Excel 8.0;DATABASE=" & ExcelFile & "]." & _
    '    HojaExcel & " FROM tempo"
    
'    strsqlp = "Select * from Tempo"
'    If Exportar_ADO_Excel(strsqlp, "c:\libro1.xLS") Then
'       MsgBox "Ok", vbInformation
'    End If

    
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "delete from tb_Entrega_H where ds_batch='" & cmbBatch.Text & "'"
    rsSQL.Open
    
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    
    lstMensajes.AddItem "DIRECTORIO DE ENVIO"
    lstMensajes.AddItem UCase(strDirectorio)
    lstMensajes.AddItem ""
    lstMensajes.AddItem "ARCHIVOS GENERADOS"
    
'*------------------------------------------------------------------------------------------------*
    strZona = "0000"
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from vw_liquidacionentrega where ds_batch='" & cmbBatch.Text & "' ORDER BY csid"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
         
        Borrar_Tempo_Historico_Sql
        strZona = "" & rsSQL!ds_año
        intPedidosValidados = rsSQL.RecordCount
        
        'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".txt"
        'If Dir(strArchivo, vbArchive) <> "" Then Kill strArchivo
        'lstMensajes.AddItem "Pedidos_" & ("" & rsSQL!ds_zona) & "_" & cmbBatch.Text & ".txt"
        'lstMensajes.AddItem "Descargo_Masivo_Empleadores_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".txt"
        
        rsSQL.MoveFirst
        
        strBatch = ""
        strAsesora = ""
        strCampaña = ""
        strCodigoImportacion = "01"
        strCliente = "000"
        strFolletoSiguiente = ""
        strFolletoPosterior = ""
        strFiller = "0000"
        strCL = ""
        strCantidad = ""
        'lngFile = FreeFile
        'Open strArchivo For Output As #lngFile
        
        'Print #lngFile, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"
                        
            Do While Not rsSQL.EOF
                ReDim xcuspp(1 To intProductos) As String
                ReDim Xcodigo(1 To intProductos) As String
                ReDim xfecha1(1 To intProductos) As String
                ReDim xfecha2(1 To intProductos) As String
                ReDim xcomen(1 To intProductos) As String
                ReDim xFere(1 To intProductos) As String
                ReDim xmonto(1 To intProductos) As String
                ReDim xPlani(1 To intProductos) As String
                ReDim XRechazo(1 To intProductos) As String
                ReDim CAMPO(1 To intProductos) As String
                ReDim xobs(1 To intProductos) As String
                                      
                      StrFirma = False
                      If rsFormulario.State = 1 Then rsFormulario.Close
                         
                          If StrBarra_Chek = "0" Then
                          StrDato1 = "HO" & Trim$(Mid(rsSQL.Fields("ds_barra"), 1, 9))
                          StrDato2 = Format(Trim$(Mid(rsSQL.Fields("ds_barra"), 10, 5)), "00000")
                          Else
                          StrDato1 = Trim$(rsSQL.Fields("ds_nroliq"))
                          StrDato2 = Format(Trim$(rsSQL.Fields("ds_pag")), , "00000")
                          End If
                                
                      StrDato = StrDato1 & StrDato2
                      rsFormulario.Source = "select * from tb_formularios where nroliqpag='" & StrDato & "'"
                      rsFormulario.Open
                      
                If rsFormulario.RecordCount > 0 Then
                        StrFirma_Bo = Verificar_Firma("" & rsSQL.Fields("ds_dni"), "" & rsSQL.Fields("ds_nombre"))
                        Strperiodo_HO = Verificar_Periodo("" & rsFormulario!Periodo)
                        For intIndice = 1 To intProductos
                                  
                                    StrCampo = "cuspp" & intIndice
                                    strProducto = Trim(("" & rsFormulario.Fields(StrCampo)))
                                    If strProducto = "" Then
                                      Exit For
                                    Else
                                        xcuspp(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        StrCampo = "ds_cod" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        Xcodigo(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        StrCampo = "ds_fec_ini" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xfecha1(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        xfecha1(intIndice) = Formatear_Fecha(xfecha1(intIndice))
                                        
                                        
                                        StrCampo = "ds_fec_ce" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xfecha2(intIndice) = (IIf(IsNull(strProducto), "", strProducto))
                                        xfecha2(intIndice) = Formatear_Fecha(xfecha2(intIndice))
                                        
                                                
                                        StrCampo = "ds_rem_as" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xmonto(intIndice) = Format(IIf(IsNull(strProducto), "", strProducto & "00"), "000000000")
                                        
                                        StrCampo = "ds_pla" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xPlani(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        
                                        If Xcodigo(intIndice) = "" Then
                                        Xcodigo(intIndice) = Codigo_Nulo(xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice))
                                        End If
                                        XRechazo(intIndice) = ""
                                        'xFere(intIndice) = Mostrar_Fecha_Ref(Xcodigo(intIndice), xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice), xPlani(intIndice))
                                        xcomen(intIndice) = Mostrar_Comentario(Xcodigo(intIndice), xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice), xPlani(intIndice))
                                        XRechazo(intIndice) = Mostrar_Rechazo(Xcodigo(intIndice), xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice), xPlani(intIndice))
                                        'XRechazo(intIndice) = Validar_Fecha(Xfecha1(intIndice))
                                        'XRechazo(intIndice) = Validar_Fecha(Xfecha2(intIndice))
                                        'StrFecha_Bo = Verificar_Fecha(xfecha1(intIndice))
'                                        If Xcodigo(intIndice) = "99" And xcomen(intIndice) = "EMPLEADOR CONFIRMA DEUDA" Then
'                                        Xcodigo(intIndice) = "77"
'                                        End If
'                                            If Xcodigo(intIndice) = "99" Then
'                                            MsgBox "PRUEBA"
'                                            End If

                                            If StrFirma_Bo = False Then
                                              StrDupli = "4"
                                              XRechazo(intIndice) = "LP SIN FIRMA"
                                            End If
                                             
                                            If XRechazo(intIndice) <> "" Then
                                              StrDupli = "4"
                                            Else
                                                XRechazo(intIndice) = Validar_Fecha(xfecha1(intIndice), xfecha2(intIndice))
                                                'XRechazo(intIndice) = Validar_Fecha(Xfecha2(intIndice))
                                                If XRechazo(intIndice) <> "" Then
                                                StrDupli = "4"
                                                End If
                                            End If
                                        
                                        Xcodigo(intIndice) = Verificar_No_Existe(Xcodigo(intIndice))
                                        'xfecha2(intIndice) = Formatear_Fecha(IIf(IsNull(strProducto), "", strProducto))
                                        'StrFecha_Bo = Verificar_Fecha(xfecha1(intIndice))
                                        'If Xcodigo(intIndice) <> "" Or xfecha1(intIndice) <> "" Or xfecha2(intIndice) <> "" Or xmonto(intIndice) <> "" Then
                                        
                                        If Xcodigo(intIndice) = "" Or Xcodigo(intIndice) = "50" Or Xcodigo(intIndice) = "60" Or Xcodigo(intIndice) = "70" Or Xcodigo(intIndice) = "80" Or Xcodigo(intIndice) = "90" Then
                                        xobs(intIndice) = Validar_Observacion(Xcodigo(intIndice))
                                        StrDupli = "4"
                                        End If
                                        ''''''
                                        
                                        If Xcodigo(intIndice) = "06" Or Xcodigo(intIndice) = "09" Or Xcodigo(intIndice) = "10" Or Xcodigo(intIndice) = "11" Or Xcodigo(intIndice) = "12" Then
                                        'xobs(intIndice) = Validar_Observacion(Xcodigo(intIndice))
                                        StrDupli = "3"
                                        End If
                                        '''''
                                        
                                        StrTramiteR = ("" & rsSQL!ds_numerotra)
                                        
                                        If ("" & rsSQL!ds_numerotra) = "" Then
                                        StrTramiteR = "0000000"
                                        StrDupli = "4"
                                        End If
                                        
                                        
                                        If xfecha1(intIndice) <> "" Or xfecha2(intIndice) <> "" Or Val(xmonto(intIndice)) <> 0 Or xPlani(intIndice) <> "" Or xcomen(intIndice) <> "" Or xobs(intIndice) <> "" Then
                                        strCadenaSQL = "EXEC sp_InsertarEntrega '" & StrDato & "','" & cmbBatch.Text & "','" & Replace(("" & Mid(rsFormulario!Ruc, 1, 11)), "'", "`") & "','" & ("" & Replace(rsFormulario!razon_social, "'", "`")) & "','" & xcomen(intIndice) & "','" & Xcodigo(intIndice) & "','" & xcuspp(intIndice) & "','" & ("" & Strperiodo_HO) & "','" & xfecha1(intIndice) & "','" & xfecha2(intIndice) & "','" & Replace(StrTramiteR, "'", "`") & "','" & xmonto(intIndice) & "','" & XRechazo(intIndice) & "','" & xobs(intIndice) & "','" & StrDupli & "','" & ("" & rsSQL.Fields("csid")) & "'"
                                        cnServer.Execute strCadenaSQL
                                        
                                        strCadenaSQL = "EXEC sp_InsertarEntrega_H '" & StrDato & "','" & cmbBatch.Text & "','" & Replace(("" & Mid(rsFormulario!Ruc, 1, 11)), "'", "`") & "','" & ("" & Replace(rsFormulario!razon_social, "'", "`")) & "','" & xcomen(intIndice) & "','" & Xcodigo(intIndice) & "','" & xcuspp(intIndice) & "','" & ("" & Strperiodo_HO) & "','" & xfecha1(intIndice) & "','" & xfecha2(intIndice) & "','" & Replace(StrTramiteR, "'", "`") & "','" & xmonto(intIndice) & "','" & XRechazo(intIndice) & "','" & xobs(intIndice) & "','" & StrDupli & "','" & ("" & rsSQL.Fields("csid")) & "'"
                                        cnServer.Execute strCadenaSQL
                                        
'                                        If StrFirma_Bo = False Then
'                                        strEjecutar = "Update tb_ENTREGA set ds_estado='4',ds_rechazo='LP SIN FIRMA' where id_entrega = " & rsSQL.Fields("id_entrega") & ""
'                                        cnServer.Execute strEjecutar
'                                        End If
'                                        If XOBS(intIndice) <> "" Then
'                                         strEjecutar = "Update tb_ENTREGA set ds_estado='4',ds_rechazo='" & XOBS(intIndice) & "' where id_entrega = " & rsSQL.Fields("id_entrega") & ""
'                                         cnServer.Execute strEjecutar
'                                        End If
                                        'If StrFecha_Bo = True Then
                                        'strEjecutar = "Update tb_ENTREGA set ds_estado='4' where id_entrega = " & rsSQL.Fields("id_entrega") & ""
                                        'cnServer.Execute strEjecutar
                                        'End If
                                        End If
                                    End If
                         
                        Next intIndice
                End If
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                rsSQL.MoveNext
            Loop
        Close #lngFile
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    If rsFormulario.State = 1 Then rsFormulario.Close
              
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''Archivo 1 Entrega Horizonte1''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If rsSQL.State = 1 Then rsSQL.Close
If rsFormulario.State = 1 Then rsFormulario.Close
rsSQL.Source = "Select * from vw_entregaHorizonte1 order by id_entrega"
rsSQL.Open
Borrar_Tempo_Sql
If rsSQL.RecordCount > 0 Then
     'Existe_Temporal
     'Create_TemporalDBF
     
     'Borrar_tempo
     '
     'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     StrArchivo = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_MOVPERDI" & "_" & cmbBatch & ".XLS"
     StrArchivo_Texto = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_MOVPERDI" & "_" & cmbBatch & ".TXT"
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\F" & Format(Now, "YYMMDD") & "_MOVPERDI" & cmbBatch & ".XLS"
     
     If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        lstMensajes.AddItem "\F" & Format(Now, "YYMMDD") & "_MOVPERDI" & cmbBatch & ".TXT"
     
     lngFile = FreeFile
     Open StrArchivo_Texto For Output As #lngFile
        
          Print #lngFile, "CUENTA AFILIADO"; vbTab; "CUSPP AFILIADO"; vbTab; "RUC EMPLEADOR"; vbTab; "TIPO MOV LABORAL"; vbTab; "FECHA INICIO MOV"; vbTab; "FECHA FIN MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO DEVENGE"; vbTab; "TIPO TRAMO LABORAL"; vbTab; "NRO REQUERIMIENTO"; vbTab; "FECHA"
     
     Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL!ds_ruc)
              StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              StrComentario = ("" & rsSQL!ds_comentario)
              StrCodigo = ("" & rsSQL!ds_cod_descargo)
              Strcuspp = Replace(("" & rsSQL!ds_cuspp), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL!ds_periodo)
              StrFechainicio = ("" & rsSQL!ds_fecha_inicio)
              StrFechaCese = ("" & rsSQL!ds_fecha_cese)
              StrArchivotif = ("" & rsSQL!ds_archivo)
              StrRequerimiento = ("" & rsSQL!ds_nrorequerimiento)
              'JN 16/04/2013
              strPath = "C:\HORIZONTE\IMAGENES\" & Mid(StrArchivotif, 1, 6) & "-" & Mid(StrArchivotif, 7, 4) & "-" & Mid(StrArchivotif, 11, 7) & "\" & StrArchivotif 'JN 16/04/2013
              If StrCodigo = "01" Then
              StrFechaCese = "99991231"
              End If
              
              If StrCodigo = "02" Then
              StrFechainicio = "00000000"
              End If
              
              If ("" & rsSQL!ds_estado) = "1" Then
                       
                       'xxGrabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, Trim$("CESE DE RELACION LABORAL"), Trim$("02"), Strcuspp, strPeriodo, Trim$("" & rsSQL!ds_fecha_cese)
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & Trim$("CESE DE RELACION LABORAL") & "','" & Trim$("02") & "','" & Strcuspp & "','" & strPeriodo & "','" & Trim$("" & rsSQL!ds_fecha_cese) & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$("CESE DE RELACION LABORAL"); vbTab; Trim$("02"); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$("" & rsSQL!ds_fecha_cese)
              Else
               
               
               
               'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
               strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & Format(StrRequerimiento, "0000000000000") & "','" & StrArchivotif & "'"
               cnServer.Execute strCadenaSQL
               
               'JN 16/04/2013
               strCadenaSQL = "EXEC sp_Insertartempo_envio_historico '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & Format(StrRequerimiento, "0000000000000") & "','" & "" & "','" & Mid(("" & rsSQL!nroliqpag), 1, 11) & "','" & Mid(("" & rsSQL!nroliqpag), 12, 5) & "','" & "1" & "','" & "" & "','" & "" & "','" & "" & "','" & "F" & Format(Now, "YYMMDD") & "_MOVPERDI" & ".XLS" & "','" & StrArchivotif & "','" & strPath & "'"
               cnServer.Execute strCadenaSQL
               
               'nrorequerimiento , observacion, nroliquidacion, PAGINA, registro, cambio_rs_ruc, direccion, email, Tipoarchivo, archivo
               
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               Print #lngFile, ""; vbTab; Strcuspp; vbTab; StrRucArchivo; vbTab; StrCodigo; vbTab; StrFechainicio; vbTab; StrFechaCese; vbTab; ""; vbTab; strPeriodo; vbTab; "N"; vbTab; Format(StrRequerimiento, "000000000000"); vbTab; Format(Now, "yyyymmdd")
              End If
              
              strLinea = ""
          
          
     rsSQL.MoveNext
     Loop
     Close #lngFile
     
        'strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        strsqlp = "Select cuenta as [CUENTA AFILIADO],CUSPP AS [CUSPP AFILIADO],RUC AS [RUC EMPLEADOR],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA] , ARCHIVO from Tempo order by id_tempo"
        

'        strCadenaSQL = "drop table [" & "A1" & cmbBatch & "]"
'        cnServer.Execute strCadenaSQL
'        strCadenaSQL = "Select cuenta as [CUENTA AFILIADO],CUSPP AS [CUSPP AFILIADO],RUC AS [RUC EMPLEADOR],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA] , ARCHIVO  into " & "A1" & cmbBatch & " from Tempo order by id_tempo"
'        cnServer.Execute strCadenaSQL
        If Option1(0).Value = True Then
           If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
           End If
           If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        Else
            
        End If
        
     
         If rsSQL.State = 1 Then rsSQL.Close
         rsSQL.Source = "Select * from Tempo"
         rsSQL.Open
         intRegistroValidos = 0
         If rsSQL.RecordCount > 1 Then
         intRegistroValidos = rsSQL.RecordCount
         End If
     
'      Crtxls.DataFiles(0) = App.Path & "\TEMPO.DBF"
'      Crtxls.ReportFileName = App.Path & "\renvio.rpt"
'      Crtxls.Destination = crptToFile
'      Crtxls.PrintFileName = strArchivo
'      Crtxls.PrintFileType = crptExcel50
'      Crtxls.Action = 1
     
     
End If
        strCadenaSQL = "drop table [" & "A1" & cmbBatch & "]"
        cnServer.Execute strCadenaSQL
        strCadenaSQL = "Select cuenta as [CUENTA AFILIADO],CUSPP AS [CUSPP AFILIADO],RUC AS [RUC EMPLEADOR],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA] , ARCHIVO  into " & "A1" & cmbBatch & " from Tempo order by id_tempo"
        cnServer.Execute strCadenaSQL





''''''''''''''''Archivo 2  Entrega Horizonte 2'''''''''''''''''''''''''''''''''''''
If rsSQL.State = 1 Then rsSQL.Close
If rsFormulario.State = 1 Then rsFormulario.Close
rsSQL.Source = "Select * from vw_entregaHorizonte2 order by id_entrega"
rsSQL.Open
     Borrar_Tempo_Sql
If rsSQL.RecordCount > 0 Then
     'Existe_Temporal
     'Create_TemporalDBF
     
     'Borrar_tempo
     'Borrar_Tempo_Sql
     'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     StrArchivo = strDirectorio & "\SINCERA_" & Format(Now, "YYMMDD") & "_" & cmbBatch & ".XLS"
     StrArchivo_Texto = strDirectorio & "\SINCERA_" & Format(Now, "YYMMDD") & "_" & cmbBatch & ".TXT"
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\SINCERA_" & Format(Now, "YYMMDD") & cmbBatch & ".XLS"
     
     If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        lstMensajes.AddItem "\SINCERA_" & Format(Now, "YYMMDD") & cmbBatch & ".TXT"
     
     lngFile = FreeFile
     Open StrArchivo_Texto For Output As #lngFile
        
          Print #lngFile, "RUC EMPLEADOR"; vbTab; "CUSPP AFILIADO"; vbTab; "PERIODO DEVENGE"; vbTab; "CUENTA AFILIADO "; vbTab; "REMUNERACION"; vbTab; "MOTIVO SINCERAMIENTO"
     
     Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL!ds_ruc)
              StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              StrComentario = ("" & rsSQL!ds_comentario)
              StrCodigo = ("" & rsSQL!ds_cod_descargo)
              Strcuspp = Replace(("" & rsSQL!ds_cuspp), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL!ds_periodo)
              StrFechainicio = ("" & rsSQL!ds_fecha_inicio)
              StrFechaCese = ("" & rsSQL!ds_fecha_cese)
              StrArchivotif = ("" & rsSQL!ds_archivo)
              StrRemuneracion = ("" & rsSQL!ds_remu)
              'JN 16/04/2013
              strPath = "C:\HORIZONTE\IMAGENES\" & Mid(StrArchivotif, 1, 6) & "-" & Mid(StrArchivotif, 7, 4) & "-" & Mid(StrArchivotif, 11, 7) & "\" & StrArchivotif
              
              If ("" & rsSQL!ds_estado) = "1" Then
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, Trim$("CESE DE RELACION LABORAL"), Trim$("02"), Strcuspp, strPeriodo, Trim$("" & rsSQL!ds_fecha_cese)
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & Trim$("CESE DE RELACION LABORAL") & "','" & Trim$("02") & "','" & Strcuspp & "','" & strPeriodo & "','" & Trim$("" & rsSQL!ds_fecha_cese) & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$("CESE DE RELACION LABORAL"); vbTab; Trim$("02"); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$("" & rsSQL!ds_fecha_cese)
              Else
               
               
               'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
               strCadenaSQL = "EXEC sp_Insertartempo_envio2 '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & "" & "','" & StrRemuneracion & "','" & StrArchivotif & "'"
               cnServer.Execute strCadenaSQL
               
               'JN 16/04/2013
               strCadenaSQL = "EXEC sp_Insertartempo_envio_historico '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & "" & "','" & "" & "','" & Mid(("" & rsSQL!nroliqpag), 1, 11) & "','" & Mid(("" & rsSQL!nroliqpag), 12, 5) & "','" & "1" & "','" & "" & "','" & "" & "','" & "" & "','" & "SINCERA_" & Format(Now, "YYMMDD") & ".XLS" & "','" & StrArchivotif & "','" & strPath & "'"
               cnServer.Execute strCadenaSQL
               
               
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               Print #lngFile, StrRucArchivo; vbTab; Strcuspp; vbTab; strPeriodo; vbTab; ""; vbTab; StrRemuneracion; vbTab; "Declaración Jurada"
              End If
              
              strLinea = ""
          
          
     rsSQL.MoveNext
     Loop
     Close #lngFile
     
        'strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        strsqlp = "Select RUC AS [RUC EMPLEADOR],CUSPP AS [CUSPP AFILIADO],PERIODO AS [PERIODO DEVENGE],cuenta as [CUENTA AFILIADO],REMU AS [REMUNERACION], '10' as [MOTIVO SINCERAMIENTO], ARCHIVO from Tempo order by id_tempo"
        
'        strCadenaSQL = "drop table [" & "A2" & cmbBatch & "]"
'        cnServer.Execute strCadenaSQL
'        strCadenaSQL = "Select RUC AS [RUC EMPLEADOR],CUSPP AS [CUSPP AFILIADO],PERIODO AS [PERIODO DEVENGE],cuenta as [CUENTA AFILIADO],REMU AS [REMUNERACION], '10' as [MOTIVO SINCERAMIENTO], ARCHIVO INTO " & "A2" & cmbBatch & " from Tempo order by id_tempo"
'        cnServer.Execute strCadenaSQL
                
        If Option1(0).Value = True Then
           If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
           End If
           If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        Else
            
        End If
        
     
         If rsSQL.State = 1 Then rsSQL.Close
         rsSQL.Source = "Select * from Tempo"
         rsSQL.Open
         intRegistroValidos = 0
         If rsSQL.RecordCount > 1 Then
         intRegistroValidos = rsSQL.RecordCount
         End If
End If

        strCadenaSQL = "drop table [" & "A2" & cmbBatch & "]"
        cnServer.Execute strCadenaSQL
        strCadenaSQL = "Select RUC AS [RUC EMPLEADOR],CUSPP AS [CUSPP AFILIADO],PERIODO AS [PERIODO DEVENGE],cuenta as [CUENTA AFILIADO],REMU AS [REMUNERACION], '10' as [MOTIVO SINCERAMIENTO], ARCHIVO INTO " & "A2" & cmbBatch & " from Tempo order by id_tempo"
        cnServer.Execute strCadenaSQL


''''''''''''''''Archivo 3  Entrega Horizonte 3'''''''''''''''''''''''''''''''''''''
If rsSQL.State = 1 Then rsSQL.Close
If rsFormulario.State = 1 Then rsFormulario.Close
'strCadenaSQL = "drop table [" & "A3" & cmbBatch & "]"
'cnServer.Execute strCadenaSQL
rsSQL.Source = "Select * from vw_entregaHorizonte3 order by id_entrega"
rsSQL.Open
Borrar_Tempo_Sql
If rsSQL.RecordCount > 0 Then
     'Existe_Temporal
     'Create_TemporalDBF
     
     'Borrar_tempo
     Borrar_Tempo_Sql
     'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     StrArchivo = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_OBSERVADOS" & "_" & cmbBatch & ".XLS"
     StrArchivo_Texto = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_OBSERVADOS" & "_" & cmbBatch & ".TXT"
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\F" & Format(Now, "YYMMDD") & "_OBSERVADOS" & cmbBatch & ".XLS"
     
     If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        lstMensajes.AddItem "\F" & Format(Now, "YYMMDD") & "_OBSERVADOS" & cmbBatch & ".TXT"
     
     lngFile = FreeFile
     Open StrArchivo_Texto For Output As #lngFile
        
          Print #lngFile, "CUSPP"; vbTab; "RUC"; vbTab; "TIPO MOV LABORAL"; vbTab; "FECHA INICIO MOV"; vbTab; "FECHA FIN MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO DEVENGE"; vbTab; "TIPO TRAMO LABORAL"; vbTab; "NRO REQUERIMIENTO"; vbTab; "FECHA"; vbTab; "OBSERVACIONES"
     
     Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL!ds_ruc)
              StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              StrComentario = ("" & rsSQL!ds_comentario)
              StrCodigo = ("" & rsSQL!ds_cod_descargo)
              Strcuspp = Replace(("" & rsSQL!ds_cuspp), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL!ds_periodo)
              StrFechainicio = ("" & rsSQL!ds_fecha_inicio)
              StrFechaCese = ("" & rsSQL!ds_fecha_cese)
              StrArchivotif = ("" & rsSQL!ds_archivo)
              StrRequerimiento = ("" & rsSQL!ds_nrorequerimiento)
              StrObservacion = ("" & rsSQL!ds_obs)
              'JN 16/04/2013
              strPath = "C:\HORIZONTE\IMAGENES\" & Mid(StrArchivotif, 1, 6) & "-" & Mid(StrArchivotif, 7, 4) & "-" & Mid(StrArchivotif, 11, 7) & "\" & StrArchivotif
              
              If ("" & rsSQL!ds_estado) = "1" Then
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, Trim$("CESE DE RELACION LABORAL"), Trim$("02"), Strcuspp, strPeriodo, Trim$("" & rsSQL!ds_fecha_cese)
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & Trim$("CESE DE RELACION LABORAL") & "','" & Trim$("02") & "','" & Strcuspp & "','" & strPeriodo & "','" & Trim$("" & rsSQL!ds_fecha_cese) & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$("CESE DE RELACION LABORAL"); vbTab; Trim$("02"); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$("" & rsSQL!ds_fecha_cese)
              Else
               
               'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
               strCadenaSQL = "EXEC sp_Insertartempo_envio3 '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & Format(StrRequerimiento, "0000000000000") & "',' & "" & ','" & StrObservacion & "','" & StrArchivotif & "'"
               cnServer.Execute strCadenaSQL
               
               'JN 16/04/2013
               strCadenaSQL = "EXEC sp_Insertartempo_envio_historico '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & Format(StrRequerimiento, "0000000000000") & "','" & StrObservacion & "','" & Mid(("" & rsSQL!nroliqpag), 1, 11) & "','" & Mid(("" & rsSQL!nroliqpag), 12, 5) & "','" & "1" & "','" & "" & "','" & "" & "','" & "" & "','" & "F" & Format(Now, "YYMMDD") & "_OBSERVADOS" & ".XLS" & "','" & StrArchivotif & "','" & strPath & "'"
               cnServer.Execute strCadenaSQL
               
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               Print #lngFile, Strcuspp; vbTab; StrRucArchivo; vbTab; StrCodigo; vbTab; StrFechainicio; vbTab; StrFechaCese; vbTab; ""; vbTab; strPeriodo; vbTab; "N"; vbTab; Format(StrRequerimiento, "000000000000"); vbTab; Format(Now, "yyyymmdd"); vbTab; StrObservacion
              End If
              
              strLinea = ""
          
          
     rsSQL.MoveNext
     Loop
     Close #lngFile
     
        'strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        strsqlp = "Select CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], observacion as [OBSERVACIONES],ARCHIVO from Tempo order by id_tempo"
        
'        strCadenaSQL = "drop table [" & "A3" & cmbBatch & "]"
'        cnServer.Execute strCadenaSQL
'        strCadenaSQL = "Select CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], observacion as [OBSERVACIONES],ARCHIVO INTO " & "A3" & cmbBatch & " FROM Tempo order by id_tempo"
'        cnServer.Execute strCadenaSQL
                
        
        If Option1(0).Value = True Then
           If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
           End If
           If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        Else
            
        End If
        
        
     
         If rsSQL.State = 1 Then rsSQL.Close
         rsSQL.Source = "Select * from Tempo"
         rsSQL.Open
         intRegistroValidos = 0
         If rsSQL.RecordCount > 1 Then
         intRegistroValidos = rsSQL.RecordCount
         End If
     
End If

        strCadenaSQL = "drop table [" & "A3" & cmbBatch & "]"
        cnServer.Execute strCadenaSQL
        strCadenaSQL = "Select CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], observacion as [OBSERVACIONES],ARCHIVO INTO " & "A3" & cmbBatch & " FROM Tempo order by id_tempo"
        cnServer.Execute strCadenaSQL


''''''''''''''''Archivo 4  Entrega Horizonte 4'''''''''''''''''''''''''''''''''''''
If rsSQL.State = 1 Then rsSQL.Close
If rsFormulario.State = 1 Then rsFormulario.Close
' strCadenaSQL = "drop table [" & "A4" & cmbBatch & "]"
' cnServer.Execute strCadenaSQL
rsSQL.Source = "Select * from vw_entregaHorizonte4 order by id_entrega"
rsSQL.Open
     Borrar_Tempo_Sql
If rsSQL.RecordCount > 0 Then
     
     'Borrar_tempo
     'Borrar_Tempo_Sql
     'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     StrArchivo = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_RECHAZOS" & "_" & cmbBatch & ".XLS"
     StrArchivo_Texto = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_RECHAZOS" & "_" & cmbBatch & ".TXT"
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\F" & Format(Now, "YYMMDD") & "_RECHAZOS" & cmbBatch & ".XLS"
     
     If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        lstMensajes.AddItem "\F" & Format(Now, "YYMMDD") & "_RECHAZOS" & cmbBatch & ".TXT"
     
     lngFile = FreeFile
     Open StrArchivo_Texto For Output As #lngFile
        
          Print #lngFile, "CUENTA"; vbTab; "CUSPP"; vbTab; "RUC"; vbTab; "TIPO MOV LABORAL"; vbTab; "FECHA INICIO MOV"; vbTab; "FECHA FIN MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO DEVENGE"; vbTab; "TIPO TRAMO LABORAL"; vbTab; "NRO REQUERIMIENTO"; vbTab; "FECHA"; vbTab; "MOTIVO RECHAZOS"
     
     Do While Not rsSQL.EOF
              StrRucArchivo = Mid(("" & rsSQL!ds_ruc), 1, 11)
              StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              StrComentario = ("" & rsSQL!ds_comentario)
              StrCodigo = ("" & rsSQL!ds_cod_descargo)
              Strcuspp = Replace(("" & rsSQL!ds_cuspp), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL!ds_periodo)
              StrFechainicio = ("" & rsSQL!ds_fecha_inicio)
              StrFechaCese = ("" & rsSQL!ds_fecha_cese)
              StrArchivotif = ("" & rsSQL!ds_archivo)
              StrRequerimiento = ("" & rsSQL!ds_nrorequerimiento)
              'JN 16/04/2013
              strPath = "C:\HORIZONTE\IMAGENES\" & Mid(StrArchivotif, 1, 6) & "-" & Mid(StrArchivotif, 7, 4) & "-" & Mid(StrArchivotif, 11, 7) & "\" & StrArchivotif
              
              If StrCodigo = "50" Or StrCodigo = "60" Or StrCodigo = "70" Or StrCodigo = "80" Or StrCodigo = "90" Then
              StrObservacion = ("" & rsSQL!ds_obs)
              Else
              StrObservacion = ("" & rsSQL!ds_rechazo)
              End If
              
              
              If ("" & rsSQL!ds_estado) = "1" Then
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrObservacion & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, Trim$("CESE DE RELACION LABORAL"), Trim$("02"), Strcuspp, strPeriodo, Trim$("" & rsSQL!ds_fecha_cese)
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & Trim$("CESE DE RELACION LABORAL") & "','" & Trim$("02") & "','" & Strcuspp & "','" & strPeriodo & "','" & Trim$("" & rsSQL!ds_fecha_cese) & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$("CESE DE RELACION LABORAL"); vbTab; Trim$("02"); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$("" & rsSQL!ds_fecha_cese)
              Else
               
               'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
               strCadenaSQL = "EXEC sp_Insertartempo_envio3 '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & Format(StrRequerimiento, "0000000000000") & "','" & "" & "','" & StrObservacion & "','" & StrArchivotif & "'"
               cnServer.Execute strCadenaSQL
               
               'JN 16/04/2013
               strCadenaSQL = "EXEC sp_Insertartempo_envio_historico '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechainicio & "','" & StrFechaCese & "','N','" & Format(Now, "YYYYMMDD") & "','" & Format(StrRequerimiento, "0000000000000") & "','" & StrObservacion & "','" & Mid(("" & rsSQL!nroliqpag), 1, 11) & "','" & Mid(("" & rsSQL!nroliqpag), 12, 5) & "','" & "1" & "','" & "" & "','" & "" & "','" & "" & "','" & "F" & Format(Now, "YYMMDD") & "_RECHAZOS" & ".XLS" & "','" & StrArchivotif & "','" & strPath & "'"
               cnServer.Execute strCadenaSQL
               
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               Print #lngFile, ""; vbTab; Strcuspp; vbTab; StrRucArchivo; vbTab; StrCodigo; vbTab; StrFechainicio; vbTab; StrFechaCese; vbTab; ""; vbTab; strPeriodo; vbTab; "N"; vbTab; Format(StrRequerimiento, "000000000000"); vbTab; Format(Now, "yyyymmdd"); vbTab; StrObservacion
              End If
              
              strLinea = ""
          
          
     rsSQL.MoveNext
     Loop
     Close #lngFile
     
        'strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        strsqlp = "Select cuenta as [CUENTA], CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], observacion as [MOTIVO RECHAZOS],ARCHIVO from Tempo order by id_tempo"
        
'        strCadenaSQL = "drop table [" & "A4" & cmbBatch & "]"
'        cnServer.Execute strCadenaSQL
'        strCadenaSQL = "Select cuenta as [CUENTA], CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], observacion as [MOTIVO RECHAZOS],ARCHIVO INTO " & "A4" & cmbBatch & " from Tempo order by id_tempo"
'        cnServer.Execute strCadenaSQL
        
                
        If Option1(0).Value = True Then
           If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
           End If
           If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        Else
            
        End If
        
     
         If rsSQL.State = 1 Then rsSQL.Close
         rsSQL.Source = "Select * from Tempo"
         rsSQL.Open
         intRegistroValidos = 0
         If rsSQL.RecordCount > 1 Then
         intRegistroValidos = rsSQL.RecordCount
         End If
     
End If

        strCadenaSQL = "drop table [" & "A4" & cmbBatch & "]"
        cnServer.Execute strCadenaSQL
        strCadenaSQL = "Select cuenta as [CUENTA], CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], observacion as [MOTIVO RECHAZOS],ARCHIVO INTO " & "A4" & cmbBatch & " from Tempo order by id_tempo"
        cnServer.Execute strCadenaSQL


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Archivo 5  Entrega Horizonte 5'''''''''''''''''''''''''''''''''''''

If rsSQL.State = 1 Then rsSQL.Close
If rsFormulario.State = 1 Then rsFormulario.Close
'rsSQL.Source = "Select * from TB_FORMULARIOS A,TB_LIQUIDACION B WHERE A.nroliqpag=" & "HO" & "order by id_entrega"
' strCadenaSQL = "drop table [" & "A5" & cmbBatch & "]"
' cnServer.Execute strCadenaSQL
 
rsSQL.Source = "select A.RUC,a.razon_social,a.nroliqpag,A.DIRECCION,B.DS_EMAIL,B.DS_TELEFONO,B.DS_DNI,B.DS_NOMBRE,B.DS_RAZON_RUC,B.CSID  AS DS_ARCHIVO from tb_formularios A, TB_LIQUIDACION B where A.BARRA = B.DS_BARRA AND B.DS_BATCH='" & cmbBatch.Text & "' AND (LEN(B.DS_RAZON_RUC)> 0 OR LEN(B.DS_EMAIL)> 0 OR LEN(B.DS_TELEFONO)> 0) ORDER BY B.DS_ARCHIVO"
rsSQL.Open
    Borrar_Tempo_Horizonte5_Sql
If rsSQL.RecordCount > 0 Then
     
     'Borrar_tempo
     Borrar_Tempo_Horizonte5_Sql
     'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     StrArchivo = strDirectorio & "\DATOS_GENERALES_" & Format(Now, "YYMMDD") & "_" & cmbBatch & ".XLS"
     StrArchivo_Texto = strDirectorio & "\DATOS_GENERALES_" & Format(Now, "YYMMDD") & "_" & cmbBatch & ".TXT"
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\DATOS_GENERALES_" & Format(Now, "YYMMDD") & cmbBatch & ".XLS"
     
     If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        lstMensajes.AddItem "\DATOS_GENERALES_" & Format(Now, "YYMMDD") & cmbBatch & ".TXT"
     
     lngFile = FreeFile
     Open StrArchivo_Texto For Output As #lngFile
        
          Print #lngFile, "RUC"; vbTab; "DIRECCION"; vbTab; "E-MAIL"; vbTab; "TELEFONO"; vbTab; "DOC_IDENTIDAD"; vbTab; "NOMBRE_APELLIDO"; vbTab; "FECHA"; vbTab; "CAMBIO RSOCIAL"; vbTab; "ENVIO"
     
     Do While Not rsSQL.EOF
              StrRucArchivo = Trim$(("" & rsSQL!Ruc))
              StrDireccion = Trim$(("" & rsSQL!DIRECCION))
              StrEmail = Trim$(("" & rsSQL!ds_email))
              StrTelefono = Trim$(("" & rsSQL!ds_telefono))
              StrDocumento = Trim$(("" & rsSQL!ds_dni))
              'StrNom_Ape = Replace(Trim$(("" & rsSQL!ds_nombre)), Chr(13), "")
              StrNom_Ape = Replace(Replace("" & rsSQL!ds_nombre, Chr(10), ""), Chr(13), "")
              StrArchivotif = ("" & rsSQL!ds_archivo)
              StrCambio_RS = Trim$(("" & rsSQL!ds_razon_ruc))
              'JN 16/04/2013
              strPath = "C:\HORIZONTE\IMAGENES\" & Mid(StrArchivotif, 1, 6) & "-" & Mid(StrArchivotif, 7, 4) & "-" & Mid(StrArchivotif, 11, 7) & "\" & StrArchivotif
              
              'StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
'              If ("" & rsSQL!ds_estado) = "1" Then
'
'                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
'                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrObservacion & "','" & StrArchivotif & "'"
'                       cnServer.Execute strCadenaSQL
'
'                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, Trim$("CESE DE RELACION LABORAL"), Trim$("02"), Strcuspp, strPeriodo, Trim$("" & rsSQL!ds_fecha_cese)
'                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & Trim$("CESE DE RELACION LABORAL") & "','" & Trim$("02") & "','" & Strcuspp & "','" & strPeriodo & "','" & Trim$("" & rsSQL!ds_fecha_cese) & "','" & StrArchivotif & "'"
'                       cnServer.Execute strCadenaSQL
'              Else
               
               'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
               strCadenaSQL = "EXEC sp_Insertartempo_envio4 '" & Mid(StrRucArchivo, 1, 11) & "','" & StrDireccion & "','" & StrEmail & "','" & StrTelefono & "','" & StrDocumento & "','" & StrNom_Ape & "','" & StrCambio_RS & "','" & StrArchivotif & "'"
               cnServer.Execute strCadenaSQL
               
               'JN 16/04/2013
               strCadenaSQL = "EXEC sp_Insertartempo_envio_historico '" & Mid(StrRucArchivo, 1, 11) & "','" & ("" & rsSQL!razon_social) & "','" & "" & "','" & "" & "','" & "" & "','" & "" & "','" & "" & "','" & "" & "','N','" & Format(Now, "YYYYMMDD") & "','" & "" & "','" & "" & "','" & Mid(("" & rsSQL!nroliqpag), 1, 11) & "','" & Mid(("" & rsSQL!nroliqpag), 12, 5) & "','" & "1" & "','" & StrCambio_RS & "','" & StrDireccion & "','" & StrEmail & "','" & "DATOS_GENERALES_" & Format(Now, "YYMMDD") & ".XLS" & "','" & StrArchivotif & "','" & strPath & "'"
               cnServer.Execute strCadenaSQL
               
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               Print #lngFile, StrRucArchivo; vbTab; StrDireccion; vbTab; StrEmail; vbTab; StrTelefono; vbTab; StrDocumento; vbTab; StrNom_Ape; vbTab; Format(Now, "YYMMDD"); vbTab; StrCambio_RS; vbTab; ""
'              End If
              
              strLinea = ""
          
          
     rsSQL.MoveNext
     Loop
     Close #lngFile
     
        'strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        strsqlp = "Select DS_ruc as [RUC], ds_direccion AS [DIRECCION],ds_email AS [E-MAIL],ds_telefono AS [TELEFONO],ds_dni AS [DOC_IDENTIDAD], DS_NOM_APE AS [NOMBRE APELLIDO]," & Format(Now, "YYYYMMDD") & " AS [FECHA],DS_RAZON_RUC AS [CAMBIO RSOCIAL],'' AS [ENVIO],DS_ARCHIVO AS ARCHIVO  from Tempo_cab order by id_tempo"
        
'        strCadenaSQL = "drop table [" & "A5" & cmbBatch & "]"
'        cnServer.Execute strCadenaSQL
'        strCadenaSQL = "Select DS_ruc as [RUC], ds_direccion AS [DIRECCION],ds_email AS [E-MAIL],ds_telefono AS [TELEFONO],ds_dni AS [DOC_IDENTIDAD], DS_NOM_APE AS [NOMBRE APELLIDO]," & Format(Now, "YYYYMMDD") & " AS [FECHA],DS_RAZON_RUC AS [CAMBIO RSOCIAL],' ' AS [ENVIO],DS_ARCHIVO AS ARCHIVO INTO " & "A5" & cmbBatch & " from Tempo_cab order by id_tempo"
'        cnServer.Execute strCadenaSQL
        
        
        
        
        If Option1(0).Value = True Then
           If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
           End If
           If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        Else
        End If
     
         If rsSQL.State = 1 Then rsSQL.Close
         rsSQL.Source = "Select * from Tempo_CAB"
         rsSQL.Open
         intRegistroValidos = 0
         If rsSQL.RecordCount > 1 Then
         intRegistroValidos = rsSQL.RecordCount
         End If
     
End If

        strCadenaSQL = "drop table [" & "A5" & cmbBatch & "]"
        cnServer.Execute strCadenaSQL
        strCadenaSQL = "Select DS_ruc as [RUC], ds_direccion AS [DIRECCION],ds_email AS [E-MAIL],ds_telefono AS [TELEFONO],ds_dni AS [DOC_IDENTIDAD], DS_NOM_APE AS [NOMBRE APELLIDO]," & Format(Now, "YYYYMMDD") & " AS [FECHA],DS_RAZON_RUC AS [CAMBIO RSOCIAL],' ' AS [ENVIO],DS_ARCHIVO AS ARCHIVO INTO " & "A5" & cmbBatch & " from Tempo_cab order by id_tempo"
        cnServer.Execute strCadenaSQL



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Archivo 6  Entrega Horizonte 6'''''''''''''''''''''''''''''''''''''

If rsSQL.State = 1 Then rsSQL.Close
If rsFormulario.State = 1 Then rsFormulario.Close
'rsSQL.Source = "Select * from TB_FORMULARIOS A,TB_LIQUIDACION B WHERE A.nroliqpag=" & "HO" & "order by id_entrega"
'strCadenaSQL = "drop table [" & "A6" & cmbBatch & "]"
'cnServer.Execute strCadenaSQL

rsSQL.Source = "select * from tempo_historico"
rsSQL.Open
Borrar_Tempo_Horizonte5_Sql
If rsSQL.RecordCount > 0 Then
     
     'Borrar_tempo
'     Borrar_Tempo_Horizonte5_Sql
     'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     StrArchivo = strDirectorio & "\CONSOLIDADO_" & Format(Now, "YYMMDD") & "_" & cmbBatch & ".XLS"
     StrArchivo_Texto = strDirectorio & "\CONSOLIDADO_" & Format(Now, "YYMMDD") & "_" & cmbBatch & ".TXT"
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\CONSOLIDADO_" & Format(Now, "YYMMDD") & cmbBatch & ".XLS"
     
     If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        lstMensajes.AddItem "\CONSOLIDADO_" & Format(Now, "YYMMDD") & cmbBatch & ".TXT"
     
     lngFile = FreeFile
     Open StrArchivo_Texto For Output As #lngFile
        
        'JN 16/04/2013
          Print #lngFile, "CUENTA"; vbTab; "CUSPP"; vbTab; "RUC"; vbTab; "TIPO_MOV_LABORAL"; vbTab; "FECHA_INICIO_MOV"; vbTab; "FECHA_FIN_MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO_DEVENGE"; vbTab; "TIPO_TRAMO_LABORAL"; vbTab; "NRO_REQUERIMIENTO"; vbTab; "FECHA"; vbTab; "NRO_LP"; vbTab; "PAGINA"; vbTab; "REGISTROS"; vbTab; "CAMBIO RSOCIAL"; vbTab; "OBSERVACIONES"; vbTab; "ENVIO"; vbTab; "DIRECCION"; vbTab; "E-MAIL"; vbTab; "FECHA_EMISION"; vbTab; "ARCHIVO_DESTINO"; vbTab; "RUTA_IMAGEN"; vbTab; "RAZON_SOCIAL"
     
     Do While Not rsSQL.EOF
             Dim strCuenta As String
             Dim strCUSPPX As String
             Dim strRucx As String
             Dim strTipoLab As String
             Dim strFechaI As String
             Dim strFechaF As String
             Dim strCorrelativox As String
             Dim StrPeriodox As String
             Dim strTipotramo As String
             Dim strReqx As String
             Dim strFechax As String
             Dim strNrlq As String
             Dim strPagx As String
             Dim strRegx As String
             Dim strRazonx As String
             Dim strObservacionesx As String
             Dim strEnviox As String
             Dim strDireccionx As String
             Dim strEmailx As String
             Dim strFechaEmisionx As String
             Dim strTipoArchivox As String
             Dim strPath1 As String 'JN 16/04/2013
             Dim strRazonSoc1 As String 'JN 16/04/2013
              
              strCuenta = Trim$(("" & rsSQL!CUENTA))
              strCUSPPX = Trim$(("" & rsSQL!cuspp))
              strRucx = Trim$(("" & rsSQL!Ruc))
              strTipoLab = Trim$(("" & rsSQL!cod_descargo))
              strFechaI = Trim$(("" & rsSQL!FECHA_INICIO))
              strFechaF = Trim$(("" & rsSQL!FECHA_CESE))
              strCorrelativox = Trim$(("" & rsSQL!correlativo))
              StrPeriodox = Trim$(("" & rsSQL!Periodo))
              strTipotramo = Trim$(("" & rsSQL!tipotramo))
              strReqx = Trim$(("" & rsSQL!NROREQUERIMIENTO))
              strFechax = Trim$(("" & rsSQL!fecha_proceso))
              strNrlq = Trim$(("" & rsSQL!NROLIQUIDACION))
              strPagx = Trim$(("" & rsSQL!PAGINA))
              strRegx = Trim$(("" & rsSQL!registro))
              strRazonx = Trim$(("" & rsSQL!CAMBIO_RS_RUC))
              strObservacionesx = Trim$(("" & rsSQL!observacion))
              strEnviox = ""
              strDireccionx = Trim$(("" & rsSQL!DIRECCION))
              strEmailx = Trim$(("" & rsSQL!EMAIL))
              strFechaEmisionx = Format(Now, "YYMMDD")
              strTipoArchivox = Trim$(("" & rsSQL!tipoarchivo))
              'JN 16/04/2013
              strPath1 = Trim$(("" & rsSQL!RUTA))
              strRazonSoc1 = Trim$(("" & rsSQL!RASON_SOC))
            
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               'JN 16/04/2013
               Print #lngFile, strCuenta; vbTab; strCUSPPX; vbTab; strRucx; vbTab; strTipoLab; vbTab; strFechaI; vbTab; strFechaF; vbTab; strCorrelativox; vbTab; StrPeriodox; vbTab; strTipotramo; vbTab; strReqx; vbTab; strFechax; vbTab; strNrlq; vbTab; strPagx; vbTab; strRegx; vbTab; strRazonx; vbTab; strObservacionesx; vbTab; strEnviox; vbTab; strDireccionx; vbTab; strEmailx; vbTab; strFechaEmisionx; vbTab; strTipoArchivox; vbTab; strPath1; vbTab; strRazonSoc1
'              End If

              strLinea = ""


     rsSQL.MoveNext
     Loop
     Close #lngFile
     
        'strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        'strsqlp = "Select CUENTA AS [CUENTA],DS_ruc as [RUC], ds_direccion AS [DIRECCION],ds_email AS [E-MAIL],ds_dni AS [DOC_IDENTIDAD], DS_NOM_APE AS [NOMBRE APELLIDO]," & Format(Now, "YYMMDD") & " AS [FECHA],DS_RAZON_RUC AS [CAMBIO RSOCIAL],'' AS [ENVIO]  from Tempo_cab order by id_tempo"
        'JN 16/04/2013
        strsqlp = "Select cuenta as [CUENTA], CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], NROLIQUIDACION as [NRO LP], PAGINA as [PAGINA],REGISTRO as [REGISTROS],CAMBIO_RS_RUC as [CAMBIO RSOCIAL],OBSERVACION as [OBSERVACIONES],'' as [ENVIO],DIRECCION,EMAIL AS [E-MAIL]," & Format(Now, "YYMMDD") & " AS [FECHA EMISION],TIPOARCHIVO AS [ARCHIVO DE DESTINO],RUTA AS [RUTA_IMAGEN]',RASON_SOC AS [RAZON_SOCIAL] from Tempo_HISTORICO order by id_tempo"
        
'        strCadenaSQL = "drop table [" & "A6" & cmbBatch & "]"
'        cnServer.Execute strCadenaSQL
'        strCadenaSQL = "Select cuenta as [CUENTA], CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], NROLIQUIDACION as [NRO LP], PAGINA as [PAGINA],REGISTRO as [REGISTROS],CAMBIO_RS_RUC as [CAMBIO RSOCIAL],OBSERVACION as [OBSERVACIONES],' ' as [ENVIO],DIRECCION,EMAIL AS [E-MAIL]," & Format(Now, "YYMMDD") & " AS [FECHA EMISION],TIPOARCHIVO AS [ARCHIVO DE DESTINO] INTO " & "A6" & cmbBatch & " from Tempo_HISTORICO order by id_tempo"
'        cnServer.Execute strCadenaSQL
        
               
        If Option1(0).Value = True Then
           If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
           End If
           If Dir(StrArchivo_Texto, vbArchive) <> "" Then Kill StrArchivo_Texto
        Else
        End If
     
         If rsSQL.State = 1 Then rsSQL.Close
         rsSQL.Source = "Select * from Tempo_Historico"
         rsSQL.Open
         intRegistroValidos = 0
         If rsSQL.RecordCount > 1 Then
         intRegistroValidos = rsSQL.RecordCount
         End If
     
End If

        strCadenaSQL = "drop table [" & "A6" & cmbBatch & "]"
        cnServer.Execute strCadenaSQL
        'JN 16/04/2013
        strCadenaSQL = "Select cuenta as [CUENTA], CUSPP AS [CUSPP],RUC AS [RUC],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA], NROLIQUIDACION as [NRO LP], PAGINA as [PAGINA],REGISTRO as [REGISTROS],CAMBIO_RS_RUC as [CAMBIO RSOCIAL],OBSERVACION as [OBSERVACIONES],' ' as [ENVIO],DIRECCION,EMAIL AS [E-MAIL]," & Format(Now, "YYMMDD") & " AS [FECHA EMISION],TIPOARCHIVO AS [ARCHIVO DE DESTINO],RUTA AS [RUTA_IMAGEN],RASON_SOC AS [RAZON_SOCIAL] INTO " & "A6" & cmbBatch & " from Tempo_HISTORICO order by id_tempo"
        cnServer.Execute strCadenaSQL



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Generar Archivo Duplicados--Archivo 2'''''''''''''''''''''''''''''

'                                    If rsSQL.State = 1 Then rsSQL.Close
'                                    If rsDuplica.State = 1 Then rsDuplica.Close
'                                    If rsFormulario.State = 1 Then rsFormulario.Close
'                                    rsSQL.Source = "Select * from vw_entregaarchivo2"
'                                    rsSQL.Open
'                                     Borrar_Tempo_Sql
'                                     StrArchivo = strDirectorio & "\Archivo2_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'                                     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
'                                     lstMensajes.AddItem "Archivo2_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'                                    If rsSQL.RecordCount > 0 Then
'                                         Do While Not rsSQL.EOF
'                                         '''''''''''''''''''''''''''''''''''''''''''
'                                         '''''''''''''''''''''''''''''''''''''''''
'                                              If rsDuplica.State = 1 Then rsDuplica.Close
'                                              rsDuplica.Source = "select * from tb_entrega where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_ruc='" & ("" & rsSQL!ds_ruc) & "' ORDER BY ID_ENTREGA"
'                                              rsDuplica.Open
'                                              If rsDuplica.RecordCount > 0 Then
'                                               For Z = 1 To rsDuplica.RecordCount
'                                                    If Z <> 1 Then
'                                                    strEjecutar = "Update tb_ENTREGA set ds_estadodu='1' where id_entrega = " & rsDuplica.Fields("id_entrega") & ""
'                                                    cnServer.Execute strEjecutar
'                                                    End If
'                                               rsDuplica.MoveNext
'                                               Next Z
'                                              End If
'                                         '''''''''''''''''''''''''''''''''''''''''''''''
'                                         ''''''''''''''''''''''''''''''''''''''''''''''
'                                              If rsFormulario.State = 1 Then rsFormulario.Close
'                                              rsFormulario.Source = "select ds_ruc, ds_raz_soc,ds_comentario, ds_cod_descargo, ds_cuspp , ds_fecha_refe, ds_periodo, ds_estado,ds_estadodu ,id_entrega,ds_archivo from tb_entrega where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_ruc='" & ("" & rsSQL!ds_ruc) & "' order by id_entrega"
'                                              rsFormulario.Open
'                                              If rsFormulario.RecordCount > 0 Then
'                                                Do While Not rsFormulario.EOF
'                                                 If ("" & rsFormulario!ds_estadodu) <> "0" Then
'                                                    StrRucArchivo = ("" & rsFormulario!ds_ruc)
'                                                    StrRaSocArchivo = Replace(("" & rsFormulario!ds_raz_soc), "Ð", "Ñ")
'                                                    StrComentario = ("" & rsFormulario!ds_comentario)
'                                                    StrCodigo = ("" & rsFormulario!ds_cod_descargo)
'                                                    Strcuspp = Replace(("" & rsFormulario!ds_cuspp), "Ð", "Ñ")
'                                                    strPeriodo = ("" & rsFormulario!ds_periodo)
'                                                    StrFechaRef = ("" & rsFormulario!ds_fecha_refe)
'                                                    StrArchivotif = ("" & rsFormulario!ds_archivo)
'
'                                                    strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
'                                                    cnServer.Execute strCadenaSQL
'                                                    strLinea = ""
'                                                  End If
'                                                rsFormulario.MoveNext
'                                                Loop
'                                              End If
'                                          rsSQL.MoveNext
'                                         Loop
'                                     End If
'                                          ''''''''''''''''''''''''''''''
'                                          '''''Duplicados historico'''''''''''''''''''''''''
'                                             If rsSQL.State = 1 Then rsSQL.Close
'                                             rsSQL.Source = "Select * from TB_ENTREGA ORDER BY ID_ENTREGA"
'                                             rsSQL.Open
'                                             If rsSQL.RecordCount > 0 Then
'
'                                              Do While Not rsSQL.EOF
'                                                  If rsDuplica.State = 1 Then rsDuplica.Close
'                                                    'rsFormulario.Source = "select * from tb_entrega where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_periodo='" & ("" & rsSQL!ds_periodo) & "'"
'                                                    rsDuplica.Source = "select ds_ruc, ds_raz_soc,ds_comentario, ds_cod_descargo, ds_cuspp , ds_fecha_refe, ds_periodo,ds_archivo from tb_entrega_H where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_ruc='" & ("" & rsSQL!ds_ruc) & "'  AND (DS_BATCH <> '" & cmbBatch.Text & "'  and DS_BATCH < '" & cmbBatch.Text & "')"
'                                                    rsDuplica.Open
'                                                    If rsDuplica.RecordCount > 0 Then
'                                                            strEjecutar = "Update tb_ENTREGA set ds_estadodu='1' where id_entrega = " & rsSQL.Fields("id_entrega") & ""
'                                                            cnServer.Execute strEjecutar
'                                                            For Z = 1 To rsDuplica.RecordCount
'                                                               StrRucArchivo = ("" & rsDuplica!ds_ruc)
'                                                               StrRaSocArchivo = Replace(("" & rsDuplica!ds_raz_soc), "Ð", "Ñ")
'                                                               StrComentario = ("" & rsDuplica!ds_comentario)
'                                                               StrCodigo = ("" & rsDuplica!ds_cod_descargo)
'                                                               Strcuspp = Replace(("" & rsDuplica!ds_cuspp), "Ð", "Ñ")
'                                                               strPeriodo = ("" & rsDuplica!ds_periodo)
'                                                               StrFechaRef = ("" & rsDuplica!ds_fecha_refe)
'                                                               StrArchivotif = ("" & rsDuplica!ds_archivo)
'
'                                                               strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
'                                                               cnServer.Execute strCadenaSQL
'                                                               strLinea = ""
'                                                            rsDuplica.MoveNext
'                                                            Next Z
'                                                    Else
'                                                    End If
'                                              rsSQL.MoveNext
'                                              Loop
'                                             End If
'                                          '''''''''''''''''''''''''''''
'                                             If rsSQL.State = 1 Then rsSQL.Close
'                                             rsSQL.Source = "Select * from Tempo"
'                                             rsSQL.Open
'                                             intRegistroDuplicados = 0
'                                             If rsSQL.RecordCount > 1 Then
'                                             intRegistroDuplicados = rsSQL.RecordCount
'                                             End If
'                                          '''''''''''''''''''''''''''''
'
'                                    strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER from Tempo order by id_tempo"
'                                    If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
'                                    End If
'
'
'                                    ''''                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    ''''                            ''''''''''''''''''''''''''''''''''''''ARCHIVO1 CODIGOS QUE NO SE CARGAN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    ''''                            If rsSQL.State = 1 Then rsSQL.Close
'                                    ''''                            If rsFormulario.State = 1 Then rsFormulario.Close
'                                    ''''                            rsSQL.Source = "Select * from vw_entregaarchivo1 order by id_entrega"
'                                    ''''                            rsSQL.Open
'                                    ''''                            If rsSQL.RecordCount > 0 Then
'                                    ''''                                 'Existe_Temporal
'                                    ''''                                 'Create_TemporalDBF
'                                    ''''                                 Borrar_tempo
'                                    ''''                                 strArchivo = strDirectorio & "\Archivo1_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'                                    ''''
'                                    ''''                                 If Dir(strArchivo, vbArchive) <> "" Then Kill strArchivo
'                                    ''''                                    lstMensajes.AddItem "\Archivo1_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'                                    ''''                                 'lngFile = FreeFile
'                                    ''''
'                                    ''''                                 'Open strArchivo For Output As #lngFile
'                                    ''''
'                                    ''''                                 'Print #lngFile, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"
'                                    ''''                                 Do While Not rsSQL.EOF
'                                    ''''                                          StrRucArchivo = ("" & rsSQL!ds_ruc)
'                                    ''''                                          StrRaSocArchivo = ("" & rsSQL!ds_raz_soc)
'                                    ''''                                          StrComentario = ("" & rsSQL!ds_comentario)
'                                    ''''                                          StrCodigo = ("" & rsSQL!ds_cod_descargo)
'                                    ''''                                          Strcuspp = ("" & rsSQL!ds_cuspp)
'                                    ''''                                          strPeriodo = ("" & rsSQL!ds_periodo)
'                                    ''''                                          StrFechaRef = ("" & rsSQL!ds_fecha_refe)
'                                    ''''
'                                    ''''                                          Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
'                                    ''''                                          'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
'                                    ''''                                          'strLinea = ""
'                                    ''''
'                                    ''''
'                                    ''''                                 rsSQL.MoveNext
'                                    ''''                                 Loop
'                                    ''''                                 'Close #lngFile
'                                    ''''                                  crar1.DataFiles(0) = App.Path & "\TEMPO.DBF"
'                                    ''''                                  crar1.ReportFileName = App.Path & "\renvio.rpt"
'                                    ''''                                  crar1.Destination = crptToFile
'                                    ''''                                  'Crtxls.PrintFileName = Dirxls & "\" & "CPT_" & Xfecha & "_" & XposIni + 1 & "_" & XposFin & "_" & flxls.FileName
'                                    ''''                                  'Xarchivo = Dirxls & "\" & "CPT_" & Xfecha & "_" & XposIni + 1 & "_" & XposFin & "_" & flxls.FileName
'                                    ''''                                  crar1.PrintFileName = strArchivo
'                                    ''''                                  crar1.PrintFileType = crptExcel50
'                                    ''''                                  crar1.Action = 1
'                                    ''''                            End If
'
'                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    ''''''''''''''''''''''''''''''''''''''ARCHIVO3 TODOS LOS CODIGOS CAPTURADOS PARA EL SOFTWARE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    If rsSQL.State = 1 Then rsSQL.Close
'                                    If rsFormulario.State = 1 Then rsFormulario.Close
'                                    rsSQL.Source = "Select * from vw_entregaarchivo3 order by id_entrega"
'                                    rsSQL.Open
'                                    If rsSQL.RecordCount > 0 Then
'                                         'Existe_Temporal
'                                         'Create_TemporalDBF
'
'                                         'Borrar_tempo
'                                         Borrar_Tempo_Sql
'                                         'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'                                         StrArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_inicio_mod_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'
'
'                                         If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
'                                            lstMensajes.AddItem "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'
'                                         'lngFile = FreeFile
'                                         'Open strArchivo For Output As #lngFile
'
'                                         'Print #lngFile, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"
'                                         Do While Not rsSQL.EOF
'                                                  StrRucArchivo = ("" & rsSQL!ds_ruc)
'                                                  StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
'                                                  StrComentario = ("" & rsSQL!ds_comentario)
'                                                  StrCodigo = ("" & rsSQL!ds_cod_descargo)
'                                                  Strcuspp = Replace(("" & rsSQL!ds_cuspp), "Ð", "Ñ")
'                                                  strPeriodo = ("" & rsSQL!ds_periodo)
'                                                  StrFechaRef = ("" & rsSQL!ds_fecha_refe)
'                                                  StrArchivotif = ("" & rsSQL!ds_archivo)
'
'                                                  If ("" & rsSQL!ds_estado) = "1" Then
'
'                                                           'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
'                                                           strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
'                                                           cnServer.Execute strCadenaSQL
'
'                                                           'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, Trim$("CESE DE RELACION LABORAL"), Trim$("02"), Strcuspp, strPeriodo, Trim$("" & rsSQL!ds_fecha_cese)
'                                                           strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & Trim$("CESE DE RELACION LABORAL") & "','" & Trim$("02") & "','" & Strcuspp & "','" & strPeriodo & "','" & Trim$("" & rsSQL!ds_fecha_cese) & "','" & StrArchivotif & "'"
'                                                           cnServer.Execute strCadenaSQL
'                                                   'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
'                                                   'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$("CESE DE RELACION LABORAL"); vbTab; Trim$("02"); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$("" & rsSQL!ds_fecha_cese)
'                                                  Else
'
'
'
'                                                   'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
'                                                   strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
'                                                   cnServer.Execute strCadenaSQL
'
'
'                                                   'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
'                                                  End If
'
'                                                  strLinea = ""
'
'
'                                         rsSQL.MoveNext
'                                         Loop
'                                         'Close #lngFile
'
'                                            strsqlp = "Select RUC,RASON_SOC AS RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
'                                            If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
'                                            End If
'
'                                             If rsSQL.State = 1 Then rsSQL.Close
'                                             rsSQL.Source = "Select * from Tempo"
'                                             rsSQL.Open
'                                             intRegistroValidos = 0
'                                             If rsSQL.RecordCount > 1 Then
'                                             intRegistroValidos = rsSQL.RecordCount
'                                             End If
'
'                                    '      Crtxls.DataFiles(0) = App.Path & "\TEMPO.DBF"
'                                    '      Crtxls.ReportFileName = App.Path & "\renvio.rpt"
'                                    '      Crtxls.Destination = crptToFile
'                                    '      Crtxls.PrintFileName = strArchivo
'                                    '      Crtxls.PrintFileType = crptExcel50
'                                    '      Crtxls.Action = 1
'
'
'                                    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''LASER FISH ARCHIVO LST''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    If rsSQL.State = 1 Then rsSQL.Close
'                    rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch"
'
'                    rsSQL.Open
'                    If rsSQL.RecordCount > 0 Then
'                        intPedidosRechazados = rsSQL.RecordCount
'
'                        strArchivo = strDirectorioLST & "\Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
'                        If Dir(strArchivo, vbArchive) <> "" Then Kill strArchivo
'
'                        lstMensajes.AddItem "Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
'                        rsSQL.MoveFirst
'
'                        lngFile = FreeFile
'
'
'                        Dim strFolderPadre As String
'                        Dim strFolderActual As String
'
'                        Open strArchivo For Output As #lngFile
'
'                            Print #lngFile, "LASERFICHE IMPORT LIST"
'                            Print #lngFile, "FOLDERS(Liquidacion Previa)"
'                            Print #lngFile, "FOLDERS(" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ")"
'                            strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                            strFolderActual = strFolderPadre
'                            Print #lngFile, "FOLDERS(" & strFolderActual & ")"
'                            Do While Not rsSQL.EOF
'                                strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                                If strFolderPadre <> strFolderActual Then
'                                    Print #lngFile, "FOLDERS()"
'                                    strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                                    strFolderActual = strFolderPadre
'                                    Print #lngFile, "FOLDERS(" & strFolderActual & ")"
'                                End If
'
'                                Print #lngFile, "STARTFIELDS"
'                                Print #lngFile, strPlantillaLF
'                                Print #lngFile, ("Liquidacion Previa")
'                                Print #lngFile, ("" & rsSQL!ID_liq)
'                                Print #lngFile, ("" & rsSQL!ds_nroliq)
'                                Print #lngFile, Format(Year(Date), "0000")
'
'                                Print #lngFile, (rsSQL!ds_periodo)
'
'                                'Print #lngFile, (rsSQL!priapell) & " " & (rsSQL!segapell) & " " & (rsSQL!nombres)   '''APELLIDOS Y NOMBRE  80 CARACTERES
'                                Print #lngFile, ("Vigente")
'                                Print #lngFile, "ENDFIELDS"
'
'                                Print #lngFile, "DOCUMENT(" & "LP_" & Format(rsSQL!ID_liq, "0000000000") & "_" & ("" & rsSQL!ds_nroliq) & ")"
'                                Print #lngFile, "STARTLIST"
'                                Print #lngFile, rsSQL!Orig_file
'                                Print #lngFile, "ENDLIST"
'                                rsSQL.MoveNext
'                            Loop
'                        Close #lngFile
'
'                    End If
'                    If rsSQL.State = 1 Then rsSQL.Close


If rsSQL.State = 1 Then rsSQL.Close
rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch,orig_file"

rsSQL.Open
Dim StrTramite As String
If rsSQL.RecordCount > 0 Then

   StrArchivo = strDirectorioLST & "\Lst_" & Format(Now, "YYYY") & "_General.lst"
'lstMensajes.AddItem "Se genero el archivo LST: " & StrArchivo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  rsSQL.MoveLast
  rsSQL.MoveFirst
'   xMin = "0000" 'Select_Imagen_Min(XBATCH)
'   xMax = Select_Imagen_Max(XBATCH)
strPeriodo = rsSQL!ds_periodo
strAnno = rsSQL!ds_año
If rsSQL.RecordCount > 0 Then
                intPedidosRechazados = rsSQL.RecordCount
                        

                StrArchivo = strDirectorioLST & "\Lst_" & Format(Now, "YYYY") & "_" & "General.lst"
                If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
                'lstMensajes.AddItem "Lst_" & ("" & strAnno) & "_" & ("" &rsSql!LOTE) & ".lst"
                lstMensajes.AddItem "Se genero el archivo LST: " & StrArchivo
                'rsSQL.MoveFirst
                lngFile = FreeFile
                
                Open StrArchivo For Output As #lngFile

                    Print #lngFile, "LASERFICHE IMPORT LIST"
                    Print #lngFile, "FOLDERS(LIQUIDACIONES PREVIAS)"
                    Print #lngFile, "FOLDERS(" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ")"
                    strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & rsSQL!DS_batch)
                    strFolderActual = strFolderPadre
                    'Print #lngFile, "FOLDERS(" & strFolderActual & ")"
                    
                    xMin = "0000" 'Select_Imagen_Min(XBATCH)
                    xMax = Select_Imagen_Max(cmbBatch)
                    rsSQL.MoveLast
                    xExcellUltimo = Mid(rsSQL.Fields("orig_file"), Len(rsSQL!Orig_file) - 7, 4)
                    rsSQL.MoveFirst
                        Do While Not rsSQL.EOF
                        xMin = xMin + 1
                                  If Format(xMin, "0000") = Mid(rsSQL!Orig_file, Len(rsSQL!Orig_file) - 7, 4) Then
                                    If strFolderPadre <> strFolderActual Then
                                        Print #lngFile, "FOLDERS()"
                                        strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & rsSQL!DS_batch)
                                        strFolderActual = strFolderPadre
                                        Print #lngFile, "FOLDERS(" & strFolderActual & ")"
                                    End If
                                    
                                    If Trim$(("" & rsSQL!ds_numerotra)) = "" Then
                                    StrTramite = "0000001"
                                    Else
                                    StrTramite = ("" & rsSQL!ds_numerotra)
                                    End If
'                                    If StrTramite = "1706362" Then
'                                    MsgBox "prueba"
'                                    End If
                                    strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & rsSQL!DS_batch)
                                    Print #lngFile, "DOCUMENT(" & "" & StrTramite & ")"
                                    StrReclamo = "" & StrTramite
                                    Print #lngFile, "STARTFIELDS"
                                    Print #lngFile, ("LIQUIDACIONES PREVIAS")
                                    Print #lngFile, ("" & rsSQL!razon_social)
                                    StrRazonAd = ("" & rsSQL!razon_social)
                                    Print #lngFile, ("" & rsSQL!Ruc)
                                    StrRucAd = ("" & rsSQL!Ruc)
                                    Print #lngFile, ("" & rsSQL!Periodo)
                                    StrPeriodoAd = ("" & rsSQL!Periodo)
                                    Print #lngFile, ("" & StrTramite)
                                    Print #lngFile, "ENDFIELDS"
                                    Print #lngFile, "STARTLIST"
                                    Print #lngFile, ("" & rsSQL!Orig_file)
                                    RutaArchivo = Mid(("" & rsSQL!Orig_file), 1, Len(rsSQL!Orig_file) - 25)
                                    Print #lngFile, "ENDLIST"
                                    
                                    If Mid(rsSQL.Fields("orig_file"), Len(rsSQL!Orig_file) - 7, 4) = xExcellUltimo Then
                                    
                                        If xMax - xExcellUltimo > 0 Then
                                             For IntContador = 1 To xMax - xExcellUltimo
                                                xExcellUltimo = xExcellUltimo + 1
                                                xxRuta = Select_Imagen_Existe("" & rsSQL!DS_batch, Format(xExcellUltimo, "0000"))
                                                Print #lngFile, "DOCUMENT(" & "" & StrTramite & ")"
                                                StrReclamo = "" & StrTramite
                                                Print #lngFile, "STARTFIELDS"
                                                Print #lngFile, ("LIQUIDACIONES PREVIAS")
                                                Print #lngFile, ("" & rsSQL!razon_social)
                                                Print #lngFile, ("" & rsSQL!Ruc)
                                                Print #lngFile, ("" & rsSQL!Periodo)
                                                
                                                Print #lngFile, ("" & StrTramite)
                                                Print #lngFile, "ENDFIELDS"
                                                Print #lngFile, "STARTLIST"
                                                'Print #lngFile, ("" & rsSQL!Orig_file)
                                                Print #lngFile, ("" & RutaArchivo & xxRuta)
                                                'RutaArchivo = Mid(("" & rsSQL!Orig_file), 1, Len(rsSQL!Orig_file) - 25)
                                                Print #lngFile, "ENDLIST"
                                             Next IntContador
                                        End If
                                    End If
                                    rsSQL.MoveNext
                               Else
                                                'xxRuta = Select_Imagen_Existe("" & rsSQL!ds_batch, Format(xExcellUltimo, "0000"))
                                                xxRuta = Select_Imagen_Existe("" & rsSQL!DS_batch, Format(xMin, "0000"))
                                                Print #lngFile, "DOCUMENT(" & "" & StrTramite & ")"
                                                StrReclamo = "" & StrTramite
                                                Print #lngFile, "STARTFIELDS"
                                                Print #lngFile, ("LIQUIDACIONES PREVIAS")
                                                Print #lngFile, ("" & StrRazonAd)
                                                Print #lngFile, ("" & StrRucAd)
                                                Print #lngFile, ("" & StrPeriodoAd)
                                                
                                                Print #lngFile, ("" & StrTramite)
                                                Print #lngFile, "ENDFIELDS"
                                                Print #lngFile, "STARTLIST"
                                                'Print #lngFile, ("" & rsSQL!Orig_file)
                                                Print #lngFile, ("" & RutaArchivo & xxRuta)
                                                'RutaArchivo = Mid(("" & rsSQL!Orig_file), 1, Len(rsSQL!Orig_file) - 25)
                                                Print #lngFile, "ENDLIST"
                                             
                               'rsSQL.MoveNext
                               End If
                        
                        Loop
                    
  End If
                
                Close #lngFile
  
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''' LASER FISH LST LIQUIDACION PREVIA'''''''''''''''''''''''''''''''''''''''''''
'        If rsSQL.State = 1 Then rsSQL.Close
'                    rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch"
'
'                    rsSQL.Open
'                    If rsSQL.RecordCount > 0 Then
'                        intPedidosRechazados = rsSQL.RecordCount
'
'                        StrArchivo = strDirectorioLST & "\Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
'                        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
'
'                        lstMensajes.AddItem "Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
'                        rsSQL.MoveFirst
'
'                        lngFile = FreeFile
'
'
'                        Dim strFolderPadre As String
'                        Dim strFolderActual As String
'
'                        Open StrArchivo For Output As #lngFile
'
'                            Print #lngFile, "LASERFICHE IMPORT LIST"
'                            Print #lngFile, "FOLDERS(RECLAMOS)"
'                            Print #lngFile, "FOLDERS(" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ")"
'                            strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                            strFolderActual = strFolderPadre
'                            'Print #lngFile, "FOLDERS(" & strFolderActual & ")"
'                            Do While Not rsSQL.EOF
'                                strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                                If strFolderPadre <> strFolderActual Then
'                                    Print #lngFile, "FOLDERS()"
'                                    strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                                    strFolderActual = strFolderPadre
'                                    Print #lngFile, "FOLDERS(" & strFolderActual & ")"
'                                End If
'                                Print #lngFile, "DOCUMENT(" & "LP_" & Format(rsSQL!id_liq, "0000000000") & "_" & ("" & rsSQL!ds_nroliq) & ")"
'                                Print #lngFile, "STARTFIELDS"
'                                'Print #lngFile, strPlantillaLF
'                                Print #lngFile, ("RECLAMOS")
'                                Print #lngFile, ("")
'                                Print #lngFile, ("" & rsSQL!id_liq)
'                                'Print #lngFile, ("" & rsSQL!ds_nroliq)
'                                'Print #lngFile, Format(Year(Date), "0000")
'
'                                'Print #lngFile, (rsSQL!ds_periodo)
'
'                                'Print #lngFile, (rsSQL!priapell) & " " & (rsSQL!segapell) & " " & (rsSQL!nombres)   '''APELLIDOS Y NOMBRE  80 CARACTERES
'                                'Print #lngFile, ("Vigente")
'                                Print #lngFile, "ENDFIELDS"
'
''                                Print #lngFile, "DOCUMENT(" & "LP_" & Format(rsSQL!ID_liq, "0000000000") & "_" & ("" & rsSQL!ds_nroliq) & ")"
'                                Print #lngFile, "STARTLIST"
'                                Print #lngFile, rsSQL!Orig_file
'                                Print #lngFile, "ENDLIST"
'                                rsSQL.MoveNext
'                            Loop
'                        Close #lngFile
'
'                    End If
'                    If rsSQL.State = 1 Then rsSQL.Close

        
'*------------------------------------------------------------------------------------------------*
'                    ''''''''''''''''''''''''''''''''''''ARCHIVO CARGO'''''''''''''''''''''''''''''''''''''''''''''''''
'                     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    If rsSQL.State = 1 Then rsSQL.Close
'                    'If rsFormulario.State = 1 Then rsFormulario.Close
'                    'rsSQL.Source = "Select * from vw_cargo where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch"
'                    rsSQL.Source = "Select * from vw_cargo where ds_batch='" & cmbBatch.Text & "' order by csid"
'                    rsSQL.Open
'                    If rsSQL.RecordCount > 0 Then
'
'                    'Borrar_tempoCargo
'                    Borrar_Tempo_Cargo_Sql
'                    StrArchivo = strDirectorio & "\Archivo_Cargo_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'
'                         If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
'                            lstMensajes.AddItem "\Archivo_Cargo_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
'
'                         lngFile = FreeFile
'                         'Open strArchivo For Output As #lngFile
'
'                         'Print #lngFile, "NROLIQ"; vbTab; "LOTE"; vbTab; "PAGINA"; vbTab; "RUTA"; vbTab; "NRORECLAMO"
'                         Do While Not rsSQL.EOF
'                                  StrNrolircargo = ("" & rsSQL!ds_nroliq)
'                                  StrLoteCargo = ("" & rsSQL!DS_batch)
'                                  StrRutaCargo = ("" & rsSQL!Orig_file)
'
'                                  If rsFormulario.State = 1 Then rsFormulario.Close
'                                  rsFormulario.Source = "select * from tb_formularios where nroliqui='" & StrNrolircargo & "'"
'                                  rsFormulario.Open
'                                  If rsFormulario.RecordCount > 0 Then
'                                  StrRucArchivo = ("" & rsFormulario!campo7)
'                                  Else
'                                  StrRucArchivo = ""
'                                  End If
'
'                                  If IIf(IsNull(rsSQL.Fields("ds_barra")), "", rsSQL.Fields("ds_barra")) = "" Then
'                                  StrPagCargo = Mid((rsSQL!ds_pag), 6, 3)
'                                  Else
'                                  StrPagCargo = Mid(rsSQL!ds_barra, 10, 3)
'                                  End If
'
'
'                                  'Grabar_TemporalCargo Trim$(StrNrolircargo), Trim$(StrLoteCargo), Trim$(StrPagCargo), Trim$(StrRutaCargo), "", Trim$(StrRucArchivo)
'
'                                  strCadenaSQL = "EXEC sp_InsertarTempo_Cargo_Envio '" & Trim$(StrNrolircargo) & "','" & Trim$(StrLoteCargo) & "','" & Trim$(StrPagCargo) & "','" & Trim$(StrRutaCargo) & "','" & Trim$(StrRucArchivo) & "'"
'                                  cnServer.Execute strCadenaSQL
'
'
'
'                                  'Print #lngFile, Trim$(StrNrolircargo); vbTab; Trim$(StrLoteCargo); vbTab; Trim$(StrPagCargo); vbTab; Trim$(StrRutaCargo); vbTab; ""
'
'                          '        strLinea = ""
'
'                         rsSQL.MoveNext
'                         Loop
'                         'Close #lngFile
'
'                    End If
'
'                            strsqlp = "Select NROLIQ,LOTE,PAGINA,RUTA,NRORECLAMO,RUC from Tempoca order by id_tempo"
'                            If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
'                            End If
      
      
'      crCargo.DataFiles(0) = App.Path & "\TEMPOca.DBF"
'      crCargo.ReportFileName = App.Path & "\rcargo.rpt"
'      crCargo.Destination = crptToFile
'      crCargo.PrintFileName = strArchivo
'      crCargo.PrintFileType = crptExcel50
'      crCargo.Action = 1
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If rsSQL.State = 1 Then rsSQL.Close
    
    'rsSQL.Source = "Select * from vw_afiliacionesRechazadas where ds_batch='" & cmbBatch.Text & "' order by id_hojapedido"
    rsSQL.Source = "Select * from vw_liquidacionesRechazadas where ds_batch='" & cmbBatch.Text & "' order by id_hojapedido"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        intPedidosRechazados = rsSQL.RecordCount
        
        StrArchivo = strDirectorio & "\Errores_Liquidacion_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        
        lstMensajes.AddItem "Errores_Liquidacion_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".txt"
        rsSQL.MoveFirst
        
        lngFile = FreeFile
        
        strErrZona = ""
        strErrIDHojaPedido = ""
        strErrAsesora = ""
        strErrRechazo = ""
        strErrValor = ""
        
        
        
        Open StrArchivo For Output As #lngFile
            strErrZona = "AÑO"
            strErrIDHojaPedido = "ID_lIQUIDACION"
            'strErrAsesora = "ASESORA"
            strErrRechazo = "MOTIVO ERROR"
            'strErrValor = "VALOR DEL CAMPO"
            
            'strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrAsesora & "|" & strErrRechazo '& "|" & strErrValor
            strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo
            Print #lngFile, strLinea

            Do While Not rsSQL.EOF
                strErrZona = rsSQL!ds_año
                strErrIDHojaPedido = Mid(Right(rsSQL!Orig_file, 8), 1, 4) ' rsSQL!id_hojapedido
                'strErrAsesora = rsSQL!ds_asesora
                strErrRechazo = rsSQL!ds_rechazo
                'strErrValor = rsSQL!ds_valor
                
                'strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrAsesora & "|" & strErrRechazo '& "|" & strErrValor
                strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo  '& "|" & strErrValor
                Print #lngFile, strLinea
                rsSQL.MoveNext
            Loop
        Close #lngFile
        
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    
'*------------------------------------------------------------------------------------------------*
    
                        If rsSQL.State = 1 Then rsSQL.Close
                        rsSQL.Source = "Select * from vw_liqVERIFICAR where ds_batch='" & cmbBatch.Text & "'"
                        rsSQL.Open
                        If rsSQL.RecordCount > 0 Then
        
                            StrArchivo = strDirectorio & "\Errores_Liquidacion_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
                            If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        
                            lstMensajes.AddItem "Errores_Liquidacion_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
                            rsSQL.MoveFirst
        
                            lngFile = FreeFile
        
                            strErrZona = ""
                            strErrIDHojaPedido = ""
                            strErrAsesora = ""
                            strErrRechazo = ""
                            strErrValor = ""
        
        
        
                            Open StrArchivo For Output As #lngFile
                                strErrZona = "AÑO"
                                strErrIDHojaPedido = "ID_LIQUIDACION"
                                'strErrAsesora = "ASESORA"
                                strErrRechazo = "MOTIVO ERROR"
                                'strErrProducto = "POSICION"
                                strErrValor = "DATO"
        
                                'strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" &  strErrRechazo & "|" & strErrValor '& "|" & strErrProducto & "|" & strErrValor
                                strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo & "|" & strErrValor  '& "|"  & strErrValor
                                Print #lngFile, strLinea
        
                                Do While Not rsSQL.EOF
                                    strErrZona = rsSQL!ds_año
                                    strErrIDHojaPedido = Mid(Right(rsSQL!Orig_file, 8), 1, 4) 'rsSQL!id_hojapedido
                                    'strErrAsesora = rsSQL!ds_asesora
                                    strErrRechazo = rsSQL!ds_rechazo
                                    'strErrProducto = rsSQL!nu_posicion
                                    strErrValor = rsSQL!ds_valor
        
                                    strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo & "|" & strErrValor
                                    Print #lngFile, strLinea
                                    rsSQL.MoveNext
                                Loop
                            Close #lngFile
        
                        End If
                        If rsSQL.State = 1 Then rsSQL.Close
    
'*------------------------------------------------------------------------------------------------*
    
    'If intPedidosValidados > 0 Then
    
    
        StrArchivo = strDirectorio & "\Resumen_" & cmbBatch.Text & ".txt"
        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "Resumen_" & strZona & "_" & cmbBatch.Text & ".txt"
        
        
        If rsZonas.State = 1 Then rsZonas.Close
        'rsZonas.Source = "select distinct ds_zona,ds_campaña,ds_batch from pedidos where ds_batch='" & cmbBatch.Text & "' order by 1,2"
         rsZonas.Source = "select distinct ds_año,ds_periodo,ds_batch from tb_liquidacion where ds_batch='" & cmbBatch.Text & "' order by 1,2"
        rsZonas.Open
        If rsZonas.RecordCount > 0 Then
            Open StrArchivo For Output As #lngFile
            strLinea = "Archivos resumen"
            Print #lngFile, strLinea
            Print #lngFile, "***************************************************************************"
            Print #lngFile, " "
            Do While Not rsZonas.EOF()
                strZona = "" & rsZonas!ds_año
                
                intPedidosValidados = 0
                intPedidosRechazados = 0
                
                If rsSQL.State = 1 Then rsSQL.Close
                'rsSQL.Source = "Select * from vw_PedidosAxis where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & " ' and ds_batch='" & rsZonas!DS_batch & "'"
                'rsSQL.Source = "Select distinct id  from vw_afiliacionAxis where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & " ' and ds_batch='" & rsZonas!DS_batch & "'"
                 rsSQL.Source = "Select distinct id_liq  from vw_Liquidacionentrega where ds_año='" & rsZonas!ds_año & "' and ds_periodo='" & rsZonas!ds_periodo & " ' and ds_batch='" & rsZonas!DS_batch & "'"
                rsSQL.Open
                If rsSQL.RecordCount > 0 Then
                    intPedidosValidados = rsSQL.RecordCount
                End If
                
                If rsSQL.State = 1 Then rsSQL.Close
                'rsSQL.Source = "Select * from vw_PedidosRechazados where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & "' and ds_batch='" & rsZonas!DS_batch & "'"
                'rsSQL.Source = "Select distinct id_hojapedido from vw_afiliacionesRechazadas where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & "' and ds_batch='" & rsZonas!DS_batch & "'"
                rsSQL.Source = "Select distinct id_hojapedido from vw_liquidacionesRechazadas where ds_año='" & rsZonas!ds_año & "' and ds_periodo='" & rsZonas!ds_periodo & "' and ds_batch='" & rsZonas!DS_batch & "'"
                rsSQL.Open
                If rsSQL.RecordCount > 0 Then
                    intPedidosRechazados = rsSQL.RecordCount
                End If
            
                Print #lngFile, "---------------------------------------------------------------------------"
                strLinea = "Año     :" & "" & rsZonas!ds_año
                Print #lngFile, strLinea
                strLinea = "Periodo  :" & "" & rsZonas!ds_periodo
                Print #lngFile, strLinea
                strLinea = "Batch    :" & "" & rsZonas!DS_batch
                Print #lngFile, strLinea
                Print #lngFile, " "
                'strLinea = "Hojas de Pedido procesados :" & intPedidosValidados + intPedidosRechazados
                strLinea = "Formatos de Liquidación procesados :" & intPedidosValidados + intPedidosRechazados
                Print #lngFile, strLinea
                'strLinea = "Hojas de Pedido validados  :" & intPedidosValidados
                strLinea = "Formatos de Liquidación validos  :" & intPedidosValidados
                Print #lngFile, strLinea
                'strLinea = "Hojas de Pedido rechazados :" & intPedidosRechazados
                strLinea = "Formatos de Liquidación rechazados :" & intPedidosRechazados
                Print #lngFile, strLinea
                Print #lngFile, " "
                strLinea = "Registros Validos :" & intRegistroValidos
                Print #lngFile, strLinea
                strLinea = "Registros Duplicados :" & intRegistroDuplicados
                Print #lngFile, strLinea
                Print #lngFile, " "
                Print #lngFile, "---------------------------------------------------------------------------"
                rsZonas.MoveNext
            Loop
            Close #lngFile
        End If
    
'    rsZonas.Source = "select id_cl from cl where id_cl='" & strCL & "'"
'    .Openm
    If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch, Orig_file  "

        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
        StrArchivo = strDirectorio & "\Cargo_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        
        
        End If
        
'    End If
   'Cambio 07/04/2011
   '''''''''''''''''''
    If MsgBox("Se Realizara el envio individual de este lote al cliente en estos momentos ?", vbYesNo + vbQuestion, "Finalizado") = vbYes Then
        
        If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "Select * from tb_envios where lotes='" & cmbBatch.Text & "'"

        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
           If MsgBox("Este Lote ya genero envio desea volver a generar envio ?", vbYesNo + vbQuestion, "Horizonte") = vbYes Then
           
            strEjecutar = "delete from tb_envios where lotes='" & cmbBatch.Text & "'"
            cnServer.Execute strEjecutar
           
            Strorden = InputBox("Ingrese el Numeo de Orden:", "Horizonte")
            strEjecutar = "Insert into tb_envios  (orden,lotes,cant,fecha) values ('" & Format(Strorden, "0000") & "','" & cmbBatch.Text & "','" & "1" & "','" & Format(Now, "dd/mm/yyyy") & "')"
            cnServer.Execute strEjecutar
           
           End If
        Else
            Strorden = InputBox("Ingrese el Numeo de Orden:", "Horizonte")
            strEjecutar = "Insert into tb_envios (orden,lotes,cant,fecha) values  ('" & Format(Strorden, "0000") & "','" & cmbBatch.Text & "','" & "1" & "','" & Format(Now, "dd/mm/yyyy") & "')"
            cnServer.Execute strEjecutar
        End If
        
        
    End If
   '''''''''''''''''''
    
    If MsgBox("Se han exportado los Archivo de texto, desea Generar los backup's y limpiar las tablas ?", vbYesNo + vbQuestion, "Finalizado") = vbYes Then
        strEjecutar = "drop table [" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "_Año_" & strZona & "_Envio_" & cmbBatch.Text & "]"
        cnServer.Execute strEjecutar
        strEjecutar = "select * into [" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "_Año_" & strZona & "_Envio_" & cmbBatch.Text & "] from tb_liquidacion where ds_batch='" & cmbBatch.Text & "'"
        cnServer.Execute strEjecutar
        strEjecutar = "delete from tb_liquidacion where ds_batch='" & cmbBatch.Text & "'"
        cnServer.Execute strEjecutar
    End If
    
    MsgBox "Fin del proceso", vbInformation, "Generaciòn del archivo de retorno"
    picAvance.Picture = LoadPicture()
    picAvance.BorderStyle = 1
    cmdProcesar.Enabled = True
    cmbBatch.Enabled = True
    Screen.MousePointer = 0
    ListarBatch
    Exit Sub
Err:
    If Err.Number = -2147217865 Then
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, App.EXEName
        Me.MousePointer = 0
        Screen.MousePointer = 0
        Unload Me
    End If
End Sub

Function Selec_Tipo_doc(StrTipo) As String
Select Case StrTipo
  Case "1": Selec_Tipo_doc = "AV"
  Case "2": Selec_Tipo_doc = "CL"
  Case "3": Selec_Tipo_doc = "JR"
  Case "4": Selec_Tipo_doc = "PR"
  Case "5": Selec_Tipo_doc = "PRG"
End Select
End Function

Function Selec_Tipo_Viv(StrTipo) As String
Select Case StrTipo
  Case "1": Selec_Tipo_Viv = "INV"
  Case "2": Selec_Tipo_Viv = "AAHH"
  Case "3": Selec_Tipo_Viv = "PPJJ"
  Case "4": Selec_Tipo_Viv = "URBANIZACION"
  Case "5": Selec_Tipo_Viv = "COOP"
  Case "6": Selec_Tipo_Viv = "UNID VEC"
End Select
End Function

Function Selec_Tipo_SN(StrTipo) As String
Select Case StrTipo
  Case "S": Selec_Tipo_SN = "06" & ";" & "TRUE"
  Case "N": Selec_Tipo_SN = "06" & ";" & "FALSE"
  Case Else
  Selec_Tipo_SN = "06" & ";" & "FALSE"
End Select
End Function

Function Selec_Est_Civil(StrTipo) As String
Select Case StrTipo
  Case "1": Selec_Est_Civil = "01" & ";" & "SOLTERA"
  Case "2": Selec_Est_Civil = "01" & ";" & "CASADA"
  Case "3": Selec_Est_Civil = "01" & ";" & "SEPARADA"
  Case "4": Selec_Est_Civil = "01" & ";" & "DIVORCIADA"
  Case "5": Selec_Est_Civil = "01" & ";" & "CONVIVIENTE"
  Case "6": Selec_Est_Civil = "01" & ";" & "OTRO"
  Case Else
  Selec_Est_Civil = ";"
End Select
End Function

Function Selec_Pro_Viv(StrTipo) As String
Select Case StrTipo
  Case "1": Selec_Pro_Viv = "04" & ";" & "CASA"
  Case "2": Selec_Pro_Viv = "04" & ";" & "CARRO"
  Case "3": Selec_Pro_Viv = "04" & ";" & "ESTUDIO"
  Case "4": Selec_Pro_Viv = "04" & ";" & "VIAJE"
  Case "5": Selec_Pro_Viv = "04" & ";" & "OTRO"
  Case Else
  Selec_Pro_Viv = ";"
End Select
End Function

Function Selec_Tiempo_Lib(StrTipo) As String
Select Case StrTipo
  Case "1": Selec_Tiempo_Lib = "05" & ";" & "DEPORTE"
  Case "2": Selec_Tiempo_Lib = "05" & ";" & "FAMILIA"
  Case "3": Selec_Tiempo_Lib = "05" & ";" & "LEER"
  Case "4": Selec_Tiempo_Lib = "05" & ";" & "MANUALIDADES"
  Case "5": Selec_Tiempo_Lib = "05" & ";" & "OTRO"
  Case Else
  Selec_Tiempo_Lib = ";"
End Select
End Function

Function Selec_Tipo_Suministro(StrTipo) As String
Select Case StrTipo
  Case "EE": Selec_Tipo_Suministro = "01"
  Case "AGUA": Selec_Tipo_Suministro = "02"
  Case Else
  Selec_Tipo_Suministro = ""
End Select
End Function

Function Selec_Tipo_Correo(StrTipo) As String
Select Case StrTipo
  Case "S": Selec_Tipo_Correo = "07" & ";" & "TRUE"
  Case "N": Selec_Tipo_Correo = "07" & ";" & "FALSE"
  Case Else
  Selec_Tipo_Correo = "07" & ";" & "FALSE"
End Select
End Function

Function Selec_Tipo_Credito(StrTipo) As String
Select Case StrTipo
  Case "S": Selec_Tipo_Credito = "50" & ";" & "TRUE"
  Case "N": Selec_Tipo_Credito = "50" & ";" & "FALSE"
  Case Else
  Selec_Tipo_Credito = "50" & ";" & "FALSE"
End Select
End Function

Private Function Exportar_ADO_Excel(Sql As String, sOutputPathXLS As String) As Boolean
       
    On Error GoTo errSub
       
    Dim cn          As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
       
    Me.Enabled = False
       
   ' -- Abrir la base
   
   rec.CursorType = adOpenDynamic
    rec.LockType = adLockBatchOptimistic
    rec.ActiveConnection = cnServer
   
   
   If rec.State = 1 Then rec.Close
          'rsFormulario.Source = "select * from tb_entrega where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_periodo='" & ("" & rsSQL!ds_periodo) & "'"
          rec.Source = Sql
          rec.Open
   
   
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPathDB & ";"
           
    ' -- Abrir el Recordset pasándole la cadena sql
    'rec.Open Sql, cn
       
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
       
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
       
    'Excel.Visible = True: Excel.UserControl = True
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
        Hoja.Cells(1, iCol).Font.Bold = True
        Hoja.Cells(1, iCol).NumberFormat = "@"
    Next
       
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(2, 1).CopyFromRecordset rec
    Else
  
        arrData = rec.GetRows
  
        iRec = UBound(arrData, 2) + 1
           
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
                    'Hoja.Selection.NumberFormat = "@"
  
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
               
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
    End If
  
    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.rows.AutoFit
  
    ' -- Cierra el recordset y la base de datos y los objetos ADO
    rec.Close
    'cn.Close
       
    Set rec = Nothing
    Set cn = Nothing
    ' -- guardar el libro
    Libro.SaveAs sOutputPathXLS
    Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    Excel.quit
    Set Excel = Nothing
       
    Exportar_ADO_Excel = True
    Me.Enabled = True
    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel = False
    Me.Enabled = True
End Function


Private Function GetData(vValue As Variant) As Variant
    Dim X As Long, Y As Long, xMax As Long, yMax As Long, T As Variant
       
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
       
    ReDim T(xMax, yMax)
    For X = 0 To xMax
        For Y = 0 To yMax
            T(X, Y) = vValue(Y, X)
        Next Y
    Next X
       
    GetData = T
End Function

Sub GenerarEntregableTXT_TODOS()
On Local Error GoTo Err:

Dim lngFile As Long
Dim strLinea As String
Dim intOk As Integer

Dim strErrZona As String * 4
Dim strErrIDHojaPedido As String * 8
Dim strErrAsesora As String * 14
Dim strErrRechazo As String * 30
Dim strErrProducto As String * 10
Dim strErrValor As String * 20

Dim bolZip As Boolean
Dim strTXT As String
Dim strZip As String

Dim intProductos As Integer
Dim intIndice As Integer

Dim strProd As String
Dim strCant As String
Dim intTotProd As Integer

Dim intPedidosProcesados As Integer
Dim intPedidosRechazados As Integer
Dim intPedidosValidados As Integer
Dim rsZonas As New Recordset

Dim rsFormulario As New Recordset
Dim rsDuplica As New Recordset
Dim StrCampo As String
Dim StrDato1 As String
Dim StrDato2 As String
Dim StrDato As String

Dim StrRucArchivo As String
Dim StrRaSocArchivo As String
Dim StrComentario As String
Dim StrCodigo As String
Dim Strcuspp As String
Dim strPeriodo As String
Dim StrFechaRef As String

Dim StrNrolircargo As String
Dim StrLoteCargo As String
Dim StrPagCargo As String
Dim StrRutaCargo As String
Dim StrArchivotif As String

Dim StrArchivo_Texto As String
Dim Z As Integer

    rsZonas.CursorType = adOpenDynamic
    rsZonas.LockType = adLockBatchOptimistic
    rsZonas.ActiveConnection = cnServer
    
    rsFormulario.CursorType = adOpenDynamic
    rsFormulario.LockType = adLockBatchOptimistic
    rsFormulario.ActiveConnection = cnServer
    
    rsDuplica.CursorType = adOpenDynamic
    rsDuplica.LockType = adLockBatchOptimistic
    rsDuplica.ActiveConnection = cnServer
    
    Screen.MousePointer = 11
    
    lstMensajes.Clear
    
    If Trim(cmbBatch.Text) = "" Then
        MsgBox "Seleccione un número de batch", vbCritical, "Generaciòn de archivos de retorno"
       Screen.MousePointer = 0
        Exit Sub
    End If
    
    'cmbbatch.Text = Padln(Trim(mskEnvio.Text), 4, "0")
    
    cmdProcesar.Enabled = False
    cmbBatch.Enabled = False
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "delete from tb_Entrega"
    rsSQL.Open
    
    
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from tb_Parametros"
    rsSQL.Open
    
    strCampañaActual = "" & rsSQL!ds_periodo
    strZonaActual = "" & rsSQL!ds_año
    strDirectorio = "" & rsSQL!ds_ftpaxis
    intProductos = 0 & rsSQL!nu_cl
    
    
    Dim strsqlp As String
    Dim HojaExcel As String
    'ExcelFile = App.Path & "\Libro1.xls"
    'HojaExcel = "[Hoja1]"
    'cnServer.Execute "SELECT * INTO [Excel 8.0;DATABASE=" & ExcelFile & "]." & _
    '    HojaExcel & " FROM tempo"
    
'    strsqlp = "Select * from Tempo"
'    If Exportar_ADO_Excel(strsqlp, "c:\libro1.xLS") Then
'       MsgBox "Ok", vbInformation
'    End If

    
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "delete from tb_Entrega_H where ds_batch='" & cmbBatch.Text & "'"
    rsSQL.Open
    
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    
    lstMensajes.AddItem "DIRECTORIO DE ENVIO"
    lstMensajes.AddItem UCase(strDirectorio)
    lstMensajes.AddItem ""
    lstMensajes.AddItem "ARCHIVOS GENERADOS"
    
'*------------------------------------------------------------------------------------------------*
    strZona = "0000"
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from vw_liquidacionentrega where ds_batch='" & cmbBatch.Text & "' ORDER BY csid"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
    
        strZona = "" & rsSQL!ds_año
        intPedidosValidados = rsSQL.RecordCount
        
        'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".txt"
        'If Dir(strArchivo, vbArchive) <> "" Then Kill strArchivo
        'lstMensajes.AddItem "Pedidos_" & ("" & rsSQL!ds_zona) & "_" & cmbBatch.Text & ".txt"
        'lstMensajes.AddItem "Descargo_Masivo_Empleadores_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".txt"
        
        rsSQL.MoveFirst
        
        strBatch = ""
        strAsesora = ""
        strCampaña = ""
        strCodigoImportacion = "01"
        strCliente = "000"
        strFolletoSiguiente = ""
        strFolletoPosterior = ""
        strFiller = "0000"
        strCL = ""
        strCantidad = ""
        'lngFile = FreeFile
        'Open strArchivo For Output As #lngFile
        
        'Print #lngFile, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"
                        
            Do While Not rsSQL.EOF
                ReDim xcuspp(1 To intProductos) As String
                ReDim Xcodigo(1 To intProductos) As String
                ReDim xfecha1(1 To intProductos) As String
                ReDim xfecha2(1 To intProductos) As String
                ReDim xcomen(1 To intProductos) As String
                ReDim xFere(1 To intProductos) As String
                ReDim xmonto(1 To intProductos) As String
                ReDim xPlani(1 To intProductos) As String
                ReDim CAMPO(1 To intProductos) As String
                                      
                      
                      If rsFormulario.State = 1 Then rsFormulario.Close
                      StrDato1 = Trim$(rsSQL.Fields("ds_nroliq"))
                      If IIf(IsNull(rsSQL.Fields("ds_barra")), "", rsSQL.Fields("ds_barra")) = "" Then
                      StrDato2 = Trim$(Mid(rsSQL.Fields("ds_pag"), 6, 3))
                      Else
                      StrDato2 = Trim$(Mid(rsSQL.Fields("ds_barra"), 11, 3))
                      End If
                      StrDato = StrDato1 & StrDato2
                      rsFormulario.Source = "select * from tb_formularios where nroliqpag='" & StrDato & "'"
                      rsFormulario.Open
                
                If rsFormulario.RecordCount > 0 Then
                        For intIndice = 1 To intProductos
                                  
                                    StrCampo = "cuspp" & intIndice
                                    strProducto = Trim(("" & rsFormulario.Fields(StrCampo)))
                                    If strProducto = "" Then
                                      Exit For
                                    Else
                                        xcuspp(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        StrCampo = "ds_cod" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        Xcodigo(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        StrCampo = "ds_fec_ini" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xfecha1(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        StrCampo = "ds_fec_ce" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xfecha2(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                                
                                        StrCampo = "ds_rem_as" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xmonto(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        StrCampo = "ds_pla" & intIndice
                                        strProducto = Trim(("" & rsSQL.Fields(StrCampo)))
                                        xPlani(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        
                                        If Xcodigo(intIndice) = "" Then
                                        Xcodigo(intIndice) = Codigo_Nulo(xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice))
                                        End If
                                        
                                        xFere(intIndice) = Mostrar_Fecha_Ref(Xcodigo(intIndice), xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice), xPlani(intIndice))
                                        xcomen(intIndice) = Mostrar_Comentario(Xcodigo(intIndice), xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice), xPlani(intIndice))
                                        
                                        If Xcodigo(intIndice) = "99" And xcomen(intIndice) = "EMPLEADOR CONFIRMA DEUDA" Then
                                        Xcodigo(intIndice) = "77"
                                        End If
                                        
                                        Xcodigo(intIndice) = Verificar_No_Existe(Xcodigo(intIndice))
                                        
                                        'If Xcodigo(intIndice) <> "" Or xfecha1(intIndice) <> "" Or xfecha2(intIndice) <> "" Or xmonto(intIndice) <> "" Then
                                        If xfecha1(intIndice) <> "" Or xfecha2(intIndice) <> "" Or xmonto(intIndice) <> "" Or xPlani(intIndice) <> "" Or xcomen(intIndice) <> "" Then
                                        strCadenaSQL = "EXEC sp_InsertarEntrega '" & StrDato & "','" & cmbBatch.Text & "','" & Replace(("" & rsFormulario!campo7), "'", "`") & "','" & ("" & Replace(rsFormulario!campo6, "'", "`")) & "','" & xcomen(intIndice) & "','" & Xcodigo(intIndice) & "','" & xcuspp(intIndice) & "','" & ("" & rsFormulario!campo5) & "','" & xFere(intIndice) & "','" & xfecha2(intIndice) & "','" & StrDupli & "','" & ("" & rsSQL.Fields("csid")) & "'"
                                        cnServer.Execute strCadenaSQL
                                        
                                        
                                        strCadenaSQL = "EXEC sp_InsertarEntrega_H '" & StrDato & "','" & cmbBatch.Text & "','" & Replace(("" & rsFormulario!campo7), "'", "`") & "','" & ("" & Replace(rsFormulario!campo6, "'", "`")) & "','" & xcomen(intIndice) & "','" & Xcodigo(intIndice) & "','" & xcuspp(intIndice) & "','" & ("" & rsFormulario!campo5) & "','" & xFere(intIndice) & "','" & xfecha2(intIndice) & "','" & StrDupli & "','" & ("" & rsSQL.Fields("csid")) & "'"
                                        cnServer.Execute strCadenaSQL
                                        
                                        End If
                                    End If
                         
                        Next intIndice
                End If
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
'                strLinea = Trim$(strBatch) & ";" & Trim$(StrTipDoc) & ";" & Trim$(StrNumerDoc) & ";" & Trim$(StrTipDocLe) & ";" & Trim$(StrNroTipDocLe) & ";" & Trim$(strZona) & ";" & Trim$(StrSector) & ";" & Trim$(StrTerrito) & ";" & Trim$(StrNivelA3) & ";" & Trim$(StrNivelB1) & ";" & Trim$(StrNivelB2) & ";" & Trim$(StrNivelB3) & ";" & Trim$(StrAnnio) & ";" & Trim$(StrCamp) & ";" & Trim$(StrSexo) & ";" & Trim$(StrNomnbres) & ";" & Trim$(StrPriApell) & ";" & Trim$(StrSegApell) & ";" & Trim$(StrDirRecibe) & ";" & Trim$(StrRef) & ";" & Trim$(StrNumSumi) & ";" & Trim$(StrTipSumi) & ";" & Trim$(StrDirProv) & ";" & Trim$(StrDirDistri) & ";" & Trim$(StrDirVive) & ";" & Trim$(StrVivRef) & ";" & Trim$(StrVivNumSumi) & ";" & Trim$(StrVivTipSumi) & ";" & Trim$(StrVivProv) & ";" & Trim$(StrVivDistri) & ";" & Trim$(StrTelf1) & ";" & Trim$(StrTelf2) & ";" & Trim$(StrFecha) & ";" & Trim$(StrDirCorreo)
'                strLinea = Trim$(strLinea) & ";" & Trim$(StrDirCred) & ";" & Trim$(StrNse) & ";" & Trim$(StrFuente) & ";" & Trim$(StrCodigo) & ";" & Trim$(StrMercado) & ";" & Trim$(StrCobro) & ";" & Trim$(StrModalidad) & ";" & Trim$(StrTipoAsesora) & ";" & Trim$(StrTisue) & ";" & Trim$(StrEsCivil) & ";" & Trim$(StrNroHijos) & ";" & Trim$(StrProyVi) & ";" & Trim$(StrTiLibr) & ";" & Trim$(StrInter) & ";" & Trim$(StrInforcorreo)
'
'                intTotProd = 0
'
'                Print #lngFile, strLinea
'                strLinea = ""

                rsSQL.MoveNext
            Loop
        Close #lngFile
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    If rsFormulario.State = 1 Then rsFormulario.Close
              
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Generar Archivo Duplicados--Archivo 2'''''''''''''''''''''''''''''

If rsSQL.State = 1 Then rsSQL.Close
If rsDuplica.State = 1 Then rsDuplica.Close
If rsFormulario.State = 1 Then rsFormulario.Close
rsSQL.Source = "Select * from vw_entregaarchivo2"
rsSQL.Open
 Borrar_Tempo_Sql
 StrArchivo = strDirectorio & "\Archivo2_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
 If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
 lstMensajes.AddItem "Archivo2_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
If rsSQL.RecordCount > 0 Then
     Do While Not rsSQL.EOF
     '''''''''''''''''''''''''''''''''''''''''''
     '''''''''''''''''''''''''''''''''''''''''
          If rsDuplica.State = 1 Then rsDuplica.Close
          rsDuplica.Source = "select * from tb_entrega where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_ruc='" & ("" & rsSQL!ds_ruc) & "' ORDER BY ID_ENTREGA"
          rsDuplica.Open
          If rsDuplica.RecordCount > 0 Then
           For Z = 1 To rsDuplica.RecordCount
                If Z <> 1 Then
                strEjecutar = "Update tb_ENTREGA set ds_estadodu='1' where id_entrega = " & rsDuplica.Fields("id_entrega") & ""
                cnServer.Execute strEjecutar
                End If
           rsDuplica.MoveNext
           Next Z
          End If
     '''''''''''''''''''''''''''''''''''''''''''''''
     ''''''''''''''''''''''''''''''''''''''''''''''
          If rsFormulario.State = 1 Then rsFormulario.Close
          rsFormulario.Source = "select ds_ruc, ds_raz_soc,ds_comentario, ds_cod_descargo, ds_cuspp , ds_fecha_refe, ds_periodo, ds_estado, ds_estadodu ,id_entrega,ds_archivo from tb_entrega where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_ruc='" & ("" & rsSQL!ds_ruc) & "' order by id_entrega"
          rsFormulario.Open
          If rsFormulario.RecordCount > 0 Then
            Do While Not rsFormulario.EOF
             If ("" & rsFormulario!ds_estadodu) <> "0" Then
                StrRucArchivo = ("" & rsFormulario!ds_ruc)
                StrRaSocArchivo = Replace(("" & rsFormulario!ds_raz_soc), "Ð", "Ñ")
                StrComentario = ("" & rsFormulario!ds_comentario)
                StrCodigo = ("" & rsFormulario!ds_cod_descargo)
                Strcuspp = Replace(("" & rsFormulario!ds_cuspp), "Ð", "Ñ")
                strPeriodo = ("" & rsFormulario!ds_periodo)
                StrFechaRef = ("" & rsFormulario!ds_fecha_refe)
                StrArchivotif = ("" & rsFormulario!ds_archivo)
                
                strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
                cnServer.Execute strCadenaSQL
                strLinea = ""
              End If
            rsFormulario.MoveNext
            Loop
          End If
      rsSQL.MoveNext
     Loop
 End If
      ''''''''''''''''''''''''''''''
      '''''Duplicados historico'''''''''''''''''''''''''
         If rsSQL.State = 1 Then rsSQL.Close
         rsSQL.Source = "Select * from TB_ENTREGA ORDER BY ID_ENTREGA"
         rsSQL.Open
         If rsSQL.RecordCount > 0 Then
         
          Do While Not rsSQL.EOF
              If rsDuplica.State = 1 Then rsDuplica.Close
                'rsFormulario.Source = "select * from tb_entrega where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_periodo='" & ("" & rsSQL!ds_periodo) & "'"
                rsDuplica.Source = "select ds_ruc, ds_raz_soc,ds_comentario, ds_cod_descargo, ds_cuspp , ds_fecha_refe, ds_periodo,ds_archivo from tb_entrega_H where ds_cuspp='" & ("" & rsSQL!ds_cuspp) & "' and ds_comentario='" & ("" & rsSQL!ds_comentario) & "' and ds_fecha_refe='" & ("" & rsSQL!ds_fecha_refe) & "' and ds_ruc='" & ("" & rsSQL!ds_ruc) & "'  AND DS_BATCH <> '" & cmbBatch.Text & "'"
                rsDuplica.Open
                If rsDuplica.RecordCount > 0 Then
                        strEjecutar = "Update tb_ENTREGA set ds_estadodu='1' where id_entrega = " & rsSQL.Fields("id_entrega") & ""
                        cnServer.Execute strEjecutar
                        For Z = 1 To rsDuplica.RecordCount
                           StrRucArchivo = ("" & rsDuplica!ds_ruc)
                           StrRaSocArchivo = Replace(("" & rsDuplica!ds_raz_soc), "Ð", "Ñ")
                           StrComentario = ("" & rsDuplica!ds_comentario)
                           StrCodigo = ("" & rsDuplica!ds_cod_descargo)
                           Strcuspp = Replace(("" & rsDuplica!ds_cuspp), "Ð", "Ñ")
                           strPeriodo = ("" & rsDuplica!ds_periodo)
                           StrFechaRef = ("" & rsDuplica!ds_fecha_refe)
                           StrArchivotif = ("" & rsDuplica!ds_archivo)
                           
                           strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
                           cnServer.Execute strCadenaSQL
                           strLinea = ""
                        rsDuplica.MoveNext
                        Next Z
                Else
                End If
          rsSQL.MoveNext
          Loop
         End If
      '''''''''''''''''''''''''''''
      
strsqlp = "Select RUC,RASON_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER from Tempo order by id_tempo"
If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
End If


''''                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''                            ''''''''''''''''''''''''''''''''''''''ARCHIVO1 CODIGOS QUE NO SE CARGAN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''                            If rsSQL.State = 1 Then rsSQL.Close
''''                            If rsFormulario.State = 1 Then rsFormulario.Close
''''                            rsSQL.Source = "Select * from vw_entregaarchivo1 order by id_entrega"
''''                            rsSQL.Open
''''                            If rsSQL.RecordCount > 0 Then
''''                                 'Existe_Temporal
''''                                 'Create_TemporalDBF
''''                                 Borrar_tempo
''''                                 strArchivo = strDirectorio & "\Archivo1_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
''''
''''                                 If Dir(strArchivo, vbArchive) <> "" Then Kill strArchivo
''''                                    lstMensajes.AddItem "\Archivo1_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
''''                                 'lngFile = FreeFile
''''
''''                                 'Open strArchivo For Output As #lngFile
''''
''''                                 'Print #lngFile, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"
''''                                 Do While Not rsSQL.EOF
''''                                          StrRucArchivo = ("" & rsSQL!ds_ruc)
''''                                          StrRaSocArchivo = ("" & rsSQL!ds_raz_soc)
''''                                          StrComentario = ("" & rsSQL!ds_comentario)
''''                                          StrCodigo = ("" & rsSQL!ds_cod_descargo)
''''                                          Strcuspp = ("" & rsSQL!ds_cuspp)
''''                                          strPeriodo = ("" & rsSQL!ds_periodo)
''''                                          StrFechaRef = ("" & rsSQL!ds_fecha_refe)
''''
''''                                          Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
''''                                          'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
''''                                          'strLinea = ""
''''
''''
''''                                 rsSQL.MoveNext
''''                                 Loop
''''                                 'Close #lngFile
''''                                  crar1.DataFiles(0) = App.Path & "\TEMPO.DBF"
''''                                  crar1.ReportFileName = App.Path & "\renvio.rpt"
''''                                  crar1.Destination = crptToFile
''''                                  'Crtxls.PrintFileName = Dirxls & "\" & "CPT_" & Xfecha & "_" & XposIni + 1 & "_" & XposFin & "_" & flxls.FileName
''''                                  'Xarchivo = Dirxls & "\" & "CPT_" & Xfecha & "_" & XposIni + 1 & "_" & XposFin & "_" & flxls.FileName
''''                                  crar1.PrintFileName = strArchivo
''''                                  crar1.PrintFileType = crptExcel50
''''                                  crar1.Action = 1
''''                            End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''ARCHIVO3 TODOS LOS CODIGOS CAPTURADOS PARA EL SOFTWARE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If rsSQL.State = 1 Then rsSQL.Close
If rsFormulario.State = 1 Then rsFormulario.Close
rsSQL.Source = "Select * from vw_entregaarchivo3 order by id_entrega"
rsSQL.Open
If rsSQL.RecordCount > 0 Then
     'Existe_Temporal
     'Create_TemporalDBF
     
     'Borrar_tempo
     Borrar_Tempo_Sql
     'strArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     StrArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_inicio_mod_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\Descargo_Masivo_Empleadores_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     
     'lngFile = FreeFile
     'Open strArchivo For Output As #lngFile
        
     'Print #lngFile, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"
     Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL!ds_ruc)
              StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              StrComentario = ("" & rsSQL!ds_comentario)
              StrCodigo = ("" & rsSQL!ds_cod_descargo)
              Strcuspp = Replace(("" & rsSQL!ds_cuspp), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL!ds_periodo)
              StrFechaRef = ("" & rsSQL!ds_fecha_refe)
              StrArchivotif = ("" & rsSQL!ds_archivo)
              
              If ("" & rsSQL!ds_estado) = "1" Then
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
                       
                       'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, Trim$("CESE DE RELACION LABORAL"), Trim$("02"), Strcuspp, strPeriodo, Trim$("" & rsSQL!ds_fecha_cese)
                       strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & Trim$("CESE DE RELACION LABORAL") & "','" & Trim$("02") & "','" & Strcuspp & "','" & strPeriodo & "','" & Trim$("" & rsSQL!ds_fecha_cese) & "','" & StrArchivotif & "'"
                       cnServer.Execute strCadenaSQL
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$("CESE DE RELACION LABORAL"); vbTab; Trim$("02"); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$("" & rsSQL!ds_fecha_cese)
              Else
               
               
               
               'Grabar_TemporalDBF StrRucArchivo, StrRaSocArchivo, StrComentario, StrCodigo, Strcuspp, strPeriodo, StrFechaRef
               strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
               cnServer.Execute strCadenaSQL
               
               
               'Print #lngFile, Trim$(StrRucArchivo); vbTab; Trim$(StrRaSocArchivo); vbTab; Trim$(StrComentario); vbTab; Trim$(StrCodigo); vbTab; Trim$(Strcuspp); vbTab; Trim$(strPeriodo); vbTab; Trim$(StrFechaRef)
              End If
              
              strLinea = ""
          
          
     rsSQL.MoveNext
     Loop
     'Close #lngFile
     
        strsqlp = "Select RUC,RASON_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
        End If
     
'      Crtxls.DataFiles(0) = App.Path & "\TEMPO.DBF"
'      Crtxls.ReportFileName = App.Path & "\renvio.rpt"
'      Crtxls.Destination = crptToFile
'      Crtxls.PrintFileName = strArchivo
'      Crtxls.PrintFileType = crptExcel50
'      Crtxls.Action = 1
     
     
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''LASER FISH ARCHIVO LST''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    If rsSQL.State = 1 Then rsSQL.Close
'                    rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch"
'
'                    rsSQL.Open
'                    If rsSQL.RecordCount > 0 Then
'                        intPedidosRechazados = rsSQL.RecordCount
'
'                        strArchivo = strDirectorioLST & "\Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
'                        If Dir(strArchivo, vbArchive) <> "" Then Kill strArchivo
'
'                        lstMensajes.AddItem "Lst_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".lst"
'                        rsSQL.MoveFirst
'
'                        lngFile = FreeFile
'
'
'                        Dim strFolderPadre As String
'                        Dim strFolderActual As String
'
'                        Open strArchivo For Output As #lngFile
'
'                            Print #lngFile, "LASERFICHE IMPORT LIST"
'                            Print #lngFile, "FOLDERS(Liquidacion Previa)"
'                            Print #lngFile, "FOLDERS(" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ")"
'                            strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                            strFolderActual = strFolderPadre
'                            Print #lngFile, "FOLDERS(" & strFolderActual & ")"
'                            Do While Not rsSQL.EOF
'                                strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                                If strFolderPadre <> strFolderActual Then
'                                    Print #lngFile, "FOLDERS()"
'                                    strFolderPadre = rsSQL!ds_periodo & "_" & rsSQL!ds_año & "_" & rsSQL!DS_batch
'                                    strFolderActual = strFolderPadre
'                                    Print #lngFile, "FOLDERS(" & strFolderActual & ")"
'                                End If
'
'                                Print #lngFile, "STARTFIELDS"
'                                Print #lngFile, strPlantillaLF
'                                Print #lngFile, ("Liquidacion Previa")
'                                Print #lngFile, ("" & rsSQL!ID_liq)
'                                Print #lngFile, ("" & rsSQL!ds_nroliq)
'                                Print #lngFile, Format(Year(Date), "0000")
'
'                                Print #lngFile, (rsSQL!ds_periodo)
'
'                                'Print #lngFile, (rsSQL!priapell) & " " & (rsSQL!segapell) & " " & (rsSQL!nombres)   '''APELLIDOS Y NOMBRE  80 CARACTERES
'                                Print #lngFile, ("Vigente")
'                                Print #lngFile, "ENDFIELDS"
'
'                                Print #lngFile, "DOCUMENT(" & "LP_" & Format(rsSQL!ID_liq, "0000000000") & "_" & ("" & rsSQL!ds_nroliq) & ")"
'                                Print #lngFile, "STARTLIST"
'                                Print #lngFile, rsSQL!Orig_file
'                                Print #lngFile, "ENDLIST"
'                                rsSQL.MoveNext
'                            Loop
'                        Close #lngFile
'
'                    End If
'                    If rsSQL.State = 1 Then rsSQL.Close
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''' LASER FISH LST LIQUIDACION PREVIA'''''''''''''''''''''''''''''''''''''''''''
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

        
'*------------------------------------------------------------------------------------------------*
''''''''''''''''''''''''''''''''''''ARCHIVO CARGO'''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If rsSQL.State = 1 Then rsSQL.Close
'If rsFormulario.State = 1 Then rsFormulario.Close
'rsSQL.Source = "Select * from vw_cargo where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch"
rsSQL.Source = "Select * from vw_cargo where ds_batch='" & cmbBatch.Text & "' order by csid"
rsSQL.Open
If rsSQL.RecordCount > 0 Then

'Borrar_tempoCargo
Borrar_Tempo_Cargo_Sql
StrArchivo = strDirectorio & "\Archivo_Cargo_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     
     If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "\Archivo_Cargo_" & strZonaActual & "_" & cmbBatch.Text & ".XLS"
     
     lngFile = FreeFile
     'Open strArchivo For Output As #lngFile
        
     'Print #lngFile, "NROLIQ"; vbTab; "LOTE"; vbTab; "PAGINA"; vbTab; "RUTA"; vbTab; "NRORECLAMO"
     Do While Not rsSQL.EOF
              StrNrolircargo = ("" & rsSQL!ds_nroliq)
              StrLoteCargo = ("" & rsSQL!DS_batch)
              StrRutaCargo = ("" & rsSQL!Orig_file)
              
              If rsFormulario.State = 1 Then rsFormulario.Close
              rsFormulario.Source = "select * from tb_formularios where nroliqui='" & StrNrolircargo & "'"
              rsFormulario.Open
              If rsFormulario.RecordCount > 0 Then
              StrRucArchivo = ("" & rsFormulario!campo7)
              Else
              StrRucArchivo = ""
              End If
              
              If IIf(IsNull(rsSQL.Fields("ds_barra")), "", rsSQL.Fields("ds_barra")) = "" Then
              StrPagCargo = Mid((rsSQL!ds_pag), 6, 3)
              Else
              StrPagCargo = Mid(rsSQL!ds_BARRA, 10, 3)
              End If
              
              
              'Grabar_TemporalCargo Trim$(StrNrolircargo), Trim$(StrLoteCargo), Trim$(StrPagCargo), Trim$(StrRutaCargo), "", Trim$(StrRucArchivo)
              
              strCadenaSQL = "EXEC sp_InsertarTempo_Cargo_Envio '" & Trim$(StrNrolircargo) & "','" & Trim$(StrLoteCargo) & "','" & Trim$(StrPagCargo) & "','" & Trim$(StrRutaCargo) & "','" & Trim$(StrRucArchivo) & "'"
              cnServer.Execute strCadenaSQL
               
              
              
              'Print #lngFile, Trim$(StrNrolircargo); vbTab; Trim$(StrLoteCargo); vbTab; Trim$(StrPagCargo); vbTab; Trim$(StrRutaCargo); vbTab; ""
              
      '        strLinea = ""
          
     rsSQL.MoveNext
     Loop
     'Close #lngFile
     
End If
      
        strsqlp = "Select NROLIQ,LOTE,PAGINA,RUTA,NRORECLAMO,RUC from Tempoca order by id_tempo"
        If Exportar_ADO_Excel(strsqlp, StrArchivo) Then
        End If
      
      
'      crCargo.DataFiles(0) = App.Path & "\TEMPOca.DBF"
'      crCargo.ReportFileName = App.Path & "\rcargo.rpt"
'      crCargo.Destination = crptToFile
'      crCargo.PrintFileName = strArchivo
'      crCargo.PrintFileType = crptExcel50
'      crCargo.Action = 1
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If rsSQL.State = 1 Then rsSQL.Close
    
    'rsSQL.Source = "Select * from vw_afiliacionesRechazadas where ds_batch='" & cmbBatch.Text & "' order by id_hojapedido"
    rsSQL.Source = "Select * from vw_liquidacionesRechazadas where ds_batch='" & cmbBatch.Text & "' order by id_hojapedido"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        intPedidosRechazados = rsSQL.RecordCount
        
        StrArchivo = strDirectorio & "\Errores_Liquidacion_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        
        lstMensajes.AddItem "Errores_Liquidacion_" & ("" & rsSQL!ds_año) & "_" & cmbBatch.Text & ".txt"
        rsSQL.MoveFirst
        
        lngFile = FreeFile
        
        strErrZona = ""
        strErrIDHojaPedido = ""
        strErrAsesora = ""
        strErrRechazo = ""
        strErrValor = ""
        
        
        
        Open StrArchivo For Output As #lngFile
            strErrZona = "AÑO"
            strErrIDHojaPedido = "ID_lIQUIDACION"
            'strErrAsesora = "ASESORA"
            strErrRechazo = "MOTIVO ERROR"
            'strErrValor = "VALOR DEL CAMPO"
            
            'strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrAsesora & "|" & strErrRechazo '& "|" & strErrValor
            strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo
            Print #lngFile, strLinea

            Do While Not rsSQL.EOF
                strErrZona = rsSQL!ds_año
                strErrIDHojaPedido = Mid(Right(rsSQL!Orig_file, 8), 1, 4) ' rsSQL!id_hojapedido
                'strErrAsesora = rsSQL!ds_asesora
                strErrRechazo = rsSQL!ds_rechazo
                'strErrValor = rsSQL!ds_valor
                
                'strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrAsesora & "|" & strErrRechazo '& "|" & strErrValor
                strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo  '& "|" & strErrValor
                Print #lngFile, strLinea
                rsSQL.MoveNext
            Loop
        Close #lngFile
        
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    
'*------------------------------------------------------------------------------------------------*
    
                        If rsSQL.State = 1 Then rsSQL.Close
                        rsSQL.Source = "Select * from vw_liqVERIFICAR where ds_batch='" & cmbBatch.Text & "'"
                        rsSQL.Open
                        If rsSQL.RecordCount > 0 Then
        
                            StrArchivo = strDirectorio & "\Errores_Liquidacion_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
                            If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        
                            lstMensajes.AddItem "Errores_Liquidacion_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
                            rsSQL.MoveFirst
        
                            lngFile = FreeFile
        
                            strErrZona = ""
                            strErrIDHojaPedido = ""
                            strErrAsesora = ""
                            strErrRechazo = ""
                            strErrValor = ""
        
        
        
                            Open StrArchivo For Output As #lngFile
                                strErrZona = "AÑO"
                                strErrIDHojaPedido = "ID_LIQUIDACION"
                                'strErrAsesora = "ASESORA"
                                strErrRechazo = "MOTIVO ERROR"
                                'strErrProducto = "POSICION"
                                strErrValor = "DATO"
        
                                'strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" &  strErrRechazo & "|" & strErrValor '& "|" & strErrProducto & "|" & strErrValor
                                strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo & "|" & strErrValor  '& "|"  & strErrValor
                                Print #lngFile, strLinea
        
                                Do While Not rsSQL.EOF
                                    strErrZona = rsSQL!ds_año
                                    strErrIDHojaPedido = Mid(Right(rsSQL!Orig_file, 8), 1, 4) 'rsSQL!id_hojapedido
                                    'strErrAsesora = rsSQL!ds_asesora
                                    strErrRechazo = rsSQL!ds_rechazo
                                    'strErrProducto = rsSQL!nu_posicion
                                    strErrValor = rsSQL!ds_valor
        
                                    strLinea = strErrZona & "|" & strErrIDHojaPedido & "|" & strErrRechazo & "|" & strErrValor
                                    Print #lngFile, strLinea
                                    rsSQL.MoveNext
                                Loop
                            Close #lngFile
        
                        End If
                        If rsSQL.State = 1 Then rsSQL.Close
    
'*------------------------------------------------------------------------------------------------*
    
    'If intPedidosValidados > 0 Then
    
    
        StrArchivo = strDirectorio & "\Resumen_" & cmbBatch.Text & ".txt"
        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        lstMensajes.AddItem "Resumen_" & strZona & "_" & cmbBatch.Text & ".txt"
        
        
        If rsZonas.State = 1 Then rsZonas.Close
        'rsZonas.Source = "select distinct ds_zona,ds_campaña,ds_batch from pedidos where ds_batch='" & cmbBatch.Text & "' order by 1,2"
         rsZonas.Source = "select distinct ds_año,ds_periodo,ds_batch from tb_liquidacion where ds_batch='" & cmbBatch.Text & "' order by 1,2"
        rsZonas.Open
        If rsZonas.RecordCount > 0 Then
            Open StrArchivo For Output As #lngFile
            strLinea = "Archivos resumen"
            Print #lngFile, strLinea
            Print #lngFile, "***************************************************************************"
            Print #lngFile, " "
            Do While Not rsZonas.EOF()
                strZona = "" & rsZonas!ds_año
                
                intPedidosValidados = 0
                intPedidosRechazados = 0
                
                If rsSQL.State = 1 Then rsSQL.Close
                'rsSQL.Source = "Select * from vw_PedidosAxis where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & " ' and ds_batch='" & rsZonas!DS_batch & "'"
                'rsSQL.Source = "Select distinct id  from vw_afiliacionAxis where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & " ' and ds_batch='" & rsZonas!DS_batch & "'"
                 rsSQL.Source = "Select distinct id_liq  from vw_Liquidacionentrega where ds_año='" & rsZonas!ds_año & "' and ds_periodo='" & rsZonas!ds_periodo & " ' and ds_batch='" & rsZonas!DS_batch & "'"
                rsSQL.Open
                If rsSQL.RecordCount > 0 Then
                    intPedidosValidados = rsSQL.RecordCount
                End If
                
                If rsSQL.State = 1 Then rsSQL.Close
                'rsSQL.Source = "Select * from vw_PedidosRechazados where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & "' and ds_batch='" & rsZonas!DS_batch & "'"
                'rsSQL.Source = "Select distinct id_hojapedido from vw_afiliacionesRechazadas where ds_zona='" & rsZonas!ds_zona & "' and ds_campaña='" & rsZonas!ds_campaña & "' and ds_batch='" & rsZonas!DS_batch & "'"
                rsSQL.Source = "Select distinct id_hojapedido from vw_liquidacionesRechazadas where ds_año='" & rsZonas!ds_año & "' and ds_periodo='" & rsZonas!ds_periodo & "' and ds_batch='" & rsZonas!DS_batch & "'"
                rsSQL.Open
                If rsSQL.RecordCount > 0 Then
                    intPedidosRechazados = rsSQL.RecordCount
                End If
            
                Print #lngFile, "---------------------------------------------------------------------------"
                strLinea = "Año     :" & "" & rsZonas!ds_año
                Print #lngFile, strLinea
                strLinea = "Periodo  :" & "" & rsZonas!ds_periodo
                Print #lngFile, strLinea
                strLinea = "Batch    :" & "" & rsZonas!DS_batch
                Print #lngFile, strLinea
                Print #lngFile, " "
                'strLinea = "Hojas de Pedido procesados :" & intPedidosValidados + intPedidosRechazados
                strLinea = "Formatos de Liquidación procesados :" & intPedidosValidados + intPedidosRechazados
                Print #lngFile, strLinea
                'strLinea = "Hojas de Pedido validados  :" & intPedidosValidados
                strLinea = "Formatos de Liquidación validos  :" & intPedidosValidados
                Print #lngFile, strLinea
                'strLinea = "Hojas de Pedido rechazados :" & intPedidosRechazados
                strLinea = "Formatos de Liquidación rechazados :" & intPedidosRechazados
                Print #lngFile, strLinea
                Print #lngFile, " "
                Print #lngFile, "---------------------------------------------------------------------------"
                rsZonas.MoveNext
            Loop
            Close #lngFile
        End If
    
'    rsZonas.Source = "select id_cl from cl where id_cl='" & strCL & "'"
'    .Openm
    If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "Select * from vw_lstliquidacion where ds_batch='" & cmbBatch.Text & "' order by ds_perioDO,ds_año,ds_batch"

        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
        StrArchivo = strDirectorio & "\Cargo_" & strZonaActual & "_" & cmbBatch.Text & ".txt"
        If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
        
        
        End If
        
'    End If
    
    If MsgBox("Se han exportado los Archivo de texto, desea Generar los backup's y limpiar las tablas ?", vbYesNo + vbQuestion, "Finalizado") = vbYes Then
        strEjecutar = "drop table [" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "_Año_" & strZona & "_Envio_" & cmbBatch.Text & "]"
        cnServer.Execute strEjecutar
        strEjecutar = "select * into [" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "_Año_" & strZona & "_Envio_" & cmbBatch.Text & "] from tb_liquidacion where ds_batch='" & cmbBatch.Text & "'"
        cnServer.Execute strEjecutar
        strEjecutar = "delete from tb_liquidacion where ds_batch='" & cmbBatch.Text & "'"
        cnServer.Execute strEjecutar
    End If
    MsgBox "Fin del proceso", vbInformation, "Generaciòn del archivo de retorno"
    picAvance.Picture = LoadPicture()
    picAvance.BorderStyle = 1
    cmdProcesar.Enabled = True
    cmbBatch.Enabled = True
    Screen.MousePointer = 0
    ListarBatch
    Exit Sub
Err:
    If Err.Number = -2147217865 Then
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, App.EXEName
        Me.MousePointer = 0
        Screen.MousePointer = 0
        Unload Me
    End If
End Sub

Public Sub lst_otros()
If rsSQL.State = 1 Then rsSQL.Close
rsSQL.Source = "Select * from  horizontal "

rsSQL.Open
Dim StrTramite As String
If rsSQL.RecordCount > 0 Then

   StrArchivo = strDirectorioLST & "\Lst_" & Format(Now, "YYYY") & "_General.lst"
'lstMensajes.AddItem "Se genero el archivo LST: " & StrArchivo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  rsSQL.MoveLast
  rsSQL.MoveFirst
'   xMin = "0000" 'Select_Imagen_Min(XBATCH)
'   xMax = Select_Imagen_Max(XBATCH)
'strPeriodo = rsSQL!ds_periodo
'strAnno = rsSQL!ds_año
If rsSQL.RecordCount > 0 Then
                intPedidosRechazados = rsSQL.RecordCount
                        

                StrArchivo = strDirectorioLST & "\Lst_" & Format(Now, "YYYY") & "_" & "General.lst"
                If Dir(StrArchivo, vbArchive) <> "" Then Kill StrArchivo
                'lstMensajes.AddItem "Lst_" & ("" & strAnno) & "_" & ("" &rsSql!LOTE) & ".lst"
                lstMensajes.AddItem "Se genero el archivo LST: " & StrArchivo
                'rsSQL.MoveFirst
                lngFile = FreeFile
                
                Open StrArchivo For Output As #lngFile

                    Print #lngFile, "LASERFICHE IMPORT LIST"
                    Print #lngFile, "FOLDERS(FICHA DE DATOS)"
                    Print #lngFile, "FOLDERS(" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ")"
                    'strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & rsSQL!DS_batch)
                    'strFolderActual = strFolderPadre
                    'Print #lngFile, "FOLDERS(" & strFolderActual & ")"
                Do While Not rsSQL.EOF
                    If strFolderPadre <> strFolderActual Then
                        Print #lngFile, "FOLDERS()"
                        strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & rsSQL!DS_batch)
                        strFolderActual = strFolderPadre
                        Print #lngFile, "FOLDERS(" & strFolderActual & ")"
                    End If
                    
'                    If Trim$(("" & rsSQL!ds_numerotra)) = "" Then
'                    StrTramite = "0000001"
'                    Else
                    StrTramite = ("" & rsSQL!CóDIGO)
'                    End If
                    
                    'strFolderPadre = strPeriodo & "_" & strAnno & "_" & ("" & rsSQL!DS_batch)
                    Print #lngFile, "DOCUMENT(" & "" & StrTramite & ")"
                    StrReclamo = "" & StrTramite
                    Print #lngFile, "STARTFIELDS"
                    Print #lngFile, ("FICHA DE DATOS")
                    Print #lngFile, ("" & rsSQL![CóDIGO])
                    Print #lngFile, ("" & rsSQL![APELLIDO PATERNO]) & " " & ("" & rsSQL![APELLIDO MATERNO]) & " " & ("" & rsSQL![Nombres])
                    Print #lngFile, "ENDFIELDS"
                    Print #lngFile, "STARTLIST"
                    Print #lngFile, ("" & rsSQL!csid)
                    'RutaArchivo = Mid(("" & rsSQL!Orig_file), 1, Len(rsSQL!Orig_file) - 25)
                    Print #lngFile, "ENDLIST"
                rsSQL.MoveNext
                Loop
                    
  End If
                
                Close #lngFile
  
End If


End Sub

VERSION 5.00
Begin VB.Form Icr_FrmControl 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Control Entregable-Protocolos"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   Icon            =   "Icr_FrmControl.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
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
         ItemData        =   "Icr_FrmControl.frx":030A
         Left            =   120
         List            =   "Icr_FrmControl.frx":030C
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2640
         Width           =   6495
      End
      Begin VB.CommandButton cmdprocesar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Procesar- Control de Envios (Protocolos)"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   6495
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         Begin VB.CheckBox chktodos 
            BackColor       =   &H80000005&
            Caption         =   "Todos Los formatos"
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.PictureBox picAvance 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            DrawMode        =   7  'Invert
            FillColor       =   &H00400000&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   840
            ScaleHeight     =   270
            ScaleWidth      =   4785
            TabIndex        =   4
            Top             =   840
            Width           =   4815
         End
         Begin VB.ComboBox cmbBatch 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   3255
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
            Left            =   840
            TabIndex        =   5
            Top             =   285
            Width           =   1410
         End
      End
   End
End
Attribute VB_Name = "Icr_FrmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprocesar_Click()
''Procesar el control de calidad con los protocolos establecidos por Horizonte
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

    rsZonas.CursorType = adOpenDynamic
    rsZonas.LockType = adLockBatchOptimistic
    rsZonas.ActiveConnection = cnServer
    
    rsFormulario.CursorType = adOpenDynamic
    rsFormulario.LockType = adLockBatchOptimistic
    rsFormulario.ActiveConnection = cnServer
    
    Screen.MousePointer = 11
    
    lstMensajes.Clear
    
    If Trim(cmbBatch.Text) = "" Then
        MsgBox "Seleccione un número de batch", vbCritical, "Generaciòn de archivos de retorno"
       Screen.MousePointer = 0
        Exit Sub
    End If
    
    'cmbbatch.Text = Padln(Trim(mskEnvio.Text), 4, "0")
    
    cmdprocesar.Enabled = False
    cmbBatch.Enabled = False
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "delete from tb_verifica"
    rsSQL.Open
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from tb_Parametros"
    rsSQL.Open
    
    strCampañaActual = "" & rsSQL!ds_periodo
    strZonaActual = "" & rsSQL!ds_año
    strDirectorio = "" & rsSQL!ds_ftpaxis
    intProductos = 0 & rsSQL!nu_cl
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    
'    lstMensajes.AddItem "DIRECTORIO DE ENVIO"
'    lstMensajes.AddItem UCase(strDirectorio)
'    lstMensajes.AddItem ""
'    lstMensajes.AddItem "ARCHIVOS GENERADOS"
    
'*------------------------------------------------------------------------------------------------*
    strZona = "0000"
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from vw_liquidacionentrega where ds_batch='" & cmbBatch.Text & "'"
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
                ReDim Xcuspp(1 To intProductos) As String
                ReDim Xcodigo(1 To intProductos) As String
                ReDim xfecha1(1 To intProductos) As String
                ReDim xfecha2(1 To intProductos) As String
                ReDim xcomen(1 To intProductos) As String
                ReDim xFere(1 To intProductos) As String
                ReDim xmonto(1 To intProductos) As String
                ReDim xmonto(1 To intProductos) As String
                ReDim XPla(1 To intProductos) As String
                      
                      If rsFormulario.State = 1 Then rsFormulario.Close
'                      StrDato1 = Trim$(rsSQL.Fields("ds_nroliq"))
'                      StrDato2 = Trim$(Mid(IIf(IsNull(rsSQL.Fields("ds_barra")), "", rsSQL.Fields("ds_barra")), 11, 3))
'                      If StrDato2 = "" Then
'                      StrDato2 = Trim$(Mid(IIf(IsNull(rsSQL.Fields("ds_pag")), "", rsSQL.Fields("ds_pag")), 6, 3))
'                      Else
'
'                      End If
                      
                      
                      
                          If StrBarra_Chek = "0" Then
                          StrDato1 = "HO" & Trim$(Mid(rsSQL.Fields("ds_barra"), 1, 9))
                          StrDato2 = Format(Trim$(Mid(rsSQL.Fields("ds_barra"), 10, 5)), "00000")
                          Else
                          StrDato1 = Trim$(rsSQL.Fields("ds_nroliq"))
                          StrDato2 = Format(Trim$(rsSQL.Fields("ds_pag")), , "00000")
                          End If
                      StrDato = StrDato1 & StrDato2
                      If rsFormulario.State = 1 Then rsFormulario.Close
                      rsFormulario.Source = "select * from tb_formularios where nroliqpag='" & StrDato & "'"
                      rsFormulario.Open
                
                If rsFormulario.RecordCount > 0 Then
                        For intIndice = 1 To intProductos
                                  
                                    StrCampo = "cuspp" & intIndice
                                    strProducto = Trim(("" & rsFormulario.Fields(StrCampo)))
                                    If strProducto = "" Then
                                      Exit For
                                    Else
                                        Xcuspp(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
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
                                        XPla(intIndice) = IIf(IsNull(strProducto), "", strProducto)
                                        
                                        If Xcodigo(intIndice) = "" Then
                                        Xcodigo(intIndice) = Codigo_Nulo(xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice))
                                        End If
                                        
                                        xFere(intIndice) = Mostrar_Fecha_Ref(Xcodigo(intIndice), xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice), XPla(intIndice))
                                        xcomen(intIndice) = Mostrar_Comentario(Xcodigo(intIndice), xfecha1(intIndice), xfecha2(intIndice), xmonto(intIndice), XPla(intIndice))
                                        
                                        If Xcodigo(intIndice) <> "" Or xfecha1(intIndice) <> "" Or xfecha2(intIndice) <> "" Or xmonto(intIndice) <> "" Then
                                        strCadenaSQL = "EXEC sp_Insertarverifica '" & StrDato & "','" & cmbBatch.Text & "','" & ("" & rsFormulario!Ruc) & "','" & ("" & Replace(rsFormulario!razon_social, "'", "''")) & "','" & xcomen(intIndice) & "','" & Xcodigo(intIndice) & "','" & Xcuspp(intIndice) & "','" & ("" & rsFormulario!Periodo) & "','" & xFere(intIndice) & "','" & xfecha1(intIndice) & "','" & xfecha2(intIndice) & "','" & xmonto(intIndice) & "','" & XPla(intIndice) & "','" & StrDupli & "'," & Trim(("" & rsSQL.Fields("id_liq"))) & ",'',''," & intIndice & ""
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
       ' Close #lngFile
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    If rsFormulario.State = 1 Then rsFormulario.Close
    
    picAvance.Picture = LoadPicture()
    picAvance.BorderStyle = 1
    cmdprocesar.Enabled = True
    cmbBatch.Enabled = True
    Screen.MousePointer = 0
    ''''''''''''''''''''''''''''''''''Verificar Control de Calidad de Enviados''''''''''''''''''''
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from Vw_LiquidacionVerificar where ds_batch='" & cmbBatch.Text & "'"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
      Unload Me
       Icr_FrmVerificar.Show vbModal
    Else
       MsgBox "No Existen Datos a verificar", vbCritical, "Revisar Datos de Envio "
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Err:
    If Err.Number = -2147217865 Then
        Resume Next
    Else
       If Len(Err.Description) > 0 Then
        MsgBox Err.Description, vbCritical, App.EXEName
        Me.MousePointer = 0
        Screen.MousePointer = 0
        Unload Me
       End If
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Form_Load()
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    ListarBatch
End Sub
Sub ListarBatch()
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Not (rsSQL Is Nothing) Then
        If rsSQL.State = 1 Then rsSQL.Close
        Set rsSQL = Nothing
    End If
End Sub

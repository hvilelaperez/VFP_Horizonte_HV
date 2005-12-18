VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form Icr_FrmRechazos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rechazos Liquidación Previa"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   Icon            =   "Icr_Rechazos.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10500
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   10215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   8400
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   9570
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   435
         Left            =   7440
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   9570
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   6480
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   9570
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdSiguiente 
         Height          =   630
         Left            =   13890
         Picture         =   "Icr_Rechazos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Reverso"
         Top             =   4530
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   630
         Left            =   13860
         Picture         =   "Icr_Rechazos.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anverso"
         Top             =   3810
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Actualizar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1815
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9720
         TabIndex        =   9
         Text            =   "Text4"
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdEof 
         Caption         =   "&Primero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   9570
         Width           =   975
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "An&terior"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1335
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   9570
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   " &Siguiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4185
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   9570
         Width           =   975
      End
      Begin VB.CommandButton cmdBOF 
         Caption         =   "&Ultimo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5205
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   9570
         Width           =   975
      End
      Begin VB.CommandButton cmdNormal 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_Rechazos.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Normal"
         Top             =   2295
         Width           =   630
      End
      Begin VB.CommandButton cmdZommout 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_Rechazos.frx":1BA8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "ZoomOut"
         Top             =   855
         Width           =   630
      End
      Begin VB.CommandButton cmdZommin 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_Rechazos.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "ZoomIn"
         Top             =   1575
         Width           =   630
      End
      Begin VB.CommandButton cmdRotar 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_Rechazos.frx":40BC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Rotar"
         Top             =   3015
         Width           =   630
      End
      Begin ImgeditLibCtl.ImgEdit imgVisor 
         Height          =   8535
         Left            =   240
         TabIndex        =   16
         Top             =   810
         Width           =   13335
         _Version        =   131076
         _ExtentX        =   23521
         _ExtentY        =   15055
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   2
         UndoBufferSize  =   462500352
         OcrZoneVisibility=   -4004
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3624
         sReserved2      =   -3624
         lReserved1      =   2010226912
         lReserved2      =   2010226912
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin VB.Label lblCampaña 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Dato :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8760
         TabIndex        =   19
         Top             =   315
         Width           =   690
      End
      Begin VB.Label lblError 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   8415
      End
      Begin VB.Shape Shape10 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   8580
         Left            =   13680
         Shape           =   4  'Rounded Rectangle
         Top             =   810
         Width           =   975
      End
      Begin VB.Label lblImagen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2355
         TabIndex        =   17
         Top             =   9585
         Width           =   1800
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   9450
         Width           =   14415
      End
   End
End
Attribute VB_Name = "Icr_FrmRechazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ZoomMax = 6554
Const ZoomMin = 2

Dim rsSQL As New Recordset
Dim iSeleccion As Integer
Dim strDirectorio As String
Dim strImagen As String
Dim strActual As String
Dim strSiguiente As String
Dim strArchivoImagen As String
Private Sub cmdAnterior_Click()
    If (iSeleccion = 1) Then
        imgVisor.Image = strActual
        imgVisor.Display

    End If
End Sub

Private Sub cmdCancelar_Click()
Dim strEjecutar As String
Dim lngPosition As Long
On Local Error GoTo Err:
    If rsSQL.RecordCount > 0 Then
        Select Case rsSQL!ds_rechazo
        
           Case "COD.BARRA.NOEXISTE"
                If ExisteBarra(Trim("" & txtdato.Text)) = True Then
                   If Verificar_Pag(Format(Mid(txtdato.Text, 10, 5), "00000")) Then
                    ActualizarDatos
                   Else
                     MsgBox "El Formato de pagina dentro del codigo de barra no es adecuado  5 caracteres ", vbCritical, "Liquidación"
                   End If
                    'If ExisteBarra("" & txtDato.Text) = False Then
                    '    MsgBox "El codigo de Barra no existe en la tabla de LIQUIDACION IMPRESAS", vbCritical, "Liquidación"
                    'Else
                    'ActualizarDatos
                    'End If
                Else
                    MsgBox "El Formato de barra no es adecuado  9 caracteres ", vbCritical, "Liquidación"
                End If
            
           Case "PAG.NULL", "PAG.ERROR"
                If Verificar_Pag(Trim("" & Mid(txtdato.Text, 10, 5))) = True Then
                    ActualizarDatos
                    'If ExisteBarra("" & txtDato.Text) = False Then
                    '    MsgBox "El codigo de Barra no existe en la tabla de LIQUIDACION IMPRESAS", vbCritical, "Liquidación"
                    'Else
                    'ActualizarDatos
                    'End If
                Else
                    MsgBox "El nro de Pagina Errada  Ejemplo [PAG.:001]", vbCritical, "Liquidación"
                End If
            
           Case "COD.BARRA.NULL", "COD.BARRA.NOEXISTE"
                If Trim("" & txtdato.Text) <> "" Then
                    If ExisteBarra("" & txtdato.Text) = False Then
                        MsgBox "El codigo de Barra no existe en la tabla de LIQUIDACION IMPRESAS", vbCritical, "Liquidación"
                    Else
                    ActualizarDatos
                    End If
                End If
                
            
           Case "NRO.LIQUIDACION.NOEXISTE", "NRO.LIQUIDACION.NULL", "NRO.LIQUIDACION.ERROR"
                If Trim("" & txtdato.Text) <> "" Then
                    If ExisteNroliq("" & txtdato.Text) = False Then
                        MsgBox "El Nro de Liquidación no existe en la tabla de LIQUIDACION IMPRESAS", vbCritical, "Liquidación"
                    Else
                    ActualizarDatos
                    End If
                End If
                
                
           Case "CODIGO.NOEXISTE"
                If Trim("" & txtdato.Text) <> "" Then
                    If Verificar_Codigo("" & txtdato.Text) = False Then
                        MsgBox "El Codigo no es valido", vbCritical, "Liquidación"
                    Else
                    ActualizarDatos
                    End If
                Else
                    ActualizarDatos
                End If
                
            
           Case "FECHA.ERROR.INI"
                If Trim("" & txtdato.Text) <> "" Then
                    If Verificar_Fecha("" & txtdato.Text) = False Then
                        MsgBox "La fecha ingresada no es valida", vbCritical, "Liquidación"
                    Else
                    ActualizarDatos
                    End If
                Else
                    ActualizarDatos
                End If
                
            
           Case "FECHA.ERROR.CESE"
                If Trim("" & txtdato.Text) <> "" Then
                    If Verificar_Fecha("" & txtdato.Text) = False Then
                        MsgBox "La fecha ingresada no es valida", vbCritical, "Liquidación"
                    Else
                    ActualizarDatos
                    End If
                Else
                    ActualizarDatos
                End If
                
            
            Case "REMUNERACION.ERROR"
                If Trim("" & txtdato.Text) <> "" Then
                    If IsNumeric("" & txtdato.Text) = False Then
                        MsgBox "Los datos nos son numericos", vbCritical, "Liquidación"
                    Else
                    ActualizarDatos
                    End If
                Else
                    ActualizarDatos
                End If
            
            Case "NRO.PLANILLA.ERROR"
                If Trim("" & txtdato.Text) <> "" Then
                    If IsNumeric("" & txtdato.Text) = False Then
                        MsgBox "Los datos nos son numericos", vbCritical, "Liquidación"
                    Else
                    ActualizarDatos
                    End If
                Else
                    ActualizarDatos
                End If
                
                
'            Case "NRO.LIQUIDACION.NOEXISTE"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If ExisteNroliq("" & txtDato.Text) = False Then
'                        MsgBox "El Nro de Liquidación no existe en la tabla de LIQUIDACION IMPRESAS", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
        End Select
        
    End If
    
        If rsSQL.RecordCount > 0 Then
        rsSQL.MoveNext
        Call Reposition
        End If
    
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Error"
End Sub
Sub ActualizarDatosAsesora()
    txtdato.Text = UCase("" & Trim(txtdato.Text))
    strEjecutar = "Update PEDIDOS set " & rsSQL!ds_campo & "='" & txtdato.Text & "', codbar='" & txtdato.Text & "' where id_hojapedido = " & rsSQL!id_hojapedido
    cnServer.Execute strEjecutar
    strEjecutar = "Update rechazos set ds_valor='" & txtdato.Text & "' where id_rechazo = " & rsSQL!id_rechazo
    cnServer.Execute strEjecutar
    
    lngPosition = rsSQL.AbsolutePosition
    rsSQL.Requery
    rsSQL.AbsolutePosition = lngPosition
    Reposition
End Sub
Sub ActualizarDatos()
    txtdato.Text = UCase("" & Trim(txtdato.Text))
    'strEjecutar = "Update AFILIACION set " & rsSQL!ds_campo & "='" & txtDato.Text & "'  where id = " & rsSQL!id_hojapedido
    strEjecutar = "Update tb_liquidacion set " & rsSQL!ds_campo & "='" & txtdato.Text & "'  where id_liq = " & rsSQL!id_hojapedido
    cnServer.Execute strEjecutar
    strEjecutar = "Update tb_rechazos set ds_valor='" & txtdato.Text & "' where id_rechazo = " & rsSQL!id_rechazo
    cnServer.Execute strEjecutar
    
    lngPosition = rsSQL.AbsolutePosition
    rsSQL.Requery
    rsSQL.AbsolutePosition = lngPosition
    Reposition
End Sub

Private Sub cmdNormal_Click()
On Error GoTo Err:
    imgVisor.FitTo 1
Err:
End Sub

Private Sub cmdRotar_Click()
On Local Error GoTo Err:
    imgVisor.FitTo 1
    imgVisor.RotateRight
Err:
End Sub

Private Sub cmdSiguiente_Click()
    If (iSeleccion = 1) Then
        imgVisor.Image = strSiguiente
        imgVisor.Display
    End If
End Sub

Private Sub cmdZommin_Click()
Dim zoomval As Single
On Error GoTo Err:
    zoomval = imgVisor.Zoom
    zoomval = zoomval / 1.05
    If zoomval >= ZoomMin Then
        imgVisor.Zoom = zoomval
        Text1.Text = zoomval
        imgVisor.Refresh
    Else
        MsgBox "Se llego a su minima expresión", vbInformation, "Disminuir tamaño de la imagen"
    End If
    Exit Sub
Err:
End Sub

Private Sub cmdZommout_Click()
Dim zoomval As Single
On Local Error GoTo Err:
    zoomval = imgVisor.Zoom
    zoomval = zoomval * 1.05
    If zoomval < ZoomMax Then
        imgVisor.Zoom = zoomval
        Text1.Text = zoomval
        imgVisor.Refresh
    Else
        MsgBox "Se llego al maximo tamaño", vbInformation, "Aumentar tamaño de la imagen"
    End If
Err:
End Sub

Private Sub Form_Load()
  
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    rsSQL.Source = "Select * from vw_Rechazos order by id_hojapedido"
    rsSQL.Open
    
    If rsSQL.RecordCount > 0 Then
        Reposition
    Else
        MsgBox "No existen rechazos", vbCritical, "Revisar errores "
        Unload Icr_FrmRechazos
    End If

End Sub
Sub Verimagen(strImagen As String)
Dim lngZomm As Long
Dim lngScrollY As Long
Dim lngScrollX As Long
Dim lngPosicion As Long
Dim lngPosicionFecha As Long

Dim lngRex As Long
Dim lngRey As Long
Dim lngReAncho As Long
Dim lngReAlto As Long

    If FileLen(strImagen) > 0 Then
        
        imgVisor.Image = strImagen
        
        lngPosicion = 0 '2765
        lngPosicionFecha = 950
        Select Case rsSQL!ds_rechazo
                 
                 
           Case "COD.BARRA.NOEXISTE"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 115
                    lngScrollX = 496
                    lngScrollY = 1731
                    lngRex = CInt(lngScrollX)
                    lngRey = CInt(lngScrollY)
                    lngReAncho = 5000
                    lngReAlto = 1000
                 
                 
           Case "PAG.NULL"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 115
                    lngScrollX = 496
                    lngScrollY = 1731
                    lngRex = CInt(lngScrollX) + 1500
                    lngRey = CInt(lngScrollY) * 14
                    lngReAncho = 5000
                    lngReAlto = 1000
                    'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
           Case "NRO.LIQUIDACION.NOEXISTE"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 344
                    lngScrollX = 4032
                    lngScrollY = 246
                    'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
                    lngRex = CInt(lngScrollX) + 1500
                    lngRey = CInt(lngScrollY) * 30
                    lngReAncho = 5000
                    lngReAlto = 1000
           
           Case "PAG.ERROR"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 115
                    lngScrollX = 516
                    lngScrollY = 1731
                    lngRex = 3200 'CInt(lngScrollX)
                    lngRey = 6400 'CInt(lngScrollY)
                    lngReAncho = 7300
                    lngReAlto = 1500
                    
                 'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
           Case "COD.BARRA.NULL"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 220
                    lngScrollX = 1551
                    lngScrollY = 3180
                    
                  'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
           Case "NRO.LIQUIDACION.NULL"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 315
                    lngScrollX = 3605
                    lngScrollY = 142
                    lngRex = CInt(lngScrollX) + 1500
                    lngRey = CInt(lngScrollY) * 30
                    lngReAncho = 5000
                    lngReAlto = 1000
                    'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
           
           Case "NRO.LIQUIDACION.ERROR"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 203
                    lngScrollX = 1005
                    lngScrollY = 15
                    lngRex = CInt(lngScrollX) + 1500
                    lngRey = CInt(lngScrollY) * 30
                    lngReAncho = 5000
                    lngReAlto = 1000
           
           Case "DNI.RUC.MULTIMARCA"
                  lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                    lngZomm = 160
                    lngScrollX = 360
                    lngScrollY = 80
                    'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
           Case "CODIGO.NOEXISTE"
                    lblError.Caption = "ITEM " & rsSQL!nu_posicion & "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                  Select Case rsSQL!nu_posicion
                  
                  Case 1
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 300 '300 'lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                  Case 2
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 336 '300
                  Case 3
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 374 '300
                  Case 4
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 408 '300
                  Case 5
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 442 '300
                  Case 6
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 476 '300
                  Case 7
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 510 '300
                  Case 8
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 544 '300
                  Case 9
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 578 '300
                  Case 10
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 612 '300
                  Case 11
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 646 '300
                  Case 12
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 680 '300
                  Case 13
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 714 '300
                  Case 14
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 748 '300
                  Case 15
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 782 '300
                  Case 16
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 816 '300
                  Case 17
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 1550 '300
                  Case 18
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2100 '300
                  Case 19
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2700 '300
                  Case 20
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 3200 '300
                        
'                        lngScrollX = 0 '1480
'                        lngScrollY = lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
                  End Select
                        lngRex = CInt(lngScrollX)
                        lngRey = CInt(lngScrollY)
                        lngReAncho = 12000
                        lngReAlto = 600
                     
           Case "FECHA.ERROR.INI"
                    lblError.Caption = "ITEM " & rsSQL!nu_posicion & "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
                  Select Case rsSQL!nu_posicion
                  
                  Case 1
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 300 '300 'lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                  Case 2
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 336 '300
                  Case 3
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 374 '300
                  Case 4
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 408 '300
                  Case 5
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 442 '300
                  Case 6
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 476 '300
                  Case 7
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 510 '300
                  Case 8
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 544 '300
                  Case 9
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 578 '300
                  Case 10
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 612 '300
                  Case 11
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 646 '300
                  Case 12
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 680 '300
                  Case 13
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 714 '300
                  Case 14
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 748 '300
                  Case 15
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 782 '300
                  Case 16
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 816 '300
                  Case 17
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 1550 '300
                  Case 18
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2100 '300
                  Case 19
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2700 '300
                  Case 20
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 3200 '300
                        
'                        lngScrollX = 0 '1480
'                        lngScrollY = lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
                  End Select
                        lngRex = CInt(lngScrollX)
                        lngRey = CInt(lngScrollY)
                        lngReAncho = 12000
                        lngReAlto = 600
                  
'                  Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
'                        lngZomm = 294
'                        lngScrollX = 862
'                        lngScrollY = lngPosicionFecha + (160 * rsSQL!nu_posicion)  '372
'                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
'                  End Select
          
           Case "FECHA.ERROR.CESE"
                    lblError.Caption = "ITEM " & rsSQL!nu_posicion & "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
'                  Select Case rsSQL!nu_posicion
'                  Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
'                        lngZomm = 294
'                        lngScrollX = 1616
'                        lngScrollY = lngPosicionFecha + (160 * rsSQL!nu_posicion)  '372
'                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
'                  End Select
                  Select Case rsSQL!nu_posicion
                  
                  Case 1
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 300 '300 'lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                  Case 2
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 336 '300
                  Case 3
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 374 '300
                  Case 4
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 408 '300
                  Case 5
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 442 '300
                  Case 6
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 476 '300
                  Case 7
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 510 '300
                  Case 8
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 544 '300
                  Case 9
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 578 '300
                  Case 10
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 612 '300
                  Case 11
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 646 '300
                  Case 12
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 680 '300
                  Case 13
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 714 '300
                  Case 14
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 748 '300
                  Case 15
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 782 '300
                  Case 16
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 816 '300
                  Case 17
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 1550 '300
                  Case 18
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2100 '300
                  Case 19
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2700 '300
                  Case 20
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 3200 '300
                        
'                        lngScrollX = 0 '1480
'                        lngScrollY = lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
                  End Select
                        lngRex = CInt(lngScrollX)
                        lngRey = CInt(lngScrollY)
                        lngReAncho = 12000
                        lngReAlto = 600
          
          
          Case "REMUNERACION.ERROR"
                    lblError.Caption = "ITEM " & rsSQL!nu_posicion & "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
'                  Select Case rsSQL!nu_posicion
'                  Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
'                        lngZomm = 294
'                        lngScrollX = 2420
'                        lngScrollY = lngPosicionFecha + (160 * rsSQL!nu_posicion)  '372
'                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
'                  End Select

                 Select Case rsSQL!nu_posicion
                  
                  Case 1
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 300 '300 'lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                  Case 2
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 336 '300
                  Case 3
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 374 '300
                  Case 4
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 408 '300
                  Case 5
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 442 '300
                  Case 6
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 476 '300
                  Case 7
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 510 '300
                  Case 8
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 544 '300
                  Case 9
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 578 '300
                  Case 10
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 612 '300
                  Case 11
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 646 '300
                  Case 12
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 680 '300
                  Case 13
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 714 '300
                  Case 14
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 748 '300
                  Case 15
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 782 '300
                  Case 16
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 816 '300
                  Case 17
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 1550 '300
                  Case 18
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2100 '300
                  Case 19
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2700 '300
                  Case 20
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 3200 '300
                        
'                        lngScrollX = 0 '1480
'                        lngScrollY = lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
                  End Select
                        lngRex = CInt(lngScrollX)
                        lngRey = CInt(lngScrollY)
                        lngReAncho = 12000
                        lngReAlto = 600
          
          Case "NRO.PLANILLA.ERROR"
                    lblError.Caption = "ITEM " & rsSQL!nu_posicion & "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                    Me.txtdato.Text = "" & rsSQL!ds_valor
'                  Select Case rsSQL!nu_posicion
'                  Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
'                        lngZomm = 294
'                        lngScrollX = 3379
'                        lngScrollY = lngPosicionFecha + (160 * rsSQL!nu_posicion)  '372
'                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
'                  End Select

                 Select Case rsSQL!nu_posicion
                  
                  Case 1
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 300 '300 'lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                  Case 2
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 336 '300
                  Case 3
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 374 '300
                  Case 4
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 408 '300
                  Case 5
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 442 '300
                  Case 6
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 476 '300
                  Case 7
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 510 '300
                  Case 8
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 544 '300
                  Case 9
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 578 '300
                  Case 10
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 612 '300
                  Case 11
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 646 '300
                  Case 12
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 680 '300
                  Case 13
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 714 '300
                  Case 14
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 748 '300
                  Case 15
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 782 '300
                  Case 16
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 816 '300
                  Case 17
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 1550 '300
                  Case 18
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2100 '300
                  Case 19
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2700 '300
                  Case 20
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 3200 '300
                        
'                        lngScrollX = 0 '1480
'                        lngScrollY = lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
                  End Select
                        lngRex = CInt(lngScrollX)
                        lngRey = CInt(lngScrollY)
                        lngReAncho = 12000
                        lngReAlto = 600
          
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''
           
        End Select
        imgVisor.Zoom = lngZomm
        imgVisor.Display
        imgVisor.ScrollPositionY = lngScrollY
        imgVisor.ScrollPositionX = lngScrollX
        imgVisor.ScrollBars = True
'        Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
         Resaltado1 imgVisor, CInt(lngRex), CInt(lngRey), CInt(lngReAncho), CInt(lngReAlto)
         
    End If
    txtdato.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not (rsSQL Is Nothing) Then
        If rsSQL.State = 1 Then rsSQL.Close
        Set rsSQL = Nothing
    End If
End Sub


Private Sub imgVisor_DblClick()
On Error GoTo Err:
    imgVisor.FitTo 1
Err:
End Sub

Private Sub imgVisor_Scroll()
  Text2.Text = "Y: " & imgVisor.ScrollPositionY
  Text3.Text = "X: " & imgVisor.ScrollPositionX
End Sub

Private Sub imgVisor_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
On Error GoTo Err:
    imgVisor.ZoomToSelection
    Exit Sub
Err:
End Sub
Sub Reposition()
On Error GoTo Error:
    If (rsSQL.RecordCount > 0) Then
        If (rsSQL.EOF) Then
            rsSQL.MovePrevious
        End If
        If (rsSQL.BOF) Then
            rsSQL.MoveNext
        End If
        lblimagen.Caption = rsSQL.AbsolutePosition & " de " & rsSQL.RecordCount
        AsignaDatos
        txtdato.SetFocus
    End If
Error:
End Sub
Sub AsignaDatos()
    strArchivoImagen = IIf(IsNull(rsSQL!Orig_file), " ", rsSQL!Orig_file)
    If Trim(strArchivoImagen) <> "" Then
        If Dir(strArchivoImagen, vbArchive) <> "" Then
           iSeleccion = 0
            Verimagen (strArchivoImagen)
            
        ''''''''''''''''''''''''''''''''''''
'        imgVisor.AnnotationFont = "Courier"
'        imgVisor.AnnotationFontColor = vbRed
'        'This stamp uses embedded text macros for date and time
'        imgVisor.AnnotationStampText = "Numero de Reclamo on %B %d, %Y at %H"
        
        'imgVisor.AnnotationType = wiTextStamp  '8
        'imgVisor.Draw 10, 10
        'imgvisor.SaveAnnotations app.Path & "\" &
        'imgVisor.LoadAnnotations App.Path & "\prueba.tif", 1, 1, 0
        'imgVisor.SaveAnnotations App.Path & "\savemark.jpg", 1, 1, REMOVE_ANNOTATIONS
        'imgVisor.AnnotationType = wiOcrRegion
        'imgVisor.AnnotationOcrType = wiOcrTypeText
        '''''''''''''''''''''''''''''''''''''
        
        End If
    Else
        MsgBox "No hay imagen para el registro verifique", vbCritical + vbOKOnly, "Visor"
    End If
    
End Sub
Private Sub cmdprev_Click()
    If rsSQL.RecordCount > 0 Then
        rsSQL.MovePrevious
        Call Reposition
    End If
End Sub
Private Sub cmdEof_Click()
    If rsSQL.RecordCount > 0 Then
        rsSQL.MoveFirst
        Call Reposition
    End If
End Sub
Private Sub cmdBOF_Click()
    If rsSQL.RecordCount > 0 Then
        rsSQL.MoveLast
        Call Reposition
    End If
End Sub
Private Sub cmdNext_Click()
    If rsSQL.RecordCount > 0 Then
        rsSQL.MoveNext
        Call Reposition
    End If
End Sub

Private Sub txtDato_GotFocus()
    txtdato.SelStart = 0
    txtdato.SelLength = Len(txtdato.Text)
End Sub
Sub Seleccionar()
Dim intActual As Integer
Dim intsiguiente As Integer
Dim strTemporal As String
    intActual = Mid(strActual, Len(strActual) - 2)
    strTemporal = strDirectorio & Mid(strActual, 1, Len(strActual) - 3) & "000.TIF"
    If Dir(strTemporal, vbArchive) <> "" Then
        If ((intActual Mod 2) = 0) Then
            strSiguiente = strDirectorio & Mid(strActual, 1, Len(strActual) - 3) & Padln(intActual + 1, 3, "0") & ".TIF"
            strActual = strArchivoImagen
        Else
            strSiguiente = strArchivoImagen
            strTemporal = strDirectorio & Mid(strActual, 1, Len(strActual) - 3) & Padln(intActual - 1, 3, "0") & ".TIF"
            strActual = strTemporal
        End If
    Else
        If ((intActual Mod 2) = 0) Then
            strSiguiente = strArchivoImagen
            strTemporal = strDirectorio & Mid(strActual, 1, Len(strActual) - 3) & Padln(intActual - 1, 3, "0") & ".TIF"
            strActual = strTemporal
        Else
            strSiguiente = strDirectorio & Mid(strActual, 1, Len(strActual) - 3) & Padln(intActual + 1, 3, "0") & ".TIF"
            strActual = strArchivoImagen
        End If
    End If
    iSeleccion = 1
End Sub

Sub Resaltado(Izq As Integer, Top As Integer, Width As Integer, Height As Integer, X As Integer, Y As Integer)
Dim Resaltado As String
On Local Error Resume Next
        imgVisor.DeleteAnnotationGroup ("Resaltado")
        imgVisor.AddAnnotationGroup ("Resaltado")
        imgVisor.AnnotationType = 4  '4
        'ImgEdit1.Draw 700, 100, 1200, 150
        'imgVisor.ScrollPositionX = x
        'ImgEdit1.ScrollPositionY = 100
        'imgVisor.ScrollPositionY = Y + Val(txtFactor.Text)
        'imgVisor.ScrollPositionY = y '+ Val(txtFactor.Text)
        imgVisor.Draw Izq / 15, Top / 15, Width / 15, Height / 15
        imgVisor.ShowAnnotationGroup ("Resaltado")
        imgVisor.AnnotationType = wiNone
        imgVisor.Enabled = False
        imgVisor.Enabled = True
        imgVisor.MousePointer = 0
        'imgVisor.DrawSelectionRect 1, 1, 1, 1
        imgVisor.DrawSelectionRect Izq, Top, Width, Height
        imgVisor.Refresh
   
End Sub


Sub Resaltado1(Control As ImgEdit, X As Integer, Y As Integer, w As Integer, h As Integer)

On Error Resume Next
vbzoom = False

        'ImgEdit1.DeleteImageData
       ' Control.DrawSelectionRect x / 15, y / 15, w, h
               
        Control.DeleteAnnotationGroup ("Resaltado")
        Control.AddAnnotationGroup ("Resaltado")
        Control.AnnotationType = 4  '4
        Control.AnnotationFillColor = vbYellow
        vbzoom = False
        Control.Draw X / 15, Y / 15, CInt(w / 15), CInt(h / 15)
        Control.ShowAnnotationGroup ("Resaltado")
        Control.AnnotationType = wiNone
        Control.Refresh
End Sub

VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "IMGEDIT.OCX"
Begin VB.Form Icr_FrmVerificar 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Verificar Liquidación Previa"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   Icon            =   "Icr_FrmVerificar.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10410
   ScaleWidth      =   14970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   10335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.Frame fravaidar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   13335
         Begin VB.TextBox txtnropla 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   10920
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtremu 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   8790
            TabIndex        =   25
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Txtcuspp 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   240
            MaxLength       =   12
            TabIndex        =   24
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txtcodigo 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2970
            MaxLength       =   2
            TabIndex        =   23
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtfechaini 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4515
            MaxLength       =   6
            TabIndex        =   22
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtfechafin 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6645
            MaxLength       =   6
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NRO DE PLANILLA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   10965
            TabIndex        =   32
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "REMUNERACION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   8760
            TabIndex        =   31
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "FECHA CESE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   6720
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblcuspp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "CUSPP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   480
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "CODIGO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   3000
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "FECHA INICO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   4560
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdRotar 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_FrmVerificar.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Rotar"
         Top             =   4215
         Width           =   630
      End
      Begin VB.CommandButton cmdZommin 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_FrmVerificar.frx":1594
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "ZoomIn"
         Top             =   2775
         Width           =   630
      End
      Begin VB.CommandButton cmdZommout 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_FrmVerificar.frx":281E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "ZoomOut"
         Top             =   2055
         Width           =   630
      End
      Begin VB.CommandButton cmdNormal 
         Height          =   630
         Left            =   13845
         Picture         =   "Icr_FrmVerificar.frx":3AA8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Normal"
         Top             =   3495
         Width           =   630
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
         TabIndex        =   11
         Top             =   9690
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
         TabIndex        =   10
         Top             =   9690
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
         TabIndex        =   9
         Top             =   9690
         Width           =   975
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
         Top             =   9690
         Width           =   975
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   9720
         Width           =   1815
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   630
         Left            =   13860
         Picture         =   "Icr_FrmVerificar.frx":4D32
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Anverso"
         Top             =   5010
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton cmdSiguiente 
         Height          =   630
         Left            =   13890
         Picture         =   "Icr_FrmVerificar.frx":503C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Reverso"
         Top             =   5730
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   6480
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   9690
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   435
         Left            =   7440
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   9690
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   8400
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   9690
         Visible         =   0   'False
         Width           =   975
      End
      Begin ImgeditLibCtl.ImgEdit imgVisor 
         Height          =   7695
         Left            =   240
         TabIndex        =   16
         Top             =   1770
         Width           =   13335
         _Version        =   131076
         _ExtentX        =   23521
         _ExtentY        =   13573
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         AutoRefresh     =   -1  'True
         DisplayScaleAlgorithm=   1
         UndoBufferSize  =   412402176
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
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
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
         Width           =   13335
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
         TabIndex        =   19
         Top             =   9705
         Width           =   1800
      End
      Begin VB.Shape Shape10 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   7725
         Left            =   13680
         Shape           =   4  'Rounded Rectangle
         Top             =   1770
         Width           =   975
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
         TabIndex        =   17
         Top             =   360
         Width           =   690
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   9570
         Width           =   14415
      End
   End
End
Attribute VB_Name = "Icr_FrmVerificar"
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
    
    If Len(txtfechaini.Text) > 0 Then
            If Len(txtfechaini.Text) = 6 Then
                        If Verificar_Fecha(Trim$(txtfechaini.Text)) Then
                        Else
                        MsgBox "verifique la fecha Inicio", vbCritical, "Liquidación"
                        Exit Sub
                        End If
                
            Else
            MsgBox "verifique la fecha Inicio", vbCritical, "Liquidación"
            Exit Sub
            End If
    Else
    'MsgBox "verifique la fecha Inicio", vbCritical, "Liquidación"
    'Exit Sub
    End If
    
    If Len(txtfechafin.Text) > 0 Then
            If Len(txtfechafin.Text) = 6 Then
                        If Verificar_Fecha(Trim$(txtfechafin.Text)) Then
                        Else
                        MsgBox "verifique la fecha Cese", vbCritical, "Liquidación"
                        Exit Sub
                        End If
                
            Else
            MsgBox "verifique la fecha Cese", vbCritical, "Liquidación"
            Exit Sub
            End If
    Else
    'MsgBox "verifique la fecha Inicio", vbCritical, "Liquidación"
    'Exit Sub
    End If
    
    If Len(Format(txtcodigo.Text, "00")) > 0 Then
       If Len(Format(txtcodigo.Text, "00")) = 2 Then
         If IsNumeric(Val(Trim$(txtcodigo.Text))) Then
         Else
         MsgBox "Ingresar Correctamente el codigo ", vbCritical, "Liquidación"
         Exit Sub
         End If
       Else
       MsgBox "Ingresar Correctamente el codigo ", vbCritical, "Liquidación"
       Exit Sub
       End If
    Else
    End If
    
    
    If Len(Trim$(txtremu.Text)) > 0 Then
      If IsNumeric(Val(Trim$(txtremu.Text))) Then
      
      Else
      MsgBox "Ingresar Correctamente la Remuneracion ", vbCritical, "Liquidación"
       Exit Sub
      End If
    Else
    
    End If
    
    
    If Len(Trim$(txtnropla.Text)) > 0 Then
      If IsNumeric(Val(Trim$(txtnropla.Text))) Then
      
      Else
      MsgBox "Ingresar Correctamente el Nro de Planilla", vbCritical, "Liquidación"
       Exit Sub
      End If
    Else
    
    End If
    
    
   If Len(Trim$(txtcodigo.Text)) <> 0 Or Len(Trim$(txtfechaini.Text)) <> 0 Or Len(Trim$(txtfechafin.Text)) <> 0 Or Len(Trim$(txtremu.Text)) <> 0 Or Len(Trim$(txtnropla.Text)) <> 0 Then
    ActualizarDatos_Verificados "" & rsSQL!nu_posicion
   End If
    
    
    
'        Select Case rsSQL!ds_rechazo
'
'           Case "COD.BARRA.NULL", "COD.BARRA.NOEXISTE"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If ExisteBarra("" & txtDato.Text) = False Then
'                        MsgBox "El codigo de Barra no existe en la tabla de LIQUIDACION IMPRESAS", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
'
'
'           Case "NRO.LIQUIDACION.NOEXISTE", "NRO.LIQUIDACION.NULL"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If ExisteNroliq("" & txtDato.Text) = False Then
'                        MsgBox "El Nro de Liquidación no existe en la tabla de LIQUIDACION IMPRESAS", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
'
'
'           Case "CODIGO.NOEXISTE"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If Verificar_Codigo("" & txtDato.Text) = False Then
'                        MsgBox "El Codigo no es valido", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
'
'
'           Case "FECHA.ERROR.INI"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If Verificar_Fecha("" & txtDato.Text) = False Then
'                        MsgBox "La fecha ingresada no es valida", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
'
'
'           Case "FECHA.ERROR.CESE"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If Verificar_Fecha("" & txtDato.Text) = False Then
'                        MsgBox "La fecha ingresada no es valida", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
'
'
'            Case "REMUNERACION.ERROR"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If IsNumeric("" & txtDato.Text) = False Then
'                        MsgBox "Los datos nos son numericos", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
'
'            Case "NRO.PLANILLA.ERROR"
'                If Trim("" & txtDato.Text) <> "" Then
'                    If IsNumeric("" & txtDato.Text) = False Then
'                        MsgBox "Los datos nos son numericos", vbCritical, "Liquidación"
'                    Else
'                    ActualizarDatos
'                    End If
'                End If
'
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        End Select
'
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

Sub ActualizarDatos_Verificados(strNumero As String)
Dim StrCodigo As String
Dim strFechaIni As String
Dim strFechaFi As String
Dim strRemunera As String
Dim strPlanilla As String

Dim strCodigoDat As String
Dim strFechaIniDat As String
Dim strFechaFiDat As String
Dim strRemuneraDat As String
Dim strPlanillaDat As String

StrCodigo = "ds_cod" & strNumero
strFechaIni = "ds_fec_ini" & strNumero
strFechaFi = "ds_fec_ce" & strNumero
strRemunera = "ds_rem_as" & strNumero
strPlanilla = "ds_pla" & strNumero

strCodigoDat = UCase("" & Trim(txtcodigo.Text))
strFechaIniDat = UCase("" & Trim(txtfechaini.Text))
strFechaFiDat = UCase("" & Trim(txtfechafin.Text))
strRemuneraDat = UCase("" & Trim(txtremu.Text))
strPlanillaDat = UCase("" & Trim(txtnropla.Text))

    'txtDato.Text = UCase("" & Trim(txtDato.Text))
    'strEjecutar = "Update AFILIACION set " & rsSQL!ds_campo & "='" & txtDato.Text & "'  where id = " & rsSQL!id_hojapedido
    strEjecutar = "Update tb_liquidacion set " & StrCodigo & " ='" & strCodigoDat & "'," & strFechaIni & "='" & strFechaIniDat & "'," & strFechaFi & "='" & strFechaFiDat & "'," & strRemunera & "='" & strRemuneraDat & "'," & strPlanilla & "='" & strPlanillaDat & "'  where id_liq = " & rsSQL!id_hojapedido
    cnServer.Execute strEjecutar
    'strEjecutar = "Update tb_rechazos set ds_valor='" & txtDato.Text & "' where id_rechazo = " & rsSQL!id_rechazo
    'cnServer.Execute strEjecutar
    
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
    'rsSQL.Source = "Select * from vw_Rechazos order by id_hojapedido"
    
    If rsSQL.State = 1 Then rsSQL.Close
    
    rsSQL.Source = "Select * from vw_verificacion order by id_hojapedido"
    rsSQL.Open
    
    If rsSQL.RecordCount > 0 Then
        Reposition
    Else
        MsgBox "No existen Datos a Verificar", vbCritical, "Revisar Verificaciones "
        Unload Icr_FrmVerificar
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
        
        lngPosicion = 0
        lngPosicionFecha = 950
        Select Case rsSQL!nu_posicion
                       
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                    'lblError.Caption = "MOTIVO : [ " & rsSQL!ds_rechazo & " ]"
                  Me.txtdato.Text = "" & rsSQL!ds_valor
                  fravaidar.Caption = "Datos de Registro Numero " & rsSQL!nu_posicion
                  Txtcuspp.Text = "" & rsSQL!ds_cuspp
                  txtcodigo.Text = "" & rsSQL!ds_cod_descargo
                  txtfechaini.Text = "" & rsSQL!ds_fec_ini
                  txtfechafin.Text = "" & rsSQL!ds_fecha_cese
                  txtremu.Text = "" & rsSQL!ds_rem
                  txtnropla.Text = "" & rsSQL!ds_pla
                  'txtfechaini.Text = "" & rsSQL!ds_fecha
                  lblError.Caption = "Razon Social " & "" & rsSQL!ds_raz_soc & "-Nro Liq. " & Mid(rsSQL!nroliqpag, 1, 10)
                  
                 ' Select Case rsSQL!nu_posicion
                    
                 ' Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                  Select Case rsSQL!nu_posicion
                  
                  Case 1
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 280 '300 'lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                  Case 2
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 316 '300
                  Case 3
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 364 '300
                  Case 4
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 395 '300
                  Case 5
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 432 '300
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
                        lngScrollY = 826 '300
                  Case 17
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 860 '300
                  Case 18
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 1300 '300
                  Case 19
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 1800 '300
                  Case 20
                        lngZomm = 69 '708
                        lngScrollX = 0 '1480
                        lngScrollY = 2460 '300
                        
'                        lngScrollX = 0 '1480
'                        lngScrollY = lngPosicion + (300 * rsSQL!nu_posicion)  '376 '36 '3160
                        'Resaltado imgVisor.Left, imgVisor.Top, imgVisor.Width, imgVisor.Height, CInt(lngScrollX), CInt(lngScrollY)
                  End Select
                        lngRex = CInt(lngScrollX)
                        lngRey = CInt(lngScrollY)
                        lngReAncho = 12000
                        lngReAlto = 600
                        
'                        lngZomm = 61.74
'                        lngScrollX = 0
'                        lngScrollY = lngPosicion + (22 * rsSQL!nu_posicion)  '372
                        
                  
        End Select
        
        imgVisor.Zoom = lngZomm
        imgVisor.Display
        imgVisor.ScrollPositionY = lngScrollY
        imgVisor.ScrollPositionX = lngScrollX
        imgVisor.ScrollBars = True
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
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'imgVisor.AnnotationFont = "Courier"
        'imgVisor.AnnotationFontColor = vbRed
        'This stamp uses embedded text macros for date and time
        'imgVisor.AnnotationStampText = "Numero de Reclamo on %B %d, %Y at %H"
        'imgVisor.AnnotationType = wiTextStamp  '8
        'imgVisor.Draw 10, 10
        'imgvisor.SaveAnnotations app.Path & "\" &
        'imgVisor.LoadAnnotations App.Path & "\prueba.tif", 1, 1, 0
        'imgVisor.SaveAnnotations App.Path & "\savemark.jpg", 1, 1, REMOVE_ANNOTATIONS
        'imgVisor.AnnotationType = wiOcrRegion
        'imgVisor.AnnotationOcrType = wiOcrTypeText
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
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


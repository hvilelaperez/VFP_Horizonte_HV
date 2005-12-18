VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLote 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del lote"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPrefijo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   15
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtFolioInicial 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   2910
      Width           =   2895
   End
   Begin VB.TextBox txtFolioFinal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   19
      Top             =   3285
      Width           =   2895
   End
   Begin VB.TextBox txtLoteFisico 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   435
      Width           =   2895
   End
   Begin VB.ComboBox cmbOperadorPreparacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   795
      Width           =   2895
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   240
      TabIndex        =   41
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   1680
      TabIndex        =   44
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtFechaFin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdFechaFin 
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
      Left            =   240
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Seleccione el directorio ..."
      Top             =   1605
      Width           =   255
   End
   Begin MSMask.MaskEdBox mskHoraInicio 
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Top             =   1185
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm AM/PM"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdFechaInicio 
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
      Left            =   240
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Seleccione el directorio ..."
      Top             =   1230
      Width           =   255
   End
   Begin VB.TextBox txtFechaInicio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   1185
      Width           =   1575
   End
   Begin MSMask.MaskEdBox mskHoraFin 
      Height          =   315
      Left            =   3840
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm AM/PM"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
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
      Left            =   240
      TabIndex        =   43
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label txtTotalFolios 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2160
      TabIndex        =   46
      Top             =   3660
      Width           =   2925
   End
   Begin VB.Label lblTotalFolios 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total folios"
      Height          =   210
      Left            =   1230
      TabIndex        =   45
      Top             =   3712
      Width           =   780
   End
   Begin VB.Label lblFoliofinal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folio final"
      Height          =   210
      Left            =   1335
      TabIndex        =   18
      Top             =   3345
      Width           =   675
   End
   Begin VB.Label lblFolioInicial 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folio inicial"
      Height          =   210
      Left            =   1245
      TabIndex        =   16
      Top             =   2955
      Width           =   765
   End
   Begin VB.Label lblPrefijo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prefijo del folio"
      Height          =   210
      Left            =   960
      TabIndex        =   14
      Top             =   2565
      Width           =   1050
   End
   Begin VB.Label lblTotalImagenes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total imagenes"
      Height          =   210
      Left            =   930
      TabIndex        =   24
      Top             =   4965
      Width           =   1080
   End
   Begin VB.Label txtTotalImagenes 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2160
      TabIndex        =   25
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label lblFinClean 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin"
      Height          =   210
      Left            =   3840
      TabIndex        =   39
      Top             =   6450
      Width           =   210
   End
   Begin VB.Label lblInicioClean 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      Height          =   210
      Left            =   3840
      TabIndex        =   37
      Top             =   5805
      Width           =   360
   End
   Begin VB.Label txtFinClean 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3840
      TabIndex        =   40
      Top             =   6720
      Width           =   1605
   End
   Begin VB.Label txtInicioClean 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3840
      TabIndex        =   38
      Top             =   6075
      Width           =   1605
   End
   Begin VB.Label lblFinIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin"
      Height          =   210
      Left            =   2040
      TabIndex        =   35
      Top             =   6450
      Width           =   210
   End
   Begin VB.Label lblInicioIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      Height          =   210
      Left            =   2040
      TabIndex        =   33
      Top             =   5805
      Width           =   360
   End
   Begin VB.Label txtFinIndex 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2040
      TabIndex        =   36
      Top             =   6720
      Width           =   1605
   End
   Begin VB.Label txtInicioIndex 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2040
      TabIndex        =   34
      Top             =   6075
      Width           =   1605
   End
   Begin VB.Label lblFinScan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin"
      Height          =   210
      Left            =   240
      TabIndex        =   31
      Top             =   6450
      Width           =   210
   End
   Begin VB.Label lbInicioScan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      Height          =   210
      Left            =   240
      TabIndex        =   29
      Top             =   5805
      Width           =   360
   End
   Begin VB.Label txtFinScan 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   32
      Top             =   6720
      Width           =   1605
   End
   Begin VB.Label txtInicioScan 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   30
      Top             =   6075
      Width           =   1605
   End
   Begin VB.Label lblClean 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso CLEAN"
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
      Left            =   3840
      TabIndex        =   28
      Top             =   5520
      Width           =   1275
   End
   Begin VB.Label lblIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso INDEX"
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
      Left            =   2040
      TabIndex        =   27
      Top             =   5520
      Width           =   1185
   End
   Begin VB.Label lblScan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso SCAN"
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
      Left            =   240
      TabIndex        =   26
      Top             =   5520
      Width           =   1185
   End
   Begin VB.Label lblLoteFisico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote Fisico"
      Height          =   210
      Left            =   1230
      TabIndex        =   1
      Top             =   480
      Width           =   780
   End
   Begin VB.Label lblOperador2 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2160
      TabIndex        =   23
      Top             =   4530
      Width           =   2895
   End
   Begin VB.Label lblOperador 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2160
      TabIndex        =   42
      Top             =   795
      Width           =   2895
   End
   Begin VB.Label lblLoteActual 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2160
      TabIndex        =   21
      Top             =   4155
      Width           =   2895
   End
   Begin VB.Label lblOperadorScan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador"
      Height          =   210
      Left            =   1320
      TabIndex        =   22
      Top             =   4590
      Width           =   690
   End
   Begin VB.Label lblScaneo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la digitalización"
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
      TabIndex        =   13
      Top             =   2160
      Width           =   2100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha - Hora Inicio"
      Height          =   210
      Left            =   660
      TabIndex        =   5
      Top             =   1230
      Width           =   1350
   End
   Begin VB.Label lblPreparacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la preparación"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha - Hora fin"
      Height          =   210
      Left            =   840
      TabIndex        =   9
      Top             =   1605
      Width           =   1170
   End
   Begin VB.Label lblOperadorPreparacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador"
      Height          =   210
      Left            =   1320
      TabIndex        =   3
      Top             =   855
      Width           =   690
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote Digitalizacion"
      Height          =   210
      Left            =   720
      TabIndex        =   20
      Top             =   4200
      Width           =   1290
   End
End
Attribute VB_Name = "frmLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strPrefijoINI As String

Private Sub cmdAceptar_Click()
Dim strCampos As String
Dim intTotalFolios As Integer
    mskHoraInicio.Text = Replace(mskHoraInicio.Text, "_", "0")
    mskHoraFin.Text = Replace(mskHoraFin.Text, "_", "0")
    
    If txtFolioInicial.Text = "" Then txtFolioInicial.Text = 0
    If txtFolioFinal.Text = "" Then txtFolioFinal.Text = 0
    txtTotalFolios.Caption = 0

    
    If Val(txtFolioFinal) > 0 And Val(txtFolioInicial) > 0 Then
        If Val(txtFolioFinal.Text) < Val(txtFolioInicial.Text) Then
            MsgBox "El número de folio inicial debe ser menor que el final", vbCritical, "Nuevo lote"
            Exit Sub
        Else
            intTotalFolios = Val(txtFolioFinal.Text) - Val(txtFolioInicial.Text)
            txtTotalFolios.Caption = intTotalFolios + 1
            'GuardarIni App.Path & "\digitacion.ini", "seteos", "PrefijoFolio", "" & "" & txtPrefijo.Text
        End If
    End If
    
    
    If txtFechaInicio.Text <> "" Then
        If txtFechaFin.Text <> "" Then
            If mskHoraInicio.Text <> "" Then
                If mskHoraFin.Text <> "" Then
                    If mskHoraFin.Text = mskHoraInicio.Text Then
                        MsgBox "La hora final no debe ser menor igual a la hora de inicio", vbCritical, "Nuevo lote"
                    Else
                        
                        If VerificaHora(mskHoraInicio.Text) = True Then
                            If VerificaHora(mskHoraFin.Text) = True Then
                                Dim rsLoteActual As New Recordset
                                Dim bolExisteLote As Boolean
                                
                                bolExisteLote = False
                                Set rsLoteActual = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Scan,[*],'where lotefisico=" & Trim(txtLoteFisico.Text) & "',''")
                                If rsLoteActual.RecordCount > 0 Then bolExisteLote = True Else bolExisteLote = False
                                If rsLoteActual.State = 1 Then rsLoteActual.Close
                                Set rsLoteActual = Nothing
                                
                                If bolExisteLote = True Then
                                    MsgBox "El lote fisico " & Trim(txtLoteFisico.Text) & " ya fué registrado", vbCritical, "Nuevo lote"
                                Else
                                    lngLoteActual = lngLoteActual + 1
                                    clsDataScan.Conexion.BeginTrans
                                    strCampos = lngIdVolumenActual & "," & lngLoteActual & ",'1'," & "'" & txtFechaInicio.Text & " " & mskHoraInicio & "'" & ",'" & txtFechaFin.Text & " " & mskHoraFin & "','" & Trim(txtLoteFisico.Text) & "'"
                                    clsDataScan.Conexion.Execute ("EXECUTE sp_InsertarPreScan " & strCampos)
                                    strCampos = lngIdVolumenActual & "," & lngLoteActual & ",1,'Scan','" & Trim(txtLoteFisico.Text) & "','" & Trim("" & txtPrefijo.Text) & "'," & ("" & txtFolioInicial.Text) & "," & ("" & txtFolioFinal.Text) & "," & txtTotalFolios.Caption
                                    clsDataScan.Conexion.Execute ("EXECUTE sp_InsertarScan " & strCampos)
                                    clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTabla _Volumenes,'set lote_actual=" & lngLoteActual & ",lote_fisico=" & Trim(txtLoteFisico.Text) & "','idVolumen=" & lngIdVolumenActual & "'")
                                    frmLotes.lstLotes.AddItem lngLoteActual
                                    frmLotes.lstLotes2.AddItem Trim(txtLoteFisico.Text)
                                    frmScan.lblLote.Caption = " " & lngLoteActual
                                    frmScan.lblLote2.Caption = " " & Trim(txtLoteFisico.Text)
                                    frmScan.lstImagenes.Clear
                                    frmScan.lblTotalImagenes.Caption = ""
                                    clsDataScan.Conexion.CommitTrans
                                    Unload Me
                                End If
                            Else
                                MsgBox "Ingrese una hora valida", vbCritical, "Hora final"
                            End If
                        Else
                            MsgBox "Ingrese una hora valida", vbCritical, "Hora de inicio"
                        End If
                    End If
                Else
                    MsgBox "Complete los datos", vbCritical, "Nuevo lote"
                End If
            Else
                MsgBox "Complete los datos", vbCritical, "Nuevo lote"
            End If
        Else
            MsgBox "Complete los datos", vbCritical, "Nuevo lote"
        End If
    Else
        MsgBox "Complete los datos", vbCritical, "Nuevo lote"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdFechaFin_Click()
    strFechaSeleccionada = ""
    frmFecha.Show vbModal
    If strFechaSeleccionada <> "" Then txtFechaFin.Text = strFechaSeleccionada
End Sub
Private Sub cmdFechaInicio_Click()
    strFechaSeleccionada = ""
    frmFecha.Show vbModal
    If strFechaSeleccionada <> "" Then txtFechaInicio.Text = strFechaSeleccionada
End Sub
Private Sub Form_Load()
    PrenderBotones bolModoForm
    strPrefijoINI = ""
    
    If bolModoForm = False Then
        clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTotalImagenesScan " & lngIdVolumenActual & "," & strLoteActual)
        Dim rsLoteActual As New Recordset
        Set rsLoteActual = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _PreScan,[*],'where idvolumen=" & lngIdVolumenActual & " and lote = " & strLoteActual & "',''")
        If rsLoteActual.RecordCount > 0 Then
            lblLoteActual.Caption = " " & rsLoteActual!Lote
            lblOperador.Caption = " " & rsLoteActual!idPreparador
            mskHoraInicio.Text = Format(Hour(rsLoteActual!fecha_hora_inicio), "00") & ":" & Format(Minute(rsLoteActual!fecha_hora_inicio), "00")
            mskHoraFin.Text = Format(Hour(rsLoteActual!fecha_hora_fin), "00") & ":" & Format(Minute(rsLoteActual!fecha_hora_fin), "00")
            txtFechaInicio.Text = Format(Day(rsLoteActual!fecha_hora_inicio), "00") & "/" & Format(Month(rsLoteActual!fecha_hora_inicio), "00") & "/" & Year(rsLoteActual!fecha_hora_inicio)
            txtFechaFin.Text = Format(Day(rsLoteActual!fecha_hora_fin), "00") & "/" & Format(Month(rsLoteActual!fecha_hora_fin), "00") & "/" & Year(rsLoteActual!fecha_hora_fin)
            txtLoteFisico.Text = " " & rsLoteActual!lotefisico
        End If
        If rsLoteActual.State = 1 Then rsLoteActual.Close
        Set rsLoteActual = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Scan,[*],'where idvolumen=" & lngIdVolumenActual & " and lote = " & strLoteActual & "',''")
        If rsLoteActual.RecordCount > 0 Then
            txtTotalImagenes.Caption = " " & "" & rsLoteActual!total_imagenes
            txtInicioScan.Caption = " " & "" & rsLoteActual!Inicio_Scan
            txtFinScan.Caption = " " & "" & rsLoteActual!Fin_Scan
            txtInicioIndex.Caption = " " & "" & rsLoteActual!Inicio_Indice
            txtFinIndex.Caption = " " & "" & rsLoteActual!Fin_Indice
            txtInicioClean.Caption = " " & "" & rsLoteActual!Inicio_Clean
            txtFinClean.Caption = " " & "" & rsLoteActual!Fin_Clean
            txtPrefijo.Text = " " & "" & rsLoteActual!prefijo
            txtFolioFinal.Text = " " & "" & rsLoteActual!foliofinal
            txtFolioInicial.Text = " " & "" & rsLoteActual!folioinicial
            txtTotalFolios.Caption = " " & "" & rsLoteActual!totalfolios
        End If
        If rsLoteActual.State = 1 Then rsLoteActual.Close
        
        Set rsLoteActual = Nothing
    Else
        lblLoteActual.Caption = " " & lngLoteActual + 1
        strPrefijoINI = LeerIni(App.Path & "\digitacion.ini", "seteos", "PrefijoFolio")
        txtPrefijo.Text = strPrefijoINI
    End If
End Sub
Private Sub mskHoraFin_KeyPress(KeyAscii As Integer)
    If bolModoForm = False Then KeyAscii = 0
End Sub
Private Sub mskHoraInicio_KeyPress(KeyAscii As Integer)
    If bolModoForm = False Then KeyAscii = 0
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtFechaInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtLote_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtOperadorPreparacion_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtOperadorScan_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Sub PrenderBotones(bolEstado As Boolean)
    cmdAceptar.Visible = bolEstado
    cmdCancelar.Visible = bolEstado
    cmdCerrar.Visible = Not bolEstado
    cmdFechaFin.Visible = bolEstado
    cmdFechaInicio.Visible = bolEstado
    lblOperador.Visible = Not bolEstado
    cmbOperadorPreparacion.Visible = bolEstado
    mskHoraFin.TabStop = bolEstado
    mskHoraInicio.TabStop = bolEstado
    txtFechaFin.TabStop = bolEstado
    txtFechaInicio.TabStop = bolEstado
End Sub

Function VerificaHora(strHora As String) As Boolean
Dim bolOk As Boolean
Dim intPosicion As Byte
Dim bytHora As Byte
Dim bytMinutos As Byte
    bolOk = False
    intPosicion = InStr(strHora, ":")
    bytHora = Val(Mid(strHora, 1, intPosicion - 1))
    bytMinutos = Val(Mid(strHora, intPosicion + 1))
    If bytHora > 0 And bytHora < 24 Then
        If bytMinutos >= 0 And bytMinutos < 60 Then
            bolOk = True
        End If
    End If
    VerificaHora = bolOk
End Function
Private Sub txtFolioFinal_GotFocus()
    txtFolioFinal.SelStart = 0
    txtFolioFinal.SelLength = Len(txtFolioFinal)
End Sub

Private Sub txtFolioFinal_KeyPress(KeyAscii As Integer)
    If bolModoForm = False Then
        KeyAscii = 0
    Else
        If KeyAscii <> 8 Then
            If (KeyAscii >= 48 And KeyAscii <= 57) Then
            Else
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtFolioInicial_GotFocus()
    txtFolioInicial.SelStart = 0
    txtFolioInicial.SelLength = Len(txtFolioInicial)
End Sub

Private Sub txtFolioInicial_KeyPress(KeyAscii As Integer)
    If bolModoForm = False Then
        KeyAscii = 0
    Else
        If KeyAscii <> 8 Then
            If (KeyAscii >= 48 And KeyAscii <= 57) Then
            Else
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtLoteFisico_KeyPress(KeyAscii As Integer)
    If bolModoForm = False Then KeyAscii = 0
End Sub

Private Sub txtPrefijo_GotFocus()
    txtPrefijo.SelStart = 0
    txtPrefijo.SelLength = Len(txtPrefijo)
End Sub

Private Sub txtPrefijo_KeyPress(KeyAscii As Integer)
    If bolModoForm = False Then KeyAscii = 0
End Sub

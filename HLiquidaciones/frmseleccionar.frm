VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Flow_frmseleccionar 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Seleccionar"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12135
   Icon            =   "frmseleccionar.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   12135
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox filExport 
      Height          =   480
      Left            =   8640
      Pattern         =   "*.tif"
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   11895
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Importar    imagenes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdTerminar 
         Caption         =   "Lote          terminado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Quitar Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdInicializar 
         Caption         =   "Inicializar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   3480
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar pbProceso 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   3690
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSDataGridLib.DataGrid datPC 
         Height          =   2655
         Left            =   1800
         TabIndex        =   4
         Top             =   930
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTrabajando 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes procesando localmente"
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
         Left            =   360
         TabIndex        =   6
         Top             =   450
         Width           =   2475
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   11415
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   3690
         Width           =   7695
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   3135
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   930
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11895
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Refrescar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid datDatos 
         Height          =   2895
         Left            =   1800
         TabIndex        =   1
         Top             =   810
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblLotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes a procesar"
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
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   11415
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2895
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   810
         Width           =   1455
      End
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
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
      TabIndex        =   15
      Top             =   210
      Width           =   705
   End
   Begin VB.Label lblUsuarios 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   120
      Width           =   2040
   End
End
Attribute VB_Name = "Flow_frmseleccionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnCliente As New Connection
Dim strCadena As String
Dim intPosicion As Long
Dim strDir As String
Dim strEntrega As String
Dim strLote As String

Private Sub cmdEliminar_Click()
    If rsPc.RecordCount > 0 Then
        'intRespuesta = MsgBox("Desea quitar el lote del Periodo: " & rsPc!Periodo & ", Año: " & rsPc!año & ", Batch: " & rsPc!batch, vbQuestion + vbYesNo, "Exportar")
        intRespuesta = MsgBox("Desea quitar el lote del Periodo: " & rsPc!Periodo & ", Año: " & rsPc!año & ", Lote: " & rsPc!lote, vbQuestion + vbYesNo, "Exportar")
        If intRespuesta = vbYes Then
            strCadena = "UPDATE TB_LOTES set ds_estado = 'LIBRE', ds_usuario = ' ' where id_lote=" & rsPc!id_origen
            cnServer.Execute strCadena
            rsDatos.Requery
            Columnas
            
            strCadena = "Delete from " & strUsuarioTablaLote & " where id_lote=" & rsPc!id_lote
            cnServer.Execute strCadena
            rsPc.Requery
            PcColumnas
        End If
    End If
End Sub
Private Sub cmdExportar_Click()
Dim strRutaCliente As String
Dim strRutaServidor As String
    If rsPc.RecordCount > 0 Then
        If rsPc!Estado <> "TERMINADO" Then
            'intRespuesta = MsgBox("Desea importar las imagenes del lote lote del Periodo: " & rsPc!Periodo & ", Año: " & rsPc!año & ", Batch: " & rsPc!batch, vbQuestion + vbYesNo, "Exportar")
            intRespuesta = MsgBox("Desea importar las imagenes del Periodo: " & rsPc!Periodo & ", Año: " & rsPc!año & ", Lote: " & rsPc!lote, vbQuestion + vbYesNo, "Exportar")
            If intRespuesta = vbYes Then
                strDir = Trim(rsPc!Path & " ")
                If strDriveCliente = "" Then strDriveCliente = "C"
                If strDriveServer = "" Then strDriveServer = "C"
                
                strRutaCliente = strDriveCliente & Mid(strDir, 2)
                If Dir(strRutaCliente, vbDirectory) = "" Then
                        CrearDirectorio (strRutaCliente)
                End If
                strRutaCliente = strDriveCliente & Mid(rsPc!Path, 2)
                strRutaServidor = strDriveServer & Mid(rsPc!Path, 2)
                ExportarImagenes strRutaServidor, strRutaCliente
            End If
        End If
    End If
End Sub
Private Sub cmdInicializar_Click()
    If (vbYes = MsgBox("Desea inicializar la tabla " & strUsuarioTablaLote, vbQuestion + vbYesNo, "PCLote")) Then
        strCadena = "Delete From " & strUsuarioTablaLote
        cnServer.Execute strCadena
        rsPc.Requery
        PcColumnas
        MsgBox "Tabla inicializada", vbInformation, App.EXEName
    End If
End Sub

Private Sub cmdRefrescar_Click()
    rsDatos.Requery
    Columnas
End Sub
Private Sub cmdSeleccionar_Click()
Dim strBuscar As String
    If rsDatos.RecordCount > 0 Then
        If rsDatos!Estado = "LIBRE" Then
            strBuscar = rsDatos!id_lote
            'If vbYes = MsgBox("Desea seleccionar el lote del Periodo: " & rsDatos!Periodo & ", Año: " & rsDatos!año & ", Batch: " & rsDatos!batch, vbQuestion + vbYesNo, "Exportar") Then
            If vbYes = MsgBox("Desea seleccionar el lote del Periodo: " & rsDatos!Periodo & ", Año: " & rsDatos!año & ", Lote: " & rsDatos!lote, vbQuestion + vbYesNo, "Exportar") Then
                rsDatos.Requery
                Columnas
                rsDatos.MoveFirst
                rsDatos.Find "id_lote = " & strBuscar
                If Not rsDatos.EOF Then
                    If rsDatos!Estado = "LIBRE" Then
                        strCadena = "exec sp_LoteEnProceso " & rsDatos!id_lote & ",'" & lblUsuarios.Caption & "'"
                        cnServer.Execute strCadena
                        rsDatos.Requery
                        rsDatos.MoveFirst
                        rsDatos.Find "id_lote = " & strBuscar
                        Columnas
                        If rsSQL.State = 1 Then rsSQL.Close
                        rsSQL.Source = "Select * from TB_LOTES where id_lote=" & strBuscar
                        rsSQL.Open
                        If rsSQL.RecordCount > 0 Then
                            If rsSQL!ds_Usuario = lblUsuarios.Caption Then
                                strCadena = "exec sp_InsertarLoteSeleccionado " & strBuscar & ",'" & strUsuarioTablaLote & "'"
                                cnServer.Execute strCadena
                                rsPc.Requery
                                PcColumnas
                            Else
                                MsgBox "No puede seleccionar este lote, está siendo procesado por el usuario " & rsDatos!Usuario, vbCritical, "Seleccionar lote"
                            End If
                        End If
                        If rsSQL.State = 1 Then rsSQL.Close
                    Else
                        MsgBox "No puede seleccionar este lote, está siendo procesado por el usuario " & rsDatos!Usuario, vbCritical, "Seleccionar lote"
                    End If
                Else
                    'MsgBox "No se encuentra el registro del lote de Campaña: " & rsDatos!Periodo & ", Año: " & rsDatos!año & ", Batch: " & rsDatos!batch, vbCritical, "Seleccionar lote"
                    MsgBox "No se encuentra el registro del lote de Campaña: " & rsDatos!Periodo & ", Año: " & rsDatos!año & ", Lote: " & rsDatos!lote, vbCritical, "Seleccionar lote"
                End If
            End If
        Else
            MsgBox "El lote no está disponible", vbCritical, "Seleccionar lote"
        End If
    End If
End Sub
Private Sub cmdTerminar_Click()
    If rsPc.RecordCount > 0 Then
        If rsPc!Estado <> "TERMINADO" Then
            'intRespuesta = MsgBox("Desea dar por terminado el lote del Periodo: " & rsPc!Periodo & ", Año: " & rsPc!año & ", Batch: " & rsPc!batch, vbQuestion + vbYesNo, "Exportar")
            intRespuesta = MsgBox("Desea dar por terminado el lote del Periodo: " & rsPc!Periodo & ", Año: " & rsPc!año & ", Lote: " & rsPc!lote, vbQuestion + vbYesNo, "Exportar")
            If intRespuesta = vbYes Then
                intPosicion = rsDatos.AbsolutePosition
                
                strCadena = "exec sp_TerminarProcesoLote " & rsPc!id_lote & "," & rsPc!id_origen & ",'" & strUsuarioTablaLote & "'"
                cnServer.Execute strCadena
                
                rsDatos.Requery
                rsDatos.AbsolutePosition = intPosicion
                
                Columnas
                rsPc.Requery
                PcColumnas
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    EstadoMenu False
    Me.Width = 12255 '12000
    Me.Height = 9105 '8910
    
    Me.Left = (Flow_Frmpanel.Width / 2) - (Me.Width / 2)
    Me.Top = (Flow_Frmpanel.Height / 2) - (Me.Height / 2)
    
    lblUsuarios.Caption = strUsuarioActivo
    
        
        rsDatos.CursorType = adOpenDynamic
        rsDatos.LockType = adLockBatchOptimistic
        rsDatos.ActiveConnection = cnServer
    
        rsPc.CursorType = adOpenDynamic
        rsPc.LockType = adLockBatchOptimistic
        rsPc.ActiveConnection = cnServer
    
        rsSQL.CursorType = adOpenDynamic
        rsSQL.LockType = adLockBatchOptimistic
        rsSQL.ActiveConnection = cnServer
    
        ''''''''''' General
        If rsDatos.State = 1 Then rsDatos.Close
        'rsDatos.Source = "Select fc_fecha,ds_periodo as Periodo, ds_año as Año, ds_batch as batch,nu_imagenes as imagenes,ds_estado as estado,ds_usuario as usuario,fc_hora_fin as hora_fin,ds_path as path,id_lote from TB_LOTES order by ds_periodo,ds_año,ds_batch"
        rsDatos.Source = "Select fc_fecha,ds_periodo as Periodo, ds_año as Año, ds_batch as Lote,nu_imagenes as imagenes,ds_estado as estado,ds_usuario as usuario,fc_hora_fin as hora_fin,ds_path as path,id_lote from TB_LOTES order by ds_periodo,ds_año,ds_batch"
        rsDatos.Open
        Set datDatos.DataSource = rsDatos
        Columnas
        ''''''''''' En la PC
        If rsPc.State = 1 Then rsPc.Close
        'rsPc.Source = "Select fc_hora_inicio as hora_inicio,ds_periodo as Periodo, ds_año as Año, ds_batch as batch,nu_imagenes as imagenes,ds_estado as estado,ds_usuario as usuario,fc_hora_fin as hora_fin,ds_path as path, id_lote,id_origen from " & strUsuarioTablaLote & " order by ds_periodo,ds_año,ds_batch"
        rsPc.Source = "Select fc_hora_inicio as hora_inicio,ds_periodo as Periodo, ds_año as Año, ds_batch as Lote,nu_imagenes as imagenes,ds_estado as estado,ds_usuario as usuario,fc_hora_fin as hora_fin,ds_path as path, id_lote,id_origen from " & strUsuarioTablaLote & " order by ds_periodo,ds_año,ds_batch"
        rsPc.Open
        Set datPC.DataSource = rsPc
        PcColumnas
    
End Sub
Sub Columnas()
        datDatos.Columns(0).Width = 2000
        datDatos.Columns(1).Width = 900
        datDatos.Columns(2).Width = 700
        datDatos.Columns(3).Width = 900
        datDatos.Columns(4).Width = 1200
End Sub
Sub PcColumnas()
    datPC.Columns(0).Width = 2000
    datPC.Columns(1).Width = 900
    datPC.Columns(2).Width = 700
    datPC.Columns(3).Width = 900
    datPC.Columns(4).Width = 1200
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err:
    EstadoMenu True
    If rsDatos.State = 1 Then rsDatos.Close
    If rsPc.State = 1 Then rsPc.Close
    If rsSQL.State = 1 Then rsSQL.Close
    If cnCliente.State = 1 Then cnCliente.Close
    If cnConexion.State = 1 Then cnConexion.Close

    Set cnCliente = Nothing
   
    Exit Sub
Err:

End Sub
Sub CrearDirectorio(strAcrear As String)
Dim intBuscar As Integer
Dim intsiguiente As Integer
Dim strCrea As String
    strAcrear = strAcrear & "\"
    intBuscar = InStr(Mid(strAcrear, 4), "\") + 3
    intsiguiente = intBuscar
    Do While intsiguiente <> 0
        If intsiguiente <> 0 Then
            strCrea = Mid(strAcrear, 1, intBuscar)
            If Dir(strCrea, vbDirectory) = "" Then
                MkDir strCrea
            End If
            intsiguiente = InStr(Mid(strAcrear, intBuscar + 1), "\")
            intBuscar = intBuscar + intsiguiente
        End If
    Loop
End Sub
Sub ExportarImagenes(strOrigen, strDestino)
Dim intIndice As Integer
    pbProceso.Min = 0
    pbProceso.Value = 0
    filExport.Path = strOrigen
    If filExport.ListCount > 0 Then
        pbProceso.Max = filExport.ListCount
        For intIndice = 0 To filExport.ListCount - 1
            DoEvents
            lblEstado.Caption = "Copiando archivo " & strOrigen & "\" & filExport.List(intIndice)
            FileCopy strOrigen & "\" & filExport.List(intIndice), strDestino & "\" & filExport.List(intIndice)
            pbProceso.Value = pbProceso.Value + 1
        Next
    End If
    lblEstado.Caption = ""
    pbProceso.Value = 0
End Sub



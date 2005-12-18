VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Flow_FrmImportarData 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Importar Datos"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15300
   Icon            =   "FrmImportarData.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10170
   ScaleWidth      =   15300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      Begin VB.ComboBox cmbEntrega 
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton cmdRefreshLotes 
         Caption         =   "Refrescar"
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
         Left            =   13680
         TabIndex        =   6
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Seleccionar lotes para importar"
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
         Left            =   4680
         TabIndex        =   5
         Top             =   5280
         Width           =   3015
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Importar Datos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7920
         TabIndex        =   4
         Top             =   7200
         Width           =   975
      End
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Refrescar"
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
         Left            =   120
         TabIndex        =   3
         Top             =   4800
         Width           =   1335
      End
      Begin VB.ListBox lstUsuarios 
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
         ForeColor       =   &H00000000&
         Height          =   4020
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.ListBox lstUsuarios 
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
         Height          =   4020
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid datFlow 
         Height          =   3975
         Left            =   4080
         TabIndex        =   8
         Top             =   720
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar pbAvance 
         Height          =   255
         Left            =   9120
         TabIndex        =   9
         Top             =   5280
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSDataGridLib.DataGrid datCliente 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   6240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            MarqueeStyle    =   2
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid datAvon 
         Height          =   3375
         Left            =   9120
         TabIndex        =   11
         Top             =   6120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5953
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   14895
         TabIndex        =   20
         Top             =   9960
         Width           =   45
      End
      Begin VB.Label lblOrdenes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 ordenes"
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
         TabIndex        =   19
         Top             =   9600
         Width           =   840
      End
      Begin VB.Label lblEntrega 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el número de lote a procesar :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   18
         Top             =   5160
         Width           =   1995
      End
      Begin VB.Label lblFlow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes asignados a Usuario"
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
         Left            =   8280
         TabIndex        =   17
         Top             =   360
         Width           =   2190
      End
      Begin VB.Label lblEstacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estación"
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
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 registros"
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
         Left            =   9120
         TabIndex        =   15
         Top             =   9600
         Width           =   915
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes seleccionados para importar a la tabla maestro"
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
         Left            =   1560
         TabIndex        =   14
         Top             =   5895
         Width           =   4395
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes en la base de datos tabla maestro"
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
         Left            =   10560
         TabIndex        =   13
         Top             =   5895
         Width           =   3315
      End
      Begin VB.Label lblUsuarios 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   630
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   7575
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   10935
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   9120
         Shape           =   4  'Rounded Rectangle
         Top             =   5640
         Width           =   5895
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Flow_FrmImportarData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnCliente As New Connection
Dim strDataServidor As String
Dim intRegistros As Integer
Dim rsFlow As New Recordset
Dim rsEntregas As New Recordset
Dim strEntrega As String
Dim strLote As String
Dim intPosicion As Long
Dim fso As New FileSystemObject
Dim Foldes As String

Private Sub cmdRefrescar_Click()
    If rsDatos.State = 1 Then
        rsDatos.Requery
        Llenar_Usuarios
    End If
End Sub
Private Sub cmdRefreshLotes_Click()
    If rsFlow.State = 1 Then
        rsFlow.Requery
        FlowColumnas
        Lista_Entregas
    End If
End Sub
Private Sub cmdSeleccionar_Click()
On Local Error GoTo Err:
     If (rsSQL.State = 1) Then
         If (rsSQL.RecordCount > 0) Then
            rsSQL.MoveFirst
            
            intRespuesta = MsgBox("Desea importar los datos del los lotes seleccionados a la tabla maestra LIQUIDACION", vbQuestion + vbYesNo, "Exportar")
            If intRespuesta = vbYes Then
               
                
                pbAvance.Value = 0
                pbAvance.Min = 0
                pbAvance.Max = rsSQL.RecordCount
                rsSQL.MoveFirst
                Do While Not rsSQL.EOF
                    pbAvance.Value = pbAvance.Value + 1
                    If rsSQL!Estado = "TERMINADO" Then
                            lblMensaje = "Importando datos .."
                            'strEjecutar = "INSERT INTO TB_LIQUIDACION SELECT * FROM " & rsSQL!Data & " where ds_periodo ='" & rsSQL!Periodo & "' and ds_año='" & rsSQL!año & "' and ds_batch='" & rsSQL!batch & "'"
                            strEjecutar = "INSERT INTO TB_LIQUIDACION SELECT * FROM " & rsSQL!Data & " where ds_periodo ='" & rsSQL!Periodo & "' and ds_año='" & rsSQL!año & "' and ds_batch='" & rsSQL!LOTE & "'"
                            cnServer.Execute strEjecutar
                            rsSqlDatos.Requery
                            AvonColumnas
                            
                            'cambio 05/03/2011
                            '''''''''''''''''''''''''''''''
                            strEjecutar = "INSERT INTO TB_LIQUIDACION_HISTORICO SELECT * FROM TB_LIQUIDACION  where ds_periodo ='" & rsSQL!Periodo & "' and ds_año='" & rsSQL!año & "' and ds_batch='" & rsSQL!LOTE & "'"
                            cnServer.Execute strEjecutar
                            '''''''''''''''''''''''''''''''
                            
                            lblMensaje = "Limpiando data cliente .."
                            'strEjecutar = "Delete FROM " & rsSQL!Data & " where ds_periodo ='" & rsSQL!Periodo & "' and ds_año='" & rsSQL!año & "' and ds_batch='" & rsSQL!batch & "'"
                            strEjecutar = "Delete FROM " & rsSQL!Data & " where ds_periodo ='" & rsSQL!Periodo & "' and ds_año='" & rsSQL!año & "' and ds_batch='" & rsSQL!LOTE & "'"
                            cnServer.Execute strEjecutar
                            
                            intPosicion = 0
                            intPosicion = rsSQL.AbsolutePosition
                            strEjecutar = "Update TB_LOTES set ds_estado = 'IMPORTADO' where id_lote=" & rsSQL!id_lote
                            cnServer.Execute strEjecutar
                            'cambio 05/03/2011
                            '''''''''''''''''''''''''''''''''
                            strEjecutar = "INSERT INTO TB_Lote_Historico SELECT * FROM tb_lotes  where ds_periodo ='" & rsSQL!Periodo & "' and ds_año='" & rsSQL!año & "' and ds_batch='" & rsSQL!LOTE & "'"
                            cnServer.Execute strEjecutar
                            '''''''''''''''''''''''''''''''''
                            
                                                    
                        DoEvents
                    End If
                    rsSQL.MoveNext
                Loop
                rsSQL.Requery
                MsgBox "Proceso terminado", vbInformation + vbOKOnly, "Importar data"
                strLote = ""
                strEntrega = ""
                pbAvance.Value = 0
            End If
        End If
    Else
        MsgBox "Seleccione un numero de entrega", vbInformation + vbOKOnly, "Importar data"
    End If
    lblMensaje = ""
    AvonColumnas
    PcColumnas
    cmdRefreshLotes_Click
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Error " & lblMensaje
End Sub
Private Sub cmdSelect_Click()
On Error GoTo Err:

    If rsFlow.State = 1 Then
        rsFlow.Requery
        FlowColumnas
    End If
        If Me.cmbEntrega.Text <> "" Then
            If rsSQL.State = 1 Then rsSQL.Close
            'rsSql.Source = "Select id_lote,ds_campaña as campaña,ds_zona as zona, ds_batch as batch, ds_envio as entrega , lotes.ds_usuario as usuario,nu_imagenes as Registros,lotes.ds_estado as estado, ds_data as data, ds_path as path, lotes.id_lote from LOTES,USUARIOS where LOTES.ds_usuario = USUARIOS.ds_usuario and LOTES.ds_estado = 'TERMINADO' and ds_envio='" & Me.cmbEntrega.Text & "' order by ds_campaña,ds_zona,ds_batch"
            'rsSQL.Source = "Select id_lote,ds_periodo as Periodo,ds_año as Año, ds_batch as Batch,  tb_lotes.ds_usuario as Usuario,nu_imagenes as Registros,tb_lotes.ds_estado as estado, ds_data as data, ds_path as path, tb_lotes.id_lote from TB_LOTES,TB_USUARIOS where TB_LOTES.ds_usuario = TB_USUARIOS.ds_usuario and TB_LOTES.ds_estado = 'TERMINADO' and ds_batch='" & Me.cmbEntrega.Text & "' order by ds_periodo,ds_año,ds_batch"
            rsSQL.Source = "Select id_lote,ds_periodo as Periodo,ds_año as Año, ds_batch as Lote,  tb_lotes.ds_usuario as Usuario,nu_imagenes as Registros,tb_lotes.ds_estado as estado, ds_data as data, ds_path as path, tb_lotes.id_lote from TB_LOTES,TB_USUARIOS where TB_LOTES.ds_usuario = TB_USUARIOS.ds_usuario and TB_LOTES.ds_estado = 'TERMINADO' and ds_batch='" & Me.cmbEntrega.Text & "' order by ds_periodo,ds_año,ds_batch"
            rsSQL.Open
            Set datCliente.DataSource = rsSQL
            PcColumnas
            rsFlow.Requery
        Else
            MsgBox "Seleccione un número de entrega", vbCritical, App.EXEName
        End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub datCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And rsSQL.State = 1 Then
        'PopupMenu mnuOpcion
    End If
End Sub


Private Sub datFlow_HeadClick(ByVal ColIndex As Integer)
    If rsFlow.State = 1 Then
        rsFlow.Close
'        rsFlow.Source = "Select lotes.usuario,caja as caja, lote as lote, imagenes as Registros, lotes.estado,hora_inicio,hora_fin,data from LOTES,USUARIOS where LOTES.usuario = USUARIOS.usuario order by " & ColIndex + 1
        'rsFlow.Source = "Select lotes.ds_usuario as usuario,ds_campaña as campaña,ds_zona as zona, ds_batch as batch , ds_envio as entrega, nu_imagenes as Registros, lotes.ds_estado as estado,fc_hora_inicio as hora_inicio,fc_hora_fin as hora_fin,ds_data as data from LOTES,USUARIOS where LOTES.ds_usuario = USUARIOS.ds_usuario  order by " & ColIndex + 1
        rsFlow.Source = "Select lotes.ds_usuario as usuario,ds_campaña as campaña,ds_zona as zona, ds_batch as Lote , ds_envio as entrega, nu_imagenes as Registros, lotes.ds_estado as estado,fc_hora_inicio as hora_inicio,fc_hora_fin as hora_fin,ds_data as data from LOTES,USUARIOS where LOTES.ds_usuario = USUARIOS.ds_usuario  order by " & ColIndex + 1
        rsFlow.Open
        Set datFlow.DataSource = rsFlow
        FlowColumnas
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errores
    EstadoMenu False
    lblMensaje = ""
    Me.Width = 15420 '14925
    Me.Height = 10680 '10155
    Me.Left = (Flow_Frmpanel.Width / 2) - (Me.Width / 2)
    Me.Top = (Flow_Frmpanel.Height / 2) - (Me.Height / 2)
    
    strLote = ""
    intRegistros = 0
    
    
        
'Todos los lotes
        rsFlow.CursorType = adOpenDynamic
        rsFlow.LockType = adLockBatchOptimistic
        rsFlow.ActiveConnection = cnServer
        
' Usuarios
        rsDatos.CursorType = adOpenDynamic
        rsDatos.LockType = adLockBatchOptimistic
        rsDatos.ActiveConnection = cnServer
        
' Seleccionados
        rsSQL.CursorType = adOpenDynamic
        rsSQL.LockType = adLockBatchOptimistic
        rsSQL.ActiveConnection = cnServer

        
' Entregas
        rsEntregas.CursorType = adOpenDynamic
        rsEntregas.LockType = adLockBatchOptimistic
        rsEntregas.ActiveConnection = cnServer
        
        If rsDatos.State = 1 Then rsDatos.Close
        rsDatos.Source = "Select * from TB_USUARIOS order by ds_usuario"
        rsDatos.Open
        
        If rsFlow.State = 1 Then rsFlow.Close
        
        
        'rsFlow.Source = "Select TB_lotes.ds_usuario as Usuario,ds_periodo as Periodo,ds_año as Año, ds_batch as Batch ,  nu_imagenes as Registros, tb_lotes.ds_estado as Estado,fc_hora_inicio as Hora_inicio,fc_hora_fin as Hora_fin,ds_data as Data from TB_LOTES,TB_USUARIOS where TB_LOTES.ds_usuario = TB_USUARIOS.ds_usuario and tb_lotes.ds_tipo='1'"
        rsFlow.Source = "Select TB_lotes.ds_usuario as Usuario,ds_periodo as Periodo,ds_año as Año, ds_batch as Lote ,  nu_imagenes as Registros, tb_lotes.ds_estado as Estado,fc_hora_inicio as Hora_inicio,fc_hora_fin as Hora_fin,ds_data as Data from TB_LOTES,TB_USUARIOS where TB_LOTES.ds_usuario = TB_USUARIOS.ds_usuario and tb_lotes.ds_tipo='1'"
        rsFlow.Open
        Set datFlow.DataSource = rsFlow
        FlowColumnas
        
        Llenar_Usuarios
        Lista_Entregas
        
        rsSQL.Source = "Select ds_data from tb_Usuarios where ds_usuario = 'ADMIN'"
        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
            strDataServidor = rsSQL!ds_data
                If rsSqlDatos.State = 1 Then rsSqlDatos.Close
                'rsSqlDatos.Source = "Select ds_periodo as Periodo,ds_año as Año, ds_batch as batch, count(*) as Registros from " & rsSQL!ds_data & " group by  ds_periodo,ds_año,ds_batch order by ds_periodo,ds_año,ds_batch"
                rsSqlDatos.Source = "Select ds_periodo as Periodo,ds_año as Año, ds_batch as Lote, count(*) as Registros from " & rsSQL!ds_data & " group by  ds_periodo,ds_año,ds_batch order by ds_periodo,ds_año,ds_batch"
                rsSqlDatos.Open
                Set datAvon.DataSource = rsSqlDatos
                AvonColumnas
            rsSQL.Close
        End If
Exit Sub
errores:
MsgBox Err.Description, vbCritical, Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err:
    EstadoMenu True
    If rsDatos.State = 1 Then rsDatos.Close
    If rsFlow.State = 1 Then rsFlow.Close
    If rsPc.State = 1 Then rsPc.Close
    If rsSQL.State = 1 Then rsSQL.Close
    If rsEntregas.State = 1 Then rsEntregas.Close
    If cnConexion.State = 1 Then cnConexion.Close
    If cnCliente.State = 1 Then cnCliente.Close
    Set cnCliente = Nothing
    Set rsFlow = Nothing
    Set rsEntregas = Nothing
    Exit Sub
Err:
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lstUsuarios_Scroll(Index As Integer)
Dim i As Integer
    For i = 0 To lstUsuarios.Count - 1
        If (i <> Index) Then lstUsuarios(i).TopIndex = lstUsuarios(Index).TopIndex
    Next
End Sub
Private Sub lstUsuarios_Click(Index As Integer)
Dim i As Integer
    For i = 0 To lstUsuarios.Count - 1
        If (i <> Index) Then lstUsuarios(i).ListIndex = lstUsuarios(Index).ListIndex
    Next
End Sub
Sub AvonColumnas()
    If rsSqlDatos.RecordCount > 0 Then
        lblTotal.Caption = Total & " Registros"
    Else
        lblTotal.Caption = "0 Registros"
    End If

    datAvon.Columns(0).Width = 1000
    datAvon.Columns(1).Width = 1000
End Sub
Sub PcColumnas()
    If rsSQL.State = 1 Then
        If rsSQL.RecordCount > 0 Then
            lblOrdenes.Caption = Ordenes & " Registros"
        Else
            lblOrdenes.Caption = "0 Registros"
        End If

        datCliente.Columns(0).Width = 1000
        datCliente.Columns(1).Width = 1000
        datCliente.Columns(2).Width = 1000
        datCliente.Columns(3).Width = 1000
        datCliente.Columns(4).Width = 1450
        datCliente.Columns(5).Width = 1000
    End If
End Sub
Function Total() As Long
    Total = 0
    rsSqlDatos.MoveFirst
    Do While Not rsSqlDatos.EOF
        Total = Total + rsSqlDatos!Registros
        rsSqlDatos.MoveNext
    Loop
    rsSqlDatos.MoveFirst
End Function
Sub Lista_Entregas()
    cmbEntrega.Clear
    If rsEntregas.State = 1 Then rsEntregas.Close
    'rsEntregas.Source = "Select distinct ds_envio from LOTES order by 1"
    rsEntregas.Source = "Select distinct ds_batch from TB_LOTES where ds_estado='TERMINADO'  and ds_tipo='1' order by 1"
    rsEntregas.Open
    If rsEntregas.RecordCount > 0 Then
        rsEntregas.MoveFirst
        Do While Not rsEntregas.EOF
            'cmbEntrega.AddItem rsEntregas!DS_ENVIO
            cmbEntrega.AddItem rsEntregas!DS_batch
            rsEntregas.MoveNext
        Loop
    End If
    If rsEntregas.State = 1 Then rsEntregas.Close
End Sub
Sub Llenar_Usuarios()
    If rsDatos.RecordCount > 0 Then
        lstUsuarios(0).Clear
        lstUsuarios(1).Clear

        rsDatos.MoveFirst
        Do While Not rsDatos.EOF
            lstUsuarios(0).AddItem rsDatos!ds_Usuario
            lstUsuarios(1).AddItem Trim(rsDatos!ds_estacion & " ")
            rsDatos.MoveNext
        Loop
    End If
End Sub
Sub FlowColumnas()
    If rsFlow.State = 1 Then
        datFlow.Columns(0).Width = 1000
        datFlow.Columns(1).Width = 1000
        datFlow.Columns(2).Width = 1000
        datFlow.Columns(3).Width = 1000
        datFlow.Columns(4).Width = 2000
        datFlow.Columns(5).Width = 2000
        datFlow.Columns(6).Width = 2000
      '  datFlow.Columns(7).Width = 2000
    End If
End Sub
Function Ordenes() As Long
    Ordenes = 0
    rsSQL.MoveFirst
    Do While Not rsSQL.EOF
        Ordenes = Ordenes + rsSQL!Registros
        rsSQL.MoveNext
    Loop
    rsSQL.MoveFirst
End Function
Private Sub mnuEntrega_Click()
    Dim varBmk As Variant
    Dim strNuevaEntrega As String * 2
    If datCliente.SelBookmarks.Count > 0 Then
        Do While True
            strEntrega = InputBox("Nuevo Nº de entrega : ", "Cambiar Entrega", rsSQL!Entrega)
            If Trim(strEntrega) = "" Then
                Exit Do
            Else
                If Val(strEntrega) >= 1 And Val(strEntrega) <= 999 Then
                    Exit Do
                Else
                    MsgBox "Ingrese un número del 1 al 999", vbCritical, "Error"
                End If
            End If
        Loop
        If Trim(strEntrega) <> "" Then
            If vbYes = MsgBox("Proceder a cambiar número de entrega por el número : " & strEntrega, vbQuestion + vbYesNo, "Cambiar Nº Entrega") Then
               For Each varBmk In datCliente.SelBookmarks
                   datCliente.Bookmark = varBmk
                   strEjecutar = "Update TB_LOTES set ds_envio = '" & Format(Trim(strEntrega), "000") & "' where ID_LOTE = " & rsSQL!id_lote
                   cnServer.Execute strEjecutar
               Next
                    
               rsFlow.Requery
               FlowColumnas
               rsSQL.Requery
               PcColumnas
               If rsSQL.RecordCount = 0 Then Lista_Entregas
            End If
        End If
    End If
End Sub

Sub RUTINAAVONFAX()
    ' Insertamos a partir de AvonFax a Avon
    strEjecutar = "INSERT INTO AVON SELECT * FROM ExtraerAvonFax_A DATA IN '" & rsSQL!Data & "'" & _
                  " where Mid(data.orig_file, 13, 2) = '" & rsSQL!LOTE & "' and Mid(Data.orig_file, 16, 3) = '" & rsSQL!caja & "' and mid(data.orig_file,20,2) = '" & rsSQL!LOTE & "';"
    cnConexion.Execute strEjecutar
    strEjecutar = "INSERT INTO AVON SELECT * FROM ExtraerAvonFax_B DATA IN '" & rsSQL!Data & "'" & _
                  " where Mid(data.orig_file, 13, 2) = '" & rsSQL!LOTE & "' and Mid(Data.orig_file, 16, 3) = '" & rsSQL!caja & "' and mid(data.orig_file,20,2) = '" & rsSQL!LOTE & "';"
    cnConexion.Execute strEjecutar
    strEjecutar = "INSERT INTO AVON SELECT * FROM ExtraerAvonFax_C DATA IN '" & rsSQL!Data & "'" & _
                  " where Mid(data.orig_file, 13, 2) = '" & rsSQL!LOTE & "' and Mid(Data.orig_file, 16, 3) = '" & rsSQL!caja & "' and mid(data.orig_file,20,2) = '" & rsSQL!LOTE & "';"
    cnConexion.Execute strEjecutar
    strEjecutar = "INSERT INTO AVON SELECT * FROM ExtraerAvonFax_D DATA IN '" & rsSQL!Data & "'" & _
                  " where Mid(data.orig_file, 13, 2) = '" & rsSQL!LOTE & "' and Mid(Data.orig_file, 16, 3) = '" & rsSQL!caja & "' and mid(data.orig_file,20,2) = '" & rsSQL!LOTE & "';"
    cnConexion.Execute strEjecutar
    
    ' Eliminamos los registros de AvonFax ya insertados en Avon
    strEjecutar = "DELETE * FROM AVONFAX DATA IN '" & rsSQL!Data & "'" & _
                  " where Mid(data.orig_file, 13, 2) = '" & rsSQL!LOTE & "' and Mid(Data.orig_file, 16, 3) = '" & rsSQL!caja & "' and mid(data.orig_file,20,2) = '" & rsSQL!LOTE & "';"
    cnConexion.Execute strEjecutar
End Sub
Function CopiarFile(rutaOrigen As String, RutaDestino As String) As Boolean
On Error GoTo errores
CopiarFile = True

Dim fso As FileSystemObject
fso.CopyFolder rutaOrigen, RutaDestino, True

Exit Function
errores:
CopiarFile = False
MsgBox Err.Description
End Function





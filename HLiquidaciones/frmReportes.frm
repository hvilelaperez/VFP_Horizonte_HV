VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ICR_frmReportes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reportes Horizonte"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   Icon            =   "frmReportes.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reportes Horizonte"
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin Crystal.CrystalReport crtreporte 
         Left            =   2760
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
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
         Left            =   4440
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Procesar"
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
         Left            =   480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ComboBox cboreportes 
         Height          =   315
         ItemData        =   "frmReportes.frx":030A
         Left            =   3000
         List            =   "frmReportes.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker DTPinicial 
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52428801
         CurrentDate     =   39980
      End
      Begin MSComCtl2.DTPicker DTPfinal 
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52428801
         CurrentDate     =   39980
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha Inicial"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo de Reporte"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "ICR_frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Dim StrDatos As String
StrDatos = UCase(cboreportes.Text)
  Select Case StrDatos
        Case "REGISTROS ICR"
          Reporte_icr
        Case "IMAGENES DIGITALIZADAS"
          Reporte_imagenes
        Case "ENVIOS REALIZADOS"
          Reporte_Envios
        Case "REGISTROS RECHAZADOS"
          Reporte_Rechazados
        Case "REGISTROS INCOSISTENTES"
          Reporte_Incosistentes
  End Select
  
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
cboreportes.ListIndex = 0
DTPinicial = Now
DTPfinal = Now
 If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    
    
 If rsDatos.State = 1 Then rsDatos.Close
    rsDatos.CursorType = adOpenDynamic
    rsDatos.LockType = adLockBatchOptimistic
    rsDatos.ActiveConnection = cnServer
    
End Sub
 Public Sub Reporte_icr()
 Dim Rst As New ADODB.Recordset
 Dim Xperiodo As String
 Dim Xfecha As String
 Dim Xlote As String
 Dim Xliquidacion As String
 Dim Xbarra As String
 Dim intNumero As Integer
  delete_TemporalIcr
'  Existe_Temporal_icr
'  Create_TemporalDBF
 
 If rsSQL.State = 1 Then rsSQL.Close
 rsSQL.Source = "select * from VW_ICR where fc_hora_fin between '" & Format(DTPinicial, "DD/MM/YYYY") & "' AND '" & Format(DTPfinal, "DD/MM/YYYY") & "' order by ds_barra"
 rsSQL.Open
 If rsSQL.RecordCount > 0 Then
            
    For intNumero = 1 To rsSQL.RecordCount
            Xperiodo = IIf(IsNull(rsSQL.Fields("ds_periodo")), "", (rsSQL.Fields("ds_periodo")))
            Xlote = IIf(IsNull(rsSQL.Fields("ds_batch")), "", (rsSQL.Fields("ds_batch")))
            Xfecha = IIf(IsNull(rsSQL.Fields("fc_hora_fin")), "", (rsSQL.Fields("fc_hora_fin")))
            Xliquidacion = "HO" & IIf(IsNull(rsSQL.Fields("ds_BARRA")), "", Mid((rsSQL.Fields("ds_BARRA")), 1, 9))
            Xbarra = IIf(IsNull(rsSQL.Fields("ds_barra")), "", Mid(rsSQL.Fields("ds_barra"), 10, 5))
        Grabar_TemporalIcr Xperiodo, Xlote, Format(Xfecha, "DD/MM/YYYY"), Xliquidacion, Xbarra
        rsSQL.MoveNext
    Next intNumero
 End If
 
            crtreporte.DataFiles(0) = App.Path & "\tmpprpt.dbf"
            crtreporte.ReportFileName = App.Path & "\Rpticr.rpt"
            crtreporte.Destination = crptToWindow
            crtreporte.Action = 1

 End Sub


Public Sub Reporte_imagenes()
 Dim Rst As New ADODB.Recordset
 Dim Xperiodo As String
 Dim Xfecha As String
 Dim Xlote As String
 Dim Xliquidacion As String
 Dim Xbarra As String
 Dim intNumero As Integer
 Dim Ximagenes As Integer
 Dim Ximagenest As Integer
 Dim Xanno As String
  delete_TemporalIcr
'  Existe_Temporal_icr
'  Create_TemporalDBF
 
 If rsSQL.State = 1 Then rsSQL.Close
 rsSQL.Source = "select * from tb_lote_historico where  fc_hora_fin between '" & Format(DTPinicial, "DD/MM/YYYY") & "' AND '" & Format(DTPfinal, "DD/MM/YYYY") & "'"
 rsSQL.Open
 If rsSQL.RecordCount > 0 Then
            
    For intNumero = 1 To rsSQL.RecordCount
            Xperiodo = IIf(IsNull(rsSQL.Fields("ds_periodo")), "", (rsSQL.Fields("ds_periodo")))
            Xlote = IIf(IsNull(rsSQL.Fields("ds_batch")), "", (rsSQL.Fields("ds_batch")))
            Xfecha = IIf(IsNull(rsSQL.Fields("fc_hora_fin")), "", (rsSQL.Fields("fc_hora_fin")))
            
            If rsDatos.State = 1 Then rsDatos.Close
                rsDatos.Source = "select count(*) from tb_liquidacion_historico where  ds_batch ='" & Xlote & "'"
                rsDatos.Open
                If rsDatos.RecordCount > 0 Then
                Ximagenes = rsDatos.Fields(0)
                End If
            Ximagenest = IIf(IsNull(rsSQL.Fields("nu_imagenes")), "", (rsSQL.Fields("nu_imagenes")))
            Xanno = IIf(IsNull(rsSQL.Fields("ds_año")), "", (rsSQL.Fields("ds_año")))
        Grabar_Temporal_Imagenes Xperiodo, Xlote, Xfecha, Ximagenes, Ximagenest, Xanno
        rsSQL.MoveNext
    Next intNumero
 End If
 
            crtreporte.DataFiles(0) = App.Path & "\tmpprpt.dbf"
            crtreporte.ReportFileName = App.Path & "\Rptimg.rpt"
            crtreporte.Destination = crptToWindow
            crtreporte.Action = 1

 End Sub

Public Sub Reporte_Envios()
 Dim Rst As New ADODB.Recordset
 Dim Xperiodo As String
 Dim Xfecha As String
 Dim Xlote As String
 Dim Xliquidacion As String
 Dim Xbarra As String
 Dim intNumero As Integer
 Dim Ximagenes As Integer
 Dim Xanno As String
 Dim Xorden As String
  delete_TemporalIcr
'  Existe_Temporal_icr
'  Create_TemporalDBF
 
 If rsSQL.State = 1 Then rsSQL.Close
 rsSQL.Source = "select * from tb_envios where  fecha between '" & Format(DTPinicial, "DD/MM/YYYY") & "' AND '" & Format(DTPfinal, "DD/MM/YYYY") & "'"
 rsSQL.Open
 If rsSQL.RecordCount > 0 Then
            
    For intNumero = 1 To rsSQL.RecordCount
            Xorden = IIf(IsNull(rsSQL.Fields("orden")), "", (rsSQL.Fields("orden")))
            Xlote = IIf(IsNull(rsSQL.Fields("lotes")), "", (rsSQL.Fields("lotes")))
            Xfecha = IIf(IsNull(rsSQL.Fields("fecha")), "", (rsSQL.Fields("fecha")))
            
        Grabar_Temporal_Envios Xorden, Xlote, Xfecha
        rsSQL.MoveNext
    Next intNumero
 End If
 
            crtreporte.DataFiles(0) = App.Path & "\tmpprpt.dbf"
            crtreporte.ReportFileName = App.Path & "\Rptenvio.rpt"
            crtreporte.Destination = crptToWindow
            crtreporte.Action = 1

 End Sub


Public Sub Reporte_Rechazados()
 Dim Rst As New ADODB.Recordset
 Dim Xperiodo As String
 Dim Xfecha As String
 Dim Xlote As String
 Dim Xliquidacion As String
 Dim Xbarra As String
 Dim intNumero As Integer
 
 'Dim Xliquidacion As String
 'Dim Xlote As String
 Dim xruc As String
 Dim Xrazon As String
 Dim Xcomentario As String
 Dim Xcupps As String
 'Dim Xfecha As String
 
  delete_TemporalIcr
'  Existe_Temporal_icr
'  Create_TemporalDBF
 
 If rsSQL.State = 1 Then rsSQL.Close
 rsSQL.Source = "select * from VW_Reporte_Rechazo where (fc_hora_fin between '" & Format(DTPinicial, "DD/MM/YYYY") & "' AND '" & Format(DTPfinal, "DD/MM/YYYY") & "') and ds_estado='4'"
 rsSQL.Open
 If rsSQL.RecordCount > 0 Then
            
    For intNumero = 1 To rsSQL.RecordCount
            Xfecha = IIf(IsNull(rsSQL.Fields("fc_hora_fin")), "", (rsSQL.Fields("fc_hora_fin")))
            Xliquidacion = IIf(IsNull(rsSQL.Fields("nroliqpag")), "", Mid(rsSQL.Fields("nroliqpag"), 1, 11))
            Xlote = IIf(IsNull(rsSQL.Fields("ds_batch")), "", (rsSQL.Fields("ds_batch")))
            xruc = IIf(IsNull(rsSQL.Fields("ds_ruc")), "", (rsSQL.Fields("ds_ruc")))
            Xrazon = IIf(IsNull(rsSQL.Fields("ds_raz_soc")), "", (rsSQL.Fields("ds_raz_soc")))
            Xcomentario = IIf(IsNull(rsSQL.Fields("ds_comentario")), "", (rsSQL.Fields("ds_comentario")))
            Xcupps = IIf(IsNull(rsSQL.Fields("ds_cuspp")), "", (rsSQL.Fields("ds_cuspp")))
            
        Grabar_Rechazos Xfecha, Xliquidacion, Xlote, xruc, Xrazon, Xcomentario, Xcupps
        rsSQL.MoveNext
    Next intNumero
 End If
 
            crtreporte.DataFiles(0) = App.Path & "\tmpprpt.dbf"
            crtreporte.ReportFileName = App.Path & "\RptRECHAZOS.rpt"
            crtreporte.Destination = crptToWindow
            crtreporte.Action = 1

 End Sub

Public Sub Reporte_Incosistentes()
 Dim Rst As New ADODB.Recordset
 Dim Xperiodo As String
 Dim Xfecha As String
 Dim Xlote As String
 Dim Xliquidacion As String
 Dim Xbarra As String
 Dim intNumero As Integer
 
 'Dim Xliquidacion As String
 'Dim Xlote As String
 Dim xruc As String
 Dim Xrazon As String
 Dim Xcomentario As String
 Dim Xcupps As String
 'Dim Xfecha As String
 
  delete_TemporalIcr
'  Existe_Temporal_icr
'  Create_TemporalDBF
 
 If rsSQL.State = 1 Then rsSQL.Close
 rsSQL.Source = "select * from VW_Reporte_Rechazo where (fc_hora_fin between '" & Format(DTPinicial, "DD/MM/YYYY") & "' AND '" & Format(DTPfinal, "DD/MM/YYYY") & "') and ds_estado='3'"
 rsSQL.Open
 If rsSQL.RecordCount > 0 Then
            
    For intNumero = 1 To rsSQL.RecordCount
            Xfecha = IIf(IsNull(rsSQL.Fields("fc_hora_fin")), "", (rsSQL.Fields("fc_hora_fin")))
            Xliquidacion = IIf(IsNull(rsSQL.Fields("nroliqpag")), "", Mid(rsSQL.Fields("nroliqpag"), 1, 11))
            Xlote = IIf(IsNull(rsSQL.Fields("ds_batch")), "", (rsSQL.Fields("ds_batch")))
            xruc = IIf(IsNull(rsSQL.Fields("ds_ruc")), "", (rsSQL.Fields("ds_ruc")))
            Xrazon = IIf(IsNull(rsSQL.Fields("ds_raz_soc")), "", (rsSQL.Fields("ds_raz_soc")))
            Xcomentario = IIf(IsNull(rsSQL.Fields("ds_comentario")), "", (rsSQL.Fields("ds_comentario")))
            Xcupps = IIf(IsNull(rsSQL.Fields("ds_cuspp")), "", (rsSQL.Fields("ds_cuspp")))
            
        Grabar_Rechazos Xfecha, Xliquidacion, Xlote, xruc, Xrazon, Xcomentario, Xcupps
        rsSQL.MoveNext
    Next intNumero
 End If
 
            crtreporte.DataFiles(0) = App.Path & "\tmpprpt.dbf"
            crtreporte.ReportFileName = App.Path & "\RptINCOSISTENCIAS.rpt"
            crtreporte.Destination = crptToWindow
            crtreporte.Action = 1

 End Sub


VERSION 5.00
Begin VB.Form frmRepErrorAsesor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de errores  por Asesor"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   240
      TabIndex        =   23
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton optHuella 
         Caption         =   "Huella y Firma"
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optOtro 
         Caption         =   "Otros Errores"
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame6 
      Height          =   5775
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   5895
      Begin VB.OptionButton OptResumen 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton optDetallado 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   3615
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   5655
         Begin VB.Frame Frame3 
            Caption         =   "Afiliación"
            ForeColor       =   &H00000000&
            Height          =   1455
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   4935
            Begin VB.TextBox txtAfiFin 
               Height          =   375
               Left            =   2160
               TabIndex        =   3
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtAfiIni 
               Height          =   375
               Left            =   2160
               TabIndex        =   2
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "Fecha Fin"
               Height          =   255
               Left            =   840
               TabIndex        =   21
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "Fecha Inicial"
               Height          =   255
               Left            =   840
               TabIndex        =   20
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Traspaso"
            ForeColor       =   &H00000000&
            Height          =   1455
            Left            =   360
            TabIndex        =   16
            Top             =   2040
            Width           =   4935
            Begin VB.TextBox txtTraFin 
               Height          =   375
               Left            =   2160
               TabIndex        =   5
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtTraIni 
               Height          =   375
               Left            =   2160
               TabIndex        =   4
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label6 
               Caption         =   "Fecha Fin"
               Height          =   255
               Left            =   840
               TabIndex        =   18
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fecha Inicial"
               Height          =   255
               Left            =   840
               TabIndex        =   17
               Top             =   360
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txtFecFin 
            Height          =   375
            Left            =   2400
            TabIndex        =   1
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtFecIni 
            Height          =   375
            Left            =   2400
            TabIndex        =   0
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Fin"
            Height          =   255
            Left            =   840
            TabIndex        =   14
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Inicial"
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   5895
      Begin VB.CommandButton cmdExpoExcel 
         Caption         =   "&Expo Excel"
         Height          =   495
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Visualizar"
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRepErrorAsesor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub VerReporte()

'        SQLstr = " SHAPE {Select * from TmpCabErr order by 1 }  AS Cabecera APPEND ({select * from TmpDetErr order by 1, 2, 3 , 4}  AS Detalle RELATE 'CodSup' TO 'CodSup') AS Detalle"
'        RsRepo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
'
'        Set DtRErroresAsesor.DataSource = RsRepo
'
'
'         With DtRErroresAsesor.Sections("Detalle")
'
'            .Controls("txtEnvio").DataMember = "Detalle"
'            .Controls("txtagencia").DataMember = "Detalle"
'            .Controls("txtfecsal").DataMember = "Detalle"
'            .Controls("txtnumsol").DataMember = "Detalle"
'            .Controls("txtDespromo").DataMember = "Detalle"
'            .Controls("txtCodPromo").DataMember = "Detalle"
'            .Controls("txtCodProAd").DataMember = "Detalle"
'            .Controls("txtindDoc").DataMember = "Detalle"
'            .Controls("txterror").DataMember = "Detalle"
'            .Controls("txtDescripcion").DataMember = "Detalle"
'            .Controls("txtNomafi").DataMember = "Detalle"
'
'
'            'Establece el nombre de cada campo hijo para cada control
'            .Controls("txtEnvio").DataField = "envio"
'            .Controls("txtagencia").DataField = "agencia"
'            .Controls("txtfecsal").DataField = "fecsal"
'            .Controls("txtnumsol").DataField = "numsol"
'            .Controls("txtDespromo").DataField = "Despromo"
'            .Controls("txtCodPromo").DataField = "Codpromo"
'            .Controls("txtCodProAd").DataField = "CodProAd"
'            .Controls("txtindDoc").DataField = "indDoc"
'            .Controls("txterror").DataField = "error"
'            .Controls("txtDescripcion").DataField = "Descripcion"
'            .Controls("txtNomafi").DataField = "Nomafi"
'
'
'        End With
'        DtRErroresAsesor.Sections("PageHeader").Controls("etqDesHas").Caption = " Desde " & txtFecIni & " Hasta " & txtFecFin
        
         'DETraspasos.Connections.Item(1).ConnectionString
         
        
         'SHAPE {set dateformat dmy
         'exec RRepErrorAsesor '16/07/2008','16/07/2008'}  AS Command1 COMPUTE Command1 BY 'CodSup'
         
         Dim RsReport As Recordset
         Set RsReport = New Recordset
            
         
    If optDetallado.Value Then
            
            Dim Report3 As CRErrorPromotor
            Set Report3 = New CRErrorPromotor
            
            With RsReport
                If .State = 1 Then .Close
                            
                '.Open "exec pRepTransImagenes " & vbgCliente & ", " & vbgProyecto & ", '" & frmRepTransacciones.DTPicker1.Value & "' , '" & frmRepTransacciones.DTPicker2.Value & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
                
                If optOtro.Value Then
                    sqltexto = "exec RRepErrorAsesorIINormales2  '" & txtFecIni & "','" & txtFecFin & "' "
                Else
                    sqltexto = "exec RRepErrorAsesorIIFirmasHuellas2  '" & txtFecIni & "','" & txtFecFin & "' "
                End If
                                        
                .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
                                                                            
                If .RecordCount > 0 Then
                
                    Report3.Database.SetDataSource RsReport
                    MousePointer = vbNormal
                                    
                    Report3.txtRangoFec.SetText ("Desde " & txtFecIni & " Hasta " & txtFecFin & " ")
                    
                    With frmRepCrView
                                                                    
                        .CRViewer91.ReportSource = Report3
                        .CRViewer91.ViewReport
                        .Show 1
                    
                    End With
                    
                    
                End If
            End With
    Else
    
            
            Dim Report4 As CRErrorPromotor
            Set Report4 = New CRErrorPromotor
            
            With RsReport
                If .State = 1 Then .Close
                            
                '.Open "exec pRepTransImagenes " & vbgCliente & ", " & vbgProyecto & ", '" & frmRepTransacciones.DTPicker1.Value & "' , '" & frmRepTransacciones.DTPicker2.Value & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
                
                sqltexto = "exec pRepResumenPromotorII  '" & txtFecIni & "','" & txtFecFin & "' "
                                        
                .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
                            
                
         
                If .RecordCount > 0 Then
                
                    Report4.Database.SetDataSource RsReport
                    MousePointer = vbNormal
                                    
                    'Report4.txtRangoFec.SetText ("Desde " & txtFecIni & " Hasta " & txtFecFin & " ")
                    
                    With frmRepCrView
                                                                    
                        .CRViewer91.ReportSource = Report3
                        .CRViewer91.ViewReport
                        .Show 1
                    
                    End With
                    
                    
                End If
            End With
            
    End If
      ' DtRErroresAsesor.PrintReport True
        'DtRErroresAsesor.Refresh
        
        
        Set Report3 = Nothing
        Set RsRepo = Nothing
End Sub

Private Sub cmdAceptar_Click()
VerReporte
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdExpoExcel_Click()
If optDetallado.Value Then
    ExpoExcel
Else
    ExpoExcelResumen
End If
End Sub

Private Sub optDetallado_Click()
If optDetallado.Value Then
    Frame5.Enabled = False
    Frame1.Enabled = True
    txtFecIni.SetFocus
    cmdAceptar.Enabled = True
End If

End Sub

Private Sub OptResumen_Click()
If OptResumen.Value Then
    Frame1.Enabled = False
    Frame5.Enabled = True
    txtAfiIni.SetFocus
    cmdAceptar.Enabled = False
End If
End Sub



Private Sub txtAfiFin_KeyPress(KeyAscii As Integer)
FVFecha txtAfiFin, KeyAscii
End Sub

Private Sub txtAfiFin_LostFocus()
ValidaFecha txtAfiFin
End Sub

Private Sub txtAfiIni_KeyPress(KeyAscii As Integer)
FVFecha txtAfiIni, KeyAscii
End Sub

Private Sub txtAfiIni_LostFocus()
ValidaFecha txtAfiIni
End Sub

Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
FVFecha txtFecFin, KeyAscii
End Sub

Private Sub TxtFecFin_LostFocus()
ValidaFecha txtFecFin
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
FVFecha txtFecIni, KeyAscii
End Sub

Private Sub TxtFecIni_LostFocus()
ValidaFecha txtFecIni
End Sub
Sub ExpoExcel()
On Error GoTo ERRORES

Dim rsRepExcel As Recordset
Set rsRepExcel = New ADODB.Recordset


Dim cmdErro As ADODB.Command
Set cmdErro = New ADODB.Command
        
'Realiza Proceso
With cmdErro
            .ActiveConnection = Con.Cnx
            If optOtro.Value Then
            'Muestra los Otros Tipos de Errores
            '.CommandText = "RRepErrorAsesorIIOtros ?,?"
            .CommandText = "RRepErrorAsesorIINormales2 ?,?"
            Else
             'Muestra Suspendidos
            '.CommandText = "RRepErrorAsesorIIHuella ?,?"
            .CommandText = "RRepErrorAsesorIIFirmasHuellas2 ?,?"
            End If
            .Parameters(0).Value = txtFecIni
            .Parameters(1).Value = txtFecFin
            .Execute , , adExecuteNoRecords
End With


Dim cabecera As String
Dim caddetalle As String
cabecera = " Envio " & vbTab & " cFecsal " & vbTab & "Agen" & vbTab & "Supervisor(Nombres y Apellidos)" & vbTab & "Promotor" & vbTab & "CodSBS" & vbTab & "CodAdm" & vbTab & "NumSol" & vbTab & " Nom. Afiliado " & vbTab & "Tdoc" & vbTab & "CodErr" & vbTab & "Descripcion" & vbTab & "Supervisor" & vbTab & "Detalle Errores"

'Exporta archivo
With rsRepExcel

   .Open "select * from TmpDetErr2 order by  2, 3, 4", Con.Cnx, adOpenStatic, adLockReadOnly

    Open App.Path & "\expo\RepErrorProm_" & Format(gDate, "yyyymmdd") & IIf(optOtro.Value, "_OtrosErrores", "Errores firmas y huellas") & ".xls" For Output As #1
    Print #1, cabecera
    
    If .RecordCount > 0 Then
        Do While Not .EOF
        
            DetError = ""
            If Trim(.Fields("DetError")) <> "" Then
            DetError = Left(.Fields("DetError"), Len(.Fields("DetError")) - 2)
            Else
            DetError = ""
            End If
             
            
            
            caddetalle = "'" & .Fields("Envio") & vbTab & .Fields("Fecsal") & vbTab & "'" & .Fields("agencia") & vbTab & .Fields("Nomsup") & vbTab
            caddetalle = caddetalle & .Fields("Despromo") & vbTab & "'" & .Fields("Codpromo") & vbTab
            caddetalle = caddetalle & "'" & .Fields("CodProAd") & vbTab & "'" & .Fields("numsol") & vbTab & .Fields("Nomafi") & vbTab
            caddetalle = caddetalle & .Fields("IndDoc") & vbTab & "'" & .Fields("error") & vbTab
            caddetalle = caddetalle & .Fields("Descripcion") & vbTab & .Fields("CodSup") & vbTab & DetError
            Print #1, caddetalle
            .MoveNext
        Loop
    
    End If
    Close #1
    
    MsgBox " Archivo exportado en " & App.Path & "\expo\RepErrorProm" & IIf(optOtro.Value, "_OtrosErrores", "_Errores firmas y huellas") & ".xls", 64, Caption
    
End With



Exit Sub
ERRORES:
Close #1
MsgBox Err.Description, vbCritical, Caption
End Sub


Sub ExpoExcelResumen()
On Error GoTo ERRORES

Dim rsRepExcel As Recordset
Set rsRepExcel = New ADODB.Recordset


Dim cmdErro As ADODB.Command
Set cmdErro = New ADODB.Command
        
'Realiza Proceso
With cmdErro
            .ActiveConnection = Con.Cnx
            .CommandText = "pRepResumenPromotor ?,?,?,?"
            .Parameters("@feciniT").Value = txtTraIni
            .Parameters("@fecfinT").Value = txtTraFin
            .Parameters("@feciniC").Value = txtAfiIni
            .Parameters("@fecfinC").Value = txtAfiFin
                                
            Set rsRepExcel = .Execute
End With


Dim cabecera As String
Dim caddetalle As String
'cabecera = " Envio " & vbTab & " cFecsal " & vbTab & "Agen" & vbTab & "Supervisor(Nombres y Apellidos)" & vbTab & "Promotor" & vbTab & "CodSBS" & vbTab & "CodAdm" & vbTab & "NumSol" & vbTab & " Nom. Afiliado " & vbTab & "Tdoc" & vbTab & "CodErr" & vbTab & "Descripcion" & vbTab & "Supervisor"
cabecera = " Cod Promo  " & vbTab & " DesPromo  " & vbTab & "  Cant Error Afi " & vbTab & " Total de Afil present " & vbTab & "  Total de Afil con error " & vbTab & " Afil Reingresos " & vbTab & " Cant Error Tras " & vbTab & " Total de Trasp present " & vbTab & " Total  Trasp con error " & vbTab & "  Trasp Reingresos " & vbTab & " Cant Error Tot " & vbTab & " Total de Docs  presentadas " & vbTab & "  Total de Docs con error " & vbTab & " Reprocesos o Reingresos " & vbTab & " Cod Sup " & vbTab & " Nombre Supervisor " & vbTab & "  Cod Interno"

'Exporta archivo
With rsRepExcel

   '.Open "select * from TmpDetErr order by  2, 3, 4", Con.Cnx, adOpenStatic, adLockReadOnly

    Open App.Path & "\expo\Resumen Errores x Promotor_" & Format(gDate, "yyyymmdd") & ".xls" For Output As #1
    Print #1, cabecera
    
    If .RecordCount > 0 Then
        Do While Not .EOF
            caddetalle = "'" & .Fields("CodPromo") & vbTab & .Fields("Despromo") & vbTab & .Fields("Nerr_Afi") & vbTab & .Fields("tot_Afi") & vbTab
            caddetalle = caddetalle & .Fields("Afi_err") & vbTab & .Fields("reing_Afi") & vbTab
            caddetalle = caddetalle & .Fields("Nerr_Tra") & vbTab & .Fields("tot_Tra") & vbTab & .Fields("Tra_err") & vbTab
            caddetalle = caddetalle & .Fields("reing_Tra") & vbTab & .Fields("N_error") & vbTab
            caddetalle = caddetalle & .Fields("tot_sol") & vbTab & .Fields("sol_err") & vbTab & .Fields("reingre") & vbTab & .Fields("CodSup") & vbTab & .Fields("Nomsup") & vbTab & "'" & .Fields("CodInPro")
            Print #1, caddetalle
            .MoveNext
        Loop
    
    End If
    Close #1
    
    MsgBox " Archivo exportado en " & App.Path & "\expo\RepErrorProm.xls", 64, Caption
    
End With



Exit Sub
ERRORES:
MsgBox Err.Description, vbCritical, Caption
End Sub

Private Sub txtTraFin_KeyPress(KeyAscii As Integer)
FVFecha txtTraFin, KeyAscii
End Sub

Private Sub txtTraFin_LostFocus()
ValidaFecha txtTraFin
End Sub

Private Sub txtTraIni_KeyPress(KeyAscii As Integer)
FVFecha txtTraIni, KeyAscii
End Sub

Private Sub txtTraIni_LostFocus()
ValidaFecha txtTraIni
End Sub

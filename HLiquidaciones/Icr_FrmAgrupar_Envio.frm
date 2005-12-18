VERSION 5.00
Begin VB.Form Icr_FrmAgrupar_Envio 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Agrupar Envios"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   Icon            =   "Icr_FrmAgrupar_Envio.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Tipo Archivo"
      Height          =   855
      Left            =   3000
      TabIndex        =   9
      Top             =   720
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000009&
         Caption         =   "Texto"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
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
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar Lotes Agrupados"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton CMDACEPTAR 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.ListBox lstselec 
      Height          =   1815
      Left            =   5160
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox LstLotes 
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Lblcargo 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lotes a Agrupar"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lotes Procesados"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblgeneral 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   7575
   End
End
Attribute VB_Name = "Icr_FrmAgrupar_Envio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim StrArchivoDESTINO1 As String
Dim StrArchivoDESTINO11 As String
Dim StrArchivoDESTINO2 As String
Dim StrArchivoDESTINO22 As String
Dim StrArchivoDESTINO3 As String
Dim StrArchivoDESTINO33 As String
Dim StrArchivoDESTINO4 As String
Dim StrArchivoDESTINO44 As String
Dim StrArchivoDESTINO5 As String
Dim StrArchivoDESTINO55 As String
Dim StrArchivoDESTINO6 As String
Dim StrArchivoDESTINO66 As String
'Dim StrArchivoDESTINO_D As String


Dim strCuenta As String
Dim StrTraLaboral As String
Dim strLp As String
Dim strCorrelativo As String
Dim strpagina As String
Dim strRegistros As String
Dim StrEnvio As String
Dim strFechaEm As String
Dim stradestino As String

Dim X  As Integer
Dim strsqlp As String
Dim StrArchivo_D As String
Dim StrArchivo_Cargo As String
Dim rsZonas As New Recordset

Dim StrLote_Agrupado As String
On Error GoTo Err
    StrLote_Agrupado = ""
    rsZonas.CursorType = adOpenDynamic
    rsZonas.LockType = adLockBatchOptimistic
    rsZonas.ActiveConnection = cnServer
    
     If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "Select * from tb_Parametros"
        rsSQL.Open
        
        strCampañaActual = "" & rsSQL!ds_periodo
        strZonaActual = "" & rsSQL!ds_año
        strDirectorio = "" & rsSQL!ds_ftpaxis
        intProductos = 0 & rsSQL!nu_cl
If lstselec.ListCount > 0 Then
Borrar_Tempo_Sql_Archivos

        StrArchivoDESTINO1 = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_MOVPERDI" & ".XLS"
        StrArchivoDESTINO11 = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_MOVPERDI" & ".TXT"
        StrArchivoDESTINO2 = strDirectorio & "\SINCERA_" & Format(Now, "YYMMDD") & ".XLS"
        StrArchivoDESTINO22 = strDirectorio & "\SINCERA_" & Format(Now, "YYMMDD") & ".TXT"
        StrArchivoDESTINO3 = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_OBSERVADOS" & ".XLS"
        StrArchivoDESTINO33 = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_OBSERVADOS" & ".TXT"
        StrArchivoDESTINO4 = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_RECHAZOS" & ".XLS"
        StrArchivoDESTINO44 = strDirectorio & "\F" & Format(Now, "YYMMDD") & "_RECHAZOS" & ".TXT"
        StrArchivoDESTINO5 = strDirectorio & "\DATOS_GENERALES_" & Format(Now, "YYMMDD") & ".XLS"
        StrArchivoDESTINO55 = strDirectorio & "\DATOS_GENERALES_" & Format(Now, "YYMMDD") & ".TXT"
        StrArchivoDESTINO6 = strDirectorio & "\CONSOLIDADO_" & Format(Now, "YYMMDD") & ".XLS"
        StrArchivoDESTINO66 = strDirectorio & "\CONSOLIDADO_" & Format(Now, "YYMMDD") & ".TXT"

        lngFile = FreeFile
        Open StrArchivoDESTINO11 For Output As #lngFile
        Print #lngFile, "CUENTA AFILIADO"; vbTab; "CUSPP AFILIADO"; vbTab; "RUC EMPLEADOR"; vbTab; "TIPO MOV LABORAL"; vbTab; "FECHA INICIO MOV"; vbTab; "FECHA FIN MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO DEVENGE"; vbTab; "TIPO TRAMO LABORAL"; vbTab; "NRO REQUERIMIENTO"; vbTab; "FECHA"
        
        lngFile1 = FreeFile
        Open StrArchivoDESTINO22 For Output As #lngFile1
        Print #lngFile1, "RUC EMPLEADOR"; vbTab; "CUSPP AFILIADO"; vbTab; "PERIODO DEVENGE"; vbTab; "CUENTA AFILIADO "; vbTab; "REMUNERACION"; vbTab; "MOTIVO SINCERAMIENTO"
        
        lngFile2 = FreeFile
        Open StrArchivoDESTINO33 For Output As #lngFile2
        Print #lngFile2, "CUSPP"; vbTab; "RUC"; vbTab; "TIPO MOV LABORAL"; vbTab; "FECHA INICIO MOV"; vbTab; "FECHA FIN MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO DEVENGE"; vbTab; "TIPO TRAMO LABORAL"; vbTab; "NRO REQUERIMIENTO"; vbTab; "FECHA"; vbTab; "OBSERVACIONES"
        
        lngFile3 = FreeFile
        Open StrArchivoDESTINO44 For Output As #lngFile3
        Print #lngFile3, "CUENTA"; vbTab; "CUSPP"; vbTab; "RUC"; vbTab; "TIPO MOV LABORAL"; vbTab; "FECHA INICIO MOV"; vbTab; "FECHA FIN MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO DEVENGE"; vbTab; "TIPO TRAMO LABORAL"; vbTab; "NRO REQUERIMIENTO"; vbTab; "FECHA"; vbTab; "MOTIVO RECHAZOS"
        
        lngFile4 = FreeFile
        Open StrArchivoDESTINO55 For Output As #lngFile4
        Print #lngFile4, "RUC"; vbTab; "DIRECCION"; vbTab; "E-MAIL"; vbTab; "TELEFONO"; vbTab; "DOC_IDENTIDAD"; vbTab; "NOMBRE_APELLIDO"; vbTab; "FECHA"; vbTab; "CAMBIO RSOCIAL"; vbTab; "ENVIO"
        
        lngFile5 = FreeFile
        Open StrArchivoDESTINO66 For Output As #lngFile5
        Print #lngFile5, "CUENTA"; vbTab; "cuspp"; vbTab; "Ruc"; vbTab; "TIPO MOV LABORAL"; vbTab; "FECHA INICIO MOV"; vbTab; "FECHA FIN MOV"; vbTab; "CORRELATIVO"; vbTab; "PERIODO DEVENGE"; vbTab; "TIPO TRAMO LABORAL"; vbTab; "NRO REQUERIMIENTO"; vbTab; "FECHA"; vbTab; "NRO LP"; vbTab; "PAGINA"; vbTab; "REGISTROS"; vbTab; "CAMBIO RSOCIAL"; vbTab; "OBSERVACIONES"; vbTab; "ENVIO"; vbTab; "DIRECCION"; vbTab; "E-MAIL"; vbTab; "FECHA EMISION"; vbTab; "ARCHIVO DE DESTINO"
        
For X = 1 To lstselec.ListCount
        
        StrLote_Agrupado = StrLote_Agrupado & lstselec.List(X - 1)
        
        'strCadenaSQL = "drop table [" & "A1" & cmbBatch & "]"
        'cnServer.Execute strCadenaSQL
        'strCadenaSQL = "Select cuenta as [CUENTA AFILIADO],CUSPP AS [CUSPP AFILIADO],RUC AS [RUC EMPLEADOR],COD_DESCARGO AS [TIPO MOV LABORAL], FECHA_INICIO AS [FECHA INICIO MOV], FECHA_CESE AS [FECHA FIN MOV],CORRELATIVO ,PERIODO AS [PERIODO DEVENGE],TIPOTRAMO AS [TIPO TRAMO LABORAL],NROREQUERIMIENTO AS [NRO REQUERIMIENTO],FECHA_PROCESO AS [FECHA] , ARCHIVO  into " & "A1" & cmbBatch & " from Tempo order by id_tempo"
        
'If rsSQL.State = 1 Then rsSQL.Close
'        rsSQL.Source = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'LIQUIDACIONHORIZONTE' AND TABLE_NAME = '" & "A1" & lstselec.List(X - 1) & "'"
'        rsSQL.Open
'    If rsSQL.Fields(0) > 0 Then
        
        strCadenaSQL = "insert into tempoa1 select * from " & "A1" & lstselec.List(X - 1)
        cnServer.Execute strCadenaSQL
        ''''''''''''''''''''''''''''''''''ARCHIVO TEXTO 1''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If rsSQL.State = 1 Then rsSQL.Close

        rsSQL.Source = "select * from " & "A1" & lstselec.List(X - 1)
        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
        
         Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL![RUC EMPLEADOR])
              'StrRaSocArchivo = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              'StrComentario = ("" & rsSQL!ds_comentario)
              StrCodigo = ("" & rsSQL![TIPO MOV LABORAL])
              Strcuspp = Replace(("" & rsSQL![CUSPP AFILIADO]), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL![PERIODO DEVENGE])
              StrFechainicio = ("" & rsSQL![FECHA INICIO MOV])
              StrFechaCese = ("" & rsSQL![FECHA FIN MOV])
              strFecha = ("" & rsSQL![FECHA])
              'StrArchivotif = ("" & rsSQL!ds_archivo)
              StrRequerimiento = ("" & rsSQL![NRO REQUERIMIENTO])
              Print #lngFile, ""; vbTab; Strcuspp; vbTab; StrRucArchivo; vbTab; StrCodigo; vbTab; StrFechainicio; vbTab; StrFechaCese; vbTab; ""; vbTab; strPeriodo; vbTab; "N"; vbTab; Format(StrRequerimiento, "000000000000"); vbTab; strFecha
         rsSQL.MoveNext
         Loop
        End If
          
'End If
        
'If rsSQL.State = 1 Then rsSQL.Close
'        rsSQL.Source = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'LIQUIDACIONHORIZONTE' AND TABLE_NAME = '" & "A2" & lstselec.List(X - 1) & "'"
'        rsSQL.Open
'   If rsSQL.Fields(0) > 0 Then
        
        
        strCadenaSQL = "insert into tempoa2 select * from " & "A2" & lstselec.List(X - 1)
        cnServer.Execute strCadenaSQL
        ''''''''''''''''''''''''''''''''''ARCHIVO TEXTO 2''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "select * from " & "A2" & lstselec.List(X - 1)
        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
        
         Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL![RUC EMPLEADOR])
              'StrRemuneracion = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              StrRemuneracion = ("" & rsSQL!REMUNERACION)
              'StrCodigo = ("" & rsSQL![TIPO MOV LABORAL])
              Strcuspp = Replace(("" & rsSQL![CUSPP AFILIADO]), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL![PERIODO DEVENGE])
              'StrFechainicio = ("" & rsSQL![FECHA INICIO MOV])
              'StrFechaCese = ("" & rsSQL![FECHA FIN MOV])
              'StrArchivotif = ("" & rsSQL!ds_archivo)
              'StrRequerimiento = ("" & rsSQL![NRO REQUERIMIENTO])
              'strFecha = ("" & rsSQL![FECHA])
              Print #lngFile1, StrRucArchivo; vbTab; Strcuspp; vbTab; strPeriodo; vbTab; ""; vbTab; StrRemuneracion; vbTab; "10"
         rsSQL.MoveNext
         Loop
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
 'End If
        
'  If rsSQL.State = 1 Then rsSQL.Close
'        rsSQL.Source = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'LIQUIDACIONHORIZONTE' AND TABLE_NAME = '" & "A3" & lstselec.List(X - 1) & "'"
'        rsSQL.Open
'    If rsSQL.Fields(0) > 0 Then
        
        
        strCadenaSQL = "insert into tempoa3 select * from " & "A3" & lstselec.List(X - 1)
        cnServer.Execute strCadenaSQL
        ''''''''''''''''''''''''''''''''''ARCHIVO TEXTO 3''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "select * from " & "A3" & lstselec.List(X - 1)
        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
        
         Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL![Ruc])
              'StrRemuneracion = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              'StrRemuneracion = ("" & rsSQL!REMUNERACION)
              StrCodigo = ("" & rsSQL![TIPO MOV LABORAL])
              Strcuspp = Replace(("" & rsSQL![cuspp]), "Ð", "Ñ")
              strPeriodo = ("" & rsSQL![PERIODO DEVENGE])
              StrFechainicio = ("" & rsSQL![FECHA INICIO MOV])
              StrFechaCese = ("" & rsSQL![FECHA FIN MOV])
              'StrArchivotif = ("" & rsSQL!ds_archivo)
              StrRequerimiento = ("" & rsSQL![NRO REQUERIMIENTO])
              StrObservacion = ("" & rsSQL![OBSERVACIONES])
              strFecha = ("" & rsSQL![FECHA])
              Print #lngFile2, Strcuspp; vbTab; StrRucArchivo; vbTab; StrCodigo; vbTab; StrFechainicio; vbTab; StrFechaCese; vbTab; ""; vbTab; strPeriodo; vbTab; "N"; vbTab; Format(StrRequerimiento, "000000000000"); vbTab; strFecha; vbTab; StrObservacion
         rsSQL.MoveNext
         Loop
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
  'End If
        
        
'        If rsSQL.State = 1 Then rsSQL.Close
'        rsSQL.Source = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'LIQUIDACIONHORIZONTE' AND TABLE_NAME = '" & "A4" & lstselec.List(X - 1) & "'"
'        rsSQL.Open
'           If rsSQL.Fields(0) > 0 Then
        
                strCadenaSQL = "insert into tempoa4 select * from " & "A4" & lstselec.List(X - 1)
                cnServer.Execute strCadenaSQL
                ''''''''''''''''''''''''''''''''''ARCHIVO TEXTO 4''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If rsSQL.State = 1 Then rsSQL.Close
                rsSQL.Source = "select * from " & "A4" & lstselec.List(X - 1)
                rsSQL.Open
                If rsSQL.RecordCount > 0 Then
                
                 Do While Not rsSQL.EOF
                      StrRucArchivo = ("" & rsSQL![Ruc])
                      'StrRemuneracion = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
                      'StrRemuneracion = ("" & rsSQL!REMUNERACION)
                      StrCodigo = ("" & rsSQL![TIPO MOV LABORAL])
                      Strcuspp = Replace(("" & rsSQL![cuspp]), "Ð", "Ñ")
                      strPeriodo = ("" & rsSQL![PERIODO DEVENGE])
                      StrFechainicio = ("" & rsSQL![FECHA INICIO MOV])
                      StrFechaCese = ("" & rsSQL![FECHA FIN MOV])
                      'StrArchivotif = ("" & rsSQL!ds_archivo)
                      StrRequerimiento = ("" & rsSQL![NRO REQUERIMIENTO])
                      StrObservacion = ("" & rsSQL![MOTIVO RECHAZOS])
                      strFecha = ("" & rsSQL![FECHA])
                      Print #lngFile3, ""; vbTab; Strcuspp; vbTab; StrRucArchivo; vbTab; StrCodigo; vbTab; StrFechainicio; vbTab; StrFechaCese; vbTab; ""; vbTab; strPeriodo; vbTab; "N"; vbTab; Format(StrRequerimiento, "000000000000"); vbTab; strFecha; vbTab; StrObservacion
                 rsSQL.MoveNext
                 Loop
                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
   
        'End If
  
  
  
'  If rsSQL.State = 1 Then rsSQL.Close
'        rsSQL.Source = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'LIQUIDACIONHORIZONTE' AND TABLE_NAME = '" & "A5" & lstselec.List(X - 1) & "'"
'        rsSQL.Open
'      If rsSQL.Fields(0) > 0 Then
  
  
        strCadenaSQL = "insert into tempoa5 select * from " & "A5" & lstselec.List(X - 1)
        cnServer.Execute strCadenaSQL
  
        ''''''''''''''''''''''''''''''''''ARCHIVO TEXTO 5''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "select * from " & "A5" & lstselec.List(X - 1)
        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
        
         Do While Not rsSQL.EOF
              StrRucArchivo = ("" & rsSQL![Ruc])
              'StrRemuneracion = Replace(("" & rsSQL!ds_raz_soc), "Ð", "Ñ")
              'StrRemuneracion = ("" & rsSQL!REMUNERACION)
              StrDireccion = ("" & rsSQL![DIRECCION])
              'Strcuspp = Replace(("" & rsSQL![cuspp]), "Ð", "Ñ")
              StrEmail = ("" & rsSQL![E-MAIL])
              StrTelefono = ("" & rsSQL![TELEFONO])
              StrDocumento = ("" & rsSQL![DOC_IDENTIDAD])
              'StrArchivotif = ("" & rsSQL!ds_archivo)
              StrNom_Ape = ("" & rsSQL![NOMBRE APELLIDO])
              StrCambio_RS = ("" & rsSQL![CAMBIO RSOCIAL])
              strFecha = ("" & rsSQL![FECHA])
              Print #lngFile4, StrRucArchivo; vbTab; StrDireccion; vbTab; StrEmail; vbTab; StrTelefono; vbTab; StrDocumento; vbTab; StrNom_Ape; vbTab; strFecha; vbTab; StrCambio_RS; vbTab; ""
         rsSQL.MoveNext
         Loop
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
  ' End If
        
'   If rsSQL.State = 1 Then rsSQL.Close
'        rsSQL.Source = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'LIQUIDACIONHORIZONTE' AND TABLE_NAME = '" & "A6" & lstselec.List(X - 1) & "'"
'        rsSQL.Open
'      If rsSQL.Fields(0) > 0 Then
        strCadenaSQL = "insert into tempoa6 select * from " & "A6" & lstselec.List(X - 1)
        cnServer.Execute strCadenaSQL
'   End If
        ''''''''''''''''''''''''''''''''''ARCHIVO TEXTO 6''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "select * from " & "A6" & lstselec.List(X - 1)
        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
        
         Do While Not rsSQL.EOF
              strCuenta = ("" & rsSQL![CUENTA])
              Strcuspp = Replace(("" & rsSQL![cuspp]), "Ð", "Ñ")
              StrRucArchivo = ("" & rsSQL![Ruc])
              StrCodigo = ("" & rsSQL![TIPO MOV LABORAL])
              StrFechainicio = ("" & rsSQL![FECHA INICIO MOV])
              StrFechaCese = ("" & rsSQL![FECHA FIN MOV])
              strPeriodo = ("" & rsSQL![PERIODO DEVENGE])
              StrTraLaboral = ("" & rsSQL![TIPO TRAMO LABORAL])
              StrRequerimiento = ("" & rsSQL![NRO REQUERIMIENTO])
              strFecha = ("" & rsSQL![FECHA])
              strLp = ("" & rsSQL![nro lp])
              strCorrelativo = ("" & rsSQL![correlativo])
              strpagina = ("" & rsSQL![PAGINA])
              strRegistros = ("" & rsSQL![Registros])
              StrCambio_RS = ("" & rsSQL![CAMBIO RSOCIAL])
              StrObservacion = ("" & rsSQL![OBSERVACIONES])
              StrEnvio = ("" & rsSQL![envio])
              StrDireccion = ("" & rsSQL![DIRECCION])
              StrEmail = ("" & rsSQL![E-MAIL])
              strFechaEm = ("" & rsSQL![FECHA emision])
              stradestino = ("" & rsSQL![archivo de destino])
              
              Print #lngFile5, strCuenta; vbTab; Strcuspp; vbTab; StrRucArchivo; vbTab; StrCodigo; vbTab; StrFechainicio; vbTab; StrFechaCese; vbTab; strCorrelativo; vbTab; strPeriodo; vbTab; StrTraLaboral; vbTab; StrRequerimiento; vbTab; strFecha; vbTab; strLp; vbTab; strpagina; vbTab; strRegistros; vbTab; StrCambio_RS; vbTab; StrObservacion; vbTab; StrEnvio; vbTab; StrDireccion; vbTab; StrEmail; vbTab; strFechaEm; vbTab; stradestino
         rsSQL.MoveNext
         Loop
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        StrArchivo = strDirectorio & "\Descargo_Masivo_Empleadores_inicio_mod_" & strZonaActual & "_" & lstselec.List(X - 1) & ".XLS"
'        StrArchivo_D = strDirectorio & "\Archivo2_" & strZonaActual & "_" & lstselec.List(X - 1) & ".XLS"
'        StrArchivo_Cargo = strDirectorio & "\Archivo_Cargo_" & strZonaActual & "_" & lstselec.List(X - 1) & ".XLS"
'

Next X
        
        Close lngFile
        Close lngFile1
        Close lngFile2
        Close lngFile3
        Close lngFile4
        Close lngFile5
        'strsqlp = "Select RUC,RASON_SOC as RAZ_SOC,COMENTARIO,COD_DESCARGO,CUSPP,PERIODO,FECHA_REFERENCIA,NUMRECLAMO,AREAOPER,ARCHIVO from Tempo order by id_tempo"
        
        
        If Option1(0).Value = True Then
        
        
                strsqlp = "Select  [CUENTA AFILIADO],[CUSPP AFILIADO], [RUC EMPLEADOR],[TIPO MOV LABORAL], [FECHA INICIO MOV], [FECHA FIN MOV],CORRELATIVO ,[PERIODO DEVENGE],[TIPO TRAMO LABORAL],[NRO REQUERIMIENTO],[FECHA] , ARCHIVO from TempoA1 order by id_tempo"
                If Exportar_ADO_Excel(strsqlp, StrArchivoDESTINO1) Then
                End If
                
                If Dir(StrArchivoDESTINO11, vbArchive) <> "" Then Kill StrArchivoDESTINO11
                
                strsqlp = "Select  [RUC EMPLEADOR], [CUSPP AFILIADO], [PERIODO DEVENGE],[CUENTA AFILIADO],[REMUNERACION], [MOTIVO SINCERAMIENTO], ARCHIVO from TempoA2 order by id_tempo"
                If Exportar_ADO_Excel(strsqlp, StrArchivoDESTINO2) Then
                End If
                
                If Dir(StrArchivoDESTINO22, vbArchive) <> "" Then Kill StrArchivoDESTINO22
                
                strsqlp = "Select [CUSPP],[RUC],[TIPO MOV LABORAL],[FECHA INICIO MOV],[FECHA FIN MOV],CORRELATIVO ,[PERIODO DEVENGE],[TIPO TRAMO LABORAL],[NRO REQUERIMIENTO],[FECHA], [OBSERVACIONES],ARCHIVO from TempoA3 order by id_tempo"
                If Exportar_ADO_Excel(strsqlp, StrArchivoDESTINO3) Then
                End If
                
                If Dir(StrArchivoDESTINO33, vbArchive) <> "" Then Kill StrArchivoDESTINO33
                
                strsqlp = "Select [CUENTA], [CUSPP],[RUC],[TIPO MOV LABORAL], [FECHA INICIO MOV],[FECHA FIN MOV],CORRELATIVO ,[PERIODO DEVENGE],[TIPO TRAMO LABORAL],[NRO REQUERIMIENTO],[FECHA],[MOTIVO RECHAZOS],ARCHIVO from TempoA4 order by id_tempo"
                If Exportar_ADO_Excel(strsqlp, StrArchivoDESTINO4) Then
                End If
                
                If Dir(StrArchivoDESTINO44, vbArchive) <> "" Then Kill StrArchivoDESTINO44
                
                strsqlp = "Select [RUC], [DIRECCION],[E-MAIL],[TELEFONO],[DOC_IDENTIDAD],[NOMBRE APELLIDO],[FECHA],[CAMBIO RSOCIAL],[ENVIO],ARCHIVO  from TempoA5 order by id_tempo"
                If Exportar_ADO_Excel(strsqlp, StrArchivoDESTINO5) Then
                End If
                
                If Dir(StrArchivoDESTINO55, vbArchive) <> "" Then Kill StrArchivoDESTINO55
        
                strsqlp = "Select [CUENTA], [CUSPP],[RUC],[TIPO MOV LABORAL],[FECHA INICIO MOV],[FECHA FIN MOV],CORRELATIVO ,[PERIODO DEVENGE],[TIPO TRAMO LABORAL],[NRO REQUERIMIENTO],[FECHA], [NRO LP], [PAGINA],[REGISTROS],[CAMBIO RSOCIAL], [OBSERVACIONES],[ENVIO],DIRECCION,[E-MAIL],[FECHA EMISION],[ARCHIVO DE DESTINO] from TempoA6 order by id_tempo"
                If Exportar_ADO_Excel(strsqlp, StrArchivoDESTINO6) Then
                End If
                
                If Dir(StrArchivoDESTINO66, vbArchive) <> "" Then Kill StrArchivoDESTINO66
        
        Else
            
        End If
        
        lblgeneral.Caption = "Se generaron correctamente : " & StrArchivoDESTINO1 & Chr(13)
        lblgeneral.Caption = lblgeneral.Caption & "Se generaron correctamente : " & StrArchivoDESTINO2 & Chr(13)
        lblgeneral.Caption = lblgeneral.Caption & "Se generaron correctamente : " & StrArchivoDESTINO3 & Chr(13)
        lblgeneral.Caption = lblgeneral.Caption & "Se generaron correctamente : " & StrArchivoDESTINO4 & Chr(13)
        lblgeneral.Caption = lblgeneral.Caption & "Se generaron correctamente : " & StrArchivoDESTINO5 & Chr(13)
        lblgeneral.Caption = lblgeneral.Caption & "Se generaron correctamente : " & StrArchivoDESTINO6
        
        
        'Cambio 07/04/2011
       '''''''''''''''''''
        If MsgBox("Se Realizara el envio Grupal de este lote al cliente en estos momentos ?", vbYesNo + vbQuestion, "Finalizado") = vbYes Then
            
            If rsSQL.State = 1 Then rsSQL.Close
            rsSQL.Source = "Select * from tb_envios where lotes='" & StrLote_Agrupado & "'"
    
            rsSQL.Open
            If rsSQL.RecordCount > 0 Then
               If MsgBox("Estos Lotes ya se genero envio desea volver a generar envio ?", vbYesNo + vbQuestion, "Horizonte") = vbYes Then
               
                strEjecutar = "delete from tb_envios where lotes='" & StrLote_Agrupado & "'"
                cnServer.Execute strEjecutar
               
                Strorden = InputBox("Ingrese el Numeo de Orden:", "Horizonte")
                strEjecutar = "Insert into tb_envios  (orden,lotes,cant,fecha) values ('" & Format(Strorden, "0000") & "','" & StrLote_Agrupado & "','" & lstselec.ListCount & "','" & Format(Now, "dd/mm/yyyy") & "')"
                cnServer.Execute strEjecutar
               
               End If
            Else
                Strorden = InputBox("Ingrese el Numeo de Orden:", "Horizonte")
                strEjecutar = "Insert into tb_envios (orden,lotes,cant,fecha) values  ('" & Format(Strorden, "0000") & "','" & StrLote_Agrupado & "','" & lstselec.ListCount & "','" & Format(Now, "dd/mm/yyyy") & "')"
                cnServer.Execute strEjecutar
            End If
            
            
        End If
       '''''''''''''''''''
        
        
        
        
        
        
        
MsgBox "Proceso Termidado Correctamente ", vbExclamation, "Liquidacion Horizonte"
Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Err:
    If Err.Number = -2147217865 Then
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, App.EXEName
        Me.MousePointer = 0
        Screen.MousePointer = 0
        Unload Me
    End If

    End If
End Sub

Private Sub cmdEliminar_Click()
Dim strZona As String
On Error GoTo Err
 If MsgBox("Desea Generar los backup's y limpiar las tablas de los lotes agrupados ?", vbYesNo + vbQuestion, "Finalizado") = vbYes Then
     strZona = Format(Year(Date))
     For X = 1 To lstselec.ListCount
        strEjecutar = "drop table [" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "_Año_" & strZona & "_Envio_" & lstselec.List(X - 1) & "]"
        cnServer.Execute strEjecutar
        strEjecutar = "select * into [" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "_Año_" & strZona & "_Envio_" & lstselec.List(X - 1) & "] from tb_liquidacion where ds_batch='" & lstselec.List(X - 1) & "'"
        cnServer.Execute strEjecutar
        strEjecutar = "delete from tb_liquidacion where ds_batch='" & lstselec.List(X - 1) & "'"
        cnServer.Execute strEjecutar
        lstselec.RemoveItem X - 1
     Next X
 End If
ListarBatch
MsgBox "Proceso Termidado Correctamente ", vbExclamation, "Liquidacion Horizonte"
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

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    ListarBatch
End Sub

Sub ListarBatch()
    LstLotes.Clear
    If rsSQL.State = 1 Then rsSQL.Close
    
    rsSQL.Source = "Select distinct ds_batch from TB_liquidacion order by 1"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        rsSQL.MoveFirst
        Do While Not rsSQL.EOF
            LstLotes.AddItem rsSQL!DS_batch
            rsSQL.MoveNext
        Loop
    End If
    If rsSQL.State = 1 Then rsSQL.Close
End Sub

Private Sub LstLotes_DblClick()
Dim strLibre As Boolean
Dim intVar As Integer
strLibre = False
  
For intVar = 1 To lstselec.ListCount
  If LstLotes.Text = lstselec.List(intVar - 1) Then
  strLibre = True
  End If
Next intVar
  
If strLibre = False Then
  lstselec.AddItem LstLotes.Text
End If
  
End Sub

Private Function Exportar_Excel(Sql As String, sOutputPathXLS As String) As Boolean
       
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
       
    Exportar_Excel = True
    Me.Enabled = True
    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_Excel = False
    Me.Enabled = True
End Function

Public Sub Generar_Archivo(RUTA)
Dim RsExcell As New Recordset
Dim Y As Integer
Dim StrRucArchivo As String
Dim StrRaSocArchivo As String
Dim StrComentario As String
Dim StrCodigo As String
Dim Strcuspp As String
Dim strPeriodo As String
Dim StrFechaRef As String
Dim StrArchivotif As String
Dim xhoja As String
On Error GoTo Error_Handling_Generar_Archivo
xhoja = "Hoja1"
With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RUTA & _
                ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "select * from [" & xhoja & "$]"
'.Open "INSERT " & RUTA & " select * from " & RutaArchivoExcel & ""
End With

'Icr_FrmCarga.MousePointer = vbHourglass
If RsExcell.RecordCount > 0 Then
   For Y = 1 To RsExcell.RecordCount
     DoEvents
        'Do While Not rsSQL.EOF
              StrRucArchivo = ("" & RsExcell.Fields("ruc"))
              StrRaSocArchivo = Replace(("" & RsExcell.Fields("raz_soc")), "Ð", "Ñ")
              StrComentario = ("" & RsExcell.Fields("comentario"))
              StrCodigo = ("" & RsExcell.Fields("cod_descargo"))
              Strcuspp = Replace(("" & RsExcell.Fields("cuspp")), "Ð", "Ñ")
              strPeriodo = ("" & RsExcell.Fields("periodo"))
              StrFechaRef = ("" & RsExcell.Fields("fecha_referencia"))
              StrArchivotif = ("" & RsExcell.Fields("archivo"))
              
              strCadenaSQL = "EXEC sp_Insertartempo_envio '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "','" & StrArchivotif & "'"
              cnServer.Execute strCadenaSQL
     RsExcell.MoveNext
   Next Y
End If

Error_Handling_Generar_Archivo:
End Sub

Public Sub Generar_Archivo_D(RUTA)
Dim RsExcell As New Recordset
Dim Y As Integer
Dim StrRucArchivo As String
Dim StrRaSocArchivo As String
Dim StrComentario As String
Dim StrCodigo As String
Dim Strcuspp As String
Dim strPeriodo As String
Dim StrFechaRef As String
Dim StrArchivotif As String
Dim xhoja As String
On Error GoTo Error_Handling_Generar_Archivo
xhoja = "Hoja1"
With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RUTA & _
                ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "select * from [" & xhoja & "$]"
'.Open "INSERT " & RUTA & " select * from " & RutaArchivoExcel & ""
End With

'Icr_FrmCarga.MousePointer = vbHourglass
If RsExcell.RecordCount > 0 Then
   For Y = 1 To RsExcell.RecordCount
     DoEvents
        'Do While Not rsSQL.EOF
              StrRucArchivo = ("" & RsExcell.Fields("ruc"))
              StrRaSocArchivo = Replace(("" & RsExcell.Fields("raz_soc")), "Ð", "Ñ")
              StrComentario = ("" & RsExcell.Fields("comentario"))
              StrCodigo = ("" & RsExcell.Fields("cod_descargo"))
              Strcuspp = Replace(("" & RsExcell.Fields("cuspp")), "Ð", "Ñ")
              strPeriodo = ("" & RsExcell.Fields("periodo"))
              StrFechaRef = ("" & RsExcell.Fields("fecha_referencia"))
              'StrArchivotif = ("" & RsExcell.Fields("archivo"))
              
              strCadenaSQL = "EXEC sp_Insertartempo_envio_D '" & StrRucArchivo & "','" & StrRaSocArchivo & "','" & StrComentario & "','" & StrCodigo & "','" & Strcuspp & "','" & strPeriodo & "','" & StrFechaRef & "'"
              cnServer.Execute strCadenaSQL
     RsExcell.MoveNext
   Next Y
End If

Error_Handling_Generar_Archivo:
End Sub
Public Sub Generar_Archivo_Cargo(RUTA)
Dim RsExcell As New Recordset
Dim Y As Integer
Dim StrRucArchivo As String
Dim StrRaSocArchivo As String
Dim StrComentario As String
Dim StrCodigo As String
Dim Strcuspp As String
Dim strPeriodo As String
Dim StrFechaRef As String
Dim StrArchivotif As String
Dim xhoja As String
Dim StrNrolircargo As String
Dim StrLoteCargo As String
Dim StrRutaCargo As String
'Dim StrRucArchivo As String
Dim StrPagCargo As String
Dim rsFormulario As New Recordset
On Error GoTo Error_Handling_Generar_Archivo
xhoja = "Hoja1"
    rsFormulario.CursorType = adOpenDynamic
    rsFormulario.LockType = adLockBatchOptimistic
    rsFormulario.ActiveConnection = cnServer


With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RUTA & _
                ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "select * from [" & xhoja & "$]"
'.Open "INSERT " & RUTA & " select * from " & RutaArchivoExcel & ""
End With

'Icr_FrmCarga.MousePointer = vbHourglass
If RsExcell.RecordCount > 0 Then
   For Y = 1 To RsExcell.RecordCount
     DoEvents
        'Do While Not rsSQL.EOF
              StrNrolircargo = ("" & RsExcell!NroLiq)
              StrLoteCargo = ("" & RsExcell!LOTE)
              StrRutaCargo = ("" & RsExcell!RUTA)
              StrRucArchivo = ("" & RsExcell!Ruc)
              StrPagCargo = (RsExcell!PAGINA)
              
              
              'Grabar_TemporalCargo Trim$(StrNrolircargo), Trim$(StrLoteCargo), Trim$(StrPagCargo), Trim$(StrRutaCargo), "", Trim$(StrRucArchivo)
              
              strCadenaSQL = "EXEC sp_InsertarTempo_Cargo_Envio '" & Trim$(StrNrolircargo) & "','" & Trim$(StrLoteCargo) & "','" & Trim$(StrPagCargo) & "','" & Trim$(StrRutaCargo) & "','" & Trim$(StrRucArchivo) & "'"
              cnServer.Execute strCadenaSQL
     RsExcell.MoveNext
   Next Y
End If

Error_Handling_Generar_Archivo:
End Sub


Public Sub Borrar_Tempo_Sql()
Dim cmd As Command
    Set cmd = New Command
    cmd.ActiveConnection = cnServer
    cmd.CommandText = "DELETE  FROM tempo"
    cmd.Execute
    
End Sub

Public Sub Borrar_Tempo_Sql_Archivos()
Dim cmd As Command
    Set cmd = New Command
    cmd.ActiveConnection = cnServer
    cmd.CommandText = "DELETE  FROM tempoA1"
    cmd.Execute
    cmd.CommandText = "DELETE  FROM tempoA2"
    cmd.Execute
    cmd.CommandText = "DELETE  FROM tempoA3"
    cmd.Execute
    cmd.CommandText = "DELETE  FROM tempoA4"
    cmd.Execute
    cmd.CommandText = "DELETE  FROM tempoA5"
    cmd.Execute
    cmd.CommandText = "DELETE  FROM tempoA6"
    cmd.Execute
    
End Sub

Public Sub Borrar_Tempo_Sql_D()
Dim cmd As Command
    Set cmd = New Command
    cmd.ActiveConnection = cnServer
    cmd.CommandText = "DELETE  FROM tempoD"
    cmd.Execute
End Sub
Public Sub Borrar_Tempo_Sql_Cargo()
Dim cmd As Command
    Set cmd = New Command
    cmd.ActiveConnection = cnServer
    cmd.CommandText = "DELETE  FROM tempoca"
    cmd.Execute
End Sub



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

Private Sub lstselec_DblClick()
lstselec.RemoveItem lstselec.ListIndex
'lstselec.D
End Sub

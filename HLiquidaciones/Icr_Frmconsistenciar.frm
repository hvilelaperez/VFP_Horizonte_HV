VERSION 5.00
Begin VB.Form Icr_Frmconsistenciar 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consistencias"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   Icon            =   "Icr_Frmconsistenciar.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3510
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAvance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawMode        =   7  'Invert
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   4785
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar Liquidación"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4815
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
      Height          =   1710
      ItemData        =   "Icr_Frmconsistenciar.frx":030A
      Left            =   120
      List            =   "Icr_Frmconsistenciar.frx":030C
      TabIndex        =   1
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar Afiliados"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "Icr_Frmconsistenciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSQL As Recordset

Private Sub cmdprocesar_Click()
Dim bolError As Boolean
Dim strCadenaSQL As String
Dim intIndice As Integer

Dim strCL As String
Dim strFechaI As String
Dim strRemun As String
Dim strPlanilla As String
Dim StrCampo As String

Dim rsReg As Recordset

Dim strCantidad As String
Dim strProducto As String

Dim intCantidad As Integer
Dim intNulos As Integer
    
    intCantidad = 0
    intNulos = 0
    
    lstMensajes.Clear
    cmdprocesar.Enabled = False
    strCadenaSQL = "Delete from TB_RECHAZOS"
    cnServer.Execute strCadenaSQL
    
    
    Set rsSQL = New Recordset
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    
    Set rsReg = New Recordset
    rsReg.CursorType = adOpenDynamic
    rsReg.LockType = adLockBatchOptimistic
    rsReg.ActiveConnection = cnServer
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from TB_Parametros"
    rsSQL.Open
    
    strCampañaActual = "" & rsSQL!ds_periodo
    strZonaActual = "" & rsSQL!ds_año
    strDirectorio = "" & rsSQL!ds_ftpaxis
    intProductos = 0 & rsSQL!nu_cl
    intCantidad_Minimo = 0 & rsSQL!nu_cant_minimo
    
    
    '''''''''''''''''''''''''''''SOLO POR POR NO CODIGO DE BARRA ''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If rsSQL.State = 1 Then rsSQL.Close
    cnServer.Execute "update Tb_liquidacion set ds_periodo='" & strCampañaActual & "'"
    rsSQL.Source = "Select * from tb_liquidacion order by id_liq"
    rsSQL.Open
    
    If rsSQL.RecordCount > 0 Then
        With rsSQL
            Do While Not .EOF
            strCadenaSQL = "update TB_LIQUIDACION SET ds_nroliq ='" & "HO" & Replace("" & !ds_nroliq, "O", "0", 3) & "' where id_LIQ=" & rsSQL!id_liq
            cnServer.Execute strCadenaSQL
            rsSQL.MoveNext
            Loop
            End With
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    cnServer.Execute "update Tb_liquidacion set ds_periodo='" & strCampañaActual & "'"
    rsSQL.Source = "Select * from tb_liquidacion order by id_liq"
    rsSQL.Open
    
       
    
    If rsSQL.RecordCount > 0 Then
        With rsSQL
            Do While Not .EOF
            bolError = False
            bolCuadrante = False
            
            '''''''''''''''''''Condiciones Para Consistencias y Control de Calidad''''''''''''''''''''''
            '''''''''''''''''''''''Codigos de Barra y nro de Liquidacion''''''''''''''''''''''''''''''''
'            If (!ds_BARRA) = "10C1786982" Then
'            MsgBox "PREUEBA"
'            End If
            
            
            If IIf(IsNull((!ds_BARRA)), "", (!ds_BARRA)) = "" Then
                   strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",1,'ds_barra','" & "" & !ds_BARRA & "',0"
                   cnServer.Execute strCadenaSQL
                   bolError = True

            Else
                If ExisteBarra("" & !ds_BARRA) = False Then
                      strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",3,'ds_barra','" & "" & !ds_BARRA & "',0"
                      cnServer.Execute strCadenaSQL
                      bolError = True
                End If

            End If
            
            
'            If IIf(IsNull((!ds_nroliq)), "", (!ds_nroliq)) = "" Then
'                      strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",2,'ds_nroliq','" & "" & !ds_nroliq & "',0"
'                      cnServer.Execute strCadenaSQL
'                      bolError = True
'            Else
'
'
'                If ExisteNroliq("" & !ds_nroliq) = False Then
'                      strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",4,'ds_nroliq','" & "" & !ds_nroliq & "',0"
'                      cnServer.Execute strCadenaSQL
'                      bolError = True
'                End If
'
'            End If
            
            
            
            If IIf(IsNull((!ds_BARRA)), "", (!ds_BARRA)) = "" Then
                 strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",11,'ds_pag','" & "" & !ds_pag & "',0"
                 cnServer.Execute strCadenaSQL
                 bolError = True
            Else
                If Verificar_Pag(Format(Mid(!ds_BARRA, 10, 5), "00000")) = False Then
                 strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",12,'ds_BARRA','" & "" & !ds_BARRA & "',0"
                 cnServer.Execute strCadenaSQL
                 bolError = True
                End If
            End If
            
            'If IIf(IsNull((!ds_barra)), "", (!ds_barra)) <> "" And IIf(IsNull((!ds_nroliq)), "", (!ds_nroliq)) <> "" Then
            
            If IIf(IsNull(Mid(!ds_BARRA, 10, 5)), "", Mid(!ds_BARRA, 10, 5)) <> "" And IIf(IsNull((!ds_BARRA)), "", (!ds_BARRA)) <> "" Then
                 
                 If Verificar_Pag((Mid(!ds_BARRA, 10, 5))) = True Then
                            
                            If Existe_Liqui_Pag("" & Format(Mid(!ds_BARRA, 10, 5), "00000"), "" & Mid(!ds_BARRA, 1, 9)) = False Then
                            
                                    If ExisteNroliq("" & "HO" & Mid(!ds_BARRA, 1, 9)) = False Then
                                            strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",4,'ds_barra','" & "" & "HO" & Mid(!ds_BARRA, 1, 9) & "',0"
                                            cnServer.Execute strCadenaSQL
                                            bolError = True
                                    Else
                                            'strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",13,'ds_barra','" & "" & "HO" & Mid(!ds_BARRA, 1, 9) & "',0"
                                            strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",12,'ds_BARRA','" & "" & !ds_BARRA & "',0"
                                            cnServer.Execute strCadenaSQL
                                            bolError = True
                                            
'                                            strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",12,'ds_barra','" & "" & Mid(!ds_BARRA, 10, 5) & "',0"
'                                            cnServer.Execute strCadenaSQL
'                                            bolError = True
                                    End If
                            
                            Else
                                If Liqui_pagina_Procesado("" & Format(Mid(!ds_BARRA, 10, 5), "00000"), "" & "HO" & Mid(!ds_BARRA, 1, 9)) = False Then
                                'If Liqui_pag_Procesado("" & !ds_barra, "" & !ds_nroliq) = False Then
                                   strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",5,'ds_barra','" & "" & "HO" & Mid(!ds_BARRA, 1, 9) & "',0"
                                   cnServer.Execute strCadenaSQL
                                   bolError = True
                                Else
                                   ''''''''''''''Verificar Codigo'''''''''''''''''''''''''
                                         For intIndice = 1 To intProductos
                                              ''''''VERIFICAR CODIGO''''''''
                                               strCL = "ds_cod" & intIndice
                                               strProducto = Trim(("" & .Fields(strCL)))
                                               
                                               If rsReg.State = 1 Then rsReg.Close
                                                cnServer.Execute "update Tb_liquidacion set " & strCL & "='" & Format(strProducto, "00") & "' where id_liq=" & rsSQL!id_liq
                                                rsReg.Source = "Select * from tb_liquidacion order by id_liq"
                                                rsReg.Open
                                               
                                               
                                               If Verificar_Codigo("" & .Fields(strCL)) = False Then
                                                  strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",6,'" & strCL & "','" & strProducto & "'," & intIndice
                                                  cnServer.Execute strCadenaSQL
                                                  bolError = True
                                               End If
                                              ''''''''VERIFICAR FECHA INICIAL'''''''
                                               strFechaI = "ds_fec_ini" & intIndice
                                               strProducto = Trim(("" & .Fields(strFechaI)))
                                               If Verificar_Fecha("" & .Fields(strFechaI)) = False Then
                                                  strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",7,'" & strFechaI & "','" & strProducto & "'," & intIndice
                                                  cnServer.Execute strCadenaSQL
                                                  bolError = True
                                               End If
                                              ''''''''VERIFICAR FECHA CESE'''''''
                                               strFechaI = "ds_fec_ce" & intIndice
                                               strProducto = Trim(("" & .Fields(strFechaI)))
                                               If Verificar_Fecha("" & .Fields(strFechaI)) = False Then
                                                  strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",10,'" & strFechaI & "','" & strProducto & "'," & intIndice
                                                  cnServer.Execute strCadenaSQL
                                                  bolError = True
                                               End If
                                             '''''''''REMUNARACION''''''''''
                                               StrCampo = "ds_rem_as" & intIndice
                                               strProducto = Trim(("" & .Fields(StrCampo)))
                                               If IsNumeric(IIf(("" & .Fields(StrCampo) = ""), 0, ("" & .Fields(StrCampo)))) Then
                                               Else
                                                  strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",8,'" & StrCampo & "','" & strProducto & "'," & intIndice
                                                  cnServer.Execute strCadenaSQL
                                                  bolError = True
                                               
                                               End If
                                             '''''''''''PLANILLA'''''''''''''
                                               StrCampo = "ds_pla" & intIndice
                                               strProducto = Trim(("" & .Fields(StrCampo)))
                                               If IsNumeric(IIf(("" & .Fields(StrCampo) = ""), 0, ("" & .Fields(StrCampo)))) Then
                                               Else
                                                  strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",9,'" & StrCampo & "','" & strProducto & "'," & intIndice
                                                  cnServer.Execute strCadenaSQL
                                                  bolError = True
                                               End If
                                             '''''''''''''''''''''''''''''''
                                         
                                         Next intIndice
                                     
                                End If
                            End If
                    Else
                    strCadenaSQL = "EXEC sp_Insertarliquidacion  '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_año) & "','" & ("" & rsSQL!ds_periodo) & "'," & !id_liq & ",12,'ds_BARRA','" & "" & !ds_BARRA & "',0"
                                            cnServer.Execute strCadenaSQL
                                            bolError = True
                    End If
            
            End If
                    
                        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            intCantidad = 0 & rsSQL!catsig
'            If intCantidad > intCantidad_Minimo Then
'                strCadenaSQL = "EXEC sp_InsertarRechazo '" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id_hojapedido & ",8,'catsig','" & intCantidad & "',0,'" & ("" & !asesora) & "'"
'                cnServer.Execute strCadenaSQL
'                bolError = True
'            End If
'
'            intCantidad = 0 & rsSQL!catpos
'            If intCantidad > intCantidad_Minimo Then
'                strCadenaSQL = "EXEC sp_InsertarRechazo '" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id_hojapedido & ",9,'catpos','" & intCantidad & "',0,'" & ("" & !asesora) & "'"
'                cnServer.Execute strCadenaSQL
'                bolError = True
'            End If
                                                          
            MostrarAvance picAvance, Round((.AbsolutePosition / .RecordCount) * 100), 100, True
            .MoveNext
            DoEvents
            Loop
        End With
    End If
    lstMensajes.AddItem "Liquidaciones Previas Procesadas : " & rsSQL.RecordCount
    If rsSQL.State = 1 Then rsSQL.Close
    
    'rsSQL.Source = "Select id_hojapedido from rechazos where id_tiporechazo=1 or id_tiporechazo=2 or id_tiporechazo=3"
    rsSQL.Source = "Select id_hojapedido from tb_rechazos where (id_tiporechazo=1 or id_tiporechazo=2) "
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        lstMensajes.AddItem "Liquidaciones Previas rechazados : " & rsSQL.RecordCount
    Else
        lstMensajes.AddItem "Liquidaciones Previas rechazados : 0"
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    
    'rsSQL.Source = "Select id_hojapedido from rechazos where id_tiporechazo=4 or id_tiporechazo=5 or id_tiporechazo=6 or id_tiporechazo=7"
    rsSQL.Source = "Select distinct id_hojapedido from tb_rechazos  where (id_tiporechazo <> 1 or  id_tiporechazo <> 2)"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        lstMensajes.AddItem "Liquidaciones Previas a verificar: " & rsSQL.RecordCount
    Else
        lstMensajes.AddItem "Liquidaciones Previas a verificar: 0"
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    
    MsgBox "Fin del proceso", vbInformation, "Consistencias"
    picAvance.Picture = LoadPicture()
    picAvance.BorderStyle = 1
    cmdprocesar.Enabled = True
    Set rsvalida = Nothing
End Sub

Private Sub Command1_Click()
Dim bolError As Boolean
Dim strCadenaSQL As String
Dim intIndice As Integer

Dim strCL As String

Dim strCantidad As String
Dim strProducto As String

Dim intCantidad As Integer
Dim intNulos As Integer
    
    intCantidad = 0
    intNulos = 0
    
    lstMensajes.Clear
    cmdprocesar.Enabled = False
    strCadenaSQL = "Delete from TB_RECHAZOS"
    cnServer.Execute strCadenaSQL
    
    
    Set rsSQL = New Recordset
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from Parametros"
    rsSQL.Open
    
    strCampañaActual = "" & rsSQL!ds_campaña
    strZonaActual = "" & rsSQL!ds_zona
    strDirectorio = "" & rsSQL!ds_ftpaxis
    intProductos = 0 & rsSQL!nu_cl
    intCantidad_Minimo = 0 & rsSQL!nu_cant_minimo
    If rsSQL.State = 1 Then rsSQL.Close
    
    
    'cnServer.Execute "update pedidos set ds_campaña='" & strCampañaActual & "', camp='" & strCampañaActual & "'"
    cnServer.Execute "update afiliacion set ds_campaña='" & strCampañaActual & "'"
    
    'rsSQL.Source = "Select * from PEDIDOS order by id_hojapedido"
    rsSQL.Source = "Select * from afiliacion order by id"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        With rsSQL
            Do While Not .EOF
            bolError = False
            bolCuadrante = False

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If IIf(IsNull((!tipdoc)), "", (!tipdoc)) = "" Then
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",10,'tipdoc','" & "" & !tipdoc & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
                strCadenaSQL = "update afiliacion set tipdoc ='1' where id=" & rsSQL!ID
                cnServer.Execute strCadenaSQL
                Else
                  If Not (Val(Trim$(!tipdoc)) >= 1 And Val(Trim$(!tipdoc)) <= 4) Then
                   strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",11,'tipdoc','" & "" & !tipdoc & "',0"
                   cnServer.Execute strCadenaSQL
                   bolError = True
                  End If
                End If
                                               
                If IIf(IsNull((!numdoc)), "", (!numdoc)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",12,'numdoc','" & "" & !numdoc & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                End If
                
                If IIf(IsNull((!sexo)), "", (!sexo)) = "" Then
                strCadenaSQL = "update afiliacion set sexo ='F' where id=" & rsSQL!ID
                cnServer.Execute strCadenaSQL
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",13,'sexo','" & "" & !sexo & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
                Else
                   If Not ((Trim$(!sexo)) = "F" Or (Trim$(!sexo)) = "M") Then
                   strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",14,'sexo','" & "" & !sexo & "',0"
                   cnServer.Execute strCadenaSQL
                   bolError = True
                   End If
                End If
                
                If IIf(IsNull((!ZONA)), "", (!ZONA)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",15,'zona','" & "" & !ZONA & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                Else
                   If ExisteZona("" & !ZONA) = False Then
                   strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",16,'zona','" & "" & !ZONA & "',0"
                   cnServer.Execute strCadenaSQL
                   bolError = True
                   End If
                End If
                
                If IIf(IsNull((!SECTOR)), "", (!SECTOR)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",17,'sector','" & "" & !SECTOR & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                End If
                
                If IIf(IsNull((!territorio)), "", (!territorio)) = "" Then
                
                strCadenaSQL = "update afiliacion set territorio ='1' where id=" & rsSQL!ID
                cnServer.Execute strCadenaSQL
                
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",57,'territorio','" & "" & !territorio & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
                End If
                
                If IIf(IsNull((!annio)), "", (!annio)) = "" Then
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",18,'annio','" & "" & !annio & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
                
                strCadenaSQL = "update afiliacion set annio ='" & Format(Now, "YYYY") & "'  where id=" & rsSQL!ID
                cnServer.Execute strCadenaSQL
                
                Else
                    If Len(Trim$(!annio)) = 4 Then
                        If Not (Trim$(!annio)) = Format(Now, "YYYY") Then
'                         strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",19,'annio','" & "" & !annio & "',0"
'                         cnServer.Execute strCadenaSQL
'                         bolError = True
                        
                         strCadenaSQL = "update afiliacion set annio ='" & Format(Now, "YYYY") & "'  where id=" & rsSQL!ID
                         cnServer.Execute strCadenaSQL
                        
                        End If
                    Else
                         strCadenaSQL = "update afiliacion set annio ='" & Format(Now, "YYYY") & "'  where id=" & rsSQL!ID
                         cnServer.Execute strCadenaSQL
                    End If
                End If
                
                If IIf(IsNull((!camp)), "", (!camp)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",20,'camp','" & "" & !camp & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                Else
                 If Not (Trim$(!camp) = Mid(Trim$(!ds_campaña), 1, 2)) Then
                  strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",21,'camp','" & "" & !camp & "',0"
                  cnServer.Execute strCadenaSQL
                  bolError = True
                 End If
                End If
                
                If IIf(IsNull((!priapell)), "", (!priapell)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",22,'priapell','" & "" & !priapell & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                End If
                
'                If IIf(IsNull((!segapell)), "", (!segapell)) = "" Then
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",23,'segapell','" & "" & !segapell & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
'                End If
                
                If IIf(IsNull((!Nombres)), "", (!Nombres)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",24,'nombres','" & "" & !Nombres & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If IIf(IsNull((!recdir)), "", (!recdir)) = "" And IIf(IsNull((!recdirnom)), "", (!recdirnom)) = "" And IIf(IsNull((!recnum)), "", (!recnum)) = "" And IIf(IsNull((!recmanza)), "", (!recmanza)) = "" And IIf(IsNull((!reclote)), "", (!reclote)) = "" And IIf(IsNull((!recsector)), "", (!recsector)) = "" And IIf(IsNull((!recgrupo)), "", (!recgrupo)) = "" And IIf(IsNull((!recgrupo)), "", (!recgrupo)) = "" And IIf(IsNull((!recetapa)), "", (!recetapa)) = "" And IIf(IsNull((!rectipviv)), "", (!rectipviv)) = "" And IIf(IsNull((!recnombre)), "", (!recnombre)) = "" Then
                     strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",58,'recdir','" & "" & !recdir & "',0"
                     cnServer.Execute strCadenaSQL
                     bolError = True
                Else
                            
                            If IIf(IsNull((!recdir)), "", (!recdir)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",25,'recdir','" & "" & !recdir & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
                            Else
                              If Not (Val(Trim$(!recdir)) >= 1 And Val(Trim$(!recdir)) <= 5) Then
                               strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",26,'recdir','" & "" & !recdir & "',0"
                               cnServer.Execute strCadenaSQL
                               bolError = True
                               End If
                            End If
                            
'                            If IIf(IsNull((!recdirnom)), "", (!recdirnom)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",27,'recdirnom','" & "" & !recdirnom & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
'                            End If
                            
'                            If IIf(IsNull((!recnum)), "", (!recnum)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",28,'recnum','" & "" & !recnum & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
'                            End If
                            
            '                If IIf(IsNull((!recmanza)), "", (!recmanza)) = "" And IIf(IsNull((!reclote)), "", (!reclote)) = "" And IIf(IsNull((!recsector)), "", (!recsector)) = "" And IIf(IsNull((!recgrupo)), "", (!recgrupo)) = "" And IIf(IsNull((!recgrupo)), "", (!recgrupo)) = "" Then
            '                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",29,'recmanza','" & "" & !recmanza & "',0"
            '                cnServer.Execute strCadenaSQL
            '                bolError = True
            '                Else
            '                  If Not (IIf(IsNull((!recmanza)), "", (!recmanza)) = "") Then
            '                    If IIf(IsNull((!reclote)), "", (!reclote)) <> "" Or IIf(IsNull((!recsector)), "", (!recsector)) <> "" Or IIf(IsNull((!recgrupo)), "", (!recgrupo)) <> "" Or IIf(IsNull((!recetapa)), "", (!recetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",30,'recmanza','" & "" & !recmanza & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!reclote)), "", (!reclote)) = "") Then
            '                    If IIf(IsNull((!recmanza)), "", (!recmanza)) <> "" Or IIf(IsNull((!recsector)), "", (!recsector)) <> "" Or IIf(IsNull((!recgrupo)), "", (!recgrupo)) <> "" Or IIf(IsNull((!recetapa)), "", (!recetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",30,'reclote','" & "" & !reclote & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!recsector)), "", (!recsector)) = "") Then
            '                    If IIf(IsNull((!recmanza)), "", (!recmanza)) <> "" Or IIf(IsNull((!reclote)), "", (!reclote)) <> "" Or IIf(IsNull((!recgrupo)), "", (!recgrupo)) <> "" Or IIf(IsNull((!recetapa)), "", (!recetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",30,'recsector','" & "" & !recsector & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!recgrupo)), "", (!recgrupo)) = "") Then
            '                    If IIf(IsNull((!recmanza)), "", (!recmanza)) <> "" Or IIf(IsNull((!reclote)), "", (!reclote)) <> "" Or IIf(IsNull((!recsector)), "", (!recsector)) <> "" Or IIf(IsNull((!recetapa)), "", (!recetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",30,'recgrupo','" & "" & !recgrupo & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!recetapa)), "", (!recetapa)) = "") Then
            '                    If IIf(IsNull((!recmanza)), "", (!recmanza)) <> "" Or IIf(IsNull((!reclote)), "", (!reclote)) <> "" Or IIf(IsNull((!recgrupo)), "", (!recgrupo)) <> "" Or IIf(IsNull((!recsector)), "", (!recsector)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",30,'recetapa','" & "" & !recetapa & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                End If
                            
                            If IIf(IsNull((!rectipviv)), "", (!rectipviv)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",31,'rectipviv','" & "" & !rectipviv & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
                            Else
                              If Not (Val((!rectipviv)) >= 1 And Val((!rectipviv)) <= 6) Then
                               strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",32,'rectipviv','" & "" & !rectipviv & "',0"
                               cnServer.Execute strCadenaSQL
                               bolError = True
                              End If
                            End If
                            
'                            If IIf(IsNull((!recnombre)), "", (!recnombre)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",33,'recnombre','" & "" & !recnombre & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
'                            End If
                
                End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If IIf(IsNull((!recnsum)), "", (!recnsum)) = "" And IIf(IsNull((!rectsumi)), "", (!rectsumi)) = "" Then
                Else
                   If IIf(IsNull((!recnsum)), "", (!recnsum)) <> "" And IIf(IsNull((!rectsumi)), "", (!rectsumi)) <> "" Then
                   Else
                   If IIf(IsNull((!recnsum)), "", (!recnsum)) = "" Then
'                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",60,'recnsum','" & "" & !recnsum & "',0"
'                    cnServer.Execute strCadenaSQL
'                    bolError = True
                    strCadenaSQL = "update afiliacion set rectsumi ='' where id=" & rsSQL!ID
                    cnServer.Execute strCadenaSQL
                   Else
'                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",61,'rectsumi','" & "" & !rectsumi & "',0"
'                    cnServer.Execute strCadenaSQL
'                    bolError = True
                    strCadenaSQL = "update afiliacion set rectsumi ='EE' where id=" & rsSQL!ID
                    cnServer.Execute strCadenaSQL
                    
                   End If
                   End If
                End If
                
                If IIf(IsNull((!recdistrit)), "", (!recdistrit)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",34,'recdistrit','" & "" & !recdistrit & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                Else
                  If Len(Trim$(!recdistrit)) = 6 Then
                  Else
                  strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",64,'recdistrit','" & "" & !recdistrit & "',0"
                  cnServer.Execute strCadenaSQL
                  bolError = True
                  End If
                End If
                
                If IIf(IsNull((!recprov)), "", (!recprov)) = "" Then
                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",35,'recprov','" & "" & !recprov & "',0"
                cnServer.Execute strCadenaSQL
                bolError = True
                Else
                  If Len(Trim$(!recprov)) = 4 Then
                  Else
                  strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",65,'recprov','" & "" & !recprov & "',0"
                  cnServer.Execute strCadenaSQL
                  bolError = True
                  End If
                End If
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          If IIf(IsNull((!vivdir)), "", (!vivdir)) = "" And IIf(IsNull((!vivdirnom)), "", (!vivdirnom)) = "" And IIf(IsNull((!vivnum)), "", (!vivnum)) = "" And IIf(IsNull((!vivmanza)), "", (!vivmanza)) = "" And IIf(IsNull((!vivlote)), "", (!vivlote)) = "" And IIf(IsNull((!vivsector)), "", (!vivsector)) = "" And IIf(IsNull((!vivgrupo)), "", (!vivgrupo)) = "" And IIf(IsNull((!vivetapa)), "", (!vivetapa)) = "" And IIf(IsNull((!vivtipviv)), "", (!vivtipviv)) = "" And IIf(IsNull((!vivnombre)), "", (!vivnombre)) = "" Then
'            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",59,'vivdir','" & "" & !vivdir & "',0"
'            cnServer.Execute strCadenaSQL
'            bolError = True
          Else
                            
                            If IIf(IsNull((!vivdir)), "", (!vivdir)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",37,'vivdir','" & "" & !vivdir & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
                            Else
                              If Not (Val((!vivdir)) >= 1 And Val((!vivdir)) <= 5) Then
                               strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",38,'vivdir','" & "" & !vivdir & "',0"
                               cnServer.Execute strCadenaSQL
                               bolError = True
                              End If
                            End If
                
'                            If IIf(IsNull((!vivdirnom)), "", (!vivdirnom)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",39,'vivdirnom','" & "" & !vivdirnom & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
'                            End If
'
'                            If IIf(IsNull((!vivnum)), "", (!vivnum)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",40,'vivnum','" & "" & !vivnum & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
'                            End If
                            
            '                If IIf(IsNull((!vivmanza)), "", (!vivmanza)) = "" And IIf(IsNull((!vivlote)), "", (!vivlote)) = "" And IIf(IsNull((!vivsector)), "", (!vivsector)) = "" And IIf(IsNull((!vivgrupo)), "", (!vivgrupo)) = "" And IIf(IsNull((!vivetapa)), "", (!vivetapa)) = "" Then
            '                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",41,'vivmanza','" & "" & !vivmanza & "',0"
            '                cnServer.Execute strCadenaSQL
            '                bolError = True
            '                Else
            '                  If Not (IIf(IsNull((!vivmanza)), "", (!vivmanza)) = "") Then
            '                    If IIf(IsNull((!vivlote)), "", (!vivlote)) <> "" Or IIf(IsNull((!vivsector)), "", (!vivsector)) <> "" Or IIf(IsNull((!vivgrupo)), "", (!vivgrupo)) <> "" Or IIf(IsNull((!vivetapa)), "", (!vivetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",42,'vivlote','" & "" & !vivlote & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!vivlote)), "", (!vivlote)) = "") Then
            '                    If IIf(IsNull((!vivmanza)), "", (!vivmanza)) <> "" Or IIf(IsNull((!vivsector)), "", (!vivsector)) <> "" Or IIf(IsNull((!vivgrupo)), "", (!vivgrupo)) <> "" Or IIf(IsNull((!vivetapa)), "", (!vivetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",42,'vivlote','" & "" & !vivlote & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!vivsector)), "", (!vivsector)) = "") Then
            '                    If IIf(IsNull((!vivmanza)), "", (!vivmanza)) <> "" Or IIf(IsNull((!vivlote)), "", (!vivlote)) <> "" Or IIf(IsNull((!vivgrupo)), "", (!vivgrupo)) <> "" Or IIf(IsNull((!vivetapa)), "", (!vivetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",42,'vivsector','" & "" & !vivsector & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!vivgrupo)), "", (!vivgrupo)) = "") Then
            '                    If IIf(IsNull((!vivmanza)), "", (!vivmanza)) <> "" Or IIf(IsNull((!vivlote)), "", (!vivlote)) <> "" Or IIf(IsNull((!vivsector)), "", (!vivsector)) <> "" Or IIf(IsNull((!vivetapa)), "", (!vivetapa)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",42,'vivgrupo','" & "" & !vivgrupo & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                  If Not (IIf(IsNull((!vivetapa)), "", (!vivetapa)) = "") Then
            '                    If IIf(IsNull((!recmanza)), "", (!recmanza)) <> "" Or IIf(IsNull((!reclote)), "", (!reclote)) <> "" Or IIf(IsNull((!recgrupo)), "", (!recgrupo)) <> "" Or IIf(IsNull((!recsector)), "", (!recsector)) <> "" Then
            '                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",42,'vivetapa','" & "" & !vivetapa & "',0"
            '                    cnServer.Execute strCadenaSQL
            '                    bolError = True
            '                    End If
            '                  End If
            '                End If
                                            
                       
                                            
                                            
                            If IIf(IsNull((!vivtipviv)), "", (!vivtipviv)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",43,'vivtipviv','" & "" & !vivtipviv & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
                            Else
                              If Not (Val((!vivtipviv)) >= 1 And Val((!vivtipviv)) <= 6) Then
                               strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",44,'vivtipviv','" & "" & !vivtipviv & "',0"
                               cnServer.Execute strCadenaSQL
                               bolError = True
                              End If
                            End If
                            
'                            If IIf(IsNull((!vivnombre)), "", (!vivnombre)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",45,'vivnombre','" & "" & !vivnombre & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
'                            End If
            End If
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                      If IIf(IsNull((!vivnsum)), "", (!vivnsum)) = "" And IIf(IsNull((!vivtsumi)), "", (!vivtsumi)) = "" Then
                        Else
                           If IIf(IsNull((!vivnsum)), "", (!vivnsum)) <> "" And IIf(IsNull((!vivtsumi)), "", (!vivtsumi)) <> "" Then
                           Else
                           If IIf(IsNull((!vivnsum)), "", (!vivnsum)) = "" Then
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",62,'vivnsum','" & "" & !vivnsum & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
                            strCadenaSQL = "update afiliacion set vivtsumi ='' where id=" & rsSQL!ID
                            cnServer.Execute strCadenaSQL
                           Else
'                            strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",63,'vivtsumi','" & "" & !vivtsumi & "',0"
'                            cnServer.Execute strCadenaSQL
'                            bolError = True
                            strCadenaSQL = "update afiliacion set vivtsumi ='EE' where id=" & rsSQL!ID
                            cnServer.Execute strCadenaSQL
                           End If
                           End If
                        End If





'                If IIf(IsNull((!vivdistrit)), "", (!vivdistrit)) = "" Then
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",46,'vivdistrit','" & "" & !vivdistrit & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
'                End If
'
'                If IIf(IsNull((!vivprov)), "", (!vivprov)) = "" Then
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",47,'vivprov','" & "" & !vivprov & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
'                End If
                
                If IIf(IsNull((!vivtelef1)), "", (!vivtelef1)) = "" And IIf(IsNull((!vivtelef2)), "", (!vivtelef2)) = "" Then
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",48,'vivtelef1','" & "" & !vivtelef1 & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
                strCadenaSQL = "update afiliacion set vivtelef1 ='1' where id=" & rsSQL!ID
                cnServer.Execute strCadenaSQL
                Else
                 If IIf(IsNull((!vivtelef1)), "", (!vivtelef1)) = "" Then
                 strCadenaSQL = "update afiliacion set vivtelef1 ='" & !vivtelef2 & "' where id=" & rsSQL!ID
                 cnServer.Execute strCadenaSQL
                 End If
                End If
                
                If IIf(IsNull((!vivfecha)), "", (!vivfecha)) <> "" Then
                        
                  If Len((!vivfecha)) = 8 Then
                        If Mid((!vivfecha), 1, 2) = "" Then
                        strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",49,'vivfecha','" & "" & !vivfecha & "',0"
                        cnServer.Execute strCadenaSQL
                        bolError = True
                        Else
                          If Val(Mid((!vivfecha), 1, 2)) > 31 Then
                          strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",54,'vivfecha','" & "" & !vivfecha & "',0"
                          cnServer.Execute strCadenaSQL
                          bolError = True
                          End If
                        End If
                        
                        If Mid((!vivfecha), 3, 2) = "" Then
                        strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",50,'vivfecha','" & "" & !vivfecha & "',0"
                        cnServer.Execute strCadenaSQL
                        bolError = True
                        Else
                          If Val(Mid((!vivfecha), 3, 2)) > 12 Then
                           strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",55,'vivfecha','" & "" & !vivfecha & "',0"
                           cnServer.Execute strCadenaSQL
                           bolError = True
                          End If
                        End If
                        
                        If Mid((!vivfecha), 5, 4) = "" Then
                        strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",51,'vivfecha','" & "" & !vivfecha & "',0"
                        cnServer.Execute strCadenaSQL
                        bolError = True
                        Else
                          If Val(Mid((!vivfecha), 5, 4)) > Format(Now, "YYYY") Then
                           strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",56,'vivfecha','" & "" & !vivfecha & "',0"
                           cnServer.Execute strCadenaSQL
                           bolError = True
                          End If
                        End If
                  Else
                        strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !ID & ",53,'vivfecha','" & "" & !vivfecha & "',0"
                        cnServer.Execute strCadenaSQL
                        bolError = True
                  End If
                Else
'                    strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",49,'vivfecha','" & "" & !vivfecha & "',0"
'                    cnServer.Execute strCadenaSQL
'                    bolError = True
                End If
                
'                If IsNumeric(Val((!codasesor))) = False Then
'                strCadenaSQL = "EXEC sp_InsertarRechazoaf '" & ("" & rsSQL!DS_batch) & "','" & ("" & rsSQL!ds_zona) & "','" & ("" & rsSQL!ds_campaña) & "'," & !id & ",52,'CODASESOR','" & "" & !codasesor & "',0"
'                cnServer.Execute strCadenaSQL
'                bolError = True
'                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'End If
                                                          
            MostrarAvance picAvance, Round((.AbsolutePosition / .RecordCount) * 100), 100, True
            .MoveNext
            DoEvents
            Loop
        End With
    End If
    lstMensajes.AddItem "Hojas de Afiliados procesados : " & rsSQL.RecordCount
    If rsSQL.State = 1 Then rsSQL.Close
    
    'rsSQL.Source = "Select id_hojapedido from rechazos where id_tiporechazo=1 or id_tiporechazo=2 or id_tiporechazo=3"
    'rsSQL.Source = "Select id_hojapedido from rechazos where id_tiporechazo=58"
     rsSQL.Source = "Select id_hojapedido from vw_afiliacionesRechazadas"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        lstMensajes.AddItem "Hojas de Afiliados rechazados : " & rsSQL.RecordCount
    Else
        lstMensajes.AddItem "Hojas de Afiliados rechazados : 0"
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    
    'rsSQL.Source = "Select id_hojapedido from rechazos where id_tiporechazo=4 or id_tiporechazo=5 or id_tiporechazo=6 or id_tiporechazo=7"
    'rsSQL.Source = "Select distinct id_hojapedido from rechazos
    rsSQL.Source = "Select distinct id_hojapedido from rechazos  where id_tiporechazo <> 58"
    
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        lstMensajes.AddItem "Hojas de Afiliados a verificar: " & rsSQL.RecordCount
    Else
        lstMensajes.AddItem "Hojas de Afiliados a verificar: 0"
    End If
    If rsSQL.State = 1 Then rsSQL.Close
    
    MsgBox "Fin del proceso", vbInformation, "Consistencias"
    picAvance.Picture = LoadPicture()
    picAvance.BorderStyle = 1
    cmdprocesar.Enabled = True
    Set rsvalida = Nothing

End Sub

Private Sub Form_Load()
'    Me.Height = 4200
'    Me.Width = 5200
'    Centrar mdiMain, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsSQL Is Nothing) Then
        If rsSQL.State = 1 Then rsSQL.Close
        Set rsSQL = Nothing
    End If
End Sub

Function Duplicado(strIdFormulario) As Boolean
Dim bolExiste As Boolean
Dim rsExiste As New Recordset
    rsExiste.CursorType = adOpenDynamic
    rsExiste.LockType = adLockBatchOptimistic
    rsExiste.ActiveConnection = cnServer
    rsExiste.Source = "select * from historico where formulario='" & strIdFormulario & "'"
    rsExiste.Open
    bolExiste = False
    If rsExiste.RecordCount > 0 Then bolExiste = True
    If rsExiste.State = 1 Then rsExiste.Close
    Set rsExiste = Nothing
    Duplicado = bolExiste
End Function

Function ObtenerUbigeo(strUbigeo As String, strZona As String) As String
Dim rsUbigeo As New Recordset
Dim strRetorno As String * 33
    strRetorno = ""
    rsUbigeo.CursorType = adOpenDynamic
    rsUbigeo.LockType = adLockBatchOptimistic
    rsUbigeo.ActiveConnection = cnServer
    If rsUbigeo.State = 1 Then rsUbigeo.Close
    rsUbigeo.Source = "Select disnom,depnom from ubigeo where ubigeo='" & strUbigeo & "' and zona='" & strZona & "'"
    rsUbigeo.Open
    If rsUbigeo.RecordCount > 0 Then
        strRetorno = (Trim("" & rsUbigeo!disnom) & "-" & Trim("" & rsUbigeo!depnom))
    Else
        strRetorno = ""
    End If
    If rsUbigeo.State = 1 Then rsUbigeo.Close
    Set rsUbigeo = Nothing
    ObtenerUbigeo = strRetorno
End Function




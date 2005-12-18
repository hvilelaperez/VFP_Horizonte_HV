VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Flow_FrmasignarLote 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Asignar Lotes Para Procesar"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14700
   Icon            =   "AsignarLote.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9030
   ScaleWidth      =   14700
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   3360
         TabIndex        =   4
         Top             =   120
         Width           =   11175
         Begin VB.CommandButton cmdFormat 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Formatear Carp Archivos"
            Height          =   810
            Left            =   4440
            Picture         =   "AsignarLote.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtRuta 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   10935
         End
         Begin VB.CommandButton cmdAsignar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Asignar Lote"
            Height          =   810
            Left            =   6345
            Picture         =   "AsignarLote.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdIniciar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Inicializar"
            Height          =   810
            Left            =   7545
            Picture         =   "AsignarLote.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtBatch 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2670
            TabIndex        =   9
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Refrescar"
            Height          =   810
            Left            =   9945
            Picture         =   "AsignarLote.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtCampaña 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   7
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdEliminar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Eliminar"
            Height          =   810
            Left            =   8745
            Picture         =   "AsignarLote.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtZona 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1500
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblLote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Directorio de imagenes"
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
            TabIndex        =   16
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
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
            Left            =   3225
            TabIndex        =   14
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblCampaña 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
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
            Left            =   465
            TabIndex        =   13
            Top             =   960
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
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
            Left            =   1905
            TabIndex        =   12
            Top             =   960
            Width           =   330
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   8655
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtCantidad 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   8280
            Width           =   2895
         End
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            Height          =   2175
            Left            =   120
            Pattern         =   "*.tif"
            TabIndex        =   20
            Top             =   6000
            Width           =   2895
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   2895
         End
         Begin VB.DirListBox Dir1 
            Height          =   2115
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   2895
         End
         Begin VB.ListBox lstdirectorios 
            Appearance      =   0  'Flat
            Height          =   2760
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   17
            Top             =   3000
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   6975
         Left            =   3360
         TabIndex        =   1
         Top             =   1800
         Width           =   11175
         Begin MSDataGridLib.DataGrid datDatos 
            Height          =   6615
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   11668
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
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
      End
   End
End
Attribute VB_Name = "Flow_FrmasignarLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency 'FILETIME
    ftLastAccessTime As Currency 'FILETIME
    ftLastWriteTime As Currency 'FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Declare Function FindFirstFile Lib "kernel32" _
    Alias "FindFirstFileA" (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" _
    Alias "FindNextFileA" (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long
'/*---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
'

Option Explicit
Dim strCadena As String
''''''''''''''''''''''''''''''''
Dim strZona As String * 3
Dim strLote As String * 3
Dim strTemp As String

Private Sub cmdAsignar_Click()
Dim intIndice  As Integer
    
txtCampaña.Text = txtZona.Text & txtCampaña.Text
If Len(Trim(txtCampaña.Text)) > 0 And Len(Trim(txtZona.Text)) > 0 And Len(Trim(txtBatch.Text)) > 0 Then
                
        If lstdirectorios.SelCount = 0 Then
            If (File1.ListCount < 1) Then
                 MsgBox "No puede asignar un lote con 0 imagenes", vbCritical + vbOKOnly, "Error"
            Else
                InsertarLote
            End If
        Else
        
            For intIndice = 0 To lstdirectorios.ListCount - 1
                If lstdirectorios.Selected(intIndice) = True Then
                
                    'File1.Path = Dir1.Path & "\" & lstdirectorios.List(intIndice) & "\" & Format(Val(Right(lstdirectorios.List(intIndice), 2)), "00000000")
                    File1.Path = Dir1.Path & "\" & lstdirectorios.List(intIndice) & "\"
                    
                    File1.Refresh
                    TXTCANTIDAD.Text = File1.ListCount & " IMAGENES"
                    
                    txtCampaña.Text = Mid(lstdirectorios.List(intIndice), 1, 6)
                    txtZona.Text = Mid(lstdirectorios.List(intIndice), 8, 4)
                    txtBatch.Text = Mid(lstdirectorios.List(intIndice), 13, 7)
                    
                    txtRuta = UCase(File1.Path)
                    
                    
                    InsertarLote
                End If
            Next
            'MsgBox "Se agregaron " & lstdirectorios.SelCount & " lotes", vbInformation, App.EXEName
        End If
    'End If
  Else
   MsgBox "Verificar carpeta formateada", vbCritical + vbOKOnly, "Error"
  End If
End Sub
Sub InsertarLote()
        If rsSQL.State = 1 Then rsSQL.Close
        rsSQL.Source = "Select id_lote from TB_Lotes where ds_batch = '" & Me.txtBatch.Text & "' and ds_periodo = '" & txtCampaña.Text & "' and ds_año = '" & txtZona.Text & "'"
        'rsSql.Source = "Select id_lote from LOTES where ds_batch = '" & Me.txtBatch.Text & "' and ds_campaña = '" & txtCampaña.Text & "' and ds_zona = '" & txtZona.Text & "'"
        rsSQL.Open
        If rsSQL.RecordCount > 0 Then
            MsgBox "Ya existe un registro de Periodo: " & rsDatos!Periodo & ", Año: " & rsDatos!año & ", Lote: " & rsDatos!LOTE, vbCritical + vbOKOnly, "Error"
        Else
'                strCadena = "Insert Into LOTES (ds_campaña,ds_zona,ds_batch,nu_imagenes,ds_path,ds_envio,ds_tipo) values (" & _
'                "'" & IIf(IsNull(txtCampaña.Text), " ", txtCampaña.Text) & "' ," & _
'                "'" & IIf(IsNull(txtZona.Text), " ", txtZona.Text) & "' ," & _
'                "'" & IIf(IsNull(txtBatch.Text), " ", txtBatch.Text) & "' ," & _
'                Val(txtCantidad) & ",'" & txtRuta.Text & "','" & cmbEntrega.Text & "','" & StrTipo & "')"
            
            strCadena = "Insert Into TB_Lotes (ds_periodo,ds_año,ds_batch,nu_imagenes,ds_path,ds_tipo) values (" & _
                "'" & IIf(IsNull(txtCampaña.Text), " ", txtCampaña.Text) & "' ," & _
                "'" & IIf(IsNull(txtZona.Text), " ", txtZona.Text) & "' ," & _
                "'" & IIf(IsNull(txtBatch.Text), " ", txtBatch.Text) & "' ," & _
                Val(TXTCANTIDAD) & ",'" & txtRuta.Text & "','1')"
                
                
                
                cnServer.Execute strCadena
                rsDatos.Requery
                Columnas
        End If
End Sub
Private Sub CmdEliminar_Click()
    If rsDatos.RecordCount > 0 Then
        intRespuesta = MsgBox("Desea eliminar el Lote asignado al Periodo : " & rsDatos!Periodo & ", Zona: " & rsDatos!año & ", Lote: " & rsDatos!LOTE, vbQuestion + vbYesNo, "Exportar")
        If intRespuesta = vbYes Then
            strCadena = "Delete from TB_Lotes where id_lote= " & rsDatos!id_lote
            'strCadena = "Delete from LOTES where id_lote= " & rsDatos!id_lote
            cnServer.Execute strCadena
            rsDatos.Requery
            Columnas
        End If
    End If
End Sub

Private Sub cmdFormat_Click()
Dim Intbatch As Double
If lstdirectorios.SelCount > 0 Then
    
    
      If Len(lstdirectorios) = 19 Then
            If rsSQL.State = 1 Then rsSQL.Close
      
            rsSQL.Source = "Select nro_batch from TB_batch where nro_batch=" & Val(Mid(lstdirectorios, 13, 7))
            rsSQL.Open
            If rsSQL.RecordCount > 0 Then
      
            MsgBox "Este directorio ya fue formateado", vbInformation, App.EXEName
            Else
            Flow_Frmformat.Show vbModal
                          If Len(strPeriodo) > 0 And Len(strAño) > 0 Then
                              If rsSQL.State = 1 Then rsSQL.Close
                              rsSQL.Source = "Select nro_batch from TB_batch order by nro_batch DESC"
                              'rsSql.Source = "Select id_lote from LOTES where ds_batch = '" & Me.txtBatch.Text & "' and ds_campaña = '" & txtCampaña.Text & "' and ds_zona = '" & txtZona.Text & "'"
                              rsSQL.Open
                              If rsSQL.RecordCount > 0 Then
                              Intbatch = rsSQL.Fields("nro_batch") + 1
                              strCadena = "exec sp_insertarbatch " & Intbatch
                              cnServer.Execute strCadena
                              Else
                              Intbatch = 1
                              strCadena = "exec sp_insertarbatch " & Intbatch
                              cnServer.Execute strCadena
                              End If
                          '  Archivos_No_Procesados Dir1.Path & "\" & lstdirectorios
                            Archivos_Liquidacion Dir1.Path, lstdirectorios, Intbatch
                            Dir1.Refresh
                            lstdirectorios.Refresh
                            File1.Refresh
                            'Dir1.Path
                          
                        End If
            End If
                          
        Else
           Flow_Frmformat.Show vbModal
                          If Len(strPeriodo) > 0 And Len(strAño) > 0 Then
                              If rsSQL.State = 1 Then rsSQL.Close
                              rsSQL.Source = "Select nro_batch from TB_batch order by nro_batch DESC"
                              'rsSql.Source = "Select id_lote from LOTES where ds_batch = '" & Me.txtBatch.Text & "' and ds_campaña = '" & txtCampaña.Text & "' and ds_zona = '" & txtZona.Text & "'"
                              rsSQL.Open
                              If rsSQL.RecordCount > 0 Then
                              Intbatch = rsSQL.Fields("nro_batch") + 1
                              strCadena = "exec sp_insertarbatch " & Intbatch
                              cnServer.Execute strCadena
                              Else
                              Intbatch = 1
                              strCadena = "exec sp_insertarbatch " & Intbatch
                              cnServer.Execute strCadena
                              End If
                          '  Archivos_No_Procesados Dir1.Path & "\" & lstdirectorios
                            Borrar_Imagenes_Scan Str(Intbatch)
                            Archivos_Liquidacion Dir1.Path, lstdirectorios, Str(Intbatch)
                            
                            intRespuesta = MsgBox("Desea Formatear imagenes con Drop Poup ?", vbQuestion + vbYesNo, "Lotes")
                             If (intRespuesta = vbYes) Then

                                Flow_Directorio.Show vbModal
                                Borrar_Imagenes_Scan Str(Intbatch)
                                Archivos_Liquidacion_DroPoup strRutaDP, strDirDP, Intbatch
                                Dir1.Refresh
                                lstdirectorios.Refresh
                                File1.Refresh
                             Else
                                Dir1.Refresh
                                lstdirectorios.Refresh
                                File1.Refresh
                             End If

                          
                        End If
            
        End If
   
End If
End Sub

Private Sub cmdIniciar_Click()
    intRespuesta = MsgBox("Desea iniciar la tabla de LOTES ?", vbQuestion + vbYesNo, "Lotes")
    If (intRespuesta = vbYes) Then
        strEjecutar = "Insert into TB_HISTORICO select * from TB_LOTES where ds_estado='IMPORTADO'"
        cnServer.Execute strEjecutar
        strCadena = "Delete From TB_LOTES where ds_estado='IMPORTADO'"
        cnServer.Execute strCadena
        rsDatos.Requery
        Columnas
    End If
    MsgBox "Proceso teminado", vbInformation, App.EXEName
End Sub
Private Sub cmdRefresh_Click()
    rsDatos.Requery
    Columnas
End Sub
Private Sub datDatos_HeadClick(ByVal ColIndex As Integer)
    If rsDatos.State = 1 Then
        rsDatos.Close
        'rsDatos.Source = "Select fecha,caja as caja, lote as lote,imagenes,estado,usuario,hora_inicio,hora_fin,path from LOTEs order by " & ColIndex + 1
        rsDatos.Source = "Select fc_fecha as fecha,ds_campaña as campaña,ds_zona as zona,ds_batch as batch,ds_envio as Entrega,nu_imagenes as imagenes,ds_estado as estado,ds_usuario as usuario,fc_hora_inicio as hora_inicio, fc_hora_fin as hora_fin, ds_path as path, id_lote from TB_LOTES order by " & ColIndex + 1
        rsDatos.Open
        Set datDatos.DataSource = rsDatos
        Columnas
    End If
End Sub

Private Sub EnumFilesUnder(Optional blnType As Boolean, Optional strPath As String)
    Dim hFind As Long
    Dim sPath As String, oldPath As String
    Dim wf As WIN32_FIND_DATA
    Dim n2 As Node

    sPath = strPath
    
    hFind = FindFirstFile(sPath, wf)
    Do
        DoEvents

        If InStr(wf.cFileName, Chr$(0)) > 0 Then
            sPath = Left$(wf.cFileName, _
                InStr(wf.cFileName, Chr$(0)) - 1)
        Else
            sPath = wf.cFileName
        End If
        If Len(sPath) = 0 Or StrComp(sPath, oldPath) = 0 Then
            Exit Do
        ElseIf sPath = "." Or sPath = ".." Then
            GoTo Iterate
        End If
        
        If (wf.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            lstdirectorios.AddItem wf.cFileName
            
            DoEvents
        End If
Iterate:
        FindNextFile hFind, wf
        oldPath = sPath
    Loop
    FindClose hFind
End Sub

Private Sub Dir1_Change()
Dim intInicio As Integer
Dim intFin As Integer
Dim strTempo As String
On Local Error GoTo Err:
'    If StrRutaAF = Mid(Dir1.Path, 1, Len(StrRutaAF)) Then 'StrRutaAF
       File1.Path = Dir1.Path
       lstdirectorios.Clear
       Call EnumFilesUnder(, Dir1.Path & "\*.*")
       'lstdirectorios
       txtRuta.Text = UCase(File1.Path)
       TXTCANTIDAD.Text = File1.ListCount & " IMAGENES"
       strCadena = Dir1.Path
        'intInicio = InStrRev(strCadena, "\") - 1
        'intFin = InStrRev(strCadena, "\", intInicio) + 1
        'strTempo = Mid(strCadena, InStrRev(strCadena, "\", InStrRev(strCadena, "\") - 1) + 1, 10)
       
       strTempo = Mid(strCadena, InStrRev(strCadena, "\") + 1)
       intInicio = InStr(1, strTempo, "-") - 1
       'txtCampaña.Text = IIf(IsNumeric(Mid(strTempo, 1, 6)), Mid(strTempo, 1, 6), " ")
       txtCampaña.Text = IIf(IsNumeric(Mid(strTempo, 5, 2)), Mid(strTempo, 5, 2), " ")
       txtZona.Text = IIf(IsNumeric(Mid(strTempo, 8, 4)), Mid(strTempo, 8, 4), " ")
       txtBatch.Text = IIf(IsNumeric(Mid(strTempo, 13, 7)), Mid(strTempo, 13, 7), " ")
       If lstdirectorios.ListCount > 0 And Dir1.ListIndex > 0 Then
        lstdirectorios.ListIndex = Dir1.ListIndex
       End If
       Exit Sub
 '   Else
 '   lstdirectorios.Clear
 '   File1.Path = App.Path
 '   End If
Err:
End Sub

Private Sub Dir1_Click()
On Error Resume Next
Contlista = "1"
If lstdirectorios.ListCount > 0 And Dir1.ListIndex > 0 Then
        lstdirectorios.ListIndex = Dir1.ListIndex
        'lstdirectorios.List (Dir1.ListIndex)
End If
Contlista = "0"
End Sub

Private Sub Drive1_Change()
On Error GoTo errores
    Dir1.Path = Drive1.Drive
    Exit Sub
errores:
    MsgBox "Drive no habilitado", vbCritical, "Drive"
End Sub
Private Sub Form_Load()
Dim intEntrega As Integer
    Contlista = "0"
    Me.Height = 9540
    Me.Width = 14820
    
    Me.Left = (Me.Width / 2) - (Me.Width / 2)
    Me.Top = (Me.Height / 2) - (Me.Height / 2)
    
    'Dir1.Path = App.Path
    'Drive1.Drive = Mid(Dir1.Path, 1, 1)
    
    Dir1.Path = StrRutaAF
    Drive1.Drive = Dir1.Path
    'cmbEntrega.Clear
    
'    For intEntrega = 1 To 999
'        cmbEntrega.AddItem Padln(Trim(Str(intEntrega)), 3, "0")
'    Next
    TXTCANTIDAD.Text = "0 IMAGENES"
        
    rsDatos.CursorType = adOpenDynamic
    rsDatos.LockType = adLockBatchOptimistic
    rsDatos.ActiveConnection = cnServer
    
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    
    If rsDatos.State = 1 Then rsDatos.Close
    'rsDatos.Source = "Select fc_fecha as fecha,ds_campaña as campaña,ds_zona as zona,ds_batch as batch,ds_envio as Entrega,nu_imagenes as imagenes,ds_estado as estado,ds_usuario as usuario,fc_hora_inicio as hora_inicio, fc_hora_fin as hora_fin, ds_path as path, id_lote from TB_Lotes"
    'rsDatos.Source = "Select fc_fecha as Fecha,ds_periodo as Periodo,ds_año as Año,ds_batch as Batch,nu_imagenes as Imagenes,ds_estado as Estado,ds_usuario as Usuario,fc_hora_inicio as Hora_inicio, fc_hora_fin as Hora_fin, ds_path as Path, id_lote from TB_Lotes"
    rsDatos.Source = "Select fc_fecha as Fecha,ds_periodo as Periodo,ds_año as Año,ds_batch as Lote,nu_imagenes as Imagenes,ds_estado as Estado,ds_usuario as Usuario,fc_hora_inicio as Hora_inicio, fc_hora_fin as Hora_fin, ds_path as Path, id_lote from TB_Lotes"
    rsDatos.Open
    Set datDatos.DataSource = rsDatos
    Columnas
    'Cbotipo.ListIndex = 0
End Sub

Private Sub Form_Resize()
    If Me.Width > 4000 Then datDatos.Width = Me.Width - 4000 '3600
    If Me.Height > 4000 Then datDatos.Height = Me.Height - 3500 '3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err:
    If rsDatos.State = 1 Then rsDatos.Close
    If rsSQL.State = 1 Then rsSQL.Close
    Exit Sub
Err:
End Sub

Private Sub lstdirectorios_Click()
If Contlista <> "1" Then
 Dir1.ListIndex = lstdirectorios.ListIndex
End If
End Sub

Private Sub TXTCANTIDAD_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtRuta_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Sub Columnas()
    datDatos.Columns(0).Width = 2000
    datDatos.Columns(1).Width = 1000
    datDatos.Columns(2).Width = 1200
    datDatos.Columns(3).Width = 1000
    datDatos.Columns(4).Width = 1000
    datDatos.Columns(5).Width = 1000
    datDatos.Columns(6).Width = 1500
End Sub




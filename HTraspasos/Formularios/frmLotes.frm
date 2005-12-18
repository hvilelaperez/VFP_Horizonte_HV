VERSION 5.00
Begin VB.Form frmLotes 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lotes"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReabrir 
      Caption         =   "Reabrir lote"
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
      Left            =   4200
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ListBox lstLotes2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Seleccionar"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   480
      Width           =   1455
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
      Left            =   4200
      TabIndex        =   2
      Top             =   7050
      Width           =   1455
   End
   Begin VB.CommandButton cmdInformacion 
      Caption         =   "Información"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox lstLotes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6990
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLotes As Recordset
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lstLotes.ListCount > 0 Then
       If lstLotes2.SelCount > 0 Then
            If vbYes = MsgBox("Desea eliminar el lote seleccionado ?", vbCritical + vbYesNo, "Eliminar lote") Then
                If LoteSinProceso(lstLotes.List(lstLotes2.ListIndex), frmScan.lngVolumenActual) = True Then
                    If Trim(Dir(frmScan.strDirectorioImagenes & "\" & Format(lstLotes2.List(lstLotes2.ListIndex), "00000000"), vbDirectory)) <> "" Then
                        
                        If Val(lstLotes.List(lstLotes2.ListIndex)) = Val(frmScan.lblLote) Then
                            frmScan.TerminarLote
                        End If
                        
                        clsDataScan.Conexion.Execute "delete from _imagenes where idvolumen=" & frmScan.lngVolumenActual & " and lote=" & Val(lstLotes.List(lstLotes2.ListIndex))
                        clsDataScan.Conexion.Execute "delete from _scan where idvolumen=" & frmScan.lngVolumenActual & " and lote=" & Val(lstLotes.List(lstLotes2.ListIndex))
                        clsDataScan.Conexion.Execute "delete from _prescan where idvolumen=" & frmScan.lngVolumenActual & " and lote=" & Val(lstLotes.List(lstLotes2.ListIndex))
                        clsDataScan.Conexion.Execute "update _volumenes set lote_actual=0 where idvolumen=" & frmScan.lngVolumenActual
                       
                        DelTree (frmScan.strDirectorioImagenes & "\" & Format(lstLotes2.List(lstLotes2.ListIndex), "00000000"))
                        
                        CargarLotes
                    End If
                Else
                    MsgBox "El lote seleccionado actualmente esta siendo procesado", vbCritical, "Eliminar lote"
                End If
            End If
       End If
    End If
End Sub
Public Function LoteSinProceso(lngLote As Long, lngVolumen As Long) As Boolean
Dim bolEstadoLote As Boolean
Dim rsSql As Recordset
    bolEstadoLote = False
    Set rsSql = New Recordset
    Set rsSql = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Scan,'icr,kfi,indice','where lote=" & lngLote & " and idvolumen=" & lngVolumen & "',''")
    If rsSql.RecordCount > 0 Then
        If (rsSql!icr + rsSql!kfi + rsSql!indice) = 0 Then
            bolEstadoLote = True
        Else
            bolEstadoLote = False
        End If
    End If
    If rsSql.State = 1 Then rsSql.Close
    Set rsSql = Nothing
    LoteSinProceso = bolEstadoLote
End Function

Private Sub cmdInformacion_Click()
    If lstLotes2.ListCount > 0 Then
        If lstLotes2.ListIndex > -1 Then
            bolModoForm = False
            strLoteActual = lstLotes.List(lstLotes2.ListIndex)
            frmLote.Show vbModal
        End If
    End If
End Sub
Private Sub cmdNuevo_Click()
    If lstLotes.ListCount > 0 Then
       If lstLotes2.SelCount > 0 Then
            lngLoteActual = lstLotes.List(lstLotes2.ListIndex)
            frmScan.lblLote.Caption = " " & lstLotes.List(lstLotes2.ListIndex)
            frmScan.lblLote2.Caption = " " & lstLotes2.List(lstLotes2.ListIndex)
            Unload Me
       End If
    End If
End Sub

Private Sub cmdReabrir_Click()
    If lstLotes2.ListCount > 0 Then
        If lstLotes2.SelCount > 0 Then
            If LoteActivo(lstLotes.List(lstLotes2.ListIndex)) = False Then
                MsgBox "El lote seleccionado esta abierto", vbInformation, "Reabrir lote"
            Else
                If vbYes = MsgBox("Desea reabrir el lote seleccionado ?", vbQuestion + vbYesNo, "Reabrir lote") Then
                    clsDataScan.Conexion.Execute ("EXECUTE sp_ReabrirLoteScan " & lngIdVolumenActual & "," & lstLotes.List(lstLotes2.ListIndex))
                    MsgBox "El lote ha sido reabierto para el proceso SCAN", vbInformation, "Reabrir lote"
                End If
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    CargarLotes
End Sub
Sub CargarLotes()
    If clsDataScan.ConexionActiva = True Then
        Set rsLotes = New Recordset
        If rsLotes.State = 1 Then rsLotes.Close
        Set rsLotes = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _SCan,[*],'where idVolumen=" & lngIdVolumenActual & "','order by lotefisico'")
        If rsLotes.RecordCount > 0 Then
            rsLotes.MoveFirst
            lstLotes.Clear
            lstLotes2.Clear
            Do While Not rsLotes.EOF
                lstLotes.AddItem rsLotes!Lote
                lstLotes2.AddItem rsLotes!lotefisico
                rsLotes.MoveNext
            Loop
        End If
        If rsLotes.State = 1 Then rsLotes.Close
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsLotes = Nothing
End Sub
Private Sub lstLotes2_DblClick()
    If lstLotes2.ListCount > 0 Then
        If lstLotes2.ListIndex > -1 Then
            cmdInformacion_Click
        End If
    End If
End Sub

Function DelTree(ByVal strDir As String) As Long
    Dim X As Long
    Dim intAttr As Integer
    Dim strAllDirs As String
    Dim strFile As String
    DelTree = -1
On Error Resume Next
    strDir = Trim$(strDir)
    If Len(strDir) = 0 Then Exit Function
    If Right$(strDir, 1) = "\" Then strDir = Left$(strDir, Len(strDir) - 1)
    If InStr(strDir, "\") = 0 Then Exit Function
    intAttr = GetAttr(strDir)
    If (intAttr And vbDirectory) = 0 Then Exit Function
    strFile = Dir$(strDir & "\*.*", vbSystem Or vbDirectory Or vbHidden)
    Do While Len(strFile)
    If strFile <> "." And strFile <> ".." Then
    intAttr = GetAttr(strDir & "\" & strFile)
    If (intAttr And vbDirectory) Then
    strAllDirs = strAllDirs & strFile & Chr$(0)
    Else
    If intAttr <> vbNormal Then
    SetAttr strDir & "\" & strFile, vbNormal
    If Err Then DelTree = Err: Exit Function
    End If
    Kill strDir & "\" & strFile
    If Err Then DelTree = Err: Exit Function
    End If
    End If
    strFile = Dir$
    Loop
    Do While Len(strAllDirs)
    X = InStr(strAllDirs, Chr$(0))
    strFile = Left$(strAllDirs, X - 1)
    strAllDirs = Mid$(strAllDirs, X + 1)
    X = DelTree(strDir & "\" & strFile)
    If X Then DelTree = X: Exit Function
    Loop
    RmDir strDir
    If Err Then
        DelTree = Err
    Else
        DelTree = 0
    End If
End Function

Public Function LoteActivo(lngLoteActual As Long) As Boolean
Dim bolEstadoLote As Boolean
Dim rsSql As Recordset
    bolEstadoLote = False
    Set rsSql = New Recordset
    Set rsSql = clsDataScan.Conexion.Execute("EXECUTE sp_EstadoScan " & lngIdVolumenActual & "," & lngLoteActual)
    If rsSql.RecordCount > 0 Then
        bolEstadoLote = rsSql!scan
    End If
    If rsSql.State = 1 Then rsSql.Close
    Set rsSql = Nothing
    LoteActivo = bolEstadoLote
End Function


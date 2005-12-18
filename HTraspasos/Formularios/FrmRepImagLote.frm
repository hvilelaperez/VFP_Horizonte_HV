VERSION 5.00
Begin VB.Form FrmRepImagLote 
   Caption         =   "Reporte de Imagenes"
   ClientHeight    =   2865
   ClientLeft      =   5070
   ClientTop       =   3630
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5370
   Begin VB.CommandButton CmdProceso 
      Caption         =   "&Proceso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CmdResuLot 
      Caption         =   "&Resum x Lote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdRImgSID 
      Caption         =   "I&mag S/ ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CmdRepo 
      Caption         =   "&Imag x Lote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame FraRangoLot 
      Caption         =   "Ingrese Rango de Lotes:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   3255
      Begin VB.TextBox TxtlotFin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtLotIni 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lote Final"
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
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lote Inicial"
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
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   915
      End
   End
End
Attribute VB_Name = "FrmRepImagLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsResuLot As New Recordset     'Resumen por Lote
Public RsRepo As New Recordset        'Reporte
Public RsRImagSID As New Recordset    'Reporte Imag S/ID

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    Con.ConexionSQL "sa", "", "DIGITACION\DIGITACION", "BDTraspasos2"
'    DatosServidor
'***
    'Call Con.AbrirTabla("select * from TMPLote", RsT, adOpenStatic, adLockOptimistic)
    CmdProceso.Enabled = False
    CmdResuLot.Enabled = False
    CmdRepo.Enabled = False
    CmdRImgSID.Enabled = False
    TxtLotIni = 0
    TxtlotFin = 0
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub TxtlotIni_GotFocus()
    Marcado TxtLotIni
End Sub

Private Sub TxtlotFin_GotFocus()
    Marcado TxtlotFin
End Sub

Private Sub TxtLotIni_Change()
    CmdProceso.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
    CmdResuLot.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
    CmdRepo.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
    CmdRImgSID.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
End Sub

Private Sub TxtlotFin_Change()
    CmdProceso.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
    CmdResuLot.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
    CmdRepo.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
    CmdRImgSID.Enabled = val(TxtLotIni) > 0 And val(TxtlotFin) > 0
End Sub

Private Sub CmdRepo_Click()
If val(TxtLotIni) > val(TxtlotFin) Then
    MsgBox "Rango de Lotes con error", vbCritical, "Ingreso de Rango de Lotes"
    TxtLotIni.SetFocus
Else
    If RsRepo.State = 1 Then RsRepo.Close
    SQLstr = "SELECT Lote, IDImagen, Imagen, Barcode, Posicion, Status,TipImag FROM _Imagenes WHERE Lote >=" & TxtLotIni & " AND Lote <=" & TxtlotFin & " order by Lote, Imagen"
    RsRepo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtRImagLote.DataSource = RsRepo
    DtRImagLote.Show
End If

End Sub

Private Sub CmdRImgSID_Click()
If val(TxtLotIni) > val(TxtlotFin) Then
    MsgBox "Rango de Lotes con error", vbCritical, "Ingreso de Rango de Lotes"
    TxtLotIni.SetFocus
Else
    If RsRImagSID.State = 1 Then RsRImagSID.Close
    SQLstr = "SELECT Lote, IDImagen, Imagen, Barcode, Posicion, Status,TipImag FROM _Imagenes WHERE Lote >=" & TxtLotIni & " AND Lote <=" & TxtlotFin & " AND IDImagen is null order by Lote, Imagen"
    RsRImagSID.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtRImagLote.DataSource = RsRImagSID
    DtRImagLote.Show
End If
End Sub

Private Sub CmdProceso_Click()
If val(TxtLotIni) > val(TxtlotFin) Then
    MsgBox "Rango de Lotes con error", vbCritical, "Ingreso de Rango de Lotes"
    TxtLotIni.SetFocus
Else
    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = "delete tmplote"
        .Execute
        .CommandText = "insert into tmplote (lote,TotalTif,TotSol,TotDNI,TotHC,TotOtr) " & _
        "select lote,count(*),0,0,0,0 " & _
        "from _imagenes where Lote >= " & val(TxtLotIni) & " AND Lote <= " & val(TxtlotFin) & " group by lote order by lote "
        .Execute
        .CommandText = "update tmplote " & _
        "set totsol = isnull((select count(*) from _imagenes where tipimag='S' and lote=a.lote group by lote),0)," & _
        "totdni = isnull((select count(*) from _imagenes where tipimag='D' and lote=a.lote group by lote),0), " & _
        "tothc = isnull((select count(*) from _imagenes where tipimag='H' and lote=a.lote group by lote),0), " & _
        "tototr = isnull((select count(*) from _imagenes where tipimag<>'H' and tipimag<>'S' and tipimag<>'D' and lote=a.lote group by lote),0) " & _
        "from tmplote a where Lote >= " & val(TxtLotIni) & " AND Lote <= " & val(TxtlotFin) & " "
        .Execute
        MsgBox "Proceso Ejecutado", vbInformation, "Proceso de Conteo de Imagenes"
    End With
End If
End Sub

Private Sub CmdResuLot_Click()
If val(TxtLotIni) > val(TxtlotFin) Then
    MsgBox "Rango de Lotes con error", vbCritical, "Ingreso de Rango de Lotes"
    TxtLotIni.SetFocus
Else
    If RsResuLot.State = 1 Then RsResuLot.Close
    SQLstr = "Select * from TMPLote WHERE Lote >= " & TxtLotIni & " AND Lote <= " & TxtlotFin & " ORDER BY Lote"
    'SQLstr = "Select Lote, num=count(*) from _Imagenes WHERE Lote >= " & TxtLotIni & " AND _Imagenes.Lote <= " & TxtlotFin & " GROUP BY Lote ORDER BY Lote"
    'SQLstr = "Select _Imagenes.lote, num=count(*), contador from _Imagenes, _Scan WHERE _Imagenes.Lote=_Scan.Lote AND _Imagenes.Lote >= " & TxtLotIni & " AND _Imagenes.Lote <=" & TxtlotFin & " GROUP BY _Imagenes.Lote, contador ORDER BY _Imagenes.Lote"
    RsResuLot.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtRResuLote.DataSource = RsResuLot
    DtRResuLote.Show
End If
End Sub

Private Sub TxtLotIni_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtlotFin.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtlotFin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdProceso.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub


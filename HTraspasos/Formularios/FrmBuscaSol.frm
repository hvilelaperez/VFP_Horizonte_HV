VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "IMGEDIT.OCX"
Begin VB.Form FrmBuscaSol 
   Caption         =   "Busca Solicitud"
   ClientHeight    =   10575
   ClientLeft      =   360
   ClientTop       =   450
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   14550
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9495
      Begin VB.TextBox TxtNom1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox TxtCuspp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxtApeMat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtApePat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TxtNumSol 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nombre1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7560
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cuspp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ap Materno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ap Paterno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Solicitud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.CommandButton CmdBusca 
      Caption         =   "&Buscar"
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
      Left            =   9720
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton CmdVerImag 
      Caption         =   "Ver &Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   7
      Top             =   240
      Width           =   975
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
      Left            =   13440
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DtGBuscaSol 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   7646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "fecsal"
         Caption         =   "F Salida"
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
         DataField       =   "lote"
         Caption         =   "Lote"
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
      BeginProperty Column02 
         DataField       =   "inddoc"
         Caption         =   "Tipo"
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
      BeginProperty Column03 
         DataField       =   "numsol"
         Caption         =   "N Solicitud"
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
      BeginProperty Column04 
         DataField       =   "secu"
         Caption         =   "sec"
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
      BeginProperty Column05 
         DataField       =   "cuspp"
         Caption         =   "CUSPP"
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
      BeginProperty Column06 
         DataField       =   "apepat"
         Caption         =   "Ap Paterno"
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
      BeginProperty Column07 
         DataField       =   "apemat"
         Caption         =   "Ap Materno"
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
      BeginProperty Column08 
         DataField       =   "nombre1"
         Caption         =   "Nombre 1"
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
      BeginProperty Column09 
         DataField       =   "nombre2"
         Caption         =   "Nombre 2"
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
      BeginProperty Column10 
         DataField       =   "imagen"
         Caption         =   "Imagen TIFF"
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
      BeginProperty Column11 
         DataField       =   "idimagen"
         Caption         =   "ID Imagen"
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
      BeginProperty Column12 
         DataField       =   "codusu"
         Caption         =   "Codusu"
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
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
      EndProperty
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   4695
      Left            =   120
      TabIndex        =   17
      Top             =   5760
      Width           =   14295
      _Version        =   131074
      _ExtentX        =   25215
      _ExtentY        =   8281
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   121450752
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   28
      lReserved2      =   28
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10800
      TabIndex        =   9
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label LblTotReg 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10920
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmBuscaSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsBuscaSol As New Recordset
Dim RsImag As New Recordset
Public xfrmrep As String
Public xnumsol As String
Public xidimagen As String
Public xlote As String
Dim xlote2 As String
Dim xtipimag1 As String
Dim xtipimag2 As String
Dim xrutavol As String
Dim ximagen As String
Dim xnumpag As Integer
Dim B1 As String
Dim Activado As Boolean

Private Sub DtGBuscaSol_HeadClick(ByVal ColIndex As Integer)
    RsBuscaSol.Sort = DtGBuscaSol.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "2326", "si-006", "BDTraspasosdig2"
'    Con.ConexionSQL "sa", "", "192.168.20.57", "BDTraspasosdig2"
'    DatosServidor
'***
    xfrmrep = "B"
    B1 = ""
    CmdVerImag.Enabled = False
End Sub

Private Sub CmdBusca_Click()
    B1_1 = ""
    B1_2 = ""
    B1_3 = ""
    B1_4 = ""
    B1_5 = ""
    B1_6 = ""
    B1_7 = ""
    B1_8 = ""

    B1_1 = "SELECT s.fecsal,s.lote,s.inddoc,s.numsol,s.secu,s.cuspp,s.apepat,s.apemat,s.nombre1,s.nombre2,s.idimagen,i.imagen,i.posicion,i.tipimag,s.codusu " & _
            "FROM solicitud s JOIN _imagenes i ON s.idimagen = i.idimagen " & _
            "WHERE ( i.tipimag <> 'H' ) "
            
'            "WHERE ( i.tipimag = 'S' or i.tipimag = 'E' ) "

    If TxtNumSol <> "" Then
        B1_2 = " and s.numsol = '" & TxtNumSol & "' "
    End If
    If TxtCuspp <> "" Then
        B1_3 = " and s.cuspp like '%" & TxtCuspp & "%' "
    End If
    If TxtApePat <> "" Then
        B1_4 = " and s.apepat like '%" & TxtApePat & "%' "
    End If
    If TxtApeMat <> "" Then
        B1_6 = " and s.apemat like '%" & TxtApeMat & "%' "
    End If
    If TxtNom1 <> "" Then
        B1_7 = " and s.nombre1 like '%" & TxtNom1 & "%' "
    End If
    
    B1_8 = " ORDER BY s.lote, s.numsol, s.secu, i.posicion "
    
    B1 = B1_1 + B1_2 + B1_3 + B1_4 + B1_5 + B1_6 + B1_7 + B1_8
    
    If RsBuscaSol.State = 1 Then RsBuscaSol.Close
        
'    SQLstr = "SELECT s.fecsal,s.lote,s.numsol,s.secu,s.cuspp,s.apepat,s.apemat,s.nombre1,s.nombre2,s.idimagen,i.imagen,i.posicion,i.tipimag,s.codusu " & _
'                "FROM solicitud s, _imagenes i " & _
'                "Where i.tipimag = 'S' and s.idimagen = i.idimagen  " & xcondicion & " " & _
'                " ORDER BY s.lote, s.numsol, s.secu, i.posicion "
    
    LlenaControles
End Sub

Private Sub LlenaControles()
    LblTotReg = 0
    
    RsBuscaSol.Open B1, Con.Cnx, adOpenStatic, adLockReadOnly
    
    With RsBuscaSol
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            Set DtGBuscaSol.DataSource = Nothing
            DtGBuscaSol.Refresh
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Verificacion"
            Exit Sub
        Else
            Set DtGBuscaSol.DataSource = RsBuscaSol
            DtGBuscaSol.Refresh
            LblTotReg.Caption = RsBuscaSol.RecordCount
            CargaImag
        End If
    End With
End Sub

Private Sub DtGBuscaSol_SelChange(Cancel As Integer)
    Activado = True
    CmdVerImag.Enabled = True
    CargaImag
End Sub

Private Sub CmdVerImag_Click()
    xfrmrep = "B"
    If Activado Then
        On Error GoTo Salir
        With RsBuscaSol
            xlote = .Fields("lote")
            xnumsol = .Fields("numsol")
            xidimagen = .Fields("idimagen")
            FrmMuestraImg.Show
        End With
    Else
        MsgBox "Selecciona la solicitud a buscar", vbCritical, "Error en la busqueda "
    End If
Salir:
   Exit Sub
End Sub

'Private Sub TxtFecha_GotFocus()
'    Marcado TxtFecha
'End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub TxtNumSol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtCuspp.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol_LostFocus()
    TxtNumSol = Format(TxtNumSol, "000000#")
End Sub

Private Sub TxtCuspp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtApePat_Change()
    Select Case KeyAscii
        Case 65 To 90
        Case 209
        Case 32, 8
        Case 13
            TxtApeMat.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtApeMat_Change()
    Select Case KeyAscii
        Case 65 To 90
        Case 209
        Case 32, 8
        Case 13
            TxtNom1.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtApeMat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNom1_Change()
    Select Case KeyAscii
        Case 65 To 90
        Case 209
        Case 32, 8
        Case 13
            TxtNom2.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNom1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CargaImag()
    With RsBuscaSol
        xlote = Format(!lote, "0000000#")
        xlote2 = Format(!lote, "0000000#")
        xnumsol = Format(!numsol, "000000#")
        xtipimag1 = "S"
        xtipimag2 = "E"
        xidimagen = "01" + xnumsol

    End With
    xnumpag = 1
      
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient
    
    SQLImag = "SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
    "WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and ( TipImag = '" & xtipimag1 & "' or TipImag = '" & xtipimag2 & "' ) and substring(idimagen,3,9)= '" & xnumsol & "' order by posicion "
    
    RsImag.Open SQLImag, Con.Cnx, adOpenStatic, adLockReadOnly
    
    With RsImag          'coneccion a _Imagenes
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Imagen"
            Exit Sub
        Else
            .MoveFirst
            MuestraImg
        End If
    End With
               
End Sub

Private Sub MuestraImg()
    
    With RsImag  'coneccion a _Imagenes
        .Find "[posicion]='" & xnumpag & "'"
'        .Find "[idimagen]='" & xidimagen & "'"
        
        If .EOF Then .MovePrevious
        ximagen = IIf(IsNull(!imagen), "", !imagen)
        xrutavol = .Fields("ruta")
        xrutavol = Replace(xrutavol, "\\", "\")
        xidvolumen = IIf(IsNull(!idvolumen), "", !idvolumen)
    End With
      
    ImgEdit1.Image = xrutavol & "\" & xlote2 & "\" & ximagen & ".TIF"
    ImgEdit1.Enabled = True

    ImgEdit1.FitTo 1
'    ImgEdit1.Zoom = 50
    ImgEdit1.Display
    
'    ImgEdit1.ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 3)
'    ImgEdit1.ScrollPositionX = 0
   
    ImgEdit1.ScrollPositionX = (Int((ImgEdit1.ImageScaleWidth / 3)) * 2)
    ImgEdit1.Refresh
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
' alto
    ImgEdit1.Height = FrmBuscaSol.Height - 7800
' ancho
    ImgEdit1.Width = FrmBuscaSol.Width - 300
' alto
'    DtGHV.Height = 1200 - ImgEdit1.Height
    ImgEdit1.Height = FrmBuscaSol.Height - 6500
End Sub

Sub MovImg(KeyCode As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF12 Then
        If Not (RsImag.EOF And RsImag.BOF) Then
            RsImag.MoveNext
            If RsImag.EOF Then RsImag.MoveLast
            'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
            ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If

    If KeyCode = vbKeyF11 Then
        If Not (RsImag.EOF And RsImag.BOF) Then
            RsImag.MovePrevious
            If RsImag.BOF Then RsImag.MoveFirst
            'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
            ImgEdit1.Image = Rutaimag(RsImag!Ruta, RsImag!lotefisico, RsImag!imagen)
      
            ImgEdit1.Display
            ImgEdit1.Refresh
        End If
    End If
End Sub

Private Sub ImgEdit1_DblClick()
    On Error Resume Next
    ImgEdit1.FitTo 1
End Sub

Private Sub ImgEdit1_KeyDown(KeyCode As Integer, Shift As Integer)
    MovImg KeyCode
End Sub

Private Sub ImgEdit1_MarkEnd(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MarkType As Integer, ByVal GroupName As String)
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

Private Sub ImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    ImgEdit1.ZoomToSelection
    ImgEdit1.RemoveAllOCRMarks
    ImgEdit1.Refresh
End Sub

'********************
Private Sub ImgEdit1_Click()
    On Error Resume Next
End Sub

Private Sub UpDown1_DownClick()
    On Error Resume Next
    If Not (RsImag.EOF And RsImag.BOF) Then
        RsImag.MoveNext
        If RsImag.EOF Then RsImag.MoveLast
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
        ImgEdit1.Image = Rutaimag(xrutavol, RsImag!lotefisico, RsImag!imagen)
        ImgEdit1.Display
        ImgEdit1.Refresh
    End If
End Sub

Private Sub UpDown1_UpClick()
    On Error Resume Next
    If Not (RsImag.EOF And RsImag.BOF) Then
        RsImag.MovePrevious
        If RsImag.BOF Then RsImag.MoveFirst
        'ImgEdit1.Image = Rutaimag(rsImag!Ruta, txtLoteAfp, rsImag!imagen)
        ImgEdit1.Image = Rutaimag(xrutavol, RsImag!lotefisico, RsImag!imagen)
        ImgEdit1.Display
        ImgEdit1.Refresh
    End If
End Sub



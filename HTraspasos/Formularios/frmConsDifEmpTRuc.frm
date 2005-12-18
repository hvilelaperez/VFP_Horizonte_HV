VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "Imgedit.ocx"
Begin VB.Form frmConsDifEmpTRuc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Empleadores "
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dtgEmpleadores 
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "envio"
         Caption         =   "Envio"
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
         DataField       =   "Lote"
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
         DataField       =   "numsol"
         Caption         =   "Num. Sol"
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
         DataField       =   "Fecsal"
         Caption         =   "Fecsal"
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
         DataField       =   "numruc"
         Caption         =   "Ruc"
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
         DataField       =   "RazSoc"
         Caption         =   "Raz. Soc."
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
         DataField       =   "tRazSoc"
         Caption         =   "T. Raz. Soc."
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
         DataField       =   "Direccion"
         Caption         =   "Direccion"
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
         DataField       =   "tdirecc"
         Caption         =   "T. Direccion"
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
         DataField       =   "Numero"
         Caption         =   "Numero"
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
         DataField       =   "tnumero"
         Caption         =   "T. Numero"
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
         DataField       =   "Urb"
         Caption         =   "Urba"
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
         DataField       =   "Turba"
         Caption         =   "T. Urba"
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
      BeginProperty Column13 
         DataField       =   "ubigeo"
         Caption         =   "Ubigeo"
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
      BeginProperty Column14 
         DataField       =   "codgeo"
         Caption         =   "T. Ubigeo"
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
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2250.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2624.882
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2475.213
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2745.071
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1184.882
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Criterio"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.TextBox TxtFecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Por Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptLote 
         Caption         =   "Por Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   285
         Width           =   1215
      End
      Begin VB.TextBox txtEnvio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6720
         MaxLength       =   8
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optEnvio 
         Caption         =   "Por Envio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   1
         Top             =   285
         Width           =   1215
      End
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   14895
      _Version        =   131074
      _ExtentX        =   26273
      _ExtentY        =   6376
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit"
      AutoRefresh     =   -1  'True
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   452420608
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3592
      sReserved2      =   -3592
      lReserved1      =   38
      lReserved2      =   38
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Registros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblNumreg 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12840
      TabIndex        =   10
      Top             =   300
      Width           =   735
   End
End
Attribute VB_Name = "frmConsDifEmpTRuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmpleador As ADODB.Recordset
Dim RsImag As ADODB.Recordset
Sub Buscar()
Dim sqltexto As String
Dim Condicion As String

    sqltexto = " select c.envio, s.lote, a.numsol, a.fecsal,a.inddoc, a.numruc, a.inddoc ,a.Razsoc , tRazsoc = b.Razsoc , a.Direccion , tDirecc = b.Direcc, " & _
                " a.numero, tnumero = b.numero, a.Urb , b.Urba, a.ubigeo, b.codgeo from empleadores a " & _
                " left join truc b on a.numruc = b.numruc " & _
                " left join solicitud s on a.numsol = s.numsol and a.fecsal = s.fecsal " & _
                " and a.inddoc = s.inddoc and a.secuencial = s.secu " & _
                " left join ctrldoc c on s.lote = c.lote " & _
                " where (a.Razsoc <> b.Razsoc or a.Direccion <> b.Direcc or a.numero <> b.numero or " & _
                " a.Urb <> b.Urba or a.ubigeo <> b.codgeo) "


    If OptFecha.Value Then
        If Trim(TxtFecha) <> "//" And TxtFecha <> "" Then
            Condicion = " and a.fecsal = convert(datetime,'" & TxtFecha & "',103) "
        Else
            MsgBox "Ingresar Fecha a buscar", vbCritical, "Error en la busqueda "
            TxtFecha.SetFocus
            Exit Sub
        End If
    Else
        If OptLote Then
            If val(TxtLote) > 0 Then
                Condicion = " and s.Lote = '" & TxtLote & "' "
            Else
                MsgBox "Ingresar Lote a buscar", vbCritical, "Error en la busqueda "
                TxtLote.SetFocus
                Exit Sub
            End If
        Else
            If val(txtEnvio) > 0 Then
                Condicion = " and c.envio = '" & txtEnvio & "' "
            Else
                MsgBox "Ingresar Envio a buscar", vbCritical, "Error en la busqueda "
                txtEnvio.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    sqltexto = sqltexto & Condicion & " order by c.envio, s.lote, a.numsol, a.inddoc "
    
    With rsEmpleador
        If .State = 1 Then .Close
        .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
        
        Set Me.dtgEmpleadores.DataSource = rsEmpleador
        lblnumreg.Caption = .RecordCount
        
        If .RecordCount > 0 Then CargaImag
        
    End With
    

End Sub

Private Sub CargaImag()
    With rsEmpleador
        xlote = Format(!Lote, "0000000#")
        xnumsol = Format(!Numsol, "000000#")
        xtipimag = "S"
        xidimagen = "01" + xnumsol
    End With
    
    xnumpag = 1
      
    If RsImag.State = 1 Then RsImag.Close
    RsImag.CursorLocation = adUseClient
    
    SQLImag = " SELECT * FROM _Imagenes i join _Volumenes v on i.idvolumen = v.idvolumen  " & _
              " WHERE Lotefisico = '" & xlote & "' AND BarCode = 0 and idimagen = '" & xidimagen & "' and Tipo = 'B' and TipImag = '" & xtipimag & "' and substring(idimagen,3,9)= '" & xnumsol & "' order by posicion "
    
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

Private Sub cmdBuscar_Click()
Call Buscar
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub dtgEmpleadores_Click()
CargaImag
End Sub

Private Sub Form_Load()
Set rsEmpleador = New ADODB.Recordset
Set RsImag = New ADODB.Recordset
CenterForm Me
End Sub


Private Sub MuestraImg()
    
    With RsImag  'coneccion a _Imagenes
'        .Find "[posicion]='" & xnumpag & "'"
'        If .EOF Then .MovePrevious
        ximagen = IIf(IsNull(!Imagen), "", !Imagen)
        xrutavol = .Fields("ruta")
        xrutavol = Replace(xrutavol, "\\", "\")
        xidvolumen = IIf(IsNull(!idvolumen), "", !idvolumen)
    End With
      
    With ImgEdit1
        .Image = xrutavol & "\" & rsEmpleador.Fields("Lote") & "\" & ximagen & ".TIF"
        .Enabled = True
        .FitTo 1
        .Display
        .ScrollPositionY = (Int(ImgEdit1.ImageScaleHeight / 4) * 1)
        .ScrollPositionX = 0
        .Refresh
    End With
    
End Sub

Private Sub txtEnvio_LostFocus()
txtEnvio.Text = Format(txtEnvio.Text, "000000#")
End Sub

Private Sub TxtFecha_GotFocus()
    Marcado TxtFecha
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
    If KeyAscii = 13 Then CmdCons.SetFocus
End Sub

Private Sub TxtFecha_LostFocus()
    ValidaFecha TxtFecha
End Sub

Private Sub TxtLote_GotFocus()
    Marcado TxtLote
End Sub

Private Sub TxtLote_LostFocus()
    TxtLote = Format(TxtLote, "0000000#")
End Sub

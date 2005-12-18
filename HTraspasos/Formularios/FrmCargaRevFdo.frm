VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmCargaRevFdo 
   Caption         =   "Carga Rev Fondo"
   ClientHeight    =   5310
   ClientLeft      =   4185
   ClientTop       =   4185
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8160
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   17
      Top             =   4320
      Width           =   5895
      Begin VB.CommandButton cmdIniciar 
         Caption         =   "&Iniciar"
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
         Left            =   4080
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Inicia Archivo Digitación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   3150
      End
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar"
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
      Left            =   6360
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archivo de Carga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   5895
      Begin VB.TextBox txtruta 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
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
      Left            =   6360
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame FraIngFecha 
      Caption         =   "Ingresa Mes y Año de Archivo a Revisar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   5895
      Begin VB.TextBox TxtAnno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtMes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   960
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
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
         Left            =   2160
         TabIndex        =   9
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         Left            =   240
         TabIndex        =   8
         Top             =   645
         Width           =   375
      End
   End
   Begin VB.Frame FraModi 
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   5895
      Begin VB.OptionButton OptMesAnno 
         Caption         =   "Graba Mes y Año en archivo Cargado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.OptionButton OptSol6 
         Caption         =   "Corrige Num Solicitud de menos de 7 posiciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   4815
      End
   End
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "&Ejecutar"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Contador"
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
      Left            =   6570
      TabIndex        =   15
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label lblContador 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6315
      TabIndex        =   14
      Top             =   2400
      Width           =   1380
   End
   Begin VB.Label Label3 
      Caption         =   "Borrando archivo anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   6360
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "FrmCargaRevFdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsFdoCarga As New Recordset      'FdoCarga

Private Sub Command1_Click()

End Sub

Private Sub cmdBuscar_Click()
   CommonDialog1.Filter = "Microsoft Excel *.xls | *.xls"
   CommonDialog1.FileName = ""
   CommonDialog1.ShowOpen
   
   
   txtruta = CommonDialog1.FileName
   
End Sub

Private Sub cmdCargar_Click()
ConectarExcel
End Sub

Private Sub cmdIniciar_Click()
Dim cmdEliDigitacion As New ADODB.Command

If MsgBox("¿Esta seguro que desea eliminar registros de DIGITACION ?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
    With cmdEliDigitacion
        .ActiveConnection = Con.Cnx
        .CommandText = "truncate table RevFdoDig"
        .Execute
    End With
    MsgBox "Registros Eliminados", 64, Caption
End If

End Sub

Private Sub Form_Load()
    Call Con.AbrirTabla("select * from RevFdoCarga where mes is not null and anno is not null", RsFdoCarga, adOpenStatic, adLockReadOnly)
'    Call Con.AbrirTabla("select * from RevFdoCarga", RsFdoCarga, adOpenStatic, adLockReadOnly)
    With RsFdoCarga
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            TxtMes = ""
            TxtAnno = ""
        Else
            .MoveFirst
            TxtMes = !Mes
            TxtAnno = !Anno
        End If
    End With
End Sub

Private Sub CmdEjecutar_Click()
    If OptMesAnno.Value Then
    
        If Val(TxtMes) = 0 Or Val(TxtMes) > 12 Then
            MsgBox "Debe ingresar Mes ", vbCritical, "Error en Ingreso de Mes"
            TxtMes.SetFocus
            Exit Sub
        End If
    
        If Val(TxtAnno) < 2006 Then
            MsgBox "Debe ingresar Mes ", vbCritical, "Error en Ingreso de Año"
            TxtAnno.SetFocus
            Exit Sub
        End If
    
        TxtMes = Format(TxtMes, "0#")
        TxtAnno = Format(TxtAnno, "000#")
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update RevFdoCarga " & _
            "set Mes = '" & TxtMes & "', " & _
            " Anno = '" & TxtAnno & "' " & _
            " Where Mes Is Null And Anno Is Null "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Graba Mes y Año"
    End If
'***
    If OptSol6.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "UPDATE RevFdoCarga " & _
            "set N_Solici = isnull(replicate ('0', 7 - len(rtrim(N_Solici))),'') + rtrim(N_Solici) " & _
            "WHERE len(N_Solici)<7 and N_Solici<>'' "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Modifica Num Solicitud de menos de 7 posiciones "
    End If
End Sub

Private Sub TxtMes_GotFocus()
    Marcado TxtMes
End Sub

Private Sub TxtMes_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtAnno.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtAnno_GotFocus()
    Marcado TxtAnno
End Sub

Private Sub TxtAnno_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdEjecutar.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Sub ConectarExcel()

'On Local Error Resume Next

Dim cmdDelete As New Command
Dim cnExcel As New Connection
Dim RsExcel As New Recordset
Dim rsLEXCEL As New Recordset
Dim strinsercion As String
Dim NombHoja As String
Dim comandoXLS As New ADODB.Command
Dim contador As Double
'Eliminacion de RevFdoCarga
With cmdDelete
    Label3.Visible = True
    Label3.Refresh
    
    .ActiveConnection = Con.Cnx
    .CommandText = "truncate table RevFdoCarga"
    
    .Execute
End With

Label3.Visible = False
Label3.Refresh

contador = 0

Set cnExcel = New ADODB.Connection
cnExcel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CommonDialog1.FileName & ";Persist Security Info=False; Extended Properties=Excel 8.0;"
cnExcel.Open

If RsExcel.State = 1 Then RsExcel.Close
RsExcel.CursorLocation = adUseClient

RsExcel.LockType = adLockReadOnly
RsExcel.CursorType = adOpenStatic



Set RsExcel = cnExcel.OpenSchema(4)

If Err.Number > 0 Then
   MsgBox Err.Description, vbCritical, Caption
   Exit Sub
End If
           
NombHoja = Trim(RsExcel.Fields("table_name"))

campos = " estn21,estc41,estno4, estaf7,estfdo,estfd3 "

If cnExcel.State = 1 Then cnExcel.Close
cnExcel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CommonDialog1.FileName & ";Persist Security Info=False; Extended Properties=Excel 8.0;"
cnExcel.Open

rsLEXCEL.CursorLocation = adUseClient

rsLEXCEL.Open " select " & campos & " from [" & Trim(NombHoja) & "]", cnExcel, adOpenStatic, adLockReadOnly

If rsLEXCEL.RecordCount > 0 Then
    Do While Not rsLEXCEL.EOF
    DoEvents
    
    
    
      If fgDatosql(" Select n_solici from RevFdoCarga where n_solici = '" & IIf(IsNull(rsLEXCEL.Fields("estn21").Value), "", rsLEXCEL.Fields("estn21").Value) & "'") = "" Then
    
    
        strinsercion = " Insert Into RevFdoCarga ( n_solici, cuspp, ape_pat, afp, ori, des ) values ( " & _
        " '" & IIf(IsNull(rsLEXCEL.Fields("estn21").Value), "", rsLEXCEL.Fields("estn21").Value) & "' ," & _
        " '" & IIf(IsNull(rsLEXCEL.Fields("estc41").Value), "", rsLEXCEL.Fields("estc41").Value) & "' ," & _
        " '" & Replace(IIf(IsNull(rsLEXCEL.Fields("estno4").Value), "", rsLEXCEL.Fields("estno4").Value), "'", "''") & "' ," & _
        " '" & IIf(IsNull(rsLEXCEL.Fields("estaf7").Value), "", rsLEXCEL.Fields("estaf7").Value) & "' ," & _
        " '" & IIf(IsNull(rsLEXCEL.Fields("estfdo").Value), "", rsLEXCEL.Fields("estfdo").Value) & "' ," & _
        " '" & IIf(IsNull(rsLEXCEL.Fields("estfd3").Value), "", rsLEXCEL.Fields("estfd3").Value) & "' )"
                  
        '------------------Insercion--------------
        With comandoXLS
            .ActiveConnection = Con.Cnx
            '.CommandText = "SET QUOTED_IDENTIFIER OFF " & strinsercion
            .CommandText = strinsercion
            'Debug.Print strInsercion
            .Execute
        End With
            
        contador = contador + 1
        
        lblContador.Caption = contador
        lblContador.Refresh
      End If
        rsLEXCEL.MoveNext
    Loop
    
    
    MsgBox "Carga Completada", 64, Caption
Else
    MsgBox "Archivo no contiene registros", 64, Caption
    Exit Sub
End If


Set cmdDelete = Nothing
Set cnExcel = Nothing
Set RsExcel = Nothing
Set rsLEXCEL = Nothing
Set comandoXLS = Nothing



End Sub

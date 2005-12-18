VERSION 5.00
Begin VB.Form FrmModSol 
   Caption         =   "Modifica Datos Claves"
   ClientHeight    =   4995
   ClientLeft      =   2850
   ClientTop       =   3330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7185
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
      Left            =   5400
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame FraModVol 
      Caption         =   "Modifica Datos Lotes en Volumenes"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   6855
      Begin VB.CommandButton CmdGraba 
         Caption         =   "&Graba Modificacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox TxtLoteNew 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   5
         Top             =   1590
         Width           =   1095
      End
      Begin VB.CommandButton CmdBusca 
         Caption         =   "&Datos Actuales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label LblLoteFis 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2040
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LblLoteAct 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Lote"
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
         Left            =   1320
         TabIndex        =   15
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lote Fisico"
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
         Left            =   2040
         TabIndex        =   14
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lote Actual"
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
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.Frame FraModSol 
      Caption         =   "Modifica el Numero de Solicitud"
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
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   6855
      Begin VB.CommandButton CmdModSol 
         Caption         =   "&Modifica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtNumSol2 
         Alignment       =   2  'Center
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
         Left            =   3480
         TabIndex        =   2
         Top             =   750
         Width           =   1335
      End
      Begin VB.TextBox TxtLote1 
         Alignment       =   2  'Center
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
         Left            =   360
         TabIndex        =   0
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtNumSol1 
         Alignment       =   2  'Center
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
         Left            =   1800
         TabIndex        =   1
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Num Sol Nueva"
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
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
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
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Num Sol Actual"
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
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1470
      End
   End
End
Attribute VB_Name = "FrmModSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsVol As New Recordset
Dim xidimagen1 As String
Dim xidimagen2 As String

Private Sub Form_Load()
    CmdModSol.Enabled = False
    CmdGraba.Enabled = False
End Sub

Private Sub CmdModSol_Click()
    xidimagen1 = "01" + TxtNumSol1
    xidimagen2 = "01" + TxtNumSol2
    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = "update empleadores " & _
        "Set empleadores.lote = solicitud.lote " & _
        "FROM empleadores, solicitud WHERE empleadores.numsol=solicitud.numsol and empleadores.secuencial=solicitud.secu and empleadores.lote = null "
        .Execute
        .CommandText = "update empleadores " & _
        "Set NumSol = '" & TxtNumSol2 & "' " & _
        "WHERE NumSol = '" & TxtNumSol1 & "' and Lote = '" & TxtLote1 & "' "
        .Execute
        .CommandText = "update solicitud " & _
        "Set NumSol = '" & TxtNumSol2 & "', " & _
        "idimagen = '" & xidimagen2 & "' " & _
        "WHERE NumSol = '" & TxtNumSol1 & "' and Lote = '" & TxtLote1 & "' "
        .Execute
        .CommandText = "update _imagenes " & _
        "Set idimagen = '" & xidimagen2 & "' " & _
        "WHERE IDImagen = '" & xidimagen1 & "' and Lote = '" & TxtLote1 & "' "
        .Execute
        MsgBox " Solicitud Modificada ", vbInformation, "Modificacion de Numero de Solicitud"
    End With
End Sub

Private Sub CmdBusca_Click()
    LblLoteAct = " "
    LblLoteFis = " "
    TxtLoteNew = " "
    
    If RsVol.State = 1 Then RsVol.Close
    SQLstr = "SELECT * FROM _Volumenes "
    RsVol.Open SQLstr, Con.Cnx, adOpenStatic, adLockBatchOptimistic
    With RsVol
        LblLoteAct = IIf(IsNull(!Lote_Actual), "", !Lote_Actual)
        LblLoteFis = IIf(IsNull(!Lote_Fisico), "", !Lote_Fisico)
    End With
End Sub

Private Sub CmdGraba_Click()
    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = "update _volumenes " & _
        "Set lote_actual = " & TxtLoteNew & ", " & _
        "lote_fisico = " & TxtLoteNew & " "
        .Execute
        MsgBox " Numero de Lote Modificado ", vbInformation, "Modificacion de Numero de Lote"
    End With

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub TxtLote1_GotFocus()
    Marcado TxtLote1
End Sub

Private Sub TxtLote1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtNumSol1.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            TxtNumSol2.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol1_LostFocus()
    TxtNumSol1 = Format(TxtNumSol1, "000000#")
End Sub

Private Sub TxtNumSol2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdModSol.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtNumSol2_LostFocus()
    TxtNumSol2 = Format(TxtNumSol2, "000000#")
End Sub

Private Sub TxtNumSol2_Change()
    CmdModSol.Enabled = Val(TxtLote1) > 0 And Val(TxtNumSol1) > 0 And Val(TxtNumSol2) > 0
End Sub

Private Sub TxtLoteNew_GotFocus()
    Marcado TxtLoteNew
End Sub

Private Sub TxtLoteNew_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            CmdGraba.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLoteNew_Change()
    CmdGraba.Enabled = Val(TxtLoteNew) > 0
End Sub


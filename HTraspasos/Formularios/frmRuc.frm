VERSION 5.00
Begin VB.Form frmRuc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo RUC"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11115
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
   ScaleHeight     =   7185
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin Traspasos.UsCBusqueda2 UsCBusqueda21 
      Height          =   375
      Left            =   360
      TabIndex        =   39
      Top             =   3120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
   End
   Begin VB.Frame Frame3 
      Caption         =   "Otros"
      ForeColor       =   &H00C00000&
      Height          =   2775
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   10695
      Begin VB.TextBox TxtAPatRL 
         Height          =   375
         Left            =   240
         MaxLength       =   40
         TabIndex        =   7
         Top             =   720
         Width           =   6135
      End
      Begin VB.TextBox TxtAxoEP 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TxtTelEP 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         MaxLength       =   9
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TxtNom1EP 
         Height          =   375
         Left            =   6120
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtAMatEP 
         Height          =   375
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox TxtNDocRL 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtTdocRL 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   6960
         MaxLength       =   2
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxtAPatEP 
         Height          =   375
         Left            =   240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ap Materno"
         Height          =   225
         Left            =   4080
         TabIndex        =   36
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   225
         Left            =   7080
         TabIndex        =   35
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Documento de Identidad"
         Height          =   225
         Left            =   6960
         TabIndex        =   34
         Top             =   480
         Width           =   2025
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   225
         Left            =   6480
         TabIndex        =   33
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         Height          =   225
         Left            =   7680
         TabIndex        =   32
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Rep Legal / Apellidos y Nombres"
         Height          =   225
         Left            =   360
         TabIndex        =   31
         Top             =   480
         Width           =   2700
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Elab por / Ap Paterno"
         Height          =   225
         Left            =   360
         TabIndex        =   30
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
         Height          =   225
         Left            =   360
         TabIndex        =   29
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Anexo"
         Height          =   225
         Left            =   1920
         TabIndex        =   28
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8280
      TabIndex        =   23
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Empleador"
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   10695
      Begin VB.TextBox txtRazSoc 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtRUC 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Razón Social"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Ruc"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ubicación"
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   10695
      Begin VB.TextBox TxtUbigeo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TxtTelef 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtUrba 
         Height          =   330
         Left            =   7320
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNro 
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtDir 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   7560
         TabIndex        =   38
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Ubigeo"
         Height          =   255
         Left            =   6120
         TabIndex        =   37
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Urbanización"
         Height          =   255
         Left            =   7320
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Distrito /  Provincia / Departamento"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label Label4 
         Caption         =   "Nro"
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Ap Materno"
      Height          =   225
      Left            =   3240
      TabIndex        =   25
      Top             =   5400
      Width           =   915
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   225
      Left            =   6000
      TabIndex        =   24
      Top             =   5400
      Width           =   675
   End
End
Attribute VB_Name = "frmRuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Form As String
Public Fecsal As String
Private horafin As String
Public horaini As String

Private Sub cmdAceptar_Click()
Grabar
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
LLenarUbigeo
If Form = "frmValruc" Then
    CargarDatosRuc
    horaini = FrmValRUC.horaini
End If

End Sub

Sub CargarDatosRuc()
With FrmValRUC.RsCmaepla
    Me.txtRUC = !numruc
    Me.txtRazSoc = !razsoc
    TxtDir.Text = IIf(IsNull(!Direcc), "", !Direcc)
    txtNro = IIf(IsNull(!Numero), "", !Numero)
    txtUrba.Text = IIf(IsNull(!URBANI), "", !URBANI)
    TxtUbigeo.Text = IIf(IsNull(!codgeo), "", !codgeo)
    TxtAPatRL.Text = IIf(IsNull(!RLAPAT), "", !RLAPAT)
        
    TxtTdocRL.Text = IIf(IsNull(!DOCTIP), "", !DOCTIP)
    TxtNDocRL.Text = IIf(IsNull(!DOCNUM), "", !DOCNUM)
        
    TxtTelef.Text = IIf(IsNull(!NROTEL), "", !NROTEL)
        
    TxtAPatEP.Text = IIf(IsNull(!ELAPAT), "", !ELAPAT)
    TxtAMatEP.Text = IIf(IsNull(!ELAMAT), "", !ELAMAT)
    TxtNom1EP.Text = IIf(IsNull(!ELNOMB1), "", !ELNOMB1)
        
    TxtTelEP.Text = Trim(IIf(IsNull(!ELABTEL), "", !ELABTEL))
    TxtAxoEP.Text = Trim(IIf(IsNull(!ELABANX), "", !ELABANX))
            
    If Trim(TxtUbigeo.Text) <> "" Then
            UsCBusqueda21.cTexto = fgDatosql("Select desgeo from TUBIGEO where CODGEO = '" & Trim(TxtUbigeo.Text) & "'") & " " & Space(150) & IIf(IsNull(!codgeo), "", !codgeo)
            'frmIngresoCabecera.EDesUbigeo = UsCBusqueda21.cTexto & Space(150) & Trim(TxtUbigeo.Text)
    End If
    

    
End With


End Sub

Sub LLenarUbigeo()
On Error Resume Next
    With UsCBusqueda21
        .Enabledf = True
        
        Set UsCBusqueda21.cConexion = Con.Cnx
        .Consulta = "Select DESGEO + ' ' + space(150) + ' '+ CODGEO from TUBIGEO "
        .Campocondicion = "DESGEO"
        .Busqinicial = False
        .Enabledf = True
        .nroCaracteres = 3
    End With
End Sub

Function validar() As Boolean
validar = True

If Trim(txtRUC) = "" Then
    MsgBox "Debe ingresar el nro de Ruc", 64, Caption
    validar = False
    txtRUC.SetFocus
    Exit Function
End If

If Not Modulo11(Trim(txtRUC)) Then
    MsgBox "El Nro de Ruc no es válido", 64, Caption
    validar = False
    txtRUC.SetFocus
    Exit Function
End If

'If fgDatosql("Select numruc from truc where numruc = '" & Trim(txtRUC) & "' ") <> "" Then
'    MsgBox "El Ruc existe no puede ingresar Ruc duplicados", 64, Caption
'    validar = False
'    txtRUC.SetFocus
'    Exit Function
'End If


If Trim(txtRazSoc) = "" Then
    MsgBox "Debe ingresar la razón social", 64, Caption
    validar = False
    txtRazSoc.SetFocus
    Exit Function
End If


End Function


Private Sub TxtAMatEP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtAPatEP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtAPatRL_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtAxoEP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtDir_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtNDocRL_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtNom1EP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtNro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtRUC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtTdocRL_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtTelef_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub TxtTelEP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub txtUrba_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
TabSig KeyAscii
End Sub

Private Sub UsCBusqueda21_Pulsacion(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

On Error GoTo errores
    If KeyAscii = 13 Then
        TxtUbigeo.Text = Trim(Right(Trim(UsCBusqueda21.cTexto), 6))
        If TxtUbigeo.Text = "" Then
            If MsgBox("No seleccionó ninguno localización, ¿ Esta seguro ?", vbQuestion + vbYesNo, Caption) = vbNo Then
                UsCBusqueda21.SetFocus
            Else
                TxtTelef.SetFocus
            End If
        Else
            TxtTelef.SetFocus
        End If
        
    
    End If
    Exit Sub
errores:
 MsgBox Err.Description, 64, Caption
End Sub
Sub Grabar()
Dim CmdRuc As ADODB.Command
Set CmdRuc = New ADODB.Command

Dim CmdCmae As ADODB.Command
Set CmdCmae = New ADODB.Command

If Not validar Then Exit Sub
With CmdRuc
    .ActiveConnection = Con.Cnx
    .CommandText = "exec pGrabarTRuc ?,?,?,?,?, ?,?,?,?,?, ?,?,?,?,?, ?,?"
    .Parameters("@CODAFP") = vbGAFP
    .Parameters("@NUMRUC") = Trim(txtRUC)
    .Parameters("@RAZSOC") = Trim(txtRazSoc)
    .Parameters("@DIRECC") = Trim(TxtDir)
    .Parameters("@CODGEO") = Trim(TxtUbigeo)
    .Parameters("@TELEFONO") = Trim(TxtTelef)
    .Parameters("@NUMERO") = Trim(txtNro)
    .Parameters("@URBA") = Trim(txtUrba)
    .Parameters("@APEPAT_RL") = Trim(TxtAPatRL)
    .Parameters("@TIP_DOC_RL") = Trim(TxtTdocRL)
    .Parameters("@NUM_DOC_RL") = Trim(TxtNDocRL)
    .Parameters("@APEPAT_EL") = Trim(TxtAPatEP)
    .Parameters("@APEMAT_EL") = Trim(TxtAMatEP)
    .Parameters("@NOMBRE1_EL") = Trim(TxtNom1EP)
    .Parameters("@NOMBRE2_EL") = ""
    .Parameters("@TELEF_EL") = Trim(TxtTelEP)
    .Parameters("@ANEXO_EL") = Trim(TxtAxoEP)
    .Parameters("@FECSAL") = Fecsal
    .Execute
End With


         
    
    Unload Me
End With

End Sub

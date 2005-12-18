VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Icr_Parametros 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Parametros"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   Icon            =   "Icr_Parametros.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.DirListBox dirEnvios 
         Height          =   1665
         Left            =   480
         TabIndex        =   2
         Top             =   2580
         Width           =   4575
      End
      Begin VB.DriveListBox drvEnvios 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   2220
         Width           =   4575
      End
      Begin MSMask.MaskEdBox mskZona 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCampaña 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCantidad 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###"
         Mask            =   "99"
         PromptChar      =   "_"
      End
      Begin VB.Label lblZona 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
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
         Left            =   1785
         TabIndex        =   11
         Top             =   900
         Width           =   465
      End
      Begin VB.Label lblCampaña 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         Left            =   1785
         TabIndex        =   10
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label lblDirectorio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directorio de envio de archivos"
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
         Left            =   480
         TabIndex        =   9
         Top             =   1980
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Validar cantidad mayor a"
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
         Left            =   360
         TabIndex        =   8
         Top             =   300
         Width           =   2115
      End
   End
End
Attribute VB_Name = "Icr_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSQL As New Recordset
Private Sub cmdAceptar_Click()
On Error GoTo Err:
    rsSQL!ds_periodo = mskCampaña.Text
    rsSQL!ds_año = mskZona.Text
    rsSQL!ds_ftpaxis = dirEnvios.Path
    rsSQL!nu_cant_minimo = Val(mskCantidad.Text)
    
    rsSQL.UpdateBatch adAffectCurrent
    Unload Me
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Parametros"
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub drvEnvios_Change()
On Local Error Resume Next
    Me.dirEnvios.Path = drvEnvios.Drive
End Sub

Private Sub Form_Load()
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    CargarDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsSQL Is Nothing) Then
        If rsSQL.State = 1 Then rsSQL.Close
        Set rsSQL = Nothing
    End If
End Sub

Sub CargarDatos()
On Error Resume Next
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "Select * from TB_Parametros"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        mskCampaña.Text = "" & rsSQL!ds_periodo
        mskZona.Text = "" & rsSQL!ds_año
        mskCantidad.Text = rsSQL!nu_cant_minimo
        dirEnvios.Path = "" & rsSQL!ds_ftpaxis
        drvEnvios.Drive = dirEnvios.Path
    End If
End Sub

Private Sub mskCampaña_GotFocus()
    mskCampaña.SelStart = 0
    mskCampaña.SelLength = Len(mskCampaña.Text)
End Sub
Private Sub mskCantidad_GotFocus()
    mskCantidad.SelStart = 0
    mskCantidad.SelLength = Len(mskZona.Text)
End Sub
Private Sub mskZona_GotFocus()
    mskZona.SelStart = 0
    mskZona.SelLength = Len(mskZona.Text)
End Sub


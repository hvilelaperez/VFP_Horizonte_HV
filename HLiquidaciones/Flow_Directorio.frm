VERSION 5.00
Begin VB.Form Flow_Directorio 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Seleccionar Directorio"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form2"
   ScaleHeight     =   2775
   ScaleWidth      =   8295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.DirListBox Dirxls 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5655
   End
   Begin VB.DriveListBox Drvxls 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
   Begin VB.FileListBox flxls 
      Height          =   1455
      Left            =   5760
      Pattern         =   "*.tif"
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Flow_Directorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If flxls.ListCount > 0 Then
strRutaDP = Mid(Dirxls, 1, InStrRev(Dirxls, "\") - 1)
strDirDP = Mid(Dirxls, InStrRev(Dirxls, "\") + 1, Len(Dirxls))
Unload Me
Else
MsgBox "No existen archivos tiff"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dirxls_Change()
flxls.Path = Dirxls.Path
End Sub

Private Sub Dirxls_Click()
flxls.Path = Dirxls.Path
End Sub

Private Sub Drvxls_Change()
Dirxls.Path = Drvxls.Drive
End Sub

Private Sub Form_Load()
Dirxls.Path = Flow_FrmasignarLote.Dir1
Drvxls.Drive = Dirxls.Path
End Sub

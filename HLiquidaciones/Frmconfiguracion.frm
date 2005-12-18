VERSION 5.00
Begin VB.Form Flow_Frmconfiguracion 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Configuración"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   Icon            =   "Frmconfiguracion.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ruta de Destino De archivos Generados"
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6375
         Begin VB.CommandButton cmguardar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Guardar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Cmdeditar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Editar"
            Height          =   495
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.DirListBox Dir2 
            Enabled         =   0   'False
            Height          =   990
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   3615
         End
         Begin VB.DriveListBox Drive2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3495
         End
      End
   End
End
Attribute VB_Name = "Flow_Frmconfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdeditar_Click()
Drive2.Enabled = True
Dir2.Enabled = True
cmguardar.Enabled = True

End Sub

Private Sub cmguardar_Click()
On Error GoTo Aplica

GuardarIni ContRutatxt, "Destinoruta", "Rutad", Dir2
MsgBox "Se guardo con exito la ruta", vbInformation, "Horizonte"

cmguardar.Enabled = False

Drive2.Enabled = False
Dir2.Enabled = False
Aplica:
End Sub

Private Sub Drive2_Change()
Dir2.Path = Drive2.Drive
End Sub

Private Sub Form_Load()
Dim Rutad As String
'''Variable de Ruta de Directorio donde se almacena lso archivos Textos
ContRutatxt = App.Path & "\Horizonte.ini"
'ContRutatxt
Rutad = LeerIni(ContRutatxt, "Destinoruta", "Rutad", "")
If Dir(Rutad, vbDirectory) = "" Then
Drive2 = App.Path

Dir2 = App.Path
Else
Drive2 = Rutad
Dir2 = Rutad
End If

End Sub

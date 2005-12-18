VERSION 5.00
Begin VB.Form frmFraAyuda 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdCopiar 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Copiar Dato"
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblAyuda 
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   2
      Top             =   280
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmFraAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
SiempreVisible Me, False
Unload Me
End Sub

Private Sub cmdCopiar_Click()
On Error Resume Next
With frmVerificacionD2

If TypeOf .ControlOrig Is TextBox Then
    .ControlOrig.Text = lblAyuda.Caption
End If

If TypeOf .ControlOrig Is UserControl1 Then
    .ControlOrig.cTexto = lblAyuda.Caption
End If

End With
End Sub

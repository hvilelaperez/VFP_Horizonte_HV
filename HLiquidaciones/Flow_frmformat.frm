VERSION 5.00
Begin VB.Form Flow_Frmformat 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Periodo / Año"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   Icon            =   "Flow_frmformat.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtaño 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtperiodo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaxLength       =   6
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AÑO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Flow_Frmformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txtperiodo.Text = Format(Now, "MM")
txtaño.Text = Format(Now, "yyyy")
'txtperiodo.SetFocus
End Sub

Private Sub txtaño_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Len(txtperiodo.Text) <> 0 And Len(txtaño.Text) <> 0 Then
  strPeriodo = txtaño.Text & Format(txtperiodo.Text, "00")
  strAño = txtaño.Text
  Unload Me
  End If
End If
End Sub

Private Sub txtperiodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Len(txtperiodo.Text) > 0 And Len(txtaño.Text) > 0 Then
  strPeriodo = txtaño.Text & txtperiodo.Text
  strAño = txtaño.Text
  Unload Me
  End If
End If
End Sub

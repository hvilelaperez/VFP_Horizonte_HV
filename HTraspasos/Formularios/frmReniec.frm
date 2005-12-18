VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmReniec 
   Caption         =   "Reniec"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToolBar 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdToolBar 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdToolBar 
      Caption         =   "Detener"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdToolBar 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9735
      ExtentX         =   17171
      ExtentY         =   11456
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmReniec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdToolBar_Click(Index As Integer)
On Error GoTo mensajeError

Select Case Index
 
  Case 0
    WebBrowser1.GoBack 'ir a la página anterior
  Case 1
    WebBrowser1.GoForward 'Siguiente página
  Case 2
    WebBrowser1.Stop 'Detener la carga de la página
  Case 3
    WebBrowser1.Refresh ' Actualizar la página
  Case 4
    WebBrowser1.GoHome 'Ir a la página de inicio
End Select


Exit Sub

mensajeError:
If Err.Number = -2147467259 Then
   Resume Next
MsgBox Err.Description
End If
End Sub

Private Sub Form_Load()
WebBrowser1.AddressBar = True
WebBrowser1.Navigate ("https://cel.reniec.gob.pe/valreg/valreg.do?accion=ini")
End Sub

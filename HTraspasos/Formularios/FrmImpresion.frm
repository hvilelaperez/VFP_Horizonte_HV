VERSION 5.00
Begin VB.Form frmImpresion 
   Caption         =   "Seleccionar Imágenes para Impresión"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   5040
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Ubicación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton cmdUbicacion 
         Caption         =   "..."
         Height          =   375
         Left            =   5280
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtUbicacion 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcesar_Click()
Call Procesar(txtUbicacion.Text)
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub cmdUbicacion_Click()
txtUbicacion.Text = BrowseFolder("Selecione una Carpeta")

End Sub

Sub Procesar(CarpetaSel As String)
On Error GoTo errores
Dim RsImag As Recordset
Set RsImag = ObtenerImagenes()
Dim RutaDef As String

Dim ContCarpeta As Integer
Dim ContImagen As Integer

ContCarpeta = 1

With RsImag
If .RecordCount > 0 Then
    
    ContImagen = 1
    
    
    Open App.Path & "\expo\errimp.log" For Output As #1
            'Exporta Archivos resultados
    
    Do While Not .EOF
        
        RutaDef = CarpetaSel & "\C" & Format(ContCarpeta, "00#")
        
        If Not ExisteDirectorio(RutaDef) Then
            CrearDirectorio (RutaDef)
        End If
                     
        If ContImagen <= 200 Then
            If ExisteArchivo(.Fields("RutaImagen")) Then
                Call CopiarFile(.Fields("RutaImagen"), RutaDef & "\" & .Fields("Secu") & ".tif")
                ContImagen = ContImagen + 1
            Else
                Print #1, " imagen " & .Fields("RutaImagen") & " : Solicitud " & .Fields("Numsol")
            End If
        Else
            ContCarpeta = ContCarpeta + 1
            ContImagen = 1
        End If
        
        .MoveNext
        DoEvents
    Loop
    
    Close #1
    
    MsgBox "Proceso Terminado", vbInformation, Caption
    
End If
End With
Exit Sub
errores:
Close #1
MsgBox Err.Description, vbCritical, Caption
End Sub

Function ObtenerImagenes() As Recordset

Dim RsImag As Recordset
Set RsImag = New Recordset
    
RsImag.Open "pListImagenImp", Con.Cnx, adOpenStatic, adLockReadOnly

Set ObtenerImagenes = RsImag
End Function

Sub ContCarpeta(Ruta As String)
    
End Sub

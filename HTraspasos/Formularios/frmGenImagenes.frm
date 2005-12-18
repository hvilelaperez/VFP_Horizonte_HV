VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGenImagenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Imágenes"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3525
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton cmdFolder 
         Caption         =   "..."
         Height          =   375
         Left            =   5280
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtnroimag 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtnroLote 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Ruta"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmGenImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cExtensionImagen
Private pidVolumen As Integer

Private Sub Form_Load()
cExtensionImagen = "Tif"
pidVolumen = 1
End Sub

Private Sub cmdAceptar_Click()
    Generar
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdFolder_Click()
Dim Carpeta As String
Carpeta = BrowseFolder(" Seleccionar Carpeta  ")
If Carpeta <> "" Then
    txtCarpeta.Text = Carpeta
    txtnrolote.SetFocus
End If
End Sub
Sub Generar()
On Error GoTo errores
Dim Ruta As String
Ruta = Trim(txtCarpeta.Text) & "\" & Format(Trim(txtnrolote.Text), "0000000#")
Dim vbEliImag As Boolean 'Si se eliminaran imagenes existentes

StatusBar1.Panels(1).Text = "Procesando ... "

StatusBar1.Panels(1).Text = "Buscando imágenes ... "

If fgDatosql("Select Lote from _Imagenes where idvolumen = " & pidVolumen & " and lote = " & Val(txtnrolote.Text) & " ") <> "" Then

    Rpta = MsgBox("Existen imágenes del lote " & Val(txtnrolote.Text) & " ¿Desea Eliminarlas ? ", vbQuestion + vbYesNoCancel, Caption)
    
    Select Case Rpta
    Case vbYes
        vbEliImag = True
    Case vbNo
        vbEliImag = False
    Case Else
        Exit Sub
    End Select
                   
End If

Screen.MousePointer = vbHourglass

If vbEliImag Then
    'Eliminar Imagenes y ActualizarContador
    StatusBar1.Panels(1).Text = "Eliminando imágenes ... "
    Call EliminarImagenes
    If ExisteDirectorio(Ruta) Then Kill "" & Ruta & "\*.tif"
End If



If CrearDirectorio(Ruta) Then
    
     StatusBar1.Panels(1).Text = "Actualizando  PreScan, Scan ... "
     
    InsertarPreScan pidVolumen, Val(txtnrolote.Text)
    
    StatusBar1.Panels(1).Text = "Creando imágenes ... "

    For i = 1 To Val(Me.txtnroimag.Text)
        Call CrearImagen(Ruta, vbEliImag)
        DoEvents
    Next
    
    StatusBar1.Panels(1).Text = "Proceso completado"
    MsgBox "Proceso Completado", 64, Caption
    
    StatusBar1.Panels(1).Text = " "
    Screen.MousePointer = vbDefault
    
End If
Exit Sub
errores:
Screen.MousePointer = vbDefault
MsgBox Err.Description, vbCritical, Caption
StatusBar1.Panels(1).Text = "Proceso tuvo errores "


End Sub
Function CrearDirectorio(strRuta As String) As Boolean
On Local Error GoTo Err:
Dim bolOk As Boolean

If ExisteDirectorio(strRuta) Then CrearDirectorio = True: Exit Function

    bolOk = True
    MkDir strRuta
    CrearDirectorio = bolOk
    Exit Function
Err:
    MsgBox Err.Description, vbCritical, "Crear directorio volumen"
    bolOk = False
    CrearDirectorio = bolOk
End Function
Function CrearImagen(Ruta As String, vbEliImagen) As Boolean
On Error GoTo errores
  Dim strImagenScan As String
  Dim strNuevaImagen As String
  Dim lngImagenActual As Long
  

  CrearImagen = False
 
  lngImagenActual = Val(fgDatosql("Select contador from _Scan where idvolumen = " & pidVolumen & " and lote = " & Val(txtnrolote) & "  "))
  lngImagenActual = lngImagenActual + 1
        
  strPrefijoRuta = Format(Trim(txtnrolote.Text), "0000") & "_"
  
  strNuevaImagen = strPrefijoRuta & Format(lngImagenActual, "0000000#")
  strImagenScan = Ruta & "\" & strNuevaImagen & "." & cExtensionImagen
  
  If CopiarFile(App.Path & "\Imagen\DemoIma.Tif", strImagenScan) Then
    InsertarImagen strNuevaImagen
    ActualizaContador lngImagenActual
  End If
  
  CrearImagen = True
  Exit Function
errores:
  CrearImagen = False
  MsgBox Err.Description, vbCritical, Caption
 
End Function
Function CopiarFile(rutaOrigen As String, rutaDestino As String) As Boolean
On Error GoTo errores
CopiarFile = True
FileCopy rutaOrigen, rutaDestino
  
Exit Function
errores:
CopiarFile = False
MsgBox Err.Description & " Copiar File"
End Function


Sub ActualizaContador(contador As Long)

On Local Error GoTo errores

Dim comm As New ADODB.Command

'Con.Cnx.BeginTrans

With comm
    .ActiveConnection = Con.Cnx
    .CommandText = "Update _Scan set contador = " & contador & " where idVolumen = " & pidVolumen & " and Lote= " & Val(txtnrolote) & ""
    .Execute

End With

'Con.Cnx.CommitTrans
Set comm = Nothing
Exit Sub
errores:
'Con.Cnx.RollbackTrans
Set comm = Nothing
MsgBox Err.Description
End Sub

Sub EliminarImagenes()
Dim cmdImagenes As ADODB.Command
Set cmdImagenes = New ADODB.Command

With cmdImagenes
    .ActiveConnection = Con.Cnx
    .CommandText = "Delete from _imagenes where idvolumen = " & pidVolumen & " and Lote = " & Val(txtnrolote) & " "
    .Execute
    
    
End With

End Sub

Sub InsertarPreScan(Volumen As Integer, Lote As Integer)
Dim cmdPreScan As ADODB.Command
Set cmdPreScan = New ADODB.Command

With cmdPreScan
    .ActiveConnection = Con.Cnx
  '  .CommandType = adCmdStoredProc
    .CommandText = "exec pInsPScanVisor ?,?,?,?,?,?"
    .Parameters("@idVolumen") = Volumen
    .Parameters("@Lote") = Lote
    .Parameters("@idPreparador") = 0
    .Parameters("@FechaInicio") = gDate
    .Parameters("@FechaFin") = gDate
    .Parameters("@LoteFisico") = Lote
    .Execute
End With

End Sub
Sub InsertarImagen(Imagen As String)
Dim cmdImagen As ADODB.Command
Set cmdImagen = New ADODB.Command

With cmdImagen
    .ActiveConnection = Con.Cnx
  '  .CommandType = adCmdStoredProc
    .CommandText = "exec sp_RegistrarImagen ?,?,?,?,?,?"
    .Parameters("@idVolumen") = pidVolumen
    .Parameters("@Lote") = Val(txtnrolote)
    .Parameters("@LoteFisico") = Val(txtnrolote)
    .Parameters("@Imagen") = Imagen
    .Parameters("@Tipo") = "B"
    .Parameters("@Replicar") = 0
    .Execute
    
End With

End Sub

Function ValidarLote() As Boolean
ValidarLote = True

If Not ExisteDirectorio(txtCarpeta.Text) Then
    MsgBox "La ruta ingresada no es válida", vbCritical, Caption
    ValidarLote = False
    txtCarpeta.SetFocus
    Exit Function
End If

If txtnrolote = "" Then
    MsgBox "Debe ingresar el número de Lote", vbCritical, Caption
    ValidarLote = False
    txtnrolote.SetFocus
    Exit Function
End If


If txtnroimag = "" Then
    MsgBox "Debe ingresar el número de imágenes", vbCritical, Caption
    ValidarLote = False
    txtnroimag.SetFocus
    Exit Function
End If

End Function

Private Sub txtnroimag_KeyPress(KeyAscii As Integer)
FEntDec txtnroimag, KeyAscii, 4, 0
End Sub

Private Sub txtnroLote_KeyPress(KeyAscii As Integer)
FEntDec txtnrolote, KeyAscii, 4, 0
End Sub

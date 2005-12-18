VERSION 5.00
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltocx12n.ocx"
Begin VB.Form frmImporImag 
   Caption         =   "Importar Imagenes"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
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
   ScaleHeight     =   6270
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHojaRuta 
      BackColor       =   &H80000009&
      Caption         =   "Seleccionar Hoja Control"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Visor"
      ForeColor       =   &H000040C0&
      Height          =   4695
      Left            =   6480
      TabIndex        =   6
      Top             =   600
      Width           =   3855
      Begin LEADLib.LEAD LEAD1 
         Height          =   4095
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3495
         _Version        =   65539
         _ExtentX        =   6165
         _ExtentY        =   7223
         _StockProps     =   229
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BitonalScaling  =   2
         ScaleHeight     =   273
         ScaleWidth      =   233
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
         PanWinTitle     =   "PanWindow"
         CLeadCtrl       =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ubicación de Archivos"
      ForeColor       =   &H000000C0&
      Height          =   4695
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   6135
      Begin VB.FileListBox File1 
         Height          =   3660
         Left            =   3240
         MultiSelect     =   2  'Extended
         Pattern         =   "*.tif"
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.DirListBox Dir1 
         Height          =   3690
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Nuevo Documento"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdImportar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transferir Imágenes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Transferir Imágenes"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lblLote 
      Caption         =   "No reconocido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Lote 
      Caption         =   "Lote :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmImporImag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RutaVol As String
'Public Lote As String
Public FormOrig As Form
Public NroImaIni As Integer ' Se pasara como parametro el ultimo numero de imagen del documento
Private ImagenBarcode As String
Public vbindLoteCreado As Boolean

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Function HojaRuta(ImagenBar As String) As Boolean

            HojaRuta = True


            Dim strLoteFisico As String
            strLoteFisico = ""

            strLoteFisico = frmScan.ObtenerLoteFisico(ImagenBar)
                
            If strLoteFisico = "" Then
                        strLoteCaptura = ""
                        bolOk = False
                        MsgBox "Debe generar un número de lote válido", vbCritical, Caption
                        HojaRuta = False
                        Exit Function
            Else
                        If frmScan.lblLote2.Caption <> "Lote nuevo" Then
                            If ExisteLote(strLoteFisico) = True Then
                                MsgBox "Ya existe un lote con el mismo numero de lote a crear", vbCritical, Caption
                                bolOk = False
                                HojaRuta = False
                                Exit Function
                            End If
                        End If
            End If
                
                
                lblLote.Caption = strLoteFisico
                
                frmScan.lblLote2.Caption = strLoteFisico
                strPrefijoRuta = Format(Trim(frmScan.lblLote2.Caption), "000#") & "_"
                frmScan.lngImagenActual = 1
                
                strNuevaImagen = frmScan.strPrefijoRuta & Format(IIf(frmScan.lngImagenActual = 0, 1, frmScan.lngImagenActual), "0000000#")
                
                strImagenScan = frmScan.strDirectorioImagenes & "\" & Format(frmScan.lblLote2.Caption, "0000000#") & "\" & strNuevaImagen & "." & cExtensionImagen
                
                If Dir(frmScan.strDirectorioImagenes & "\" & Format(frmScan.lblLote2.Caption, "0000000#"), vbDirectory) = "" Then
                        MkDir frmScan.strDirectorioImagenes & "\" & Format(frmScan.lblLote2.Caption, "0000000#")
                End If
                
                    
                lngLoteActual = fgDatosql("Select isnull(max(Lote)+1,1) from _Scan where idvolumen = " & lngIdVolumenActual & " ")
                    
                Dim strCampos As String
                
                strCampos = lngIdVolumenActual & "," & lngLoteActual & ",'" & Trim(Format(strLoteFisico, "0000000#")) & "'"
                clsDataScan.Conexion.Execute ("EXECUTE sp_InsertarPreScan " & strCampos)
                clsDataScan.Conexion.Execute ("EXECUTE sp_InsertarScan " & strCampos)
                clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTabla _Volumenes ,'set lote_actual=" & lngLoteActual & ", Lote_Fisico = ''" & Trim(Format(strLoteFisico, "0000000#")) & "''','idVolumen= " & lngIdVolumenActual & "'")
    
                frmScan.lblLote.Caption = " " & lngLoteActual
                
        
End Function

Private Sub cmdHojaRuta_Click()

ImagenBarcode = File1.Path & "\" & File1.FileName

If HojaRuta(ImagenBarcode) Then
    cmdHojaRuta.Enabled = False
    cmdImportar.Enabled = True
Else
    cmdHojaRuta.Enabled = True
    cmdImportar.Enabled = False
    MsgBox "Primero se debe reconocer la hoja de control", 64, Caption
End If


End Sub

Private Sub cmdImportar_Click()
If File1.ListCount = 0 Then MsgBox "No hay imágenes para importar", 64, Caption: Exit Sub
Importar
End Sub


Private Sub Dir1_Change()

File1.Path = Dir1.Path

'Call LlenarLista

LEAD1.Bitmap = 0
End Sub

Private Sub Drive1_Change()
On Error GoTo ERRORES
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path


Exit Sub
ERRORES:
MsgBox Err.Description, vbCritical, Caption
End Sub

Private Sub File1_Click()
strImagenScan = File1.Path & "\" & File1.FileName

LEAD1.Bitmap = 0
LEAD1.GetFileInfo strImagenScan, 0, 0
LEAD1.Load strImagenScan, 0, 1, 1
End Sub

Private Sub Form_Load()
On Error Resume Next
Drive1.Drive = "D:\"
'Dir1.Path = "D:\"
'If Dir("D:\MICROFILM", vbDirectory) = "" Then CrearDirectorio ("D:\MICROFILM")

'Dir1.Path = "D:\MICROFILM"
'File1.Path = Dir1.Path

'If frmScan.Name <> "frmScan" Then
    cmdImportar.Caption = "Importar"
'End If


'L_KEY_GIFLZW = "sg8Z2XkjL"
'L_KEY_TIFLZW = "gZWEhj9ZX2j"

LEAD1.UnlockSupport L_SUPPORT_TIFLZW, "gZWEhj9ZX2j"
 


If vbindLoteCreado Then
    cmdHojaRuta.Enabled = False
    cmdImportar.Enabled = True
Else
    cmdHojaRuta.Enabled = True
    cmdImportar.Enabled = False
End If

LEAD1.PaintSizeMode = PAINTSIZEMODE_FIT
End Sub
Sub Importar()

                    
'If FormOrig.Name = "frmScan" Then

    For i = 0 To File1.ListCount - 1
    
        If i = 0 Then
           frmScan.ImportarImagen 0, ImagenBarcode
        End If
        
        If File1.Selected(i) Then
            frmScan.ImportarImagen 0, File1.Path & "\" & File1.list(i)
            'Kill File1.Path & "\" & File1.List(i)
        End If
        
    Next i
    
    frmScan.CargarImagenes
    Unload Me

End Sub




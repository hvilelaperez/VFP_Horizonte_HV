VERSION 5.00
Begin VB.Form frmVolumen 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volumenes"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   ControlBox      =   0   'False
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
   ScaleHeight     =   7425
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstEstado 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4200
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstIDVolumen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4200
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   7050
      Width           =   1215
   End
   Begin VB.CommandButton cmdInformacion 
      Caption         =   "Información"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox lstVolumen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmVolumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVolumen As Recordset
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lstVolumen.ListCount > 0 Then
        If lstVolumen.ListIndex > -1 Then
            If vbYes = MsgBox("Esta seguro de eliminar el volumen " & lstVolumen.List(lstVolumen.ListIndex), vbCritical + vbYesNo, "Eliminar") Then
            End If
        End If
    End If
End Sub
Private Sub cmdInformacion_Click()
    If lstVolumen.ListCount > 0 Then
        If lstVolumen.ListIndex > -1 Then
            bolModoForm = False
            lngIdVolumenActual = lstIDVolumen.List(lstVolumen.ListIndex)
            frmVolInfo.Show vbModal
        End If
    End If
End Sub
Private Sub cmdNuevo_Click()
    bolModoForm = True
    frmVolInfo.Show vbModal
End Sub

Private Sub cmdSeleccionar_Click()
    If lstVolumen.ListCount > 0 Then
        If lstVolumen.ListIndex > -1 Then
            If intMenu = 1 Then
                clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTotalImagenes " & Trim(Me.lstIDVolumen.List(lstVolumen.ListIndex)))
                If lstEstado.List(lstVolumen.ListIndex) = False Then
                    lngIdVolumenActual = lstIDVolumen.List(lstVolumen.ListIndex)
                    MousePointer = 11
                    clsDataScan.Conexion.Execute ("EXECUTE sp_ActualizarTotalImagenes " & lngIdVolumenActual)
                    bolVolumenSeleccionado = True
                    MousePointer = 0
                    Unload Me
                Else
                    MsgBox "El volumen ya esta cerrado, seleccione otro volumen", vbInformation, "Seleccionar"
                End If
            Else
                lngIdVolumenActual = lstIDVolumen.List(lstVolumen.ListIndex)
                bolVolumenSeleccionado = True
                Unload Me
            End If
        End If
    End If
    MousePointer = 0
End Sub

Private Sub Form_Load()
    If intMenu = 1 Then cmdNuevo.Visible = True Else cmdNuevo.Visible = False
    If clsDataScan.ConexionActiva = True Then
        Set rsVolumen = New Recordset
        If rsVolumen.State = 1 Then rsVolumen.Close
        Set rsVolumen = clsDataScan.Conexion.Execute("EXECUTE sp_ConsultarTabla _Volumenes,[*],'',''")
        If rsVolumen.RecordCount > 0 Then
            rsVolumen.MoveFirst
            lstVolumen.Clear
            lstIDVolumen.Clear
            lstEstado.Clear
            Do While Not rsVolumen.EOF
                lstVolumen.AddItem rsVolumen!nombre
                lstIDVolumen.AddItem rsVolumen!idVolumen
                lstEstado.AddItem rsVolumen!estado
                rsVolumen.MoveNext
            Loop
        End If
        If rsVolumen.State = 1 Then rsVolumen.Close
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsVolumen = Nothing
End Sub

Private Sub lstVolumen_DblClick()
    If lstVolumen.ListCount > 0 Then
        If lstVolumen.ListIndex > -1 Then
            cmdInformacion_Click
        End If
    End If
End Sub



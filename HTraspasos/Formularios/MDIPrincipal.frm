VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Profuturo - Sistema de Traspasos  "
   ClientHeight    =   9585
   ClientLeft      =   3240
   ClientTop       =   1275
   ClientWidth     =   13110
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9210
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   1111
      ButtonWidth     =   1799
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "C Doc"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Solicitud"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reportes"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Validacion"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportacion"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lst"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cruzar Data"
            Description     =   "Cruce de Datos Afiliacion"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tablas"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modifica AFP"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1EC45A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F26F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F32C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F3BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F3FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F4444
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F574E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F5A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F66BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F69D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F6E26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDigitalizacion 
      Caption         =   "Digitalización"
      Begin VB.Menu mnuDigitalizar 
         Caption         =   "Digitalizar"
      End
   End
   Begin VB.Menu MnuIng 
      Caption         =   "&Ingreso"
      Begin VB.Menu MnuIngCDoc 
         Caption         =   "Control Doc"
      End
      Begin VB.Menu MnuIngSol 
         Caption         =   "Solicitud"
      End
      Begin VB.Menu mnudigitacionD2 
         Caption         =   "Solicitud 2da Digitacion"
      End
      Begin VB.Menu MnuIngCtaEmp 
         Caption         =   "Corrige Datos Grles"
      End
      Begin VB.Menu MnuIngBusSol 
         Caption         =   "Busca Solicitud"
      End
      Begin VB.Menu MnuIngMod 
         Caption         =   "Modifica Num Solicitud"
      End
      Begin VB.Menu MnuIngDelReg 
         Caption         =   "Elimina Registros"
      End
      Begin VB.Menu mnuAcepSol 
         Caption         =   "Aceptar Solicitudes"
      End
   End
   Begin VB.Menu MnuCons 
      Caption         =   "&Consulta"
      Begin VB.Menu MnuConEnv 
         Caption         =   "Consulta por Envio"
      End
      Begin VB.Menu mnuEstLot 
         Caption         =   "Consulta de Estado de Lotes"
      End
      Begin VB.Menu MnuConVisor 
         Caption         =   "Visor"
      End
      Begin VB.Menu MnuVisorRTFond 
         Caption         =   "Visor Revisa Tipo Fondo"
      End
      Begin VB.Menu MnuConsSol 
         Caption         =   "Solicitudes"
      End
      Begin VB.Menu MnuConsEmp 
         Caption         =   "Empleadores"
      End
      Begin VB.Menu mnuEmpRuc 
         Caption         =   "Cruce Empleados vs TRUC"
      End
      Begin VB.Menu MnuConsImag 
         Caption         =   "Archivo Imágenes"
      End
      Begin VB.Menu MnuConsUbi 
         Caption         =   "Ubigeo"
      End
      Begin VB.Menu MnuConsPromo 
         Caption         =   "Promotores"
      End
      Begin VB.Menu mnuVerD2 
         Caption         =   "Verificar Segunda Digitación"
      End
      Begin VB.Menu MnuConsVeri 
         Caption         =   "Verifica Ria fechas"
      End
      Begin VB.Menu MnuVerCuspp 
         Caption         =   "Verifica CUSPP"
      End
      Begin VB.Menu MnuVeriTfon 
         Caption         =   "Verifica Tipo Fondo"
      End
      Begin VB.Menu mnuCargos 
         Caption         =   "Cargos"
      End
      Begin VB.Menu mnuNacionalidad 
         Caption         =   "Nacionalidad"
      End
      Begin VB.Menu mnuSucursal 
         Caption         =   "Sucursal"
      End
      Begin VB.Menu mnuHV 
         Caption         =   "Revision Hoja Vida"
      End
      Begin VB.Menu mnuRepFam 
         Caption         =   "Revision Familiares"
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "&Reportes"
      Begin VB.Menu MnuRepImagLot 
         Caption         =   "Imagenes x Lote"
      End
      Begin VB.Menu MnuRepImagSID 
         Caption         =   "Imagenes Sin ID"
      End
      Begin VB.Menu MnuRepSolProc 
         Caption         =   "Solicitudes Proc"
      End
      Begin VB.Menu MnuRepDigi 
         Caption         =   "Digitador"
      End
      Begin VB.Menu mnsRepErrAses 
         Caption         =   "Errores x Asesor"
      End
   End
   Begin VB.Menu MnuExpo 
      Caption         =   "&Exportacion"
   End
   Begin VB.Menu mnuPromotor 
      Caption         =   "&Importación"
      Begin VB.Menu mnuImpPro 
         Caption         =   "Promotor"
      End
      Begin VB.Menu mnuIAI1 
         Caption         =   "IAI1"
      End
   End
   Begin VB.Menu MnuTab 
      Caption         =   "&Tablas"
   End
   Begin VB.Menu mnuRevFdo 
      Caption         =   "Revisa Fondos"
      Begin VB.Menu MnuRFdoCarg 
         Caption         =   "Modifica Archivo Cargado"
      End
      Begin VB.Menu MnuRFdoIng 
         Caption         =   "Ingresa Datos x Revision"
      End
      Begin VB.Menu MnuRFdoMod 
         Caption         =   "Modifica Registros"
      End
      Begin VB.Menu MnuRFdoRep 
         Caption         =   "Proceso y Reportes"
      End
      Begin VB.Menu MnuRFdoExp 
         Caption         =   "Exporta Tipo Fondo"
      End
   End
   Begin VB.Menu mnuImpresion 
      Caption         =   "Impresión"
      Begin VB.Menu mnucopiado 
         Caption         =   "Copiado de Imágenes"
      End
   End
   Begin VB.Menu MnuSal 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
'    Con.ConexionSQL "sa", "", "DIGITACION\DIGITACION", "BDTraspasos"
'    DatosServidor
'    CargarRegistros RsF, "Facturas"
'    CargarRegistros RsI, "Pedidos"
'    CargarRegistros RsP, "Productos"
'    CargarRegistros RsC, "Clientes"
'    CargarRegistros RsU, "Usuarios"
vbCaracteresNumero = "1234567890"
Me.StatusBar1.Panels(1).Text = gNomUsu
vbGAFP = "01"
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Esta seguro de salir?", 4 + 32, "Salir") = 6 Then
        End
    Else
        Cancel = True
    End If
End Sub

Private Sub MDIForm_Resize()

'AutoRedraw = True
'DrawStyle = 6
'DrawMode = 13
'DrawWidth = 2
'ScaleMode = 3
''ScaleHeight = (256 * 2)
'For i = 0 To 255
'Me.Line (0, Y)-(Width, Y + 2), RGB(0, 0, i), BF
'Y = Y + 2
'Next i
End Sub

Private Sub mnsRepErrAses_Click()
    frmRepErrorAsesor.Show
End Sub

Private Sub mnuAcepSol_Click()
 If Mid(gNomUsu, 1, 1) = "S" Then
         frmGestionSuspencion.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Mantenimiento de Tablas"
    End If

End Sub

Private Sub mnuCargos_Click()
    FrmCargos.Show
End Sub

Private Sub MnuConEnv_Click()
   FrmConCdoc.Show
End Sub

Private Sub MnuConsEmp_Click()
 
 If Mid(gNomUsu, 1, 1) = "S" Then
        FrmRepEmple.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Mantenimiento de Tablas"
    End If
End Sub

Private Sub MnuConsImag_Click()
If Mid(gNomUsu, 1, 1) = "S" Then
    FrmRepImag.Show
 Else
    MsgBox "No está autorizado para esta Opción", vbCritical, "Error en Mantenimiento"
End If
End Sub

Private Sub MnuConsPromo_Click()
    FrmConsPromo.Show
End Sub

Private Sub MnuConsSol_Click()
  If Mid(gNomUsu, 1, 1) = "S" Then
        FrmRepSolic.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Solicitudes"
    End If
  
End Sub

Private Sub mnucopiado_Click()
frmImpresion.Show
End Sub

Private Sub mnudigitacionD2_Click()
frmSolicitudesD2.Show
End Sub

Private Sub mnuDigitalizar_Click()
frmScan.Show 1
End Sub

Private Sub mnuEmpRuc_Click()
frmConsDifEmpTRuc.Show
End Sub

Private Sub mnuEstLot_Click()
frmConsultaLote.Show
End Sub

Private Sub mnuHV_Click()

If Mid(gNomUsu, 1, 1) = "S" Then
      FrmRepHV.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Solicitudes"
    End If



End Sub

Private Sub mnuIAI1_Click()
frmHabilitado.Show 1
End Sub

Private Sub mnuImpPro_Click()
frmImportacionPro.Show
End Sub

Private Sub MnuIngMod_Click()
  If Mid(gNomUsu, 1, 1) = "S" Then
        FrmModSol.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Mantenimiento de Tablas"
    End If
End Sub

Private Sub MnuIngDelReg_Click()
  If Mid(gNomUsu, 1, 1) = "S" Then
        FrmDelReg.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Mantenimiento de Tablas"
    End If
End Sub

Private Sub MnuConsUbi_Click()
 FrmConsUbi.Show
End Sub

Private Sub MnuConsVeri_Click()

  If Mid(gNomUsu, 1, 1) = "S" Then
      FrmVeriRiaFec2.Show
  Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Solicitudes"
  End If


End Sub

Private Sub MnuConVisor_Click()
  FrmVisorArchivo.Show
End Sub

Private Sub MnuIngBusSol_Click()
    FrmBuscaSol.Show
End Sub

Private Sub MnuIngCDoc_Click()
    If Mid(gNomUsu, 1, 3) = "SUP" Then
        FrmCdoc.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If
    
End Sub

Private Sub MnuIngCtaEmp_Click()

If Mid(gNomUsu, 1, 3) = "SUP" Then
        FrmCtaEmp.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If

  
End Sub

Private Sub MnuIngSol_Click()

     frmTraspasos1.Show

End Sub

Private Sub MnuCons_Click()
    'FrmConCdoc.Show
End Sub

Private Sub mnuNacionalidad_Click()
FrmNacion.Show 1
End Sub

Private Sub MnuRepDigi_Click()
FrmRepDig.Show
End Sub

Private Sub mnuRepFam_Click()

If Mid(gNomUsu, 1, 1) = "S" Then
      FrmRepFami.Show 1
Else
      MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Solicitudes"
End If


End Sub

Private Sub MnuRepImagLot_Click()
    FrmRepImagLote.Show
End Sub

Private Sub MnuRepImagSID_Click()
    FrmRepImagLote.Show
End Sub

Private Sub MnuRepSolProc_Click()
        FrmRepSolProc.Show
End Sub

Private Sub MnuExpo_Click()
     frmExportacion.Show 1
End Sub

Private Sub MnuRFdoMod_Click()
    If Mid(gNomUsu, 1, 1) = "S" Or Mid(gNomUsu, 1, 1) = "I" Then
        FrmModFdo.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If
End Sub

Private Sub mnuSucursal_Click()
FrmSucursal.Show 1
End Sub

Private Sub MnuTab_Click()
    If Mid(gNomUsu, 1, 1) = "S" Then
        FrmUsuarios.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Mantenimiento de Tablas"
    End If
End Sub

Private Sub MnuAfp_Click()
     frmOpciones.Show
End Sub

Private Sub MnuSal_Click()
    If MsgBox("Esta seguro de salir?", 4 + 32, "Salir") = 6 Then End
End Sub

Private Sub MnuVerCuspp_Click()
        
    If Mid(gNomUsu, 1, 1) = "S" Then
      FrmVeriCuspp.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Solicitudes"
    End If
    
End Sub

Private Sub mnuVerD2_Click()
    frmVerificacionD2.Show
End Sub

Private Sub MnuVeriTfon_Click()

If Mid(gNomUsu, 1, 1) = "S" Then
      FrmConsFondo.Show 1
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Solicitudes"
    End If


    
End Sub

Private Sub MnuVisorRTFond_Click()
            
    If Mid(gNomUsu, 1, 1) = "S" Then
        FrmVisorReviTFon.Show 1
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            MnuIngCDoc_Click
        Case 2
            MnuIngSol_Click
        Case 3
            MnuCons_Click
        Case 4
            MnuRepImagLot_Click
        Case 5
            If Mid(gNomUsu, 1, 1) = "S" Or Mid(gNomUsu, 1, 1) = "I" Then
                frmValidaciones.Show 1
            Else
                MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
            End If

        Case 6
            MnuExpo_Click
        Case 7
            frmGenLst.Show 1
        Case 8
            frmCruce.Show 1
        Case 9
            MnuTab_Click
        Case 10
            MnuAfp_Click
        Case 11
            MnuSal_Click
    End Select
End Sub
Private Sub MnuRFdoCarg_Click()
    If Mid(gNomUsu, 1, 1) = "S" Or Mid(gNomUsu, 1, 1) = "I" Then
        FrmCargaRevFdo.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If
End Sub
 
Private Sub MnuRFdoExp_Click()
    If Mid(gNomUsu, 1, 1) = "S" Or Mid(gNomUsu, 1, 1) = "I" Then
        FrmExpoFdo.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If
End Sub
 
Private Sub MnuRFdoIng_Click()
    If Mid(gNomUsu, 1, 1) = "S" Or Mid(gNomUsu, 1, 1) = "I" Then
        FrmIngFondo2.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If
End Sub
 
Private Sub MnuRFdoRep_Click()
    If Mid(gNomUsu, 1, 1) = "S" Or Mid(gNomUsu, 1, 1) = "I" Then
        FrmProcFdo.Show
    Else
        MsgBox "No esta autorizado para esta Opcion ", vbCritical, "Error en Revision Mensual de Fondos"
    End If
End Sub


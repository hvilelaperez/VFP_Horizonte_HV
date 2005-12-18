VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Flow_Frmpanel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Liquidación Horizonte Flow"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12930
   Icon            =   "Frmpanel.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9030
   ScaleWidth      =   12930
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8655
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmpanel.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmpanel.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmpanel.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmpanel.frx":3228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmpanel.frx":367A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmpanel.frx":3ACC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   1535
      ButtonWidth     =   2249
      ButtonHeight    =   1376
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Generar Archivo"
            Key             =   "Generar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Asignar Lote"
            Key             =   "asiglote"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleccionar Lote"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Configuracion"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnugenarchivo 
         Caption         =   "&Generar Archivo"
      End
   End
   Begin VB.Menu mnulote 
      Caption         =   "&Lote"
      Begin VB.Menu mnuasiglote 
         Caption         =   "&Asignar Lote"
      End
      Begin VB.Menu mnuseleccionarlote 
         Caption         =   "&Seleccionar Lote"
      End
   End
   Begin VB.Menu mnuprocesos 
      Caption         =   "&Procesos-Flow"
      Begin VB.Menu mnuimportdatos 
         Caption         =   "Importar Datos"
      End
   End
   Begin VB.Menu mnuusuarios 
      Caption         =   "&Usuarios"
   End
   Begin VB.Menu mnusalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Flow_Frmpanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ContRutatxt = App.Path & "\Horizonte.ini"
     
        Me.Caption = App.EXEName
        mnulote.Visible = False
        mnuusuarios.Visible = False
        'mnuVentanas.Visible = False
        mnusalir.Visible = False
        mnuprocesos.Visible = False
        Setear
        
        'Data Source = ServidorSQL; Initial Catalog = BaseDatos; Integrated Security = True
        
        strCadena = "Driver={SQL Server};server=" & StrServidorSQL & ";database=" & strDataSQL
        If cnServer.State = 1 Then cnServer.Close
        cnServer.ConnectionTimeout = 0
        cnServer.CommandTimeout = 0
        cnServer.ConnectionString = strCadena
        cnServer.CursorLocation = adUseClient
        cnServer.Open
        
        If StrTipoApli = "1" Then
        Flow_frmbienvenida.Show vbModal
        Else
          If StrTipoApli = "2" Then
          Unload Me
          Icr_Frmpanelicr.Show vbModal
          End If
        End If
End Sub

Private Sub mnuasiglote_Click()
Flow_FrmasignarLote.Show vbModal
End Sub

Private Sub mnugenarchivo_Click()
Flow_Frmarchivo.Show vbModal
End Sub

Private Sub mnuimportdatos_Click()
Flow_FrmImportarData.Show vbModal
End Sub

Private Sub mnusalir_Click()
Unload Me
End Sub

Private Sub mnuseleccionarlote_Click()
Flow_frmseleccionar.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case UCase(Button)
        Case Is = "GENERAR ARCHIVO"
          Flow_Frmarchivo.Show vbModal
        Case Is = "ASIGNAR LOTE"
          Flow_FrmasignarLote.Show vbModal
        Case Is = "SELECCIONAR LOTE"
          Flow_frmseleccionar.Show vbModal
        Case Is = "VALIDACION"
          Flow_FrmValidacion.Show vbModal
    End Select
End Sub

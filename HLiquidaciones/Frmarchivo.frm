VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Flow_Frmarchivo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Importar Archivo Texto"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   Icon            =   "Frmarchivo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   11490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.CommandButton Cmdprocesotxt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Procesar Archivos"
         Height          =   735
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Cmdverificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Verificar Archivos"
         Height          =   735
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame FraUbicacion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Ubicación Archivos Origén Horizonte"
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8655
         Begin VB.DriveListBox Drvxls 
            Height          =   315
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Width           =   8055
         End
         Begin VB.DirListBox Dirxls 
            Height          =   1440
            Left            =   480
            TabIndex        =   8
            Top             =   720
            Width           =   2655
         End
         Begin VB.FileListBox flxls 
            Height          =   1455
            Left            =   3120
            Pattern         =   "*.LIST;*.TXT"
            TabIndex        =   7
            Top             =   720
            Width           =   5415
         End
      End
      Begin VB.CommandButton CmdProceso 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Procesar Archivos"
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame FraBarra 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   2520
         Width           =   8655
         Begin MSComctlLib.ProgressBar PrbBarra 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
      End
      Begin VB.CommandButton cmdruta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Configuracion de archivo"
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Salir"
         Height          =   735
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label LblHorizonte 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   8535
      End
      Begin VB.Label lblcarta 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   8295
      End
   End
End
Attribute VB_Name = "Flow_Frmarchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rutad As String

Private Sub cmdproceso_Click()
''Records para Archivo
'Dim Rstcob4 As Recordset
Dim Rstcob5 As Recordset
Dim Rstcob4  As Recordset
Dim Rstcob6 As Recordset
'''Records para totales Generales
Dim RstTotal As Recordset
Dim IntTotalRecord As Integer
Dim IntTotalRecord1 As Integer
Dim intPag As Boolean
Dim NroCargo As String
Dim ContcargoINT As Integer
Dim BarraCarta As String
Dim IntArchivos As Integer
Dim CmdDatos As New ADODB.Command
''''''

'''''''''''''
Dim Cantdet As Double
'''''''''''''

PrbBarra.Min = 0

 ''' Archivo Liquidación
 'Open RutaArchivos & "\Liquidacion.txt" For Output As #3
 ''''Archivo Carta
 
 If optconpago.Value = True Then
 Open RutaArchivos & "\CartaConPago.txt" For Output As #4
 End If
 
 If optsinpago.Value = True Then
 Open RutaArchivos & "\CartaSinPago.txt" For Output As #4
 End If

 If Optbajopago.Value = True Then
 Open RutaArchivos & "\BajoPago.txt" For Output As #4
 End If
Abrir_conexion

''''''''''''''''''''''''''''''''
'Set Rstcob4 = Sel_Int_CabezeraSql


For IntArchivos = 1 To 2
     
     If IntArchivos = 1 Then
     Set Rstcob4 = Sel_Int_CabezeraSql_Lima_Callao
     Else
        If IntArchivos = 2 Then
        Set Rstcob4 = Sel_Int_CabezeraSql_Provincias
        End If
     End If
     
            If Not Rstcob4.EOF Then
            Rstcob4.MoveLast
            Rstcob4.MoveFirst
              
              If IntArchivos = 1 Then
              Open RutaArchivos & "\Liquidacion_Lima_Callao.txt" For Output As #3
              Print #3, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
                                          ; vbTab; "CAMPO11"; vbTab; "CAMPO12"; vbTab; "CAMPO13"; vbTab; "CAMPO14"; vbTab; "CAMPO15"; vbTab; "CAMPO16"; vbTab; "CAMPO17"; vbTab; "CAMPO18"; vbTab; "CAMPO19" _
                                          ; vbTab; "CAMPO20"; vbTab; "CAMPO21"; vbTab; "CAMPO22"; vbTab; "CAMPO23"; vbTab; "CAMPO24"; vbTab; "CAMPO25"; vbTab; "CAMPO26"; vbTab; "CAMPO27"; vbTab; "CAMPO28" _
                                          ; vbTab; "CAMPO29"; vbTab; "CAMPO30"; vbTab; "CAMPO31"; vbTab; "CAMPO32"; vbTab; "CAMPO33"; vbTab; "CAMPO34"; vbTab; "CAMPO35"; vbTab; "CAMPO36"; vbTab; "CAMPO37"; vbTab; "CAMPO38" _
                                          ; vbTab; "CAMPO39"; vbTab; "CAMPO40"; vbTab; "CAMPO41"; vbTab; "CAMPO42"; vbTab; "CAMPO43"; vbTab; "CAMPO44"; vbTab; "CAMPO45"; vbTab; "CAMPO46" _
                                          ; vbTab; "CAMPO47"; vbTab; "CAMPO48"; vbTab; "CAMPO49"; vbTab; "CAMPO50"; vbTab; "CAMPO51"; vbTab; "CAMPO52"; vbTab; "CAMPO53"; vbTab; "CAMPO54"; vbTab; "CAMPO55"; vbTab; "CAMPO56"; vbTab; "CAMPO57" _
                                          ; vbTab; "CAMPO58"; vbTab; "CAMPO59"; vbTab; "CAMPO60"; vbTab; "CAMPO61"; vbTab; "CAMPO62"; vbTab; "CAMPO63"; vbTab; "CAMPO64"; vbTab; "CAMPO65"; vbTab; "CAMPO66"; vbTab; "CAMPO67"; vbTab; "CAMPO68" _
                                          ; vbTab; "CAMPO69"; vbTab; "CAMPO70"; vbTab; "CAMPO71"; vbTab; "CAMPO72"; vbTab; "CAMPO73"; vbTab; "CAMPO74"; vbTab; "CAMPO75"; vbTab; "CAMPO76"; vbTab; "CAMPO77"; vbTab; "CAMPO78"; vbTab; "CAMPO79" _
                                          ; vbTab; "CAMPO80"; vbTab; "CAMPO81"; vbTab; "CAMPO82"; vbTab; "CAMPO83"; vbTab; "CAMPO84"; vbTab; "CAMPO85"; vbTab; "CAMPO86"; vbTab; "CAMPO87"; vbTab; "CAMPO88"; vbTab; "CAMPO89"; vbTab; "CAMPO90" _
                                          ; vbTab; "CAMPO91"; vbTab; "CAMPO92"; vbTab; "CAMPO93"; vbTab; "CAMPO94"; vbTab; "CAMPO95"; vbTab; "CAMPO96"; vbTab; "CAMPO97"; vbTab; "CAMPO98"; vbTab; "CAMPO99"; vbTab; "CAMPO100"; vbTab; "CAMPO101" _
                                          ; vbTab; "CAMPO102"; vbTab; "CAMPO103"; vbTab; "CAMPO104"; vbTab; "CAMPO105"; vbTab; "CAMPO106"; vbTab; "CAMPO107"; vbTab; "CAMPO108"; vbTab; "CAMPO109"; vbTab; "CAMPO110"; vbTab; "CAMPO111" _
                                          ; vbTab; "CAMPO112"; vbTab; "CAMPO113"; vbTab; "CAMPO114"; vbTab; "CAMPO115"; vbTab; "CAMPO116"; vbTab; "CAMPO117"; vbTab; "CAMPO118"; vbTab; "CAMPO119"; vbTab; "CAMPO120"; vbTab; "CAMPO121" _
                                          ; vbTab; "CAMPO122"; vbTab; "CAMPO123"; vbTab; "CAMPO124"; vbTab; "CAMPO125"; vbTab; "CAMPO126"; vbTab; "CAMPO127"; vbTab; "CAMPO128"; vbTab; "CAMPO129"; vbTab; "CAMPO130"; vbTab; "CAMPO131"; vbTab; "CAMPO132" _
                                          ; vbTab; "CAMPO133"; vbTab; "CAMPO134"; vbTab; "CAMPO135"; vbTab; "CAMPO136"; vbTab; "CAMPO137"; vbTab; "CAMPO138"; vbTab; "CAMPO139"; vbTab; "CAMPO140"; vbTab; "CAMPO141"; vbTab; "CAMPO142"
            
              
              Else
                  If IntArchivos = 2 Then
                  Open RutaArchivos & "\Liquidacion_Provincias.txt" For Output As #5
                   Print #5, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
                                          ; vbTab; "CAMPO11"; vbTab; "CAMPO12"; vbTab; "CAMPO13"; vbTab; "CAMPO14"; vbTab; "CAMPO15"; vbTab; "CAMPO16"; vbTab; "CAMPO17"; vbTab; "CAMPO18"; vbTab; "CAMPO19" _
                                          ; vbTab; "CAMPO20"; vbTab; "CAMPO21"; vbTab; "CAMPO22"; vbTab; "CAMPO23"; vbTab; "CAMPO24"; vbTab; "CAMPO25"; vbTab; "CAMPO26"; vbTab; "CAMPO27"; vbTab; "CAMPO28" _
                                          ; vbTab; "CAMPO29"; vbTab; "CAMPO30"; vbTab; "CAMPO31"; vbTab; "CAMPO32"; vbTab; "CAMPO33"; vbTab; "CAMPO34"; vbTab; "CAMPO35"; vbTab; "CAMPO36"; vbTab; "CAMPO37"; vbTab; "CAMPO38" _
                                          ; vbTab; "CAMPO39"; vbTab; "CAMPO40"; vbTab; "CAMPO41"; vbTab; "CAMPO42"; vbTab; "CAMPO43"; vbTab; "CAMPO44"; vbTab; "CAMPO45"; vbTab; "CAMPO46" _
                                          ; vbTab; "CAMPO47"; vbTab; "CAMPO48"; vbTab; "CAMPO49"; vbTab; "CAMPO50"; vbTab; "CAMPO51"; vbTab; "CAMPO52"; vbTab; "CAMPO53"; vbTab; "CAMPO54"; vbTab; "CAMPO55"; vbTab; "CAMPO56"; vbTab; "CAMPO57" _
                                          ; vbTab; "CAMPO58"; vbTab; "CAMPO59"; vbTab; "CAMPO60"; vbTab; "CAMPO61"; vbTab; "CAMPO62"; vbTab; "CAMPO63"; vbTab; "CAMPO64"; vbTab; "CAMPO65"; vbTab; "CAMPO66"; vbTab; "CAMPO67"; vbTab; "CAMPO68" _
                                          ; vbTab; "CAMPO69"; vbTab; "CAMPO70"; vbTab; "CAMPO71"; vbTab; "CAMPO72"; vbTab; "CAMPO73"; vbTab; "CAMPO74"; vbTab; "CAMPO75"; vbTab; "CAMPO76"; vbTab; "CAMPO77"; vbTab; "CAMPO78"; vbTab; "CAMPO79" _
                                          ; vbTab; "CAMPO80"; vbTab; "CAMPO81"; vbTab; "CAMPO82"; vbTab; "CAMPO83"; vbTab; "CAMPO84"; vbTab; "CAMPO85"; vbTab; "CAMPO86"; vbTab; "CAMPO87"; vbTab; "CAMPO88"; vbTab; "CAMPO89"; vbTab; "CAMPO90" _
                                          ; vbTab; "CAMPO91"; vbTab; "CAMPO92"; vbTab; "CAMPO93"; vbTab; "CAMPO94"; vbTab; "CAMPO95"; vbTab; "CAMPO96"; vbTab; "CAMPO97"; vbTab; "CAMPO98"; vbTab; "CAMPO99"; vbTab; "CAMPO100"; vbTab; "CAMPO101" _
                                          ; vbTab; "CAMPO102"; vbTab; "CAMPO103"; vbTab; "CAMPO104"; vbTab; "CAMPO105"; vbTab; "CAMPO106"; vbTab; "CAMPO107"; vbTab; "CAMPO108"; vbTab; "CAMPO109"; vbTab; "CAMPO110"; vbTab; "CAMPO111" _
                                          ; vbTab; "CAMPO112"; vbTab; "CAMPO113"; vbTab; "CAMPO114"; vbTab; "CAMPO115"; vbTab; "CAMPO116"; vbTab; "CAMPO117"; vbTab; "CAMPO118"; vbTab; "CAMPO119"; vbTab; "CAMPO120"; vbTab; "CAMPO121" _
                                          ; vbTab; "CAMPO122"; vbTab; "CAMPO123"; vbTab; "CAMPO124"; vbTab; "CAMPO125"; vbTab; "CAMPO126"; vbTab; "CAMPO127"; vbTab; "CAMPO128"; vbTab; "CAMPO129"; vbTab; "CAMPO130"; vbTab; "CAMPO131"; vbTab; "CAMPO132" _
                                          ; vbTab; "CAMPO133"; vbTab; "CAMPO134"; vbTab; "CAMPO135"; vbTab; "CAMPO136"; vbTab; "CAMPO137"; vbTab; "CAMPO138"; vbTab; "CAMPO139"; vbTab; "CAMPO140"; vbTab; "CAMPO141"; vbTab; "CAMPO142"
                
                  End If
              End If
              
              
              
              
              ''Armar Encabezado
              
              Flow_Frmarchivo.MousePointer = vbHourglass
                
'                Print #3, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
'                                          ; vbTab; "CAMPO11"; vbTab; "CAMPO12"; vbTab; "CAMPO13"; vbTab; "CAMPO14"; vbTab; "CAMPO15"; vbTab; "CAMPO16"; vbTab; "CAMPO17"; vbTab; "CAMPO18"; vbTab; "CAMPO19" _
'                                          ; vbTab; "CAMPO20"; vbTab; "CAMPO21"; vbTab; "CAMPO22"; vbTab; "CAMPO23"; vbTab; "CAMPO24"; vbTab; "CAMPO25"; vbTab; "CAMPO26"; vbTab; "CAMPO27"; vbTab; "CAMPO28" _
'                                          ; vbTab; "CAMPO29"; vbTab; "CAMPO30"; vbTab; "CAMPO31"; vbTab; "CAMPO32"; vbTab; "CAMPO33"; vbTab; "CAMPO34"; vbTab; "CAMPO35"; vbTab; "CAMPO36"; vbTab; "CAMPO37"; vbTab; "CAMPO38" _
'                                          ; vbTab; "CAMPO39"; vbTab; "CAMPO40"; vbTab; "CAMPO41"; vbTab; "CAMPO42"; vbTab; "CAMPO43"; vbTab; "CAMPO44"; vbTab; "CAMPO45"; vbTab; "CAMPO46" _
'                                          ; vbTab; "CAMPO47"; vbTab; "CAMPO48"; vbTab; "CAMPO49"; vbTab; "CAMPO50"; vbTab; "CAMPO51"; vbTab; "CAMPO52"; vbTab; "CAMPO53"; vbTab; "CAMPO54"; vbTab; "CAMPO55"; vbTab; "CAMPO56"; vbTab; "CAMPO57" _
'                                          ; vbTab; "CAMPO58"; vbTab; "CAMPO59"; vbTab; "CAMPO60"; vbTab; "CAMPO61"; vbTab; "CAMPO62"; vbTab; "CAMPO63"; vbTab; "CAMPO64"; vbTab; "CAMPO65"; vbTab; "CAMPO66"; vbTab; "CAMPO67"; vbTab; "CAMPO68" _
'                                          ; vbTab; "CAMPO69"; vbTab; "CAMPO70"; vbTab; "CAMPO71"; vbTab; "CAMPO72"; vbTab; "CAMPO73"; vbTab; "CAMPO74"; vbTab; "CAMPO75"; vbTab; "CAMPO76"; vbTab; "CAMPO77"; vbTab; "CAMPO78"; vbTab; "CAMPO79" _
'                                          ; vbTab; "CAMPO80"; vbTab; "CAMPO81"; vbTab; "CAMPO82"; vbTab; "CAMPO83"; vbTab; "CAMPO84"; vbTab; "CAMPO85"; vbTab; "CAMPO86"; vbTab; "CAMPO87"; vbTab; "CAMPO88"; vbTab; "CAMPO89"; vbTab; "CAMPO90" _
'                                          ; vbTab; "CAMPO91"; vbTab; "CAMPO92"; vbTab; "CAMPO93"; vbTab; "CAMPO94"; vbTab; "CAMPO95"; vbTab; "CAMPO96"; vbTab; "CAMPO97"; vbTab; "CAMPO98"; vbTab; "CAMPO99"; vbTab; "CAMPO100"; vbTab; "CAMPO101" _
'                                          ; vbTab; "CAMPO102"; vbTab; "CAMPO103"; vbTab; "CAMPO104"; vbTab; "CAMPO105"; vbTab; "CAMPO106"; vbTab; "CAMPO107"; vbTab; "CAMPO108"; vbTab; "CAMPO109"; vbTab; "CAMPO110"; vbTab; "CAMPO111" _
'                                          ; vbTab; "CAMPO112"; vbTab; "CAMPO113"; vbTab; "CAMPO114"; vbTab; "CAMPO115"; vbTab; "CAMPO116"; vbTab; "CAMPO117"; vbTab; "CAMPO118"; vbTab; "CAMPO119"; vbTab; "CAMPO120"; vbTab; "CAMPO121" _
'                                          ; vbTab; "CAMPO122"; vbTab; "CAMPO123"; vbTab; "CAMPO124"; vbTab; "CAMPO125"; vbTab; "CAMPO126"; vbTab; "CAMPO127"; vbTab; "CAMPO128"; vbTab; "CAMPO129"; vbTab; "CAMPO130"; vbTab; "CAMPO131"; vbTab; "CAMPO132" _
'                                          ; vbTab; "CAMPO133"; vbTab; "CAMPO134"; vbTab; "CAMPO135"; vbTab; "CAMPO136"; vbTab; "CAMPO137"; vbTab; "CAMPO138"; vbTab; "CAMPO139"; vbTab; "CAMPO140"; vbTab; "CAMPO141"; vbTab; "CAMPO142"
'
                
                
'                Print #4, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
'                                          ; vbTab; "CAMPO11"; vbTab; "CAMPO12"; vbTab; "CAMPO13"; vbTab; "CAMPO14"; vbTab; "CAMPO15"; vbTab; "CAMPO16"
'
            
            
                 PrbBarra.Min = 0
                 PrbBarra.Max = Rstcob4.RecordCount + 1
                 ContIntegra = 0
              While Not Rstcob4.EOF
                 X = X + 1
                 'If (Rstcob4.Fields("cobnm2") <> "IN10056461" And Rstcob4.Fields("cobnm2") <> "IN10058374") Then
                                 DoEvents
'                                 PrbBarra.Value = X + 1
                                 FechaLiq = Trim$(Rstcob4.Fields("cobfe6"))
                                 NroLiq = Trim$(Trim$(Rstcob4.Fields("cobti7")) & "-" & Trim$(Rstcob4.Fields("cobnm1"))) & "-" & Trim$(Rstcob4.Fields("cobse2"))
                                 Razon = Replace(Trim$(IIf(IsNull(Rstcob4.Fields("cobra2")), "", Rstcob4.Fields("cobra2"))), "Ð", "Ñ")
                                 Ruc = Trim$(IIf(IsNull(Rstcob4.Fields("cobnr2")), "", Rstcob4.Fields("cobnr2")))
                                 Direct = Trim$(Rstcob4.Fields("cobdi2"))
            
                                 ContSecCargo = Format(X, "000000")
            
            
                            '     NroLiqC = Trim$(Trim$(Rstcob4.Fields("cobti5")) & "-" & Trim$(Rstcob4.Fields("cobnme")))
                            '     FechaLiqC = Trim$(Rstcob4.Fields("cobloc")) & ", " & Trim$(Rstcob4.Fields("cobfe5"))
                            '     RazonC = Trim$(IIf(IsNull(Rstcob4.Fields("cobra2")), "", Rstcob4.Fields("cobra2")))
                            '     DirectC = Trim$(Rstcob4.Fields("cobdi2"))
                            '     DistritoC = Trim$(Rstcob4.Fields("cobdi1"))
                            '     DepartamentoC = Trim$(Rstcob4.Fields("cobno1")) & "/" & Trim$(Rstcob4.Fields("cobno2"))
                            '     RucC = Trim$(IIf(IsNull(Rstcob4.Fields("cobti6")), "", Rstcob4.Fields("cobti6"))) & " " & Trim$(IIf(IsNull(Rstcob4.Fields("cobnr1")), " ", Rstcob4.Fields("cobnr1")))
                            '     Atencion = Trim$(IIf(IsNull(Rstcob4.Fields("cobate")), "", Rstcob4.Fields("cobate")))
                            '     Referencia = Trim$(IIf(IsNull(Rstcob4.Fields("cobref")), "", Rstcob4.Fields("cobref")))
                            '
                            '     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin")))
                            '     If Len(Trim$(Rstcob4.Fields("cobli3"))) > 0 Then
                            '     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli3")), "", Rstcob4.Fields("cobli3")))
                            '     End If
                            '
                            '     If Len(Trim$(Rstcob4.Fields("cobli4"))) > 0 Then
                            '     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli3")), "", Rstcob4.Fields("cobli3"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli4")), "", Rstcob4.Fields("cobli4")))
                            '     End If
                            '
                            '     If Len(Trim$(Rstcob4.Fields("cobli5"))) > 0 Then
                            '     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli3")), "", Rstcob4.Fields("cobli3"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli4")), "", Rstcob4.Fields("cobli4"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli5")), "", Rstcob4.Fields("cobli5")))
                            '     End If
                            '
                            '     If Len(Trim$(Rstcob4.Fields("cobli6"))) > 0 Then
                            '     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli3")), "", Rstcob4.Fields("cobli3"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli4")), "", Rstcob4.Fields("cobli4"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli5")), "", Rstcob4.Fields("cobli5"))) & ", " & Trim$(IIf(IsNull(Rstcob4.Fields("cobli6")), "", Rstcob4.Fields("cobli6")))
                            '     End If
                            '
            
            
                                 'Trim$(IIf(IsNull(Rstcob4.Fields("cobadj")), "", Rstcob4.Fields("cobadj"))) & ":" &
                              '''''Consulta Cobar 5
                                 'Set Rstcob5 = Sel_Int_Cabezera2(Rstcob4.Fields("cobnme"))
                                 ' If Rstcob5.RecordCount > 0 Then
                                  Ndepart = Trim$(IIf(IsNull(Rstcob4.Fields("cobdi3")), "", Rstcob4.Fields("cobdi3")))
                                  Urbani = Trim$(IIf(IsNull(Rstcob4.Fields("cobdi4")), "", Rstcob4.Fields("cobdi4")))
                                  Distrito = Trim$(Rstcob4.Fields("cobno3"))
                                  Provincia = Trim$(Rstcob4.Fields("cobno4"))
                                  Departamento = Trim$(Rstcob4.Fields("cobno5"))
                                  Pag = Trim$(Rstcob4.Fields("Cobpgi"))
            
                                  Liquipre = Trim$(Rstcob4.Fields("cobnm2"))
                                  PeriDeve = Trim$(Rstcob4.Fields("cobpe2"))
            
                                  ''' Codigo de Barras'''''
                                  'Campos Razon Social/Ruc/Tipo de Cargo/Nro Cargo
            
                            '      CodBarra = "*" & Trim$(Rstcob4.Fields("cobnm1")) & "*"
            
            
            
            
                                  'CodBarra = Trim$(Rstcob4.Fields("cobnm2"))
            
                                  ''''''''''''''''''''''''''
                                 ' End If
                                '''Grabar El Archivo Texto de Carta
                                '      Formatear_Archivo_Texto_Carta
                                '      Print #4, TXTC(1); vbTab; TXTC(2); vbTab; TXTC(3); vbTab; TXTC(4); vbTab; TXTC(5); vbTab; TXTC(6); vbTab; TXTC(7); vbTab; TXTC(8); vbTab; TXTC(9); vbTab; TXTC(10) _
                                '                              ; vbTab; TXTC(11); vbTab; TXTC(12)
            
                                '''''''
                                '''''Consulta Cobar6
                                   'Set Rstcob6 = Sel_Int_Detalle(Rstcob4.Fields("cobnme"))
                                                                  '            If IntPag = True Then
                                                                  '               If NroCargo = Rstcob4.Fields("cobnm1") Then
                                                                  '               ContcargoINT = ContcargoINT + 1
                                                                  '                  If ContPagCargo = Format(ContcargoINT, "000") Then
                                                                  '                  IntPag = False
                                                                  '                  'ContcargoINT = 0
                                                                  '                  End If
                                                                  '               Else
                                                                  '               ContcargoINT = 0
                                                                  '               IntPag = False
                                                                  '               End If
                                                                  '            Else
                                                                  '            IntTotalRecord1 = Sel_Int_Pag_Cargo(Rstcob4.Fields("cobnm1"))
                                                                  '            ContPagCargo = Format(Str(IntTotalRecord1), "000")
                                                                  '            IntPag = True
                                                                  '            ContcargoINT = 1
                                                                  '              If Val(ContPagCargo) > 0 Then
                                                                  '              NroCargo = Rstcob4.Fields("cobnm1")
                                                                  '              End If
                                                                  '                  If ContPagCargo = Format(ContcargoINT, "000") Then
                                                                  '                  IntPag = False
                                                                  '                  'ContcargoINT = 0
                                                                  '                  End If
                                                                  '
                                                                  '
                                                                  '            End If
            
                                 IntTotalRecord = Sel_Int_Detalle_Liq_Sql(Rstcob4.Fields("cobnm1"), Rstcob4.Fields("cobNM2"))
            
                                 If IntTotalRecord > 0 Then
                                  'Rstcob6.MoveLast
                                  'Rstcob6.MoveFirst
                                  '''Determinar cuantas hojas se imprimen por Razon Social
                                      Cantdet = 0
                                     If IntTotalRecord > 20 Then
                                        Cantdet = Int(IntTotalRecord / 20)
                                       If IntTotalRecord - (Cantdet * 20) > 0 Then
                                        Cantdet = Cantdet + 1
                                       Else
                                        Cantdet = Cantdet
                                       End If
                                     Else
                                        Cantdet = 1
                                     End If
            
                            '           Set RstTotal = Sel_Int_Totales_Liq(Rstcob4.Fields("cobnm1"), Rstcob4.Fields("cobNM2"))
                            '           If RstTotal.RecordCount > 0 Then
                            '           TotalFPrr = Trim$(RstTotal.Fields("cobto1"))
                            '           TotalInteMo = Trim$(RstTotal.Fields("cobto2"))
                            '           TotalAd = Trim$(RstTotal.Fields("cobto3"))
                            '           'Pagfinal = Trim(RstTotal.Fields("cobpag"))
                            '           'Pag = Pag & Pagfinal
                            '           End If
            
                                       For p = 1 To Cantdet
                                       Pagfinal = Format(Cantdet, "000")
                                       If Optbajopago.Value = True Then
                                       Set Rstcob6 = Sel_Int_Detalle_Pag_Liq_Sql(Rstcob4.Fields("cobnm1"), Format(p, "000"), Rstcob4.Fields("cobNM2"))
                                       Else
                                       Set Rstcob6 = Sel_Int_Detalle_Pag_Liq_Sql(Rstcob4.Fields("cobnm1"), Format(p, "000"), Rstcob4.Fields("cobNM2"))
                                       End If
                                          Rstcob6.MoveLast
                                          Rstcob6.MoveFirst
                                         If Rstcob6.RecordCount >= p * 20 Then
                                            For Y = 1 To 20
                                             cuspp(Y) = Trim$(Rstcob6.Fields("cobcus"))
                                             ApellNomb(Y) = Trim$(IIf(IsNull(Rstcob6.Fields("cobape")), "", Trim$(Rstcob6.Fields("cobape"))) & " / " & IIf(IsNull(Rstcob6.Fields("cobap1")), "", Trim$(Rstcob6.Fields("cobap1"))) & " / " & Trim$(Rstcob6.Fields("cobno6")) & " / " & IIf(IsNull(Rstcob6.Fields("cobno7")), "", Trim$(Rstcob6.Fields("cobno7"))))
                                             RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                                             FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                                             RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                                             TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                                             Rstcob6.MoveNext
                                            Next Y
                                         Else
                                            For Y = 1 To Rstcob6.RecordCount
                                             cuspp(Y) = Replace(Trim$(Rstcob6.Fields("cobcus")), "Ð", "Ñ")
                                             ApellNomb(Y) = Trim$(IIf(IsNull(Rstcob6.Fields("cobape")), "", Trim$(Rstcob6.Fields("cobape"))) & " / " & IIf(IsNull(Rstcob6.Fields("cobap1")), "", Trim$(Rstcob6.Fields("cobap1"))) & " / " & Trim$(Rstcob6.Fields("cobno6")) & " / " & IIf(IsNull(Rstcob6.Fields("cobno7")), "", Trim$(Rstcob6.Fields("cobno7"))))
                                             RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                                             FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                                             RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                                             TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                                             Rstcob6.MoveNext
                                            Next Y
                                                For h = Rstcob6.RecordCount + 1 To 20
                                                 cuspp(h) = ""
                                                 ApellNomb(h) = ""
                                                 RemAseg(h) = ""
                                                 FondoPens(h) = ""
                                                 RetenRetr(h) = ""
                                                 TotalFonPens(h) = ""
                                                 'cuspp(2 to 20)=  "1"
                                                Next h
                                            'End If
                                         End If
                                       ''''''''''''''''''
            
                                       PAGINA = Pag & Format(p, "000")
                                       CodBarra = "*" & Trim$(Rstcob4.Fields("cobnm2")) & Format(p, "000") & "*"
                                       If Format(p, "000") = Pagfinal Then
                                       Set RstTotal = Sel_Int_Totales_Liq_SQL(Rstcob4.Fields("cobnm1"), Rstcob4.Fields("cobNM2"))
                                            If RstTotal.RecordCount > 0 Then
                                            TotalFPrr = Trim$(RstTotal.Fields("cobto1"))
                                            TotalInteMo = Trim$(RstTotal.Fields("cobto2"))
                                            TotalAd = Trim$(RstTotal.Fields("cobto3"))
                                            End If
                                       Else
                                       TotalFPrr = "" 'Trim$(RstTotal.Fields("cobto1"))
                                       TotalInteMo = "" 'Trim$(RstTotal.Fields("cobto2"))
                                       TotalAd = "" 'Trim$(RstTotal.Fields("cobto3"))
            
            
                                       End If
                                       ''''''CREAR ARCHIVO TEXTO LIQUIDACION PREVIA INTEGRA
                                       '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                               ''''Formatear campos de Archivo texto
                                               ContcargoINT = Format(p, "000")
                                               ContPagCargo = Format(Cantdet, "000")
                                               ContPagCargoSe = Format(ContcargoINT, "000") & "/" & ContPagCargo
            
            
                                                 Formatear_Archivo_Texto_SQL
            
                                                ContIntegra = ContIntegra + 1
                                                
                                                If IntArchivos = 1 Then
                                                Print #3, TXT(1); vbTab; TXT(2); vbTab; TXT(3); vbTab; TXT(4); vbTab; TXT(5); vbTab; TXT(6); vbTab; TXT(7); vbTab; TXT(8); vbTab; TXT(9); vbTab; TXT(10) _
                                                          ; vbTab; TXT(11); vbTab; TXT(12); vbTab; TXT(13); vbTab; TXT(14); vbTab; TXT(15); vbTab; TXT(16); vbTab; TXT(17); vbTab; TXT(18); vbTab; TXT(19) _
                                                          ; vbTab; TXT(20); vbTab; TXT(21); vbTab; TXT(22); vbTab; TXT(23); vbTab; TXT(24); vbTab; TXT(25); vbTab; TXT(26); vbTab; TXT(27); vbTab; TXT(28) _
                                                          ; vbTab; TXT(29); vbTab; TXT(30); vbTab; TXT(31); vbTab; TXT(32); vbTab; TXT(33); vbTab; TXT(34); vbTab; TXT(35); vbTab; TXT(36); vbTab; TXT(37); vbTab; TXT(38) _
                                                          ; vbTab; TXT(39); vbTab; TXT(40); vbTab; TXT(41); vbTab; TXT(42); vbTab; TXT(43); vbTab; TXT(44); vbTab; TXT(45); vbTab; TXT(46) _
                                                          ; vbTab; TXT(47); vbTab; TXT(48); vbTab; TXT(49); vbTab; TXT(50); vbTab; TXT(51); vbTab; TXT(52); vbTab; TXT(53); vbTab; TXT(54); vbTab; TXT(55); vbTab; TXT(56); vbTab; TXT(57) _
                                                          ; vbTab; TXT(58); vbTab; TXT(59); vbTab; TXT(60); vbTab; TXT(61); vbTab; TXT(62); vbTab; TXT(63); vbTab; TXT(64); vbTab; TXT(65); vbTab; TXT(66); vbTab; TXT(67); vbTab; TXT(68) _
                                                          ; vbTab; TXT(69); vbTab; TXT(70); vbTab; TXT(71); vbTab; TXT(72); vbTab; TXT(73); vbTab; TXT(74); vbTab; TXT(75); vbTab; TXT(76); vbTab; TXT(77); vbTab; TXT(78); vbTab; TXT(79) _
                                                          ; vbTab; TXT(80); vbTab; TXT(81); vbTab; TXT(82); vbTab; TXT(83); vbTab; TXT(84); vbTab; TXT(85); vbTab; TXT(86); vbTab; TXT(87); vbTab; TXT(88); vbTab; TXT(89); vbTab; TXT(90) _
                                                          ; vbTab; TXT(91); vbTab; TXT(92); vbTab; TXT(93); vbTab; TXT(94); vbTab; TXT(95); vbTab; TXT(96); vbTab; TXT(97); vbTab; TXT(98); vbTab; TXT(99); vbTab; TXT(100); vbTab; TXT(101) _
                                                          ; vbTab; TXT(102); vbTab; TXT(103); vbTab; TXT(104); vbTab; TXT(105); vbTab; TXT(106); vbTab; TXT(107); vbTab; TXT(108); vbTab; TXT(109); vbTab; TXT(110); vbTab; TXT(111) _
                                                          ; vbTab; TXT(112); vbTab; TXT(113); vbTab; TXT(114); vbTab; TXT(115); vbTab; TXT(116); vbTab; TXT(117); vbTab; TXT(118); vbTab; TXT(119); vbTab; TXT(120); vbTab; TXT(121) _
                                                          ; vbTab; TXT(122); vbTab; TXT(123); vbTab; TXT(124); vbTab; TXT(125); vbTab; TXT(126); vbTab; TXT(127); vbTab; TXT(128); vbTab; TXT(129); vbTab; TXT(130); vbTab; TXT(131); vbTab; TXT(132) _
                                                          ; vbTab; TXT(133); vbTab; TXT(134); vbTab; TXT(135); vbTab; TXT(136); vbTab; TXT(137); vbTab; TXT(138); vbTab; TXT(139); vbTab; TXT(140); vbTab; TXT(141); vbTab; TXT(142)
                                                Else
                                                  If IntArchivos = 2 Then
                                                      Print #5, TXT(1); vbTab; TXT(2); vbTab; TXT(3); vbTab; TXT(4); vbTab; TXT(5); vbTab; TXT(6); vbTab; TXT(7); vbTab; TXT(8); vbTab; TXT(9); vbTab; TXT(10) _
                                                          ; vbTab; TXT(11); vbTab; TXT(12); vbTab; TXT(13); vbTab; TXT(14); vbTab; TXT(15); vbTab; TXT(16); vbTab; TXT(17); vbTab; TXT(18); vbTab; TXT(19) _
                                                          ; vbTab; TXT(20); vbTab; TXT(21); vbTab; TXT(22); vbTab; TXT(23); vbTab; TXT(24); vbTab; TXT(25); vbTab; TXT(26); vbTab; TXT(27); vbTab; TXT(28) _
                                                          ; vbTab; TXT(29); vbTab; TXT(30); vbTab; TXT(31); vbTab; TXT(32); vbTab; TXT(33); vbTab; TXT(34); vbTab; TXT(35); vbTab; TXT(36); vbTab; TXT(37); vbTab; TXT(38) _
                                                          ; vbTab; TXT(39); vbTab; TXT(40); vbTab; TXT(41); vbTab; TXT(42); vbTab; TXT(43); vbTab; TXT(44); vbTab; TXT(45); vbTab; TXT(46) _
                                                          ; vbTab; TXT(47); vbTab; TXT(48); vbTab; TXT(49); vbTab; TXT(50); vbTab; TXT(51); vbTab; TXT(52); vbTab; TXT(53); vbTab; TXT(54); vbTab; TXT(55); vbTab; TXT(56); vbTab; TXT(57) _
                                                          ; vbTab; TXT(58); vbTab; TXT(59); vbTab; TXT(60); vbTab; TXT(61); vbTab; TXT(62); vbTab; TXT(63); vbTab; TXT(64); vbTab; TXT(65); vbTab; TXT(66); vbTab; TXT(67); vbTab; TXT(68) _
                                                          ; vbTab; TXT(69); vbTab; TXT(70); vbTab; TXT(71); vbTab; TXT(72); vbTab; TXT(73); vbTab; TXT(74); vbTab; TXT(75); vbTab; TXT(76); vbTab; TXT(77); vbTab; TXT(78); vbTab; TXT(79) _
                                                          ; vbTab; TXT(80); vbTab; TXT(81); vbTab; TXT(82); vbTab; TXT(83); vbTab; TXT(84); vbTab; TXT(85); vbTab; TXT(86); vbTab; TXT(87); vbTab; TXT(88); vbTab; TXT(89); vbTab; TXT(90) _
                                                          ; vbTab; TXT(91); vbTab; TXT(92); vbTab; TXT(93); vbTab; TXT(94); vbTab; TXT(95); vbTab; TXT(96); vbTab; TXT(97); vbTab; TXT(98); vbTab; TXT(99); vbTab; TXT(100); vbTab; TXT(101) _
                                                          ; vbTab; TXT(102); vbTab; TXT(103); vbTab; TXT(104); vbTab; TXT(105); vbTab; TXT(106); vbTab; TXT(107); vbTab; TXT(108); vbTab; TXT(109); vbTab; TXT(110); vbTab; TXT(111) _
                                                          ; vbTab; TXT(112); vbTab; TXT(113); vbTab; TXT(114); vbTab; TXT(115); vbTab; TXT(116); vbTab; TXT(117); vbTab; TXT(118); vbTab; TXT(119); vbTab; TXT(120); vbTab; TXT(121) _
                                                          ; vbTab; TXT(122); vbTab; TXT(123); vbTab; TXT(124); vbTab; TXT(125); vbTab; TXT(126); vbTab; TXT(127); vbTab; TXT(128); vbTab; TXT(129); vbTab; TXT(130); vbTab; TXT(131); vbTab; TXT(132) _
                                                          ; vbTab; TXT(133); vbTab; TXT(134); vbTab; TXT(135); vbTab; TXT(136); vbTab; TXT(137); vbTab; TXT(138); vbTab; TXT(139); vbTab; TXT(140); vbTab; TXT(141); vbTab; TXT(142)
                                                
                                                  End If
                                                
                                                End If
            
                                       '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                       With CmdDatos
                                       .CommandText = "sp_InsertarDatos"
                                       .CommandType = adCmdStoredProc
                                       .Parameters.Refresh
                                       .ActiveConnection = CnIntegra.ConnectionString
                                       .Parameters.Append .CreateParameter("NROLIQUI", adVarChar, adParamInput, 100, TXT(1))
                                       .Parameters.Append .CreateParameter("CAMPO2", adVarChar, adParamInput, 100, TXT(2))
                                       .Parameters.Append .CreateParameter("NROLQP", adVarChar, adParamInput, 100, Trim$(Rstcob4.Fields("cobnm1")))
                                       .Parameters.Append .CreateParameter("CAMPO4", adVarChar, adParamInput, 100, TXT(4))
                                       .Parameters.Append .CreateParameter("CAMPO5", adVarChar, adParamInput, 100, TXT(5))
                                       .Parameters.Append .CreateParameter("CAMPO6", adVarChar, adParamInput, 100, TXT(6))
                                       .Parameters.Append .CreateParameter("CAMPO7", adVarChar, adParamInput, 100, TXT(7))
                                       .Parameters.Append .CreateParameter("CAMPO8", adVarChar, adParamInput, 100, TXT(8))
                                       .Parameters.Append .CreateParameter("CAMPO9", adVarChar, adParamInput, 100, TXT(9))
                                       .Parameters.Append .CreateParameter("CAMPO10", adVarChar, adParamInput, 100, TXT(10))
                                       .Parameters.Append .CreateParameter("CAMPO11", adVarChar, adParamInput, 100, TXT(11))
                                       .Parameters.Append .CreateParameter("CAMPO12", adVarChar, adParamInput, 100, TXT(12))
                                       .Parameters.Append .CreateParameter("CAMPO13", adVarChar, adParamInput, 100, TXT(13))
                                       .Parameters.Append .CreateParameter("CAMPO14", adVarChar, adParamInput, 100, TXT(14))
                                       .Parameters.Append .CreateParameter("CAMPO15", adVarChar, adParamInput, 100, TXT(15))
                                       .Parameters.Append .CreateParameter("CAMPO16", adVarChar, adParamInput, 100, TXT(16))
                                       .Parameters.Append .CreateParameter("CAMPO17", adVarChar, adParamInput, 100, TXT(17))
                                       .Parameters.Append .CreateParameter("CAMPO18", adVarChar, adParamInput, 100, TXT(18))
                                       .Parameters.Append .CreateParameter("CAMPO19", adVarChar, adParamInput, 100, TXT(19))
                                       .Parameters.Append .CreateParameter("CUSPP1", adVarChar, adParamInput, 100, TXT(20))
                                       .Parameters.Append .CreateParameter("NOMAPE1", adVarChar, adParamInput, 100, TXT(21))
                                       .Parameters.Append .CreateParameter("REMU1", adVarChar, adParamInput, 100, TXT(22))
                                       .Parameters.Append .CreateParameter("FONDO1", adVarChar, adParamInput, 100, TXT(23))
                                       .Parameters.Append .CreateParameter("RETEN1", adVarChar, adParamInput, 100, TXT(24))
                                       .Parameters.Append .CreateParameter("TOTFOND1", adVarChar, adParamInput, 100, TXT(25))
                                       .Parameters.Append .CreateParameter("CUSPP2", adVarChar, adParamInput, 100, TXT(26))
                                       .Parameters.Append .CreateParameter("NOMAPE2", adVarChar, adParamInput, 100, TXT(27))
                                       .Parameters.Append .CreateParameter("REMU2", adVarChar, adParamInput, 100, TXT(28))
                                       .Parameters.Append .CreateParameter("FONDO2", adVarChar, adParamInput, 100, TXT(29))
                                       .Parameters.Append .CreateParameter("RETEN2", adVarChar, adParamInput, 100, TXT(30))
                                       .Parameters.Append .CreateParameter("TOTFOND2", adVarChar, adParamInput, 100, TXT(31))
                                       .Parameters.Append .CreateParameter("CUSPP3", adVarChar, adParamInput, 100, TXT(32))
                                       .Parameters.Append .CreateParameter("NOMAPE3", adVarChar, adParamInput, 100, TXT(33))
                                       .Parameters.Append .CreateParameter("REMU3", adVarChar, adParamInput, 100, TXT(34))
                                       .Parameters.Append .CreateParameter("FONDO3", adVarChar, adParamInput, 100, TXT(35))
                                       .Parameters.Append .CreateParameter("RETEN3", adVarChar, adParamInput, 100, TXT(36))
                                       .Parameters.Append .CreateParameter("TOTFOND3", adVarChar, adParamInput, 100, TXT(37))
                                       .Parameters.Append .CreateParameter("CUSPP4", adVarChar, adParamInput, 100, TXT(38))
                                       .Parameters.Append .CreateParameter("NOMAPE4", adVarChar, adParamInput, 100, TXT(39))
                                       .Parameters.Append .CreateParameter("REMU4", adVarChar, adParamInput, 100, TXT(40))
                                       .Parameters.Append .CreateParameter("FONDO4", adVarChar, adParamInput, 100, TXT(41))
                                       .Parameters.Append .CreateParameter("RETEN4", adVarChar, adParamInput, 100, TXT(42))
                                       .Parameters.Append .CreateParameter("TOTFOND4", adVarChar, adParamInput, 100, TXT(43))
                                       .Parameters.Append .CreateParameter("CUSPP5", adVarChar, adParamInput, 100, TXT(44))
                                       .Parameters.Append .CreateParameter("NOMAPE5", adVarChar, adParamInput, 100, TXT(45))
                                       .Parameters.Append .CreateParameter("REMU5", adVarChar, adParamInput, 100, TXT(46))
                                       .Parameters.Append .CreateParameter("FONDO5", adVarChar, adParamInput, 100, TXT(47))
                                       .Parameters.Append .CreateParameter("RETEN5", adVarChar, adParamInput, 100, TXT(48))
                                       .Parameters.Append .CreateParameter("TOTFOND5", adVarChar, adParamInput, 100, TXT(49))
                                       .Parameters.Append .CreateParameter("CUSPP6", adVarChar, adParamInput, 100, TXT(50))
                                       .Parameters.Append .CreateParameter("NOMAPE6", adVarChar, adParamInput, 100, TXT(51))
                                       .Parameters.Append .CreateParameter("REMU6", adVarChar, adParamInput, 100, TXT(52))
                                       .Parameters.Append .CreateParameter("FONDO6", adVarChar, adParamInput, 100, TXT(53))
                                       .Parameters.Append .CreateParameter("RETEN6", adVarChar, adParamInput, 100, TXT(54))
                                       .Parameters.Append .CreateParameter("TOTFOND6", adVarChar, adParamInput, 100, TXT(55))
                                       .Parameters.Append .CreateParameter("CUSPP7", adVarChar, adParamInput, 100, TXT(56))
                                       .Parameters.Append .CreateParameter("NOMAPE7", adVarChar, adParamInput, 100, TXT(57))
                                       .Parameters.Append .CreateParameter("REMU7", adVarChar, adParamInput, 100, TXT(58))
                                       .Parameters.Append .CreateParameter("FONDO7", adVarChar, adParamInput, 100, TXT(59))
                                       .Parameters.Append .CreateParameter("RETEN7", adVarChar, adParamInput, 100, TXT(60))
                                       .Parameters.Append .CreateParameter("TOTFOND7", adVarChar, adParamInput, 100, TXT(61))
                                       .Parameters.Append .CreateParameter("CUSPP8", adVarChar, adParamInput, 100, TXT(62))
                                       .Parameters.Append .CreateParameter("NOMAPE8", adVarChar, adParamInput, 100, TXT(63))
                                       .Parameters.Append .CreateParameter("REMU8", adVarChar, adParamInput, 100, TXT(64))
                                       .Parameters.Append .CreateParameter("FONDO8", adVarChar, adParamInput, 100, TXT(65))
                                       .Parameters.Append .CreateParameter("RETEN8", adVarChar, adParamInput, 100, TXT(66))
                                       .Parameters.Append .CreateParameter("TOTFOND8", adVarChar, adParamInput, 100, TXT(67))
                                       .Parameters.Append .CreateParameter("CUSPP9", adVarChar, adParamInput, 100, TXT(68))
                                       .Parameters.Append .CreateParameter("NOMAPE9", adVarChar, adParamInput, 100, TXT(69))
                                       .Parameters.Append .CreateParameter("REMU9", adVarChar, adParamInput, 100, TXT(70))
                                       .Parameters.Append .CreateParameter("FONDO9", adVarChar, adParamInput, 100, TXT(71))
                                       .Parameters.Append .CreateParameter("RETEN9", adVarChar, adParamInput, 100, TXT(72))
                                       .Parameters.Append .CreateParameter("TOTFOND9", adVarChar, adParamInput, 100, TXT(73))
                                       .Parameters.Append .CreateParameter("CUSPP10", adVarChar, adParamInput, 100, TXT(74))
                                       .Parameters.Append .CreateParameter("NOMAPE10", adVarChar, adParamInput, 100, TXT(75))
                                       .Parameters.Append .CreateParameter("REMU10", adVarChar, adParamInput, 100, TXT(76))
                                       .Parameters.Append .CreateParameter("FONDO10", adVarChar, adParamInput, 100, TXT(77))
                                       .Parameters.Append .CreateParameter("RETEN10", adVarChar, adParamInput, 100, TXT(78))
                                       .Parameters.Append .CreateParameter("TOTFOND10", adVarChar, adParamInput, 100, TXT(79))
                                       .Parameters.Append .CreateParameter("CUSPP11", adVarChar, adParamInput, 100, TXT(80))
                                       .Parameters.Append .CreateParameter("NOMAPE11", adVarChar, adParamInput, 100, TXT(81))
                                       .Parameters.Append .CreateParameter("REMU11", adVarChar, adParamInput, 100, TXT(82))
                                       .Parameters.Append .CreateParameter("FONDO11", adVarChar, adParamInput, 100, TXT(83))
                                       .Parameters.Append .CreateParameter("RETEN11", adVarChar, adParamInput, 100, TXT(84))
                                       .Parameters.Append .CreateParameter("TOTFOND11", adVarChar, adParamInput, 100, TXT(85))
                                       .Parameters.Append .CreateParameter("CUSPP12", adVarChar, adParamInput, 100, TXT(86))
                                       .Parameters.Append .CreateParameter("NOMAPE12", adVarChar, adParamInput, 100, TXT(87))
                                       .Parameters.Append .CreateParameter("REMU12", adVarChar, adParamInput, 100, TXT(88))
                                       .Parameters.Append .CreateParameter("FONDO12", adVarChar, adParamInput, 100, TXT(89))
                                       .Parameters.Append .CreateParameter("RETEN12", adVarChar, adParamInput, 100, TXT(90))
                                       .Parameters.Append .CreateParameter("TOTFOND12", adVarChar, adParamInput, 100, TXT(91))
                                       .Parameters.Append .CreateParameter("CUSPP13", adVarChar, adParamInput, 100, TXT(92))
                                       .Parameters.Append .CreateParameter("NOMAPE13", adVarChar, adParamInput, 100, TXT(93))
                                       .Parameters.Append .CreateParameter("REMU13", adVarChar, adParamInput, 100, TXT(94))
                                       .Parameters.Append .CreateParameter("FONDO13", adVarChar, adParamInput, 100, TXT(95))
                                       .Parameters.Append .CreateParameter("RETEN13", adVarChar, adParamInput, 100, TXT(96))
                                       .Parameters.Append .CreateParameter("TOTFOND13", adVarChar, adParamInput, 100, TXT(97))
                                       .Parameters.Append .CreateParameter("CUSPP14", adVarChar, adParamInput, 100, TXT(98))
                                       .Parameters.Append .CreateParameter("NOMAPE14", adVarChar, adParamInput, 100, TXT(99))
                                       .Parameters.Append .CreateParameter("REMU14", adVarChar, adParamInput, 100, TXT(100))
                                       .Parameters.Append .CreateParameter("FONDO14", adVarChar, adParamInput, 100, TXT(101))
                                       .Parameters.Append .CreateParameter("RETEN14", adVarChar, adParamInput, 100, TXT(102))
                                       .Parameters.Append .CreateParameter("TOTFOND14", adVarChar, adParamInput, 100, TXT(103))
                                       .Parameters.Append .CreateParameter("CUSPP15", adVarChar, adParamInput, 100, TXT(104))
                                       .Parameters.Append .CreateParameter("NOMAPE15", adVarChar, adParamInput, 100, TXT(105))
                                       .Parameters.Append .CreateParameter("REMU15", adVarChar, adParamInput, 100, TXT(106))
                                       .Parameters.Append .CreateParameter("FONDO15", adVarChar, adParamInput, 100, TXT(107))
                                       .Parameters.Append .CreateParameter("RETEN15", adVarChar, adParamInput, 100, TXT(108))
                                       .Parameters.Append .CreateParameter("TOTFOND15", adVarChar, adParamInput, 100, TXT(109))
                                       .Parameters.Append .CreateParameter("CUSPP16", adVarChar, adParamInput, 100, TXT(110))
                                       .Parameters.Append .CreateParameter("NOMAPE16", adVarChar, adParamInput, 100, TXT(111))
                                       .Parameters.Append .CreateParameter("REMU16", adVarChar, adParamInput, 100, TXT(112))
                                       .Parameters.Append .CreateParameter("FONDO16", adVarChar, adParamInput, 100, TXT(113))
                                       .Parameters.Append .CreateParameter("RETEN16", adVarChar, adParamInput, 100, TXT(114))
                                       .Parameters.Append .CreateParameter("TOTFOND16", adVarChar, adParamInput, 100, TXT(115))
                                       .Parameters.Append .CreateParameter("CUSPP17", adVarChar, adParamInput, 100, TXT(116))
                                       .Parameters.Append .CreateParameter("NOMAPE17", adVarChar, adParamInput, 100, TXT(117))
                                       .Parameters.Append .CreateParameter("REMU17", adVarChar, adParamInput, 100, TXT(118))
                                       .Parameters.Append .CreateParameter("FONDO17", adVarChar, adParamInput, 100, TXT(119))
                                       .Parameters.Append .CreateParameter("RETEN17", adVarChar, adParamInput, 100, TXT(120))
                                       .Parameters.Append .CreateParameter("TOTFOND17", adVarChar, adParamInput, 100, TXT(121))
                                       .Parameters.Append .CreateParameter("CUSPP18", adVarChar, adParamInput, 100, TXT(122))
                                       .Parameters.Append .CreateParameter("NOMAPE18", adVarChar, adParamInput, 100, TXT(123))
                                       .Parameters.Append .CreateParameter("REMU18", adVarChar, adParamInput, 100, TXT(124))
                                       .Parameters.Append .CreateParameter("FONDO18", adVarChar, adParamInput, 100, TXT(125))
                                       .Parameters.Append .CreateParameter("RETEN18", adVarChar, adParamInput, 100, TXT(126))
                                       .Parameters.Append .CreateParameter("TOTFOND18", adVarChar, adParamInput, 100, TXT(127))
                                       .Parameters.Append .CreateParameter("CUSPP19", adVarChar, adParamInput, 100, TXT(128))
                                       .Parameters.Append .CreateParameter("NOMAPE19", adVarChar, adParamInput, 100, TXT(129))
                                       .Parameters.Append .CreateParameter("REMU19", adVarChar, adParamInput, 100, TXT(130))
                                       .Parameters.Append .CreateParameter("FONDO19", adVarChar, adParamInput, 100, TXT(131))
                                       .Parameters.Append .CreateParameter("RETEN19", adVarChar, adParamInput, 100, TXT(132))
                                       .Parameters.Append .CreateParameter("TOTFOND19", adVarChar, adParamInput, 100, TXT(133))
                                       .Parameters.Append .CreateParameter("CUSPP20", adVarChar, adParamInput, 100, TXT(134))
                                       .Parameters.Append .CreateParameter("NOMAPE20", adVarChar, adParamInput, 100, TXT(135))
                                       .Parameters.Append .CreateParameter("REMU20", adVarChar, adParamInput, 100, TXT(136))
                                       .Parameters.Append .CreateParameter("FONDO20", adVarChar, adParamInput, 100, TXT(137))
                                       .Parameters.Append .CreateParameter("RETEN20", adVarChar, adParamInput, 100, TXT(138))
                                       .Parameters.Append .CreateParameter("TOTFOND20", adVarChar, adParamInput, 100, TXT(139))
                                       .Parameters.Append .CreateParameter("BARRA", adVarChar, adParamInput, 100, TXT(140))
                                       .Parameters.Append .CreateParameter("CAMPO141", adVarChar, adParamInput, 100, TXT(141))
                                       .Parameters.Append .CreateParameter("CAMPO142", adVarChar, adParamInput, 100, TXT(142))
                                       .Parameters.Append .CreateParameter("ESTADO", adVarChar, adParamInput, 100, xxEstado)
                                       .Parameters.Append .CreateParameter("NROLIQPAG", adVarChar, adParamInput, 100, Trim$(Rstcob4.Fields("cobnm2")) & Format(p, "000"))
                                       .Parameters.Append .CreateParameter("PROCESADO", adVarChar, adParamInput, 1, "0")
                                       .Execute
                                       End With
                                       Set CmdDatos = Nothing
            '                           strCadenaSQL = "EXEC sp_InsertarDatos '" & (TXT(1)) & "','" & TXT(2) & "','" & Trim$(Rstcob4.Fields("cobnm1")) & "','" & TXT(4) & "','" & TXT(5) & "','" & Replace(TXT(6), "'", "`") & "','" & TXT(7) & "','" & TXT(8) & "','" & Replace(TXT(9), "'", "`") & "','" & TXT(10) _
            '                                              & "','" & TXT(11) & "','" & TXT(12) & "','" & TXT(13) & "','" & TXT(14) & "','" & TXT(15) & "','" & TXT(16) & "','" & TXT(17) & "','" & TXT(18) & "','" & TXT(19) _
            '                                              & "','" & TXT(20) & "','" & Replace(TXT(21), "'", "`") & "','" & TXT(22) & "','" & TXT(23) & "','" & TXT(24) & "','" & TXT(25) & "','" & TXT(26) & "','" & Replace(TXT(27), "'", "`") & "','" & TXT(28) _
            '                                              & "','" & TXT(29) & "','" & TXT(30) & "','" & TXT(31) & "','" & TXT(32) & "','" & TXT(33) & "','" & TXT(34) & "','" & TXT(35) & "','" & TXT(36) & "','" & TXT(37) & "','" & TXT(38) _
            '                                              & "','" & TXT(39) & "','" & TXT(40) & "','" & TXT(41) & "','" & TXT(42) & "','" & TXT(43) & "','" & TXT(44) & "','" & TXT(45) & "','" & TXT(46) _
            '                                              & "','" & TXT(47) & "','" & TXT(48) & "','" & TXT(49) & "','" & TXT(50) & "','" & TXT(51) & "','" & TXT(52) & "','" & TXT(53) & "','" & TXT(54) & "','" & TXT(55) & "','" & TXT(56) & "','" & TXT(57) _
            '                                              & "','" & TXT(58) & "','" & TXT(59) & "','" & TXT(60) & "','" & TXT(61) & "','" & TXT(62) & "','" & TXT(63) & "','" & TXT(64) & "','" & TXT(65) & "','" & TXT(66) & "','" & TXT(67) & "','" & TXT(68) _
            '                                              & "','" & TXT(69) & "','" & TXT(70) & "','" & TXT(71) & "','" & TXT(72) & "','" & TXT(73) & "','" & TXT(74) & "','" & TXT(75) & "','" & TXT(76) & "','" & TXT(77) & "','" & TXT(78) & "','" & TXT(79) _
            '                                              & "','" & TXT(80) & "','" & TXT(81) & "','" & TXT(82) & "','" & TXT(83) & "','" & TXT(84) & "','" & TXT(85) & "','" & TXT(86) & "','" & TXT(87) & "','" & TXT(88) & "','" & TXT(89) & "','" & TXT(90) _
            '                                              & "','" & TXT(91) & "','" & TXT(92) & "','" & Replace(TXT(93), "'", "`") & "','" & TXT(94) & "','" & TXT(95) & "','" & TXT(96) & "','" & TXT(97) & "','" & TXT(98) & "','" & TXT(99) & "','" & TXT(100) & "','" & TXT(101) _
            '                                              & "','" & TXT(102) & "','" & TXT(103) & "','" & TXT(104) & "','" & TXT(105) & "','" & TXT(106) & "','" & TXT(107) & "','" & TXT(108) & "','" & TXT(109) & "','" & TXT(110) & "','" & TXT(111) _
            '                                              & "','" & TXT(112) & "','" & TXT(113) & "','" & TXT(114) & "','" & TXT(115) & "','" & TXT(116) & "','" & TXT(117) & "','" & TXT(118) & "','" & TXT(119) & "','" & TXT(120) & "','" & TXT(121) _
            '                                              & "','" & TXT(122) & "','" & TXT(123) & "','" & TXT(124) & "','" & TXT(125) & "','" & TXT(126) & "','" & TXT(127) & "','" & TXT(128) & "','" & TXT(129) & "','" & TXT(130) & "','" & TXT(131) & "','" & TXT(132) _
            '                                              & "','" & TXT(133) & "','" & TXT(134) & "','" & TXT(135) & "','" & TXT(136) & "','" & TXT(137) & "','" & TXT(138) & "','" & TXT(139) & "','" & TXT(140) & "','" & TXT(141) & "','" & TXT(142) & "','" & xxEstado & "','" & Trim$(Rstcob4.Fields("cobnm2")) & Format(p, "000") & "','0'"
            '
            '                           CnIntegra.Execute strCadenaSQL
                                       
                                       Next p
                                  ''''
                                 End If
                              ''''
            
             'End If
            
                              Rstcob4.MoveNext
                              'Next X
            
             Wend
             
       End If
        If IntArchivos = 1 Then
        Close #3
        Else
           If IntArchivos = 2 Then
           Close #5
           End If
        End If
             
Next IntArchivos
             
     



     'Set Rstcob4 = Sel_Int_CabezeraSql_Carta
            'If Rstcob4.RecordCount > 0 Then
            
     For IntArchivos = 1 To 2
     
            If IntArchivos = 1 Then
            Set Rstcob4 = Sel_Int_CabezeraSql_Carta_Lima_Callao
            Else
               If IntArchivos = 2 Then
               Set Rstcob4 = Sel_Int_CabezeraSql_Carta_Provincia
               End If
            End If
            
            
            If Not Rstcob4.EOF Then
            Rstcob4.MoveLast
            Rstcob4.MoveFirst
            If IntArchivos = 1 Then
            Open RutaArchivos & "\Carta_Lima_callao.txt" For Output As #8
            Print #8, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
                                          ; vbTab; "CAMPO11"; vbTab; "CAMPO12"; vbTab; "CAMPO13"; vbTab; "CAMPO14"; vbTab; "CAMPO15"; vbTab; "CAMPO16"
            Else
               If IntArchivos = 2 Then
               Open RutaArchivos & "\Carta_Provincias.txt" For Output As #9
               Print #9, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
                                          ; vbTab; "CAMPO11"; vbTab; "CAMPO12"; vbTab; "CAMPO13"; vbTab; "CAMPO14"; vbTab; "CAMPO15"; vbTab; "CAMPO16"
               End If
            End If
            ContCarta = 0
             While Not Rstcob4.EOF
              'If (Rstcob4.Fields("cobnmE") <> "001383724" And Rstcob4.Fields("cobnmE") <> "001383724") Then
                             NroLiqC = Trim$(Trim$(Rstcob4.Fields("cobti5")) & "-" & Trim$(Rstcob4.Fields("cobnme")))
                             FechaLiqC = Trim$(Rstcob4.Fields("cobloc")) & ", " & Trim$(Rstcob4.Fields("cobfe5"))
                             RazonC = Trim$(IIf(IsNull(Rstcob4.Fields("cobra1")), "", Rstcob4.Fields("cobra1")))
                             DirectC = Trim$(Rstcob4.Fields("cobdir"))
                             DistritoC = Trim$(Rstcob4.Fields("cobdi1"))
                             DepartamentoC = Trim$(Rstcob4.Fields("cobno1")) & "/" & Trim$(Rstcob4.Fields("cobno2"))
                             RucC = Trim$(IIf(IsNull(Rstcob4.Fields("cobti6")), "", Rstcob4.Fields("cobti6"))) & " " & Trim$(IIf(IsNull(Rstcob4.Fields("cobnr1")), " ", Rstcob4.Fields("cobnr1")))
                             Atencion = Trim$(IIf(IsNull(Rstcob4.Fields("cobate")), "", Rstcob4.Fields("cobate")))
                             Referencia = Trim$(IIf(IsNull(Rstcob4.Fields("cobref")), "", Rstcob4.Fields("cobref")))
                             SraC = Trim$(IIf(IsNull(Rstcob4.Fields("cobqui")), "", Rstcob4.Fields("cobqui")))
                             'Sel_Int_Liq_Carta (Rstcob4.Fields("cobnme"))
                             'CodBarraC = "*" & Trim$(Rstcob4.Fields("cobnme")) & "*"
                             BarraCarta = Sel_Int_Liq_Carta(Rstcob4.Fields("cobnme"))
                             If BarraCarta = "" Then
                              CodBarraC = "*" & Trim$(Rstcob4.Fields("cobnme")) & "*"
                             Else
                              CodBarraC = "*" & BarraCarta & "001*"
                             End If
                             Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("cobadj")), "", Rstcob4.Fields("cobadj"))) & ": " & Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin")))

                             Adjunto1 = Trim$(IIf(IsNull(Rstcob4.Fields("cobli3")), "", Rstcob4.Fields("cobli3")))

                             Adjunto2 = Trim$(IIf(IsNull(Rstcob4.Fields("cobli4")), "", Rstcob4.Fields("cobli4")))

                             Adjunto3 = Trim$(IIf(IsNull(Rstcob4.Fields("cobli5")), "", Rstcob4.Fields("cobli5")))

                             Adjunto4 = Trim$(IIf(IsNull(Rstcob4.Fields("cobli6")), "", Rstcob4.Fields("cobli6")))

                             Formatear_Archivo_Texto_Carta
                             ContCarta = ContCarta + 1
                             If IntArchivos = 1 Then
                             Print #8, TXTC(1); vbTab; TXTC(2); vbTab; TXTC(3); vbTab; TXTC(4); vbTab; TXTC(5); vbTab; TXTC(6); vbTab; TXTC(7); vbTab; TXTC(8); vbTab; TXTC(9); vbTab; TXTC(10) _
                                      ; vbTab; TXTC(11); vbTab; TXTC(12); vbTab; TXTC(13); vbTab; TXTC(14); vbTab; TXTC(15); vbTab; TXTC(16)
                             
                             Else
                                If IntArchivos = 2 Then
                                Print #9, TXTC(1); vbTab; TXTC(2); vbTab; TXTC(3); vbTab; TXTC(4); vbTab; TXTC(5); vbTab; TXTC(6); vbTab; TXTC(7); vbTab; TXTC(8); vbTab; TXTC(9); vbTab; TXTC(10) _
                                      ; vbTab; TXTC(11); vbTab; TXTC(12); vbTab; TXTC(13); vbTab; TXTC(14); vbTab; TXTC(15); vbTab; TXTC(16)
                                End If
                             End If

                    Rstcob4.MoveNext
             Wend
            
            
            Close #8
            Close #9
            End If
    Next IntArchivos
   Flow_Frmarchivo.MousePointer = vbNormal
   LblIntegra.Caption = "Se procesaron la cantidad de: " & ContIntegra & " Formatos de Liquidación"
   lblcarta.Caption = "Se procesaron la cantidad de: " & ContCarta & " Formatos de Carta Liquidación Cargo"
   MsgBox "Proceso concluido Satisfactoriamente", vbInformation, "Horizonte"
   
'End If
'Close #3
Close #4

End Sub

Private Sub cmdprotocolo_Click()
On Error GoTo PRO
If Trim$(flxls.FileName) <> "" Then
        xhoja = UCase(InputBox("Ingresar el nombre de la Hoja del Archivo en excel :", "Archivo Onell"))
        RutaArchivoExcel = Dirxls.Path & "\" & flxls.FileName
        RutaDirExcel = Dirxls.Path
        
        
        Abrir_Excell
        
 Else
 MsgBox "Debe Seleccionar el Archivo", vbExclamation, "Archivo Onell"
 End If
PRO:

End Sub

Private Sub CmdProceso1_Click()

End Sub

Private Sub Cmdprocesotxt_Click()
Dim MsgArchivo As String
RUTA = Dirxls & "\" & flxls
If Dir(RUTA) = "" Then
  MsgArchivo = "No existe el archivo :" & RUTA
  MsgBox MsgArchivo, vbInformation, "Horizonte"
Else
  'Archivo_Importar RUTA
  Archivo_Importar_Enotria RUTA
  
End If
End Sub

Public Sub Archivo_Importar(StrArchivo As String)
Dim intCantidad As Double
Dim IntcantidadFormatos As Integer
Dim Intvariable As Integer
Dim intVar As Integer
Dim intPag As Integer
Dim Vartexto As String
Dim StrNroli_h  As String
lngFile = FreeFile
intCantidad = 0
Open StrArchivo For Input As #lngFile
Do Until EOF(lngFile)
   Line Input #lngFile, Vartexto
   intCantidad = intCantidad + 1
Loop
Close #lngFile
IntcantidadFormatos = (intCantidad + 1) / 25
LblHorizonte.Caption = "Se Importaran " & IntcantidadFormatos & " Formatos de Liquidacion Horizonte"

lngFile = FreeFile
Open StrArchivo For Input As #lngFile

intPag = 0
PrbBarra.Min = 1
PrbBarra.Max = IntcantidadFormatos
Do Until EOF(lngFile)
     For Intvariable = 1 To IntcantidadFormatos
               PrbBarra.Value = Intvariable
               Line Input #lngFile, Vartexto
               
               If StrNroliquidacion_h = Mid(Vartexto, 1, 11) Then
               intPag = intPag + 1
               Else
               intPag = 1
               End If
               
               'StrNroli_h = Mid(Vartexto, 1, 11)
               
               StrNroliquidacion_h = Mid(Vartexto, 1, 11)
               StrPeriodo_h = Mid(Vartexto, 12, 8)
               StrFecha_Emision_h = Mid(Vartexto, 20, 10)
               Line Input #lngFile, Vartexto
               StrRazon_Social_h = Mid(Vartexto, 1, 60)
               StrRuc_h = Mid(Vartexto, 61, 15)
               Line Input #lngFile, Vartexto
               StrDireccion_h = Mid(Vartexto, 1, 100)
               Line Input #lngFile, Vartexto
               StrDistrito_h = Mid(Vartexto, 1, 20)
               StrProvincia_h = Mid(Vartexto, 21, 15)
               StrDepartamento_h = Mid(Vartexto, 36, 15)
               For intVar = 1 To 20
                    Line Input #lngFile, Vartexto
                    cuspp_h(intVar) = Mid(Vartexto, 1, 12)
                    ApellPater_h(intVar) = Mid(Vartexto, 13, 20)
                    ApellMater_h(intVar) = Mid(Vartexto, 33, 20)
                    NombrePri_h(intVar) = Mid(Vartexto, 53, 10)
                    NombreSeg_h(intVar) = Mid(Vartexto, 63, 10)
                    RemAseg_h(intVar) = Mid(Vartexto, 73, 11)
                    FondoPens_h(intVar) = Mid(Vartexto, 84, 11)
                    RetenRetr_h(intVar) = Mid(Vartexto, 95, 11)
                    TotalRetension_h(intVar) = Mid(Vartexto, 95, 11)
                    TotalFonPens_h(intVar) = Mid(Vartexto, 106, 11)
               Next intVar
               Line Input #lngFile, Vartexto
               StrTotFondo_h = Mid(Vartexto, 1, 11)
               StrMoratorio_h = Mid(Vartexto, 12, 11)
               StrAdeudado_h = Mid(Vartexto, 23, 11)
               'strBarra_h = Mid(StrNroliquidacion_h, 3, 10) & intPag
               strBarra_h = Mid(StrNroliquidacion_h, 3, 10) & Format(intPag, "00000")
               strEstado_h = "0"
               strNroliqpag_h = "HO" & Mid(StrNroliquidacion_h, 3, 10) & Format(intPag, "00000")
               
               Insertar_Datos
     Next Intvariable
Exit Do
Loop
    MsgBox "Se Inportaron correctamente los Archivos", vbInformation, "Horizonte"
Close #lngFile
End Sub
Public Sub Archivo_Importar_Enotria(StrArchivo As String)
Dim intCantidad As Double
Dim IntcantidadFormatos As Integer
Dim Intvariable As Integer
Dim intVar As Integer
Dim intPag As Integer
Dim Vartexto As String
Dim StrNroli_h  As String
Dim strCantidad As String
Dim strCantidadReg As String
Dim TotalFonPens_h1 As String

Dim xdato As Recordset

lngFile = FreeFile
intCantidad = 0
LblHorizonte.Caption = "Se Importaran los Formatos de Liquidacion Horizonte"

lngFile = FreeFile
Open StrArchivo For Input As #lngFile

intPag = 0
'PrbBarra.Min = 1
'PrbBarra.Max = IntcantidadFormatos
Do Until EOF(lngFile)
     Line Input #lngFile, Vartexto
     strCantidad = Mid(Vartexto, 195, 7)
     strCantidad = Trim$(Mid(strCantidad, InStr(strCantidad, "/") + 1, Len(strCantidad)))
     strCantidadReg = Trim$(Mid(Vartexto, 202, 10))
     strCantidadReg1 = Trim$(Mid(Vartexto, 202, 10))
     
     
     intPag = Mid(Mid(Vartexto, 195, 7), 1, InStr(Mid(Vartexto, 195, 7), "/") - 1)
      'intCantidad = Mid(Vartexto, 195, 7)
      
     For Intvariable = 1 To strCantidad
               'PrbBarra.Value = Intvariable
               If Intvariable <> 1 Then
               Line Input #lngFile, Vartexto
               End If
               
               StrNroliquidacion_h = Mid(Vartexto, 54, 11)
               'StrPeriodo_h = Mid(Vartexto, 12, 8)
               StrPeriodo_h = "" 'Mid(Vartexto, 12, 8)
               intPag = Mid(Mid(Vartexto, 195, 7), 1, InStr(Mid(Vartexto, 195, 7), "/") - 1)
'               If StrNroliquidacion_h = "HO10C266842" Then
'               MsgBox "PRUEBA"
'               End If
               
               StrFecha_Emision_h = Mid(Vartexto, 185, 10)
               StrRazon_Social_h = Mid(Vartexto, 14, 40)
               
               StrRuc_h = Mid(Vartexto, 3, 11)
               
               StrDireccion_h = Mid(Vartexto, 65, 60)
               StrDistrito_h = Mid(Vartexto, 125, 20)
               StrProvincia_h = Mid(Vartexto, 145, 20)
               StrDepartamento_h = Mid(Vartexto, 165, 20)
               
               TotalFonPens_h1 = Mid(Vartexto, 212, 15)
               StrTotFondo_h = Mid(Vartexto, 227, 15)
               StrMoratorio_h = Mid(Vartexto, 242, 15)
               StrAdeudado_h = Mid(Vartexto, 257, 15)
               strBarra_h = Mid(Vartexto, 272, 26)
                If Val(strCantidadReg1) > 20 Then
                   If Intvariable * 20 < strCantidadReg1 Then
                   strCantidadReg = 20
                   Else
                   strCantidadReg = strCantidadReg1 - ((Intvariable - 1) * 20)
                   End If
                   'strCantidadReg = 1
                 Else
                   strCantidadReg = strCantidadReg
                End If
               For intVar = 1 To strCantidadReg
                    Line Input #lngFile, Vartexto
                    If intVar = 1 Then
                      StrPeriodo_h = Mid(Vartexto, 115, 6)
                    End If
                    cuspp_h(intVar) = Mid(Vartexto, 3, 12)
                    Separar_Apellidos_Nombres Mid(Vartexto, 15, 40), 0
                    ApellPater_h(intVar) = Mid(gApePat, 1, 20)
                    ApellMater_h(intVar) = Mid(gApeMat, 1, 20)
                    If InStr(gNombres, " ") > 0 Then
                    NombrePri_h(intVar) = Mid(Mid(gNombres, 1, InStr(gNombres, " ") - 1), 1, 10)
                    NombreSeg_h(intVar) = Mid(Mid(gNombres, InStr(gNombres, " ") + 1, Len(gNombres)), 1, 10)
                    Else
                    NombrePri_h(intVar) = Mid(Mid(gNombres, 1, Len(gNombres)), 1, 10)
                    NombreSeg_h(intVar) = ""
                    End If
                    
                    RemAseg_h(intVar) = Mid(Vartexto, 55, 15)
                    FondoPens_h(intVar) = Mid(Vartexto, 70, 15)
                    RetenRetr_h(intVar) = Mid(Vartexto, 85, 11)
                    TotalRetension_h(intVar) = Mid(Vartexto, 100, 15)
                    'TotalFonPens_h(intVar) = Mid(Vartexto, 106, 11)
                    TotalFonPens_h(intVar) = TotalFonPens_h1
               Next intVar
               If 20 - Val(strCantidadReg) > 0 Then
                   For intVar = Val(strCantidadReg) + 1 To 20
                    cuspp_h(intVar) = ""
                    'Separar_Apellidos_Nombres Mid(Vartexto, 15, 40), 0
                    ApellPater_h(intVar) = ""
                    ApellMater_h(intVar) = ""
                    NombrePri_h(intVar) = ""
                    NombreSeg_h(intVar) = ""
                    RemAseg_h(intVar) = ""
                    FondoPens_h(intVar) = ""
                    RetenRetr_h(intVar) = ""
                    TotalRetension_h(intVar) = ""
                    'TotalFonPens_h(intVar) = Mid(Vartexto, 106, 11)
                    TotalFonPens_h(intVar) = ""
                   Next intVar
               
               End If
               strEstado_h = "0"
               strNroliqpag_h = "HO" & Mid(StrNroliquidacion_h, 3, 10) & Format(intPag, "00000")
               Set xdato = Sel_Liquidacion(strNroliqpag_h)
               
               If xdato.RecordCount > 0 Then
               MsgBox "Este registro ya se encuentra en la base de datos " & strNroliqpag_h, vbInformation, "Horizonte"
               Else
               Insertar_Datos
               End If
               
     Next Intvariable
'Exit Do
Loop
    MsgBox "Se Inportaron correctamente los Archivos", vbInformation, "Horizonte"
Close #lngFile
End Sub

Private Sub cmdruta_Click()
Flow_Frmconfiguracion.Show vbModal
End Sub

Private Sub Cmdverificar_Click()
Dim MsgArchivo As String
Dim ExistArchivo As Boolean
'''verificar los archivos textos
RUTA = Dirxls & "\"
ExistArchivo = True
If Dir(RUTA & "COBAR4.TXT") = "" Then
MsgArchivo = "No existe el archivo Cobar4.txt"
ExistArchivo = False
End If

If Dir(RUTA & "COBAR5.TXT") = "" Then
MsgArchivo = MsgArchivo & "  No existe el archivo Cobar5.txt"
ExistArchivo = False
End If

If Dir(RUTA & "COBAR6.TXT") = "" Then
MsgArchivo = MsgArchivo & "  No existe el archivo Cobar6.txt"
ExistArchivo = False
End If

If Dir(RUTA & "COBAR7.TXT") = "" Then
MsgArchivo = MsgArchivo & "  No existe el archivo Cobar7.txt"
ExistArchivo = False
End If

If ExistArchivo = True Then
'''''Migrar a SQL los archivos txt
        If Abrir_conexion Then
          PrbBarra.Min = 1
          PrbBarra.Max = 400
            Borrar_Cobar4
            Importar_Sql_Cobar4 (RUTA)
            'Migrar_Cobar4
            PrbBarra.Value = 100
            Borrar_Cobar5
            Importar_Sql_Cobar5 (RUTA)
            'Migrar_Cobar5
            PrbBarra.Value = 200
            Borrar_Cobar6
            Importar_Sql_Cobar6 (RUTA)
            'Migrar_Cobar6
            PrbBarra.Value = 350
            Borrar_Cobar7
            Migrar_Cobar7
            PrbBarra.Value = 400
            
            CmdProceso.Enabled = True
        End If
      MsgBox "Se Verificaron correctamente los Archivos", vbInformation, "Horizonte"
Else
MsgBox MsgArchivo, vbInformation, "Horizonte"
End If

End Sub

Private Sub Command1_Click()
On Error GoTo PRO
If Trim$(flxls.FileName) <> "" Then
        xhoja = UCase(InputBox("Ingresar el nombre de la Hoja del Archivo en excel :", "Archivo Onell"))
        RutaArchivoExcel = Dirxls.Path & "\" & flxls.FileName
        RutaDirExcel = Dirxls.Path
        
        
        Abrir_Excell_COM
        
 Else
 MsgBox "Debe Seleccionar el Archivo", vbExclamation, "Archivo Onell"
 End If
PRO:

End Sub

Private Sub Command2_Click()
Dim Rst As Recordset
Dim PAGINA As Integer


            If Abrir_conexion = True Then
            
                    Open App.Path & "\Impresion1.txt" For Output As #6
                    Open App.Path & "\Impresion2.txt" For Output As #7
                    Open App.Path & "\Impresion3.txt" For Output As #8
                    
                    Set Rst = Sel_CebeImprenta
                     If Rst.RecordCount > 0 Then
                     
            '            Print #6, "1"; vbTab; "Pag"; vbTab; "DEPAG"; vbTab; "corre_med"; vbTab; "cod_mod"; vbTab; "anexo"; vbTab; "ie"; vbTab; "dir_ie"; vbTab; "region"; vbTab; "provincia"; vbTab; "distrito" _
            '                     ; vbTab; "cen_pob"; vbTab; "aplicador"; vbTab; "alum"
                        
                        Print #6, "1"; vbTab; "Pag"; vbTab; "DEPAG"; vbTab; "corre_med"; vbTab; "cod_mod"; vbTab; "ie"; vbTab; "dir_ie"; vbTab; "region"; vbTab; "distrito" _
                                 ; vbTab; "cen_pob"; vbTab; "aplicador"; vbTab; "alum"
                        
                        
                        Print #7, "1"; vbTab; "Pag"; vbTab; "DEPAG"; vbTab; "corre_med"; vbTab; "cod_mod"; vbTab; "ie"; vbTab; "dir_ie"; vbTab; "region"; vbTab; "distrito" _
                                 ; vbTab; "cen_pob"; vbTab; "aplicador"; vbTab; "alum"; vbTab; "11"; vbTab; "12"; vbTab; "13"; vbTab; "14"; vbTab; "15"; vbTab; "16"; vbTab; "17"; vbTab; "18"
                        
                        
                        Print #8, "1"; vbTab; "Pag"; vbTab; "DEPAG"; vbTab; "corre_med"; vbTab; "cod_mod"; vbTab; "ie"; vbTab; "dir_ie"; vbTab; "region"; vbTab; "distrito" _
                                 ; vbTab; "cen_pob"; vbTab; "aplicador"; vbTab; "alum"; vbTab; "11"; vbTab; "12"; vbTab; "13"; vbTab; "14"; vbTab; "15"; vbTab; "16"; vbTab; "17"; vbTab; "18"
                        
                        For X = 1 To Rst.RecordCount
                        
                        t1 = Trim$(IIf(IsNull(Rst.Fields("corre_med")), "", Rst.Fields("corre_med")))
                        t2 = Trim$(IIf(IsNull(Rst.Fields("cod_mod")), "", Rst.Fields("cod_mod")))
                        t3 = Trim$(IIf(IsNull(Rst.Fields("anexo")), "", Rst.Fields("anexo")))
                        t4 = Trim$(IIf(IsNull(Rst.Fields("ie")), "", Rst.Fields("ie")))
                        t5 = Trim$(IIf(IsNull(Rst.Fields("dir_ie")), "", Rst.Fields("dir_ie")))
                        t6 = Trim$(IIf(IsNull(Rst.Fields("region")), "", Rst.Fields("region")))
                        t7 = Trim$(IIf(IsNull(Rst.Fields("provincia")), "", Rst.Fields("provincia")))
                        t8 = Trim$(IIf(IsNull(Rst.Fields("distrito")), "", Rst.Fields("distrito")))
                        t9 = Trim$(IIf(IsNull(Rst.Fields("cen_pob")), "", Rst.Fields("cen_pob")))
                        t10 = Trim$(IIf(IsNull(Rst.Fields("aplicador")), "", Rst.Fields("aplicador")))
                        t11 = Trim$(IIf(IsNull(Rst.Fields("alum")), "", Rst.Fields("alum")))
                        NUM = Format(X, "0000000")
                        'PAGINA = Val(T11) / 8
                        
                        
                        If Val(t11) > 8 Then
                             Cantdet = Int(Val(t11) / 8)
                                If Val(t11) - (Cantdet * 8) > 0 Then
                                 PAGINA = Cantdet + 1
                                Else
                                 PAGINA = Cantdet
                                End If
                        Else
                        PAGINA = 1
                       End If
                        
                        
                        
                        
                        Pag = "1"
                        DEPAG = PAGINA + 1
                        
                        Print #6, NUM; vbTab; Pag; vbTab; DEPAG; vbTab; t1; vbTab; t2 & "-" & t3; vbTab; t4; vbTab; t5; vbTab; t6 & " / " & t7 _
                                          ; vbTab; t8; vbTab; t9; vbTab; t10; vbTab; Format(t11, "00")
                        
                        
                            For Y = 1 To PAGINA
                               Pag = Y + 1
                               
                               VAR = 8 * (Y - 1)
                               
                               XX1 = VAR + 1
                               XX2 = VAR + 2
                               XX3 = VAR + 3
                               XX4 = VAR + 4
                               XX5 = VAR + 5
                               XX6 = VAR + 6
                               XX7 = VAR + 7
                               XX8 = VAR + 8
                               
                               If X <= 20000 Then
                               Print #7, NUM; vbTab; Pag; vbTab; DEPAG; vbTab; t1; vbTab; t2 & "-" & t3; vbTab; t4; vbTab; t5; vbTab; t6 & " / " & t7 _
                                          ; vbTab; t8; vbTab; t9; vbTab; t10; vbTab; Format(t11, "00"); vbTab; Format(XX1, "00"); vbTab; Format(XX2, "00"); vbTab; Format(XX3, "00"); vbTab; Format(XX4, "00"); vbTab; Format(XX5, "00"); vbTab; Format(XX6, "00"); vbTab; Format(XX7, "00"); vbTab; Format(XX8, "00")
                               
                               Else
                               Print #8, NUM; vbTab; Pag; vbTab; DEPAG; vbTab; t1; vbTab; t2 & "-" & t3; vbTab; t4; vbTab; t5; vbTab; t6 & " / " & t7 _
                                          ; vbTab; t8; vbTab; t9; vbTab; t10; vbTab; Format(t11, "00"); vbTab; Format(XX1, "00"); vbTab; Format(XX2, "00"); vbTab; Format(XX3, "00"); vbTab; Format(XX4, "00"); vbTab; Format(XX5, "00"); vbTab; Format(XX6, "00"); vbTab; Format(XX7, "00"); vbTab; Format(XX8, "00")
                               End If
                            Next Y
                        
                        Rst.MoveNext
                        Next X
                     End If
                    Close #6
                    Close #7
                    Close #8
             End If
End Sub

Private Sub Command3_Click()

If Trim$(flxls.FileName) <> "" Then
        
        xhoja = UCase(InputBox("Ingresar el nombre de la Hoja del Archivo en excel :", "Archivo Onell"))
        RutaArchivoExcel = Dirxls.Path & "\" & flxls.FileName
        RutaDirExcel = Dirxls.Path
        
       If Abrir_conexion Then
        Borrar_IMPRENTA
        Abrir_Excell_IMPRENTA
       End If
 Else
 MsgBox "Debe Seleccionar el Archivo", vbExclamation, "Archivo Onell"
 End If
'If Abrir_conexion Then
'            'PrbBarra.Min = 1
'            'PrbBarra.Max = 400
'            Borrar_IMPRENTA
'            Migrar_Cobar4
'
'End If
End Sub

Private Sub Command4_Click()
FECHA = "01/20/2007"
If IsDate(FECHA) Then
mes = Month("20")
  
  MsgBox "verdadero"
Else
  MsgBox "False"
End If
End Sub

Private Sub Command5_Click()
Dim Rst1  As Recordset
Dim Rs As Recordset
Abrir_conexion

Set Rs = New Recordset
Rs.ActiveConnection = CnIntegra
        Rs.Source = "SELECT  * FROM COBAR6 WHERE  cobnli='IN10058374' and cobt01='1' order by cobpag"
        Rs.LockType = adLockOptimistic
        Rs.CursorLocation = adUseClient
        Rs.CursorType = adOpenKeyset
        Rs.Open

'Set Rs = Sel_arreglo("IN09312732")
'If Rst1.RecordCount > 0 Then
  While Not Rs.EOF
  
  For X = 1 To 118
   'If X = 118 Then MsgBox "prueba"
    For Y = 1 To 20
     Rs.Fields("cobpag") = Format(X, "000")
     Rs.Update
    Rs.MoveNext
    Next Y
  Next X
  Wend
'End If
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Dim Rst1  As Recordset
Dim Rs As Recordset
Abrir_conexion

Set Rs = Sel_PRUEBA
If Rs.RecordCount > 0 Then
  For X = 1 To Rs.RecordCount
  
      'Set Rst1 = Sel_PRUEBA1(Rs.Fields("ROTULOS"), Rs.Fields("ESPECIAL"))
      ' If Rst1.RecordCount > 0 Then
       'Rst1.MoveNext
        'For Y = 1 To Rst1.RecordCount
        
        'Borrar_PRUEBA (Rs.Fields("ROTULOS"), Rs.Fields("ESPECIAL"))
        Borrar_PRUEBA_COBAR6 Rs.Fields("ROTULOS"), Rs.Fields("ESPECIAL")
        Borrar_PRUEBA_COBAR4 Rs.Fields("ESPECIAL")
        'Rst1.MoveNext
       ' Next Y
       'End If
  Rs.MoveNext
  Next X
End If
End Sub

Private Sub Command8_Click()
Dim Rst1  As Recordset
Dim Rs As Recordset
Dim AU1, LO1, SE1 As String
Dim AU2, LO2, SE2 As String
Dim AU3, LO3, SE3 As String
Dim AU4, LO4, SE4 As String
Dim AU5, LO5, SE5 As String
Dim AU6, LO6, SE6 As String
Dim AU7, LO7, SE7 As String
Dim AU8, LO8, SE8 As String
Dim AU9, LO9, SE9 As String
Dim AU10, LO10, SE10 As String
Dim AU11, LO11, SE11 As String
Dim AU12, LO12, SE12 As String
Dim AU13, LO13, SE13 As String
Dim AU14, LO14, SE14 As String
Dim AU15, LO15, SE15 As String
Dim AUL(1 To 15) As String
Dim LOT(1 To 15) As String
Dim SES(1 To 15) As String
Abrir_conexion

Open App.Path & "\Impresion.txt" For Output As #6
Print #6, "SEDE1"; vbTab; "AULA1"; vbTab; "LOCA1"; vbTab; "SEDE2"; vbTab; "AULA2"; vbTab; "LOCAL2"; vbTab; "SEDE3"; vbTab; "AULA3"; vbTab; "LOCAL3"; vbTab; "SEDE4"; vbTab; "AULA4"; vbTab; "LOCAL4"; vbTab; "SEDE5"; vbTab; "AULA5"; vbTab; "LOCAL5"; vbTab; "SEDE6"; vbTab; "AULA6"; vbTab; "LOCAL6"; vbTab; "SEDE7"; vbTab; "AULA7"; vbTab; "LOCAL7"; vbTab; "SEDE8"; vbTab; "AULA8"; vbTab; "LOCAL8"; vbTab; "SEDE9"; vbTab; "AULA9"; vbTab; "LOCAL9"; vbTab; "SEDE10"; vbTab; "AULA10"; vbTab; "LOCAL10"; vbTab; "SEDE11"; vbTab; "AULA11"; vbTab; "LOCAL11"; vbTab; "SEDE12"; vbTab; "AULA12"; vbTab; "LOCAL12"; vbTab; "SEDE13"; vbTab; "AULA13"; vbTab; "LOCAL13"; vbTab; "SEDE14"; vbTab; "AULA14"; vbTab; "LOCAL14"; vbTab; "SEDE15"; vbTab; "AULA15"; vbTab; "LOCAL15"


Set Rs = Sel_PRUEBA
If Rs.RecordCount > 0 Then
   For X = 1 To Rs.RecordCount
      For Y = 1 To 14
      
        For p = 1 To 15
           h = h + 1
           If h <= Y * 15 Then
            AUL(p) = Rs.Fields("sedeeval")
            LOT(p) = Format(Rs.Fields("aula"), "000")
            SES(p) = Rs.Fields("localeval")
            
'            AUL(15) = ""
'            LOT(15) = ""
'            SES(15) = ""
           Else
           
           End If
        Rs.MoveNext
       Next p
      Print #6, AUL(1); vbTab; LOT(1); vbTab; SES(1); vbTab; AUL(2); vbTab; LOT(2); vbTab; SES(2); vbTab; AUL(3); vbTab; LOT(3); vbTab; SES(3); ; vbTab; AUL(4); vbTab; LOT(4); vbTab; SES(4); ; vbTab; AUL(5); vbTab; LOT(5); vbTab; SES(5); ; vbTab; AUL(6); vbTab; LOT(6); vbTab; SES(6); ; vbTab; AUL(7); vbTab; LOT(7); vbTab; SES(7); ; vbTab; AUL(8); vbTab; LOT(8); vbTab; SES(8); ; vbTab; AUL(9); vbTab; LOT(9); vbTab; SES(9); ; vbTab; AUL(10); vbTab; LOT(10); vbTab; SES(10); ; vbTab; AUL(11); vbTab; LOT(11); vbTab; SES(11); ; vbTab; AUL(12); vbTab; LOT(12); vbTab; SES(12); ; vbTab; AUL(13); vbTab; LOT(13); vbTab; SES(13); ; vbTab; AUL(14); vbTab; LOT(14); vbTab; SES(14); ; vbTab; AUL(15); vbTab; LOT(15); vbTab; SES(15)
      Next Y
   
   Next X
   Close #6
End If

End Sub

Private Sub Dirxls_Change()
flxls.Path = Dirxls.Path
End Sub

Private Sub Drvxls_Change()
Dirxls.Path = Drvxls.Drive
End Sub

Public Sub Migrar_Cobar4()
Dim Vartexto As String
Dim Lproceso As String
Dim NroCargo As String
Dim Modcarta As String
Dim Tipcargo As String
Dim Localidad  As String
Dim FechTexto As String
Dim Quien As String
Dim Razon As String
Dim direc1 As String
Dim direc2 As String
Dim TipoIdent As String
Dim NroIdent As String
'Dim TipoIdent As String
Dim NomDistr As String
Dim Nomprov As String
Dim Atencion As String
Dim Referencia As String
Dim Adjunto As String
Dim Linea1 As String
Dim Linea2 As String
Dim Linea3 As String
Dim Linea4 As String
Dim Linea5 As String
Open RUTA & "COBAR4.txt" For Input As #1
'''923 Caracteres longitud de Texto
Do Until EOF(1)
   Line Input #1, Vartexto
        DoEvents
        Lproceso = Mid(Vartexto, 1, 11)
        NroCargo = Mid(Vartexto, 12, 9)
        Modcarta = Mid(Vartexto, 21, 1)
        Tipcargo = Mid(Vartexto, 22, 4)
        Localidad = Mid(Vartexto, 26, 10)
        FechTexto = Mid(Vartexto, 36, 20)
        Quien = Mid(Vartexto, 56, 15)
        Razon = Mid(Vartexto, 71, 100)
        direc1 = Mid(Vartexto, 171, 75)
        direc2 = Mid(Vartexto, 246, 75)
        TipoIdent = Mid(Vartexto, 321, 3)
        NroIdent = Mid(Vartexto, 324, 16)
        NomDistr = Mid(Vartexto, 340, 30)
        Nomprov = Mid(Vartexto, 370, 63)
        Atencion = Mid(Vartexto, 433, 45)
        Referencia = Mid(Vartexto, 478, 75)
        Adjunto = Mid(Vartexto, 553, 21)
        Linea1 = Mid(Vartexto, 574, 70)
        Linea2 = Mid(Vartexto, 644, 70)
        Linea3 = Mid(Vartexto, 714, 70)
        Linea4 = Mid(Vartexto, 784, 70)
        Linea5 = Mid(Vartexto, 854, 70)
        
        
        If InStr(Trim$(Quien), "¤") > 0 Then
        Quien = Replace(Quien, "¤", "ñ")  'Se¤ores
        Else
        End If
        
        If InStr(Trim$(Razon), "¥") > 0 Then
        Razon = Replace(Razon, "¥", "Ñ")  'Se¤ores
        Else
        End If
     
   
   Grabar_Sql_Cobar4 Lproceso, NroCargo, Modcarta, Tipcargo, Localidad, _
              FechTexto, Quien, Razon, direc1, direc2, TipoIdent, NroIdent, _
              NomDistr, Nomprov, Atencion, Referencia, Adjunto, Linea1, Linea2, _
              Linea3, Linea4, Linea5
   
   Loop
Close #1

End Sub


Public Sub Migrar_Cobar5()
Dim Vartexto As String
Dim Lproceso As String
Dim NroCargo As String
Dim NroLiq As String
Dim Tipcargo As String
Dim Secuen  As String
Dim TipImpr As String
Dim Periodo As String
Dim FechTexto As String
Dim Pag As String
Dim Razon As String
Dim TipoIdent As String
Dim NroIdent As String
'Dim TipoIdent As String
Dim direcav As String
Dim direcnu As String
Dim direcurb As String
Dim NomDistr As String
Dim Nomprov As String
Dim NomDepar As String

Open RUTA & "COBAR5.txt" For Input As #2
'''368 Caracteres longitud de Texto
Do Until EOF(2)
   Line Input #2, Vartexto
   DoEvents
        Lproceso = Mid(Vartexto, 1, 11)
        NroCargo = Mid(Vartexto, 12, 9)
        NroLiq = Mid(Vartexto, 21, 10)
        Tipcargo = Mid(Vartexto, 31, 4)
        Secuen = Mid(Vartexto, 35, 2)
        TipImpr = Mid(Vartexto, 37, 1)
        Periodo = Mid(Vartexto, 38, 7)
        FechTexto = Mid(Vartexto, 45, 20)
        Pag = Mid(Vartexto, 65, 5)
        Razon = Mid(Vartexto, 70, 100)
        TipoIdent = Mid(Vartexto, 170, 3)
        NroIdent = Mid(Vartexto, 173, 16)
        direcav = Mid(Vartexto, 189, 50)
        direcnu = Mid(Vartexto, 239, 15)
        direcurb = Mid(Vartexto, 254, 25)
        NomDistr = Mid(Vartexto, 279, 30)
        Nomprov = Mid(Vartexto, 309, 30)
        NomDepar = Mid(Vartexto, 339, 30)
        
        
        If InStr(Trim$(Razon), "¥") > 0 Then
        Razon = Replace(Razon, "¥", "Ñ")
        Else
        End If
      
   Grabar_Sql_Cobar5 Lproceso, NroCargo, NroLiq, Tipcargo, Secuen, _
              TipImpr, Periodo, FechTexto, Pag, Razon, TipoIdent, NroIdent, _
              direcav, direcnu, direcurb, NomDistr, Nomprov, NomDepar
Loop
Close #2

End Sub


Public Sub Migrar_Cobar6()
Dim Vartexto As String
Dim Lproceso As String
Dim NroCargo As String
Dim NroLiq As String
Dim Tipdet As String
Dim Appat  As String
Dim Apmat As String
Dim Nom1 As String
Dim Nom2 As String
Dim Cusp As String
Dim Remun As String
Dim Montf As String
Dim Montr As String
'Dim TipoIdent As String
Dim Totalaf As String
Dim Totalfdo As String
Dim Totalint As String
Dim Totalgen As String
Dim Pag As String


Open RUTA & "COBAR6.txt" For Input As #3
'''199 Caracteres longitud de Texto
Do Until EOF(3)
   Line Input #3, Vartexto
   DoEvents
        Lproceso = Mid(Vartexto, 1, 11)
        NroCargo = Mid(Vartexto, 12, 9)
        NroLiq = Mid(Vartexto, 21, 10)
        Tipdet = Mid(Vartexto, 31, 1)
        Appat = Mid(Vartexto, 32, 20)
        Apmat = Mid(Vartexto, 52, 20)
        Nom1 = Mid(Vartexto, 72, 15)
        Nom2 = Mid(Vartexto, 87, 15)
        Cusp = Mid(Vartexto, 102, 12)
        Remun = Mid(Vartexto, 114, 10)
        Montf = Mid(Vartexto, 124, 10)
        Montr = Mid(Vartexto, 134, 10)
        Totalaf = Mid(Vartexto, 144, 10)
        Totalfdo = Mid(Vartexto, 154, 12)
        Totalint = Mid(Vartexto, 166, 12)
        Totalgen = Mid(Vartexto, 178, 14)
        Pag = Mid(Vartexto, 192, 8)
        
        If InStr(Trim$(Appat), "¥") > 0 Then
        Appat = Replace(Appat, "¥", "Ñ")
        Else
        End If
                
        If InStr(Trim$(Apmat), "¥") > 0 Then
        Apmat = Replace(Apmat, "¥", "Ñ")
        Else
        End If
   
        If InStr(Trim$(Cusp), "¥") > 0 Then
        Cusp = Replace(Cusp, "¥", "Ñ")
        Else
        End If
      
   
  Grabar_Sql_Cobar6 Lproceso, NroCargo, NroLiq, Tipdet, Appat, _
              Apmat, Nom1, Nom2, Cusp, Remun, Montf, Montr, _
              Totalaf, Totalfdo, Totalint, Totalgen, Pag
   
   Loop
Close #3

End Sub

Public Sub Migrar_Cobar7()
Dim Tippro As String
Dim Codmen  As String
Dim Sec As String
Dim Mens As String
Dim Vartexto As String


Open RUTA & "COBAR7.txt" For Input As #1

Do Until EOF(1)
   Line Input #1, Vartexto
   DoEvents
        Tippro = Mid(Vartexto, 1, 3)
        Codmen = Mid(Vartexto, 4, 3)
        Sec = Mid(Vartexto, 7, 3)
        Mens = Mid(Vartexto, 10, 150)
     
  Grabar_Sql_Cobar7 Tippro, Codmen, Sec, Mens
   
   Loop
Close #1

End Sub

Private Sub Form_Load()
'RutaArchivos = App.Path & "\"

'''Variable de Ruta de Directorio donde se almacena lso archivos Textos
ContRutatxt = App.Path & "\Integra.ini"
'ContRutatxt
Rutad = LeerIni(ContRutatxt, "Destinoruta", "Rutad", "")
RutaArchivos = Rutad
xxEstado = "3"
'If Dir(Rutad, vbDirectory) = "" Then
'Drive2 = App.Path
'Dir2 = App.Path
'Else
'Drive2 = Rutad
'Dir2 = Rutad
'End If

End Sub

Public Sub Abrir_Excell()
Dim Xrep As String
Dim Xrep1 As String
Dim Xnomdoc As String
Dim Xtipodoc As String
'Dim XCuspp As String
Dim Xhojas As String
Dim Xcarpeta As String
Dim Xpaterno As String
Dim Xmaterno As String
Dim Xnombres1 As String
Dim Xnombres2 As String
Dim Xcortes As String
Dim Ximp As String
Dim Xnumpag As String
Dim Xnumpre As String
Dim Xsecafi As String
Dim Xtipmov As String
Dim Xfecmov As String
Dim xremase As String
Dim xobltra As String
Dim xoblemp As String
Dim xobltot As String
Dim xvolcfi As String
Dim xvolsfi As String
Dim xvolemp As String
Dim xtotfon As String
Dim xseguro As String
Dim Xcompor As String
Dim xtotadm As String
Dim Xcampos As Integer
Dim Xposicion As Double
Dim X As Integer
Dim Y As Integer
Dim IntContador  As Integer
Dim RsExcell As New Recordset
On Error GoTo Error_Handling_Generar_Archivo

With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RutaArchivoExcel & _
                ";Extended Properties=""Excel 8.0;"""
                '";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "Select  * from [" & xhoja & "$] where CAMPO1  is not null"

End With
'Xposicion = XposIni
Open App.Path & "\Descargototal.txt" For Output As #3
Print #3, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"


If RsExcell.RecordCount > 0 Then
  If RsExcell.RecordCount > 0 Then
     'For X = 1 To RsExcell.RecordCount
      
      
       For Y = 1 To RsExcell.RecordCount
       ReDim xcuspp(1 To 20)
       ReDim Xcodigo(1 To 20)
       ReDim xfecha1(1 To 20)
       ReDim xfecha2(1 To 20)
       ReDim xcomen(1 To 20)
       ReDim xFere(1 To 20)
       IntContador = 0
       xcuspp(1) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       Xcodigo(1) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(1) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(1) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(1) = codigo(xcodigo(1), xfecha1(1), xfecha1(1))
        If Len(Trim$(xcuspp(1))) > 0 Then
        IntContador = IntContador + 1
        End If
       xcuspp(2) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO26")), "", RsExcell.Fields("CAMPO26")))
       Xcodigo(2) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(2) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(2) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(2) = codigo(xcodigo(2), xfecha1(2), xfecha1(2))
       If Len(Trim$(xcuspp(2))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       xcuspp(3) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO32")), "", RsExcell.Fields("CAMPO32")))
       Xcodigo(3) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(3) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(3) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(3) = codigo(xcodigo(3), xfecha1(3), xfecha1(3))
       If Len(Trim$(xcuspp(3))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       xcuspp(4) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO38")), "", RsExcell.Fields("CAMPO38")))
       Xcodigo(4) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(4) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(4) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(4) = codigo(xcodigo(4), xfecha1(4), xfecha1(4))
       If Len(Trim$(xcuspp(4))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(5) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO44")), "", RsExcell.Fields("CAMPO44")))
       Xcodigo(5) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(5) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(5) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(5) = codigo(xcodigo(5), xfecha1(5), xfecha1(5))
       If Len(Trim$(xcuspp(5))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(6) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO50")), "", RsExcell.Fields("CAMPO50")))
       Xcodigo(6) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(6) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(6) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(6) = codigo(xcodigo(6), xfecha1(6), xfecha1(6))
       If Len(Trim$(xcuspp(6))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(7) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO56")), "", RsExcell.Fields("CAMPO56")))
       Xcodigo(7) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(7) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(7) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(7) = codigo(xcodigo(7), xfecha1(7), xfecha1(7))
       If Len(Trim$(xcuspp(7))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(8) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO62")), "", RsExcell.Fields("CAMPO62")))
       Xcodigo(8) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(8) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(8) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(8) = codigo(xcodigo(8), xfecha1(8), xfecha1(8))
       If Len(Trim$(xcuspp(8))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(9) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO68")), "", RsExcell.Fields("CAMPO68")))
       Xcodigo(9) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(9) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(9) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(9) = codigo(xcodigo(9), xfecha1(9), xfecha1(9))
       If Len(Trim$(xcuspp(9))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       xcuspp(10) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO74")), "", RsExcell.Fields("CAMPO74")))
       Xcodigo(10) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(10) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(10) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(10) = codigo(xcodigo(10), xfecha1(10), xfecha1(10))
       If Len(Trim$(xcuspp(10))) > 0 Then
        IntContador = IntContador + 1
        End If
              
       xcuspp(11) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO80")), "", RsExcell.Fields("CAMPO80")))
       Xcodigo(11) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(11) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(11) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(11) = codigo(xcodigo(11), xfecha1(11), xfecha1(11))
       If Len(Trim$(xcuspp(11))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(12) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO86")), "", RsExcell.Fields("CAMPO86")))
       Xcodigo(12) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(12) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(12) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(12) = codigo(xcodigo(12), xfecha1(12), xfecha1(12))
       If Len(Trim$(xcuspp(12))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       xcuspp(13) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO92")), "", RsExcell.Fields("CAMPO92")))
       Xcodigo(13) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(13) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(13) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(13) = codigo(xcodigo(13), xfecha1(13), xfecha1(13))
       If Len(Trim$(xcuspp(13))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(14) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO98")), "", RsExcell.Fields("CAMPO98")))
       Xcodigo(14) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(14) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(14) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(14) = codigo(xcodigo(14), xfecha1(14), xfecha1(14))
       If Len(Trim$(xcuspp(14))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       xcuspp(15) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO104")), "", RsExcell.Fields("CAMPO104")))
       Xcodigo(15) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(15) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(15) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(15) = codigo(xcodigo(15), xfecha1(15), xfecha1(15))
       If Len(Trim$(xcuspp(15))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(16) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO110")), "", RsExcell.Fields("CAMPO110")))
       Xcodigo(16) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(16) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(16) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(16) = codigo(xcodigo(16), xfecha1(16), xfecha1(16))
       If Len(Trim$(xcuspp(16))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(17) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO116")), "", RsExcell.Fields("CAMPO116")))
       Xcodigo(17) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(17) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(17) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(17) = codigo(xcodigo(17), xfecha1(17), xfecha1(17))
       If Len(Trim$(xcuspp(17))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       
       xcuspp(18) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO122")), "", RsExcell.Fields("CAMPO122")))
       Xcodigo(18) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(18) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(18) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(18) = codigo(xcodigo(18), xfecha1(18), xfecha1(18))
       If Len(Trim$(xcuspp(18))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       xcuspp(19) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO128")), "", RsExcell.Fields("CAMPO128")))
       Xcodigo(19) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(19) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(19) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(19) = codigo(xcodigo(19), xfecha1(19), xfecha1(19))
       If Len(Trim$(xcuspp(19))) > 0 Then
        IntContador = IntContador + 1
        End If
       xcuspp(20) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO134")), "", RsExcell.Fields("CAMPO134")))
       Xcodigo(20) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(20) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha2(20) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       'Xcomen(20) = codigo(xcodigo(20), xfecha1(20), xfecha1(20))
       If Len(Trim$(xcuspp(20))) > 0 Then
        IntContador = IntContador + 1
        End If
       
       Xrazon = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO6")), "", RsExcell.Fields("CAMPO6")))
       xruc = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO7")), "", RsExcell.Fields("CAMPO7")))
       Xperiodo = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO19")), "", RsExcell.Fields("CAMPO19")))
       Xcomentario = ""
       Xfecha = ""
       
       
          For X = 1 To IntContador
          
          Print #3, xruc; vbTab; Xrazon; vbTab; Xcomentario; vbTab; Xcodigo(X); vbTab; xcuspp(X); vbTab; Xperiodo; vbTab; Xfecha
          Next X
          RsExcell.MoveNext
       Next Y
    Set RsExcell = Nothing
  End If
End If

Close #3
Error_Handling_Generar_Archivo:
End Sub

Public Function Mostrar_Comentario(Xcodigo As String, xinicio As String, xfin As String, xmonto As String) As String
xcampo = 1
Select Case Format(Xcodigo, "00")
  Case "":
        If Len(xinicio) > 0 Then
         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
        End If
        
        If Len(xfin) > 0 Then
         Mostrar_Comentario = "CESE DE RELACION LABORAL"
        End If
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
         xcampo = 2
        End If
                
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
          If Len(xmonto) > 0 Then
          Mostrar_Comentario = "ACTUALIZACION DE RIA"
          End If
        End If
         
  Case "01":
        If Len(xinicio) > 0 Then
         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
        End If
        If Len(xfin) > 0 Then
         Mostrar_Comentario = "INICIO DE RELACION LABORAL"
        End If
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
        End If
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
        xcampo = 2
        End If
        If Len(xinicio) = 4 Then
        Mostrar_Comentario = "INICIO DE RELACION LABORAL"
        End If
  Case "02":
        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
         Mostrar_Comentario = "CESE DE RELACION LABORAL"
        End If
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Comentario = "CESE DE RELACION LABORAL"
        End If
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Comentario = "CESE DE RELACION LABORAL"
        End If
        If Len(xfin) = 4 Then
        Mostrar_Comentario = "CESE DE RELACION LABORAL"
        End If

Case "04":
        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
         Mostrar_Comentario = "LICENCIA SIN GOCE DE HABER"
        End If
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
         Mostrar_Comentario = "LICENCIA SIN GOCE DE HABER"
        End If
'        If Len(xinicIo) > 0 And Len(xfin) > 0 Then
'        Mostrar_Comentario = "LICENCIA SIN GOCE DE HABER"
'        End If
'        If Len(xinicIo) = 0 And Len(xfin) = 0 Then
'        Mostrar_Comentario = "LICENCIA SIN GOCE DE HABER"
'        End If

Case "05":
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Comentario = "VACACIONES"
        Else
        Mostrar_Comentario = "VACACIONES"
        End If
        
Case Xcodigo = "18"
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Comentario = "AFILIADO NUNCA LABORO EN LA EMPRESA"
        End If

Case "20":
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Comentario = "POSTERGAR DEVENGUE Lima y provincias"
        Else
        Mostrar_Comentario = "POSTERGAR DEVENGUE Lima y provincias"
        End If

Case "21":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Comentario = "APORTE CANCELADO A AFP Horizonte"
        'End If
Case "22":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Comentario = "APORTE CANCELADO A OTRA AFP"

Case "RIA":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Comentario = "ACTUALIZACION DE RIA"

Case "99":
        
        If Len(xmonto) > 0 Then
        Mostrar_Comentario = "ACTUALIZACION DE RIA"
        End If
        If Len(xmonto) = 0 Then
        Mostrar_Comentario = "EMPLEADOR CONFIRMA DEUDA"
        End If
End Select
End Function


Public Sub Abrir_Excell_COM()
Dim Xrep As String
Dim Xrep1 As String
Dim Xnomdoc As String
Dim Xtipodoc As String
'Dim XCuspp As String
Dim Xhojas As String
Dim Xcarpeta As String
Dim Xpaterno As String
Dim Xmaterno As String
Dim Xnombres1 As String
Dim Xnombres2 As String
Dim Xcortes As String
Dim Ximp As String
Dim Xnumpag As String
Dim Xnumpre As String
Dim Xsecafi As String
Dim Xtipmov As String
Dim Xfecmov As String
Dim xremase As String
Dim xobltra As String
Dim xoblemp As String
Dim xobltot As String
Dim xvolcfi As String
Dim xvolsfi As String
Dim xvolemp As String
Dim xtotfon As String
Dim xseguro As String
Dim Xcompor As String
Dim xtotadm As String
Dim Xcampos As Integer
Dim Xposicion As Double
Dim X As Integer
Dim Y As Integer
Dim IntContador  As Integer
Dim RsExcell As New Recordset
Dim Xcome As String
On Error GoTo Error_Handling_Generar_Archivo

With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RutaArchivoExcel & _
                ";Extended Properties=""Excel 8.0;"""
                '";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "Select  * from [" & xhoja & "$] where CAMPO141  is not null"

End With
'Xposicion = XposIni
Open App.Path & "\DescargoPRUEBA.txt" For Output As #3
Print #3, "RUC"; vbTab; "RAZ_SOC"; vbTab; "COMENTARIO"; vbTab; "COD_DESCARGO"; vbTab; "CUSPP"; vbTab; "PERIODO"; vbTab; "FECHA_REFERENCIA"; vbTab; "CAMPO"


If RsExcell.RecordCount > 0 Then
  If RsExcell.RecordCount > 0 Then
     'For X = 1 To RsExcell.RecordCount
      
       For Y = 1 To RsExcell.RecordCount
       ReDim xcuspp(1 To 20)
       ReDim Xcodigo(1 To 20) As String
       ReDim xfecha1(1 To 20) As String
       ReDim xfecha2(1 To 20) As String
       ReDim xcomen(1 To 20) As String
       ReDim xFere(1 To 20) As String
       ReDim xmonto(1 To 20) As String
       ReDim CAMPO(1 To 20) As String
       IntContador = 0
       xcuspp(1) = "" 'Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       Xcodigo(1) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO1")), "", RsExcell.Fields("CODIGO1")))
       xfecha1(1) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio1")), "", RsExcell.Fields("fecha de inicio1")))
       xfecha2(1) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE1")), "", RsExcell.Fields("fecha de CESE1")))
       'Xcome = xcodigo(xcodigo(1), xfecha1(1), xfecha2(1))
       
       
       xmonto(1) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable1")), "", RsExcell.Fields("Remuneracion asegurable1")))
       xcomen(1) = Mostrar_Comentario(Xcodigo(1), xfecha1(1), xfecha2(1), xmonto(1))
       xFere(1) = Mostrar_Fecha(Xcodigo(1), xfecha1(1), xfecha2(1), xmonto(1))
       
       If Trim$(Xcodigo(1)) = "" Then
       Xcodigo(1) = Codigo_Nullo(Xcodigo(1), xfecha1(1), xfecha2(1))
       End If
       
       CAMPO(1) = xcampo
       If Len(Trim$(Xcodigo(1))) > 0 Or Len(Trim$(xfecha1(1))) > 0 Or Len(Trim$(xfecha2(1))) > 0 Or Len(Trim$(xmonto(1))) > 0 Then
       IntContador = IntContador + 1
       End If
              
       Xcodigo(2) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO2")), "", RsExcell.Fields("CODIGO2")))
       'xcodigo(2) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(2) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio2")), "", RsExcell.Fields("fecha de inicio2")))
       xfecha2(2) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE2")), "", RsExcell.Fields("fecha de CESE2")))
       xmonto(2) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable2")), "", RsExcell.Fields("Remuneracion asegurable2")))
       xcomen(2) = Mostrar_Comentario(Xcodigo(2), xfecha1(2), xfecha2(2), xmonto(2))
       xFere(2) = Mostrar_Fecha(Xcodigo(2), xfecha1(2), xfecha2(2), xmonto(2))
       
       If Trim$(Xcodigo(2)) = "" Then
       Xcodigo(2) = Codigo_Nullo(Xcodigo(2), xfecha1(2), xfecha2(2))
       End If
       
       CAMPO(2) = xcampo
       If Len(Trim$(Xcodigo(2))) > 0 Or Len(Trim$(xfecha1(2))) > 0 Or Len(Trim$(xfecha2(2))) > 0 Or Len(Trim$(xmonto(2))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(3) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO3")), "", RsExcell.Fields("CODIGO3")))
       'xcodigo(3) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(3) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio3")), "", RsExcell.Fields("fecha de inicio3")))
       xfecha2(3) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE3")), "", RsExcell.Fields("fecha de CESE3")))
       xmonto(3) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable3")), "", RsExcell.Fields("Remuneracion asegurable3")))
       xcomen(3) = Mostrar_Comentario(Xcodigo(3), xfecha1(3), xfecha2(3), xmonto(3))
       xFere(3) = Mostrar_Fecha(Xcodigo(3), xfecha1(3), xfecha2(3), xmonto(3))
       CAMPO(3) = xcampo
       If Trim$(Xcodigo(3)) = "" Then
       Xcodigo(3) = Codigo_Nullo(Xcodigo(3), xfecha1(3), xfecha2(3))
       End If
       
       If Len(Trim$(Xcodigo(3))) > 0 Or Len(Trim$(xfecha1(3))) > 0 Or Len(Trim$(xfecha2(3))) > 0 Or Len(Trim$(xmonto(3))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(4) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO4")), "", RsExcell.Fields("CODIGO4")))
       'xcodigo(4) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(4) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio4")), "", RsExcell.Fields("fecha de inicio4")))
       xfecha2(4) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE4")), "", RsExcell.Fields("fecha de CESE4")))
       xmonto(4) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable4")), "", RsExcell.Fields("Remuneracion asegurable4")))
       xcomen(4) = Mostrar_Comentario(Xcodigo(4), xfecha1(4), xfecha2(4), xmonto(4))
       xFere(4) = Mostrar_Fecha(Xcodigo(4), xfecha1(4), xfecha2(4), xmonto(4))
       CAMPO(4) = xcampo
       
       If Trim$(Xcodigo(4)) = "" Then
       Xcodigo(4) = Codigo_Nullo(Xcodigo(4), xfecha1(4), xfecha2(4))
       End If
       
       If Len(Trim$(Xcodigo(4))) > 0 Or Len(Trim$(xfecha1(4))) > 0 Or Len(Trim$(xfecha2(4))) > 0 Or Len(Trim$(xmonto(4))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(5) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO5")), "", RsExcell.Fields("CODIGO5")))
       'xcodigo(5) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(5) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio5")), "", RsExcell.Fields("fecha de inicio5")))
       xfecha2(5) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE5")), "", RsExcell.Fields("fecha de CESE5")))
       xmonto(5) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable5")), "", RsExcell.Fields("Remuneracion asegurable5")))
       xcomen(5) = Mostrar_Comentario(Xcodigo(5), xfecha1(5), xfecha2(5), xmonto(5))
       xFere(5) = Mostrar_Fecha(Xcodigo(5), xfecha1(5), xfecha2(5), xmonto(5))
       
       If Trim$(Xcodigo(5)) = "" Then
       Xcodigo(5) = Codigo_Nullo(Xcodigo(5), xfecha1(5), xfecha2(5))
       End If
       
       CAMPO(5) = xcampo
       If Len(Trim$(Xcodigo(5))) > 0 Or Len(Trim$(xfecha1(5))) > 0 Or Len(Trim$(xfecha2(5))) > 0 Or Len(Trim$(xmonto(5))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(6) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO6")), "", RsExcell.Fields("CODIGO6")))
       'xcodigo(6) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(6) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio6")), "", RsExcell.Fields("fecha de inicio6")))
       xfecha2(6) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE6")), "", RsExcell.Fields("fecha de CESE6")))
       xmonto(6) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable6")), "", RsExcell.Fields("Remuneracion asegurable6")))
       xcomen(6) = Mostrar_Comentario(Xcodigo(6), xfecha1(6), xfecha2(6), xmonto(6))
       xFere(6) = Mostrar_Fecha(Xcodigo(6), xfecha1(6), xfecha2(6), xmonto(6))
       
       If Trim$(Xcodigo(6)) = "" Then
       Xcodigo(6) = Codigo_Nullo(Xcodigo(6), xfecha1(6), xfecha2(6))
       End If
       
       CAMPO(6) = xcampo
       If Len(Trim$(Xcodigo(6))) > 0 Or Len(Trim$(xfecha1(6))) > 0 Or Len(Trim$(xfecha2(6))) > 0 Or Len(Trim$(xmonto(6))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(7) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO7")), "", RsExcell.Fields("CODIGO7")))
       'xcodigo(7) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(7) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio7")), "", RsExcell.Fields("fecha de inicio7")))
       xfecha2(7) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE7")), "", RsExcell.Fields("fecha de CESE7")))
       xmonto(7) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable7")), "", RsExcell.Fields("Remuneracion asegurable7")))
       xcomen(7) = Mostrar_Comentario(Xcodigo(7), xfecha1(7), xfecha2(7), xmonto(7))
       xFere(7) = Mostrar_Fecha(Xcodigo(7), xfecha1(7), xfecha2(7), xmonto(7))
       
       If Trim$(Xcodigo(7)) = "" Then
       Xcodigo(7) = Codigo_Nullo(Xcodigo(7), xfecha1(7), xfecha2(7))
       End If
       
       CAMPO(7) = xcampo
       If Len(Trim$(Xcodigo(7))) > 0 Or Len(Trim$(xfecha1(7))) > 0 Or Len(Trim$(xfecha2(7))) > 0 Or Len(Trim$(xmonto(7))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(8) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO8")), "", RsExcell.Fields("CODIGO8")))
       'xcodigo(8) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(8) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio8")), "", RsExcell.Fields("fecha de inicio8")))
       xfecha2(8) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE8")), "", RsExcell.Fields("fecha de CESE8")))
       xmonto(8) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable8")), "", RsExcell.Fields("Remuneracion asegurable8")))
       xcomen(8) = Mostrar_Comentario(Xcodigo(8), xfecha1(8), xfecha2(8), xmonto(8))
       xFere(8) = Mostrar_Fecha(Xcodigo(8), xfecha1(8), xfecha2(8), xmonto(8))
       
       If Trim$(Xcodigo(8)) = "" Then
       Xcodigo(8) = Codigo_Nullo(Xcodigo(8), xfecha1(8), xfecha2(8))
       End If
       
       CAMPO(8) = xcampo
       If Len(Trim$(Xcodigo(8))) > 0 Or Len(Trim$(xfecha1(8))) > 0 Or Len(Trim$(xfecha2(8))) > 0 Or Len(Trim$(xmonto(8))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(9) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO9")), "", RsExcell.Fields("CODIGO9")))
       'xcodigo(9) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(9) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio9")), "", RsExcell.Fields("fecha de inicio9")))
       xfecha2(9) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE9")), "", RsExcell.Fields("fecha de CESE9")))
       xmonto(9) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable9")), "", RsExcell.Fields("Remuneracion asegurable9")))
       xcomen(9) = Mostrar_Comentario(Xcodigo(9), xfecha1(9), xfecha2(9), xmonto(9))
       xFere(9) = Mostrar_Fecha(Xcodigo(9), xfecha1(9), xfecha2(9), xmonto(9))
       If Trim$(Xcodigo(9)) = "" Then
       Xcodigo(9) = Codigo_Nullo(Xcodigo(9), xfecha1(9), xfecha2(9))
       End If
       
       CAMPO(9) = xcampo
       
       If Len(Trim$(Xcodigo(9))) > 0 Or Len(Trim$(xfecha1(9))) > 0 Or Len(Trim$(xfecha2(9))) > 0 Or Len(Trim$(xmonto(9))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(10) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO10")), "", RsExcell.Fields("CODIGO10")))
       'xcodigo(10) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(10) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio10")), "", RsExcell.Fields("fecha de inicio10")))
       xfecha2(10) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE10")), "", RsExcell.Fields("fecha de CESE10")))
       xmonto(10) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable10")), "", RsExcell.Fields("Remuneracion asegurable10")))
       xcomen(10) = Mostrar_Comentario(Xcodigo(10), xfecha1(10), xfecha2(10), xmonto(10))
       xFere(10) = Mostrar_Fecha(Xcodigo(10), xfecha1(10), xfecha2(10), xmonto(10))
       
       If Trim$(Xcodigo(10)) = "" Then
       Xcodigo(10) = Codigo_Nullo(Xcodigo(10), xfecha1(10), xfecha2(10))
       End If
       
       CAMPO(10) = xcampo
       If Len(Trim$(Xcodigo(10))) > 0 Or Len(Trim$(xfecha1(10))) > 0 Or Len(Trim$(xfecha2(10))) > 0 Or Len(Trim$(xmonto(10))) > 0 Then
       IntContador = IntContador + 1
       End If
              
       Xcodigo(11) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO11")), "", RsExcell.Fields("CODIGO11")))
       'xcodigo(11) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(11) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio11")), "", RsExcell.Fields("fecha de inicio11")))
       xfecha2(11) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE11")), "", RsExcell.Fields("fecha de CESE11")))
       xmonto(11) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable11")), "", RsExcell.Fields("Remuneracion asegurable11")))
       xcomen(11) = Mostrar_Comentario(Xcodigo(11), xfecha1(11), xfecha2(11), xmonto(11))
       xFere(11) = Mostrar_Fecha(Xcodigo(11), xfecha1(11), xfecha2(11), xmonto(11))
       
       If Trim$(Xcodigo(11)) = "" Then
       Xcodigo(11) = Codigo_Nullo(Xcodigo(11), xfecha1(11), xfecha2(11))
       End If
       
       CAMPO(11) = xcampo
       If Len(Trim$(Xcodigo(11))) > 0 Or Len(Trim$(xfecha1(11))) > 0 Or Len(Trim$(xfecha2(11))) > 0 Or Len(Trim$(xmonto(11))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(12) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO12")), "", RsExcell.Fields("CODIGO12")))
       'xcodigo(12) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(12) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio12")), "", RsExcell.Fields("fecha de inicio12")))
       xfecha2(12) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE12")), "", RsExcell.Fields("fecha de CESE12")))
       xmonto(12) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable12")), "", RsExcell.Fields("Remuneracion asegurable12")))
       xcomen(12) = Mostrar_Comentario(Xcodigo(12), xfecha1(12), xfecha2(12), xmonto(12))
       xFere(12) = Mostrar_Fecha(Xcodigo(12), xfecha1(12), xfecha2(12), xmonto(12))
       
       If Trim$(Xcodigo(12)) = "" Then
       Xcodigo(12) = Codigo_Nullo(Xcodigo(12), xfecha1(12), xfecha2(12))
       End If
       
       CAMPO(12) = xcampo
       If Len(Trim$(Xcodigo(12))) > 0 Or Len(Trim$(xfecha1(12))) > 0 Or Len(Trim$(xfecha2(12))) > 0 Or Len(Trim$(xmonto(12))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       Xcodigo(13) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO13")), "", RsExcell.Fields("CODIGO13")))
       'xcodigo(13) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(13) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio13")), "", RsExcell.Fields("fecha de inicio13")))
       xfecha2(13) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE13")), "", RsExcell.Fields("fecha de CESE13")))
       xmonto(13) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable13")), "", RsExcell.Fields("Remuneracion asegurable13")))
       xcomen(13) = Mostrar_Comentario(Xcodigo(13), xfecha1(13), xfecha2(13), xmonto(13))
       xFere(13) = Mostrar_Fecha(Xcodigo(13), xfecha1(13), xfecha2(13), xmonto(13))
       
       If Trim$(Xcodigo(13)) = "" Then
       Xcodigo(13) = Codigo_Nullo(Xcodigo(13), xfecha1(13), xfecha2(13))
       End If
       
       CAMPO(13) = xcampo
       If Len(Trim$(Xcodigo(13))) > 0 Or Len(Trim$(xfecha1(13))) > 0 Or Len(Trim$(xfecha2(13))) > 0 Or Len(Trim$(xmonto(13))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       Xcodigo(14) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO14")), "", RsExcell.Fields("CODIGO14")))
       'xcodigo(14) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(14) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio14")), "", RsExcell.Fields("fecha de inicio14")))
       xfecha2(14) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE14")), "", RsExcell.Fields("fecha de CESE14")))
       xmonto(14) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable14")), "", RsExcell.Fields("Remuneracion asegurable14")))
       xcomen(14) = Mostrar_Comentario(Xcodigo(14), xfecha1(14), xfecha2(14), xmonto(14))
       xFere(14) = Mostrar_Fecha(Xcodigo(14), xfecha1(14), xfecha2(14), xmonto(14))
       
       If Trim$(Xcodigo(14)) = "" Then
       Xcodigo(14) = Codigo_Nullo(Xcodigo(14), xfecha1(14), xfecha2(14))
       End If
       
       CAMPO(14) = xcampo
        If Len(Trim$(Xcodigo(14))) > 0 Or Len(Trim$(xfecha1(14))) > 0 Or Len(Trim$(xfecha2(14))) > 0 Or Len(Trim$(xmonto(14))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(15) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO15")), "", RsExcell.Fields("CODIGO15")))
       'xcodigo(15) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(15) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio15")), "", RsExcell.Fields("fecha de inicio15")))
       xfecha2(15) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE15")), "", RsExcell.Fields("fecha de CESE15")))
       xmonto(15) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable15")), "", RsExcell.Fields("Remuneracion asegurable15")))
       xcomen(15) = Mostrar_Comentario(Xcodigo(15), xfecha1(15), xfecha2(15), xmonto(15))
       xFere(15) = Mostrar_Fecha(Xcodigo(15), xfecha1(15), xfecha2(15), xmonto(15))
       
       If Trim$(Xcodigo(15)) = "" Then
       Xcodigo(15) = Codigo_Nullo(Xcodigo(15), xfecha1(15), xfecha2(15))
       End If
       
       CAMPO(15) = xcampo
        If Len(Trim$(Xcodigo(15))) > 0 Or Len(Trim$(xfecha1(15))) > 0 Or Len(Trim$(xfecha2(15))) > 0 Or Len(Trim$(xmonto(15))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(16) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO16")), "", RsExcell.Fields("CODIGO16")))
       'xcodigo(16) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(16) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio16")), "", RsExcell.Fields("fecha de inicio16")))
       xfecha2(16) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE16")), "", RsExcell.Fields("fecha de CESE16")))
       xmonto(16) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable16")), "", RsExcell.Fields("Remuneracion asegurable16")))
       xcomen(16) = Mostrar_Comentario(Xcodigo(16), xfecha1(16), xfecha2(16), xmonto(16))
       xFere(16) = Mostrar_Fecha(Xcodigo(16), xfecha1(16), xfecha2(16), xmonto(16))
       
       If Trim$(Xcodigo(16)) = "" Then
       Xcodigo(16) = Codigo_Nullo(Xcodigo(16), xfecha1(16), xfecha2(16))
       End If
       
       CAMPO(16) = xcampo
       If Len(Trim$(Xcodigo(16))) > 0 Or Len(Trim$(xfecha1(16))) > 0 Or Len(Trim$(xfecha2(16))) > 0 Or Len(Trim$(xmonto(16))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(17) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO17")), "", RsExcell.Fields("CODIGO17")))
       'xcodigo(17) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(17) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio17")), "", RsExcell.Fields("fecha de inicio17")))
       xfecha2(17) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE17")), "", RsExcell.Fields("fecha de CESE17")))
       xmonto(17) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable17")), "", RsExcell.Fields("Remuneracion asegurable17")))
       xcomen(17) = Mostrar_Comentario(Xcodigo(17), xfecha1(17), xfecha2(17), xmonto(17))
       xFere(17) = Mostrar_Fecha(Xcodigo(17), xfecha1(17), xfecha2(17), xmonto(17))
       
       If Trim$(Xcodigo(17)) = "" Then
       Xcodigo(17) = Codigo_Nullo(Xcodigo(17), xfecha1(17), xfecha2(17))
       End If
       
       CAMPO(17) = xcampo
       If Len(Trim$(Xcodigo(17))) > 0 Or Len(Trim$(xfecha1(17))) > 0 Or Len(Trim$(xfecha2(17))) > 0 Or Len(Trim$(xmonto(17))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(18) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO18")), "", RsExcell.Fields("CODIGO18")))
       'xcodigo(18) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(18) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio18")), "", RsExcell.Fields("fecha de inicio18")))
       xfecha2(18) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE18")), "", RsExcell.Fields("fecha de CESE18")))
       xmonto(18) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable18")), "", RsExcell.Fields("Remuneracion asegurable18")))
       xcomen(18) = Mostrar_Comentario(Xcodigo(18), xfecha1(18), xfecha2(18), xmonto(18))
       xFere(18) = Mostrar_Fecha(Xcodigo(18), xfecha1(18), xfecha2(18), xmonto(18))
       
       If Trim$(Xcodigo(18)) = "" Then
       Xcodigo(18) = Codigo_Nullo(Xcodigo(18), xfecha1(18), xfecha2(18))
       End If
       
       CAMPO(18) = xcampo
       If Len(Trim$(Xcodigo(18))) > 0 Or Len(Trim$(xfecha1(18))) > 0 Or Len(Trim$(xfecha2(18))) > 0 Or Len(Trim$(xmonto(18))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       Xcodigo(19) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO19")), "", RsExcell.Fields("CODIGO19")))
       'xcodigo(19) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(19) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio19")), "", RsExcell.Fields("fecha de inicio19")))
       xfecha2(19) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE19")), "", RsExcell.Fields("fecha de CESE19")))
       xmonto(19) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable19")), "", RsExcell.Fields("Remuneracion asegurable19")))
       xcomen(19) = Mostrar_Comentario(Xcodigo(19), xfecha1(19), xfecha2(19), xmonto(19))
       xFere(19) = Mostrar_Fecha(Xcodigo(19), xfecha1(19), xfecha2(19), xmonto(19))
       
       If Trim$(Xcodigo(19)) = "" Then
       Xcodigo(19) = Codigo_Nullo(Xcodigo(19), xfecha1(19), xfecha2(19))
       End If
       
       CAMPO(19) = xcampo
       If Len(Trim$(Xcodigo(19))) > 0 Or Len(Trim$(xfecha1(19))) > 0 Or Len(Trim$(xfecha2(19))) > 0 Or Len(Trim$(xmonto(19))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       Xcodigo(20) = Trim$(IIf(IsNull(RsExcell.Fields("CODIGO20")), "", RsExcell.Fields("CODIGO20")))
       'xcodigo(20) = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO20")), "", RsExcell.Fields("CAMPO20")))
       xfecha1(20) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de inicio20")), "", RsExcell.Fields("fecha de inicio20")))
       xfecha2(20) = Trim$(IIf(IsNull(RsExcell.Fields("fecha de CESE20")), "", RsExcell.Fields("fecha de CESE20")))
       xmonto(20) = Trim$(IIf(IsNull(RsExcell.Fields("Remuneracion asegurable20")), "", RsExcell.Fields("Remuneracion asegurable20")))
       xcomen(20) = Mostrar_Comentario(Xcodigo(20), xfecha1(20), xfecha2(20), xmonto(20))
       xFere(20) = Mostrar_Fecha(Xcodigo(20), xfecha1(20), xfecha2(20), xmonto(20))
       
       If Trim$(Xcodigo(20)) = "" Then
       Xcodigo(20) = Codigo_Nullo(Xcodigo(20), xfecha1(20), xfecha2(20))
       End If
       
       CAMPO(20) = xcampo
       If Len(Trim$(Xcodigo(20))) > 0 Or Len(Trim$(xfecha1(20))) > 0 Or Len(Trim$(xfecha2(20))) > 0 Or Len(Trim$(xmonto(20))) > 0 Then
       IntContador = IntContador + 1
       End If
       
       
       
       'Xrazon = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO6")), "", RsExcell.Fields("CAMPO6")))
       'XRUC = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO7")), "", RsExcell.Fields("CAMPO7")))
       'XPERIODO = Trim$(IIf(IsNull(RsExcell.Fields("CAMPO19")), "", RsExcell.Fields("CAMPO19")))
       'Xcomentario = ""
       'XFECHA = ""
       
       
          For X = 1 To IntContador
          
          Print #3, xruc; vbTab; Xrazon; vbTab; xcomen(X); vbTab; Format(Xcodigo(X), "00"); vbTab; xcuspp(X); vbTab; Xperiodo; vbTab; xFere(X); vbTab; CAMPO(X)
          Next X
          RsExcell.MoveNext
       Next Y
    Set RsExcell = Nothing
  End If
End If

Close #3
Error_Handling_Generar_Archivo:
End Sub


Public Function Mostrar_Fecha(Xcodigo As String, xinicio As String, xfin As String, xmonto As String) As String

Select Case Format(Xcodigo, "00")
  Case "":
   If Len(xinicio) > 0 Then
   
       If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
         If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
          
          If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
              Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
          Else
              Mostrar_Fecha = ""
          End If
         Else
          Mostrar_Fecha = ""
         End If
       Else
          If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
          
            If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
               Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
            Else
               Mostrar_Fecha = ""
            End If
          
          Else
          Mostrar_Fecha = ""
          End If
       End If
       
   End If
   If Len(xfin) > 0 Then
    'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
    
       If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
         If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
             If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
             Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
             Else
             Mostrar_Fecha = ""
             End If
         Else
          Mostrar_Fecha = ""
         End If
       Else
          If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
            If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
            Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
            Else
            Mostrar_Fecha = ""
            End If
          Else
          Mostrar_Fecha = ""
          End If
       End If
    
   End If
    
    If Len(xinicio) > 0 And Len(xfin) > 0 Then
    'Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
    
       If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
         If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
            If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
             Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
             Else
             Mostrar_Fecha = ""
            End If
         Else
          Mostrar_Fecha = ""
         End If
       Else
          If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
            If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
             Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
            Else
             Mostrar_Fecha = ""
            End If
          Else
          Mostrar_Fecha = ""
          End If
       End If
    
    End If
    
    If Len(xinicio) = 0 And Len(xfin) = 0 Then
       If Len(xmonto) > 0 Then
       Mostrar_Fecha = "" '"RIA"
       End If
     End If


  Case "01":
        
        If Len(xinicio) > 0 Then
         'Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
                  If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
                     If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                     Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
                     Else
                     Mostrar_Fecha = ""
                     End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
                    If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                    Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                    Else
                    Mostrar_Fecha = ""
                    End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
        End If
        
        If Len(xfin) > 0 Then
         'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
'                If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
'                  If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
'                   Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
'                  Else
'                   Mostrar_Fecha = ""
'                  End If
'                Else
'                   If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
'                   Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
'                   Else
'                   Mostrar_Fecha = ""
'                   End If
'                End If
                 Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
         Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
          'Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                If Val(Mid(xinicio, 5, 2)) > 70 And Val(Mid(xinicio, 5, 2)) <= 99 Then
                  If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)) Then
                      If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                      Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "19" & Mid(xinicio, 5, 2)
                      Else
                      Mostrar_Fecha = ""
                      End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)) Then
                      If Val(Mid(xinicio, 3, 2)) > 0 And Val(Mid(xinicio, 3, 2)) <= 12 Then
                      Mostrar_Fecha = Mid(xinicio, 1, 2) & "/" & Mid(xinicio, 3, 2) & "/" & "20" & Mid(xinicio, 5, 2)
                      Else
                      Mostrar_Fecha = ""
                      End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
          
        End If
        
        If Len(xinicio) = 4 Then
        'Mostrar_Fecha = "01" & "/" & Mid(xinicio, 1, 2) & "/" & "20" & Mid(xinicio, 3, 2)
           xdia = "01" 'DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
           If Val(Mid(xinicio, 3, 2)) > 70 And Val(Mid(xinicio, 3, 2)) <= 99 Then
             If IsDate(xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)) Then
                If Val(Mid(xinicio, 1, 2)) > 0 And Val(Mid(xinicio, 1, 2)) <= 12 Then
                Mostrar_Fecha = xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)
                Else
                Mostrar_Fecha = ""
                End If
             Else
             Mostrar_Fecha = ""
             End If
           Else
             xdia = "01" 'DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
             If Val(Mid(xinicio, 3, 2)) > 70 And Val(Mid(xinicio, 3, 2)) <= 99 Then
               If IsDate(xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)) Then
                If Val(Mid(xinicio, 1, 2)) > 0 And Val(Mid(xinicio, 1, 2)) <= 12 Then
                Mostrar_Fecha = xdia & "/" & Mid(xinicio, 1, 2) & "/" & "19" & Mid(xinicio, 3, 2)
                Else
                Mostrar_Fecha = ""
                End If
               Else
                Mostrar_Fecha = ""
               End If
             Else
               If IsDate(xdia & "/" & Mid(xinicio, 1, 2) & "/" & "20" & Mid(xinicio, 3, 2)) Then
                    If Val(Mid(xinicio, 1, 2)) > 0 And Val(Mid(xinicio, 1, 2)) <= 12 Then
                    Mostrar_Fecha = xdia & "/" & Mid(xinicio, 1, 2) & "/" & "20" & Mid(xinicio, 3, 2)
                    Else
                    Mostrar_Fecha = ""

                    End If
               Else
                Mostrar_Fecha = ""
               End If
             End If
           End If
        End If
        
        If Len(xfin) = 4 Then
        Mostrar_Fecha = ""
        End If
        
  Case "02":
        
        If Len(xfin) > 0 Then
          'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
               If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
                  If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                       If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                           If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                           Mostrar_Fecha = ""
                           Else
                           Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
                           End If
                            
                       Else
                       Mostrar_Fecha = ""
                       End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                       
                       If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                           If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                           Mostrar_Fecha = ""
                           Else
                           Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
                           End If
                       Else
                       Mostrar_Fecha = ""
                       End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
        End If
        
        If Len(xinicio) > 0 Then
        Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
         Mostrar_Fecha = ""
        End If
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
         'Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
              If Val(Mid(xfin, 5, 2)) > 70 And Val(Mid(xfin, 5, 2)) <= 99 Then
                  If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                         If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                             If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)) Then
                             Mostrar_Fecha = ""
                             Else
                             Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "19" & Mid(xfin, 5, 2)
                             End If
                         Else
                         Mostrar_Fecha = ""
                         End If
                  Else
                   Mostrar_Fecha = ""
                  End If
                Else
                   If IsDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                         If Val(Mid(xfin, 3, 2)) > 0 And Val(Mid(xfin, 3, 2)) <= 12 Then
                             If Format(Now, "dd/mm/yyyy") < CDate(Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)) Then
                             Mostrar_Fecha = ""
                             Else
                             Mostrar_Fecha = Mid(xfin, 1, 2) & "/" & Mid(xfin, 3, 2) & "/" & "20" & Mid(xfin, 5, 2)
                             End If
                         Else
                         Mostrar_Fecha = ""
                         End If
                   Else
                   Mostrar_Fecha = ""
                   End If
                End If
        End If
        
        If Len(xinicio) = 4 And Len(xfin) = 4 Then
         'xdia = DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
         'Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 4)
         
           xdia = DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
           If Val(Mid(xfin, 3, 2)) > 70 And Val(Mid(xfin, 3, 2)) <= 99 Then
             If IsDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                 If Val(Mid(xfin, 1, 2)) > 0 And Val(Mid(xfin, 1, 2)) <= 12 Then
                     
                     If Format(Now, "dd/mm/yyyy") < CDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                     Mostrar_Fecha = ""
                     Else
                     Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)
                     End If
                 Else
                 Mostrar_Fecha = ""
                 End If
             Else
             Mostrar_Fecha = ""
             End If
           Else
             xdia = DateSerial(Year(Format(xfin, "yyyy")), Month(Format(xfin, "MM")) + 1, 0)
             If Val(Mid(xfin, 3, 2)) > 70 And Val(Mid(xfin, 3, 2)) <= 99 Then
               If IsDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                    
                    If Val(Mid(xfin, 1, 2)) > 0 And Val(Mid(xfin, 1, 2)) <= 12 Then
                    
                        If Format(Now, "dd/mm/yyyy") < CDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)) Then
                        Mostrar_Fecha = ""
                        Else
                        Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "19" & Mid(xfin, 3, 2)
                        End If
                    
                    
                    Else
                    Mostrar_Fecha = ""
                    End If
               Else
                Mostrar_Fecha = ""
               End If
             Else
               If IsDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 2)) Then
                    If Val(Mid(xfin, 1, 2)) > 0 And Val(Mid(xfin, 1, 2)) <= 12 Then
                        
                        If Format(Now, "dd/mm/yyyy") < CDate(xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 2)) Then
                        Mostrar_Fecha = ""
                        Else
                        Mostrar_Fecha = xdia & "/" & Mid(xfin, 1, 2) & "/" & "20" & Mid(xfin, 3, 2)
                        End If
                    Else
                    Mostrar_Fecha = ""
                    End If
               Else
                Mostrar_Fecha = ""
               End If
             End If
           End If
         
        End If

Case "04":
        If Len(xinicio) > 0 Or Len(xfin) > 0 Then
         Mostrar_Fecha = ""
        End If
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Fecha = ""
        End If
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        End If
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Fecha = ""
        End If

Case "05":
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        End If
        
        Case Xcodigo = "18"
        
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
        Mostrar_Fecha = ""
        End If

Case "20":
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        End If

Case "21":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        'End If
Case "22":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""

Case "RIA":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""

Case "99":
        
        'If Len(xinicio) > 0 And Len(xfin) > 0 Then
        Mostrar_Fecha = ""
        
        If Len(xmonto) > 0 Then
        Mostrar_Fecha = "" '"RIA"
        End If
        If Len(xmonto) = 0 Then
        Mostrar_Fecha = ""
        End If

End Select
End Function


Function Codigo_Nullo(Xcodigo As String, xinicio As String, xfin As String) As String
Select Case Format(Xcodigo, "00")
  Case "":
        If Len(xinicio) > 0 Then
         'Mostrar_Comentario = "INICIO DE RELACION LABORAL"
         Codigo_Nullo = "01"
        End If
        
        If Len(xfin) > 0 Then
         'Mostrar_Comentario = "CESE DE RELACION LABORAL"
        Codigo_Nullo = "02"
        End If
        
        If Len(xinicio) > 0 And Len(xfin) > 0 Then
         'Mostrar_Comentario = "INICIO DE RELACION LABORAL"
         xcampo = 2
         Codigo_Nullo = "01"
        End If
                
        If Len(xinicio) = 0 And Len(xfin) = 0 Then
          If Len(xmonto) > 0 Then
          'Mostrar_Comentario = "ACTUALIZACION DE RIA"
          Codigo_Nullo = "99"
          End If
        End If
  End Select
End Function


Public Sub Abrir_Excell_IMPRENTA()
Dim Xrep As String
Dim Xrep1 As String
Dim Xnomdoc As String
Dim Xtipodoc As String
'Dim XCuspp As String
Dim Xhojas As String
Dim Xcarpeta As String
Dim Xpaterno As String
Dim Xmaterno As String
Dim Xnombres1 As String
Dim Xnombres2 As String
Dim Xcortes As String
Dim Ximp As String
Dim Xnumpag As String
Dim Xnumpre As String
Dim Xsecafi As String
Dim Xtipmov As String
Dim Xfecmov As String
Dim xremase As String
Dim xobltra As String
Dim xoblemp As String
Dim xobltot As String
Dim xvolcfi As String
Dim xvolsfi As String
Dim xvolemp As String
Dim xtotfon As String
Dim xseguro As String
Dim Xcompor As String
Dim xtotadm As String
Dim Xcampos As Integer
Dim Xposicion As Double
Dim X As Integer
Dim Y As Integer
Dim IntContador  As Integer
Dim r1 As String
Dim r2 As String
Dim r3 As String
Dim r4 As String
Dim r5 As String
Dim r6 As String
Dim r7 As String
Dim r8 As String
Dim r9 As String
Dim r10 As String
Dim r11 As String
Dim RsExcell As New Recordset
On Error GoTo Error_Handling_Generar_Archivo

With RsExcell
.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & RutaArchivoExcel & _
                ";Extended Properties=""Excel 8.0;"""
                '";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""
                
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.Open "Select  * from [" & xhoja & "$] where corre_med  is not null"

End With
'Xposicion = XposIni

If RsExcell.RecordCount > 0 Then
  If RsExcell.RecordCount > 0 Then
     'For X = 1 To RsExcell.RecordCount
          
       For Y = 1 To RsExcell.RecordCount
       
       r1 = Trim$(IIf(IsNull(RsExcell.Fields("corre_med")), "", RsExcell.Fields("corre_med")))
       r2 = Trim$(IIf(IsNull(RsExcell.Fields("cod_mod")), "", RsExcell.Fields("cod_mod")))
       r3 = Trim$(IIf(IsNull(RsExcell.Fields("anexo")), "", RsExcell.Fields("anexo")))
       r4 = Trim$(IIf(IsNull(RsExcell.Fields("ie")), "", RsExcell.Fields("ie")))
       r5 = Trim$(IIf(IsNull(RsExcell.Fields("dir_ie")), "", RsExcell.Fields("dir_ie")))
       r6 = Trim$(IIf(IsNull(RsExcell.Fields("region")), "", RsExcell.Fields("region")))
       r7 = Trim$(IIf(IsNull(RsExcell.Fields("provincia")), "", RsExcell.Fields("provincia")))
       r8 = Trim$(IIf(IsNull(RsExcell.Fields("distrito")), "", RsExcell.Fields("distrito")))
       r9 = Trim$(IIf(IsNull(RsExcell.Fields("cen_pob")), "", RsExcell.Fields("cen_pob")))
       r10 = Trim$(IIf(IsNull(RsExcell.Fields("aplicador")), "", RsExcell.Fields("aplicador")))
       r11 = Trim$(IIf(IsNull(RsExcell.Fields("alum")), "", RsExcell.Fields("alum")))
       
       
       
       Grabar_Sql_IMPRENTA r1, r2, r3, r4, r5, _
              r6, r7, r8, r9, r10, r11
       
       RsExcell.MoveNext
       Next Y
   End If
End If

Error_Handling_Generar_Archivo:
End Sub

Private Sub Optbajopago_Click()
If Optbajopago.Value = True Then
    xxEstado = "1"
End If

End Sub

Private Sub optconpago_Click()
If optconpago.Value = True Then
    xxEstado = "2"
End If

End Sub

Private Sub optsinpago_Click()
    If optsinpago.Value = True Then
    xxEstado = "3"
    End If
End Sub
Public Sub Insertar_Datos()
Dim CmdDatos As New ADODB.Command
With CmdDatos
            .CommandText = "sp_InsertarDatos"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .ActiveConnection = cnServer.ConnectionString
            .Parameters.Append .CreateParameter("NROLIQUI", adVarChar, adParamInput, 50, Trim$(StrNroliquidacion_h))
            .Parameters.Append .CreateParameter("PERIODO", adVarChar, adParamInput, 50, Trim$(StrPeriodo_h))
            .Parameters.Append .CreateParameter("FECHA_EMISION", adVarChar, adParamInput, 50, Trim$(StrFecha_Emision_h))
            .Parameters.Append .CreateParameter("RAZON_SOCIAL", adVarChar, adParamInput, 50, Mid(Trim$(StrRazon_Social_h), 1, 50))
            .Parameters.Append .CreateParameter("RUC", adVarChar, adParamInput, 50, Trim$(StrRuc_h))
            .Parameters.Append .CreateParameter("DIRECCION", adVarChar, adParamInput, 150, Trim$(StrDireccion_h))
            .Parameters.Append .CreateParameter("DISTRITO", adVarChar, adParamInput, 50, Trim$(StrDistrito_h))
            .Parameters.Append .CreateParameter("PROVINCIA", adVarChar, adParamInput, 50, Trim$(StrProvincia_h))
            .Parameters.Append .CreateParameter("DEPARTAMENTO", adVarChar, adParamInput, 50, Trim$(StrDepartamento_h))
            .Parameters.Append .CreateParameter("TOTFONDO", adVarChar, adParamInput, 50, Trim$(StrTotFondo_h))
            .Parameters.Append .CreateParameter("MORATORIO", adVarChar, adParamInput, 50, Trim$(StrMoratorio_h))
            .Parameters.Append .CreateParameter("ADEUDADO", adVarChar, adParamInput, 50, Trim$(StrAdeudado_h))
            .Parameters.Append .CreateParameter("CUSPP1", adVarChar, adParamInput, 50, Trim$(cuspp_h(1)))
            .Parameters.Append .CreateParameter("APE_PAT1", adVarChar, adParamInput, 50, Trim$(ApellPater_h(1)))
            .Parameters.Append .CreateParameter("APE_MAT1", adVarChar, adParamInput, 50, Trim$(ApellMater_h(1)))
            .Parameters.Append .CreateParameter("NOMBRE1_1", adVarChar, adParamInput, 50, Trim$(NombrePri_h(1)))
            .Parameters.Append .CreateParameter("NOMBRE2_1", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(1)))
            .Parameters.Append .CreateParameter("REMU1", adVarChar, adParamInput, 50, Trim$(RemAseg_h(1)))
            .Parameters.Append .CreateParameter("FONDO1", adVarChar, adParamInput, 50, Trim$(FondoPens_h(1)))
            .Parameters.Append .CreateParameter("RETEN1", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(1)))
            .Parameters.Append .CreateParameter("TOTRET1", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(1)))
            .Parameters.Append .CreateParameter("TOTFOND1", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(1)))
            .Parameters.Append .CreateParameter("CUSPP2", adVarChar, adParamInput, 50, Trim$(cuspp_h(2)))
            .Parameters.Append .CreateParameter("APE_PAT2", adVarChar, adParamInput, 50, Trim$(ApellPater_h(2)))
            .Parameters.Append .CreateParameter("APE_MAT2", adVarChar, adParamInput, 50, Trim$(ApellMater_h(2)))
            .Parameters.Append .CreateParameter("NOMBRE1_2", adVarChar, adParamInput, 50, Trim$(NombrePri_h(2)))
            .Parameters.Append .CreateParameter("NOMBRE2_2", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(2)))
            .Parameters.Append .CreateParameter("REMU2", adVarChar, adParamInput, 50, Trim$(RemAseg_h(2)))
            .Parameters.Append .CreateParameter("FONDO2", adVarChar, adParamInput, 50, Trim$(FondoPens_h(2)))
            .Parameters.Append .CreateParameter("RETEN2", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(2)))
            .Parameters.Append .CreateParameter("TOTRET2", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(2)))
            .Parameters.Append .CreateParameter("TOTFOND2", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(2)))
            .Parameters.Append .CreateParameter("CUSPP3", adVarChar, adParamInput, 50, Trim$(cuspp_h(3)))
            .Parameters.Append .CreateParameter("APE_PAT3", adVarChar, adParamInput, 50, Trim$(ApellPater_h(3)))
            .Parameters.Append .CreateParameter("APE_MAT3", adVarChar, adParamInput, 50, Trim$(ApellMater_h(3)))
            .Parameters.Append .CreateParameter("NOMBRE1_3", adVarChar, adParamInput, 50, Trim$(NombrePri_h(3)))
            .Parameters.Append .CreateParameter("NOMBRE2_3", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(3)))
            .Parameters.Append .CreateParameter("REMU3", adVarChar, adParamInput, 50, Trim$(RemAseg_h(3)))
            .Parameters.Append .CreateParameter("FONDO3", adVarChar, adParamInput, 50, Trim$(FondoPens_h(3)))
            .Parameters.Append .CreateParameter("RETEN3", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(3)))
            .Parameters.Append .CreateParameter("TOTRET3", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(3)))
            .Parameters.Append .CreateParameter("TOTFOND3", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(3)))
            .Parameters.Append .CreateParameter("CUSPP4", adVarChar, adParamInput, 50, Trim$(cuspp_h(4)))
            .Parameters.Append .CreateParameter("APE_PAT4", adVarChar, adParamInput, 50, Trim$(ApellPater_h(4)))
            .Parameters.Append .CreateParameter("APE_MAT4", adVarChar, adParamInput, 50, Trim$(ApellMater_h(4)))
            .Parameters.Append .CreateParameter("NOMBRE1_4", adVarChar, adParamInput, 50, Trim$(NombrePri_h(4)))
            .Parameters.Append .CreateParameter("NOMBRE2_4", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(4)))
            .Parameters.Append .CreateParameter("REMU4", adVarChar, adParamInput, 50, Trim$(RemAseg_h(4)))
            .Parameters.Append .CreateParameter("FONDO4", adVarChar, adParamInput, 50, Trim$(FondoPens_h(4)))
            .Parameters.Append .CreateParameter("RETEN4", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(4)))
            .Parameters.Append .CreateParameter("TOTRET4", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(4)))
            .Parameters.Append .CreateParameter("TOTFOND4", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(4)))
            .Parameters.Append .CreateParameter("CUSPP5", adVarChar, adParamInput, 50, Trim$(cuspp_h(5)))
            .Parameters.Append .CreateParameter("APE_PAT51", adVarChar, adParamInput, 50, Trim$(ApellPater_h(5)))
            .Parameters.Append .CreateParameter("APE_MAT5", adVarChar, adParamInput, 50, Trim$(ApellMater_h(5)))
            .Parameters.Append .CreateParameter("NOMBRE1_5", adVarChar, adParamInput, 50, Trim$(NombrePri_h(5)))
            .Parameters.Append .CreateParameter("NOMBRE2_5", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(5)))
            .Parameters.Append .CreateParameter("REMU5", adVarChar, adParamInput, 50, Trim$(RemAseg_h(5)))
            .Parameters.Append .CreateParameter("FONDO5", adVarChar, adParamInput, 50, Trim$(FondoPens_h(5)))
            .Parameters.Append .CreateParameter("RETEN5", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(5)))
            .Parameters.Append .CreateParameter("TOTRET5", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(5)))
            .Parameters.Append .CreateParameter("TOTFOND5", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(5)))
            .Parameters.Append .CreateParameter("CUSPP6", adVarChar, adParamInput, 50, Trim$(cuspp_h(6)))
            .Parameters.Append .CreateParameter("APE_PAT6", adVarChar, adParamInput, 50, Trim$(ApellPater_h(6)))
            .Parameters.Append .CreateParameter("APE_MAT6", adVarChar, adParamInput, 50, Trim$(ApellMater_h(6)))
            .Parameters.Append .CreateParameter("NOMBRE1_6", adVarChar, adParamInput, 50, Trim$(NombrePri_h(6)))
            .Parameters.Append .CreateParameter("NOMBRE2_6", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(6)))
            .Parameters.Append .CreateParameter("REMU6", adVarChar, adParamInput, 50, Trim$(RemAseg_h(6)))
            .Parameters.Append .CreateParameter("FONDO6", adVarChar, adParamInput, 50, Trim$(FondoPens_h(6)))
            .Parameters.Append .CreateParameter("RETEN6", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(6)))
            .Parameters.Append .CreateParameter("TOTRET6", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(6)))
            .Parameters.Append .CreateParameter("TOTFOND6", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(6)))
            .Parameters.Append .CreateParameter("CUSPP7", adVarChar, adParamInput, 50, Trim$(cuspp_h(7)))
            .Parameters.Append .CreateParameter("APE_PAT7", adVarChar, adParamInput, 50, Trim$(ApellPater_h(7)))
            .Parameters.Append .CreateParameter("APE_MAT7", adVarChar, adParamInput, 50, Trim$(ApellMater_h(7)))
            .Parameters.Append .CreateParameter("NOMBRE1_7", adVarChar, adParamInput, 50, Trim$(NombrePri_h(7)))
            .Parameters.Append .CreateParameter("NOMBRE2_7", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(7)))
            .Parameters.Append .CreateParameter("REMU7", adVarChar, adParamInput, 50, Trim$(RemAseg_h(7)))
            .Parameters.Append .CreateParameter("FONDO7", adVarChar, adParamInput, 50, Trim$(FondoPens_h(7)))
            .Parameters.Append .CreateParameter("RETEN7", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(7)))
            .Parameters.Append .CreateParameter("TOTRET7", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(7)))
            .Parameters.Append .CreateParameter("TOTFOND7", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(7)))
            .Parameters.Append .CreateParameter("CUSPP8", adVarChar, adParamInput, 50, Trim$(cuspp_h(8)))
            .Parameters.Append .CreateParameter("APE_PAT8", adVarChar, adParamInput, 50, Trim$(ApellPater_h(8)))
            .Parameters.Append .CreateParameter("APE_MAT8", adVarChar, adParamInput, 50, Trim$(ApellMater_h(8)))
            .Parameters.Append .CreateParameter("NOMBRE1_8", adVarChar, adParamInput, 50, Trim$(NombrePri_h(8)))
            .Parameters.Append .CreateParameter("NOMBRE2_8", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(8)))
            .Parameters.Append .CreateParameter("REMU8", adVarChar, adParamInput, 50, Trim$(RemAseg_h(8)))
            .Parameters.Append .CreateParameter("FONDO8", adVarChar, adParamInput, 50, Trim$(FondoPens_h(8)))
            .Parameters.Append .CreateParameter("RETEN8", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(8)))
            .Parameters.Append .CreateParameter("TOTRET8", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(8)))
            .Parameters.Append .CreateParameter("TOTFOND8", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(8)))
            .Parameters.Append .CreateParameter("CUSPP9", adVarChar, adParamInput, 50, Trim$(cuspp_h(9)))
            .Parameters.Append .CreateParameter("APE_PAT9", adVarChar, adParamInput, 50, Trim$(ApellPater_h(9)))
            .Parameters.Append .CreateParameter("APE_MAT9", adVarChar, adParamInput, 50, Trim$(ApellMater_h(9)))
            .Parameters.Append .CreateParameter("NOMBRE1_9", adVarChar, adParamInput, 50, Trim$(NombrePri_h(9)))
            .Parameters.Append .CreateParameter("NOMBRE2_9", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(9)))
            .Parameters.Append .CreateParameter("REMU9", adVarChar, adParamInput, 50, Trim$(RemAseg_h(9)))
            .Parameters.Append .CreateParameter("FONDO9", adVarChar, adParamInput, 50, Trim$(FondoPens_h(9)))
            .Parameters.Append .CreateParameter("RETEN9", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(9)))
            .Parameters.Append .CreateParameter("TOTRET9", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(9)))
            .Parameters.Append .CreateParameter("TOTFOND9", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(9)))
            .Parameters.Append .CreateParameter("CUSPP10", adVarChar, adParamInput, 50, Trim$(cuspp_h(10)))
            .Parameters.Append .CreateParameter("APE_PAT10", adVarChar, adParamInput, 50, Trim$(ApellPater_h(10)))
            .Parameters.Append .CreateParameter("APE_MAT10", adVarChar, adParamInput, 50, Trim$(ApellMater_h(10)))
            .Parameters.Append .CreateParameter("NOMBRE1_10", adVarChar, adParamInput, 50, Trim$(NombrePri_h(10)))
            .Parameters.Append .CreateParameter("NOMBRE2_10", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(10)))
            .Parameters.Append .CreateParameter("REMU10", adVarChar, adParamInput, 50, Trim$(RemAseg_h(10)))
            .Parameters.Append .CreateParameter("FONDO10", adVarChar, adParamInput, 50, Trim$(FondoPens_h(10)))
            .Parameters.Append .CreateParameter("RETEN10", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(10)))
            .Parameters.Append .CreateParameter("TOTRET10", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(10)))
            .Parameters.Append .CreateParameter("TOTFOND10", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(10)))
            .Parameters.Append .CreateParameter("CUSPP11", adVarChar, adParamInput, 50, Trim$(cuspp_h(11)))
            .Parameters.Append .CreateParameter("APE_PAT11", adVarChar, adParamInput, 50, Trim$(ApellPater_h(11)))
            .Parameters.Append .CreateParameter("APE_MAT11", adVarChar, adParamInput, 50, Trim$(ApellMater_h(11)))
            .Parameters.Append .CreateParameter("NOMBRE1_11", adVarChar, adParamInput, 50, Trim$(NombrePri_h(11)))
            .Parameters.Append .CreateParameter("NOMBRE2_11", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(11)))
            .Parameters.Append .CreateParameter("REMU11", adVarChar, adParamInput, 50, Trim$(RemAseg_h(11)))
            .Parameters.Append .CreateParameter("FONDO11", adVarChar, adParamInput, 50, Trim$(FondoPens_h(11)))
            .Parameters.Append .CreateParameter("RETEN11", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(11)))
            .Parameters.Append .CreateParameter("TOTRET11", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(11)))
            .Parameters.Append .CreateParameter("TOTFOND11", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(11)))
            .Parameters.Append .CreateParameter("CUSPP1", adVarChar, adParamInput, 50, Trim$(cuspp_h(12)))
            .Parameters.Append .CreateParameter("APE_PAT12", adVarChar, adParamInput, 50, Trim$(ApellPater_h(12)))
            .Parameters.Append .CreateParameter("APE_MAT12", adVarChar, adParamInput, 50, Trim$(ApellMater_h(12)))
            .Parameters.Append .CreateParameter("NOMBRE1_12", adVarChar, adParamInput, 50, Trim$(NombrePri_h(12)))
            .Parameters.Append .CreateParameter("NOMBRE2_12", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(12)))
            .Parameters.Append .CreateParameter("REMU12", adVarChar, adParamInput, 50, Trim$(RemAseg_h(12)))
            .Parameters.Append .CreateParameter("FONDO12", adVarChar, adParamInput, 50, Trim$(FondoPens_h(12)))
            .Parameters.Append .CreateParameter("RETEN12", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(12)))
            .Parameters.Append .CreateParameter("TOTRET12", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(12)))
            .Parameters.Append .CreateParameter("TOTFOND12", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(12)))
            .Parameters.Append .CreateParameter("CUSPP13", adVarChar, adParamInput, 50, Trim$(cuspp_h(13)))
            .Parameters.Append .CreateParameter("APE_PAT13", adVarChar, adParamInput, 50, Trim$(ApellPater_h(13)))
            .Parameters.Append .CreateParameter("APE_MAT13", adVarChar, adParamInput, 50, Trim$(ApellMater_h(13)))
            .Parameters.Append .CreateParameter("NOMBRE1_13", adVarChar, adParamInput, 50, Trim$(NombrePri_h(13)))
            .Parameters.Append .CreateParameter("NOMBRE2_13", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(13)))
            .Parameters.Append .CreateParameter("REMU13", adVarChar, adParamInput, 50, Trim$(RemAseg_h(13)))
            .Parameters.Append .CreateParameter("FONDO13", adVarChar, adParamInput, 50, Trim$(FondoPens_h(13)))
            .Parameters.Append .CreateParameter("RETEN13", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(13)))
            .Parameters.Append .CreateParameter("TOTRET13", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(13)))
            .Parameters.Append .CreateParameter("TOTFOND13", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(13)))
            .Parameters.Append .CreateParameter("CUSPP14", adVarChar, adParamInput, 50, Trim$(cuspp_h(14)))
            .Parameters.Append .CreateParameter("APE_PAT14", adVarChar, adParamInput, 50, Trim$(ApellPater_h(14)))
            .Parameters.Append .CreateParameter("APE_MAT14", adVarChar, adParamInput, 50, Trim$(ApellMater_h(14)))
            .Parameters.Append .CreateParameter("NOMBRE1_14", adVarChar, adParamInput, 50, Trim$(NombrePri_h(14)))
            .Parameters.Append .CreateParameter("NOMBRE2_14", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(14)))
            .Parameters.Append .CreateParameter("REMU14", adVarChar, adParamInput, 50, Trim$(RemAseg_h(14)))
            .Parameters.Append .CreateParameter("FONDO14", adVarChar, adParamInput, 50, Trim$(FondoPens_h(14)))
            .Parameters.Append .CreateParameter("RETEN14", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(14)))
            .Parameters.Append .CreateParameter("TOTRET14", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(14)))
            .Parameters.Append .CreateParameter("TOTFOND14", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(14)))
            .Parameters.Append .CreateParameter("CUSPP15", adVarChar, adParamInput, 50, Trim$(cuspp_h(15)))
            .Parameters.Append .CreateParameter("APE_PAT15", adVarChar, adParamInput, 50, Trim$(ApellPater_h(15)))
            .Parameters.Append .CreateParameter("APE_MAT15", adVarChar, adParamInput, 50, Trim$(ApellMater_h(15)))
            .Parameters.Append .CreateParameter("NOMBRE1_15", adVarChar, adParamInput, 50, Trim$(NombrePri_h(15)))
            .Parameters.Append .CreateParameter("NOMBRE2_15", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(15)))
            .Parameters.Append .CreateParameter("REMU15", adVarChar, adParamInput, 50, Trim$(RemAseg_h(15)))
            .Parameters.Append .CreateParameter("FONDO15", adVarChar, adParamInput, 50, Trim$(FondoPens_h(15)))
            .Parameters.Append .CreateParameter("RETEN15", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(15)))
            .Parameters.Append .CreateParameter("TOTRET15", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(15)))
            .Parameters.Append .CreateParameter("TOTFOND15", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(15)))
            .Parameters.Append .CreateParameter("CUSPP16", adVarChar, adParamInput, 50, Trim$(cuspp_h(16)))
            .Parameters.Append .CreateParameter("APE_PAT16", adVarChar, adParamInput, 50, Trim$(ApellPater_h(16)))
            .Parameters.Append .CreateParameter("APE_MAT16", adVarChar, adParamInput, 50, Trim$(ApellMater_h(16)))
            .Parameters.Append .CreateParameter("NOMBRE1_16", adVarChar, adParamInput, 50, Trim$(NombrePri_h(16)))
            .Parameters.Append .CreateParameter("NOMBRE2_16", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(16)))
            .Parameters.Append .CreateParameter("REMU16", adVarChar, adParamInput, 50, Trim$(RemAseg_h(16)))
            .Parameters.Append .CreateParameter("FONDO16", adVarChar, adParamInput, 50, Trim$(FondoPens_h(16)))
            .Parameters.Append .CreateParameter("RETEN16", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(16)))
            .Parameters.Append .CreateParameter("TOTRET16", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(16)))
            .Parameters.Append .CreateParameter("TOTFOND16", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(16)))
            .Parameters.Append .CreateParameter("CUSPP17", adVarChar, adParamInput, 50, Trim$(cuspp_h(17)))
            .Parameters.Append .CreateParameter("APE_PAT17", adVarChar, adParamInput, 50, Trim$(ApellPater_h(17)))
            .Parameters.Append .CreateParameter("APE_MAT17", adVarChar, adParamInput, 50, Trim$(ApellMater_h(17)))
            .Parameters.Append .CreateParameter("NOMBRE1_17", adVarChar, adParamInput, 50, Trim$(NombrePri_h(17)))
            .Parameters.Append .CreateParameter("NOMBRE2_17", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(17)))
            .Parameters.Append .CreateParameter("REMU17", adVarChar, adParamInput, 50, Trim$(RemAseg_h(17)))
            .Parameters.Append .CreateParameter("FONDO17", adVarChar, adParamInput, 50, Trim$(FondoPens_h(17)))
            .Parameters.Append .CreateParameter("RETEN17", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(17)))
            .Parameters.Append .CreateParameter("TOTRET17", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(17)))
            .Parameters.Append .CreateParameter("TOTFOND17", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(17)))
            .Parameters.Append .CreateParameter("CUSPP18", adVarChar, adParamInput, 50, Trim$(cuspp_h(18)))
            .Parameters.Append .CreateParameter("APE_PAT18", adVarChar, adParamInput, 50, Trim$(ApellPater_h(18)))
            .Parameters.Append .CreateParameter("APE_MAT18", adVarChar, adParamInput, 50, Trim$(ApellMater_h(18)))
            .Parameters.Append .CreateParameter("NOMBRE1_18", adVarChar, adParamInput, 50, Trim$(NombrePri_h(18)))
            .Parameters.Append .CreateParameter("NOMBRE2_18", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(18)))
            .Parameters.Append .CreateParameter("REMU18", adVarChar, adParamInput, 50, Trim$(RemAseg_h(18)))
            .Parameters.Append .CreateParameter("FONDO18", adVarChar, adParamInput, 50, Trim$(FondoPens_h(18)))
            .Parameters.Append .CreateParameter("RETEN18", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(18)))
            .Parameters.Append .CreateParameter("TOTRET18", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(18)))
            .Parameters.Append .CreateParameter("TOTFOND18", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(18)))
            .Parameters.Append .CreateParameter("CUSPP19", adVarChar, adParamInput, 50, Trim$(cuspp_h(19)))
            .Parameters.Append .CreateParameter("APE_PAT19", adVarChar, adParamInput, 50, Trim$(ApellPater_h(19)))
            .Parameters.Append .CreateParameter("APE_MAT19", adVarChar, adParamInput, 50, Trim$(ApellMater_h(19)))
            .Parameters.Append .CreateParameter("NOMBRE1_19", adVarChar, adParamInput, 50, Trim$(NombrePri_h(19)))
            .Parameters.Append .CreateParameter("NOMBRE2_19", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(19)))
            .Parameters.Append .CreateParameter("REMU19", adVarChar, adParamInput, 50, Trim$(RemAseg_h(19)))
            .Parameters.Append .CreateParameter("FONDO19", adVarChar, adParamInput, 50, Trim$(FondoPens_h(19)))
            .Parameters.Append .CreateParameter("RETEN19", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(19)))
            .Parameters.Append .CreateParameter("TOTRET19", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(19)))
            .Parameters.Append .CreateParameter("TOTFOND19", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(19)))
            .Parameters.Append .CreateParameter("CUSPP20", adVarChar, adParamInput, 50, Trim$(cuspp_h(20)))
            .Parameters.Append .CreateParameter("APE_PAT20", adVarChar, adParamInput, 50, Trim$(ApellPater_h(20)))
            .Parameters.Append .CreateParameter("APE_MAT20", adVarChar, adParamInput, 50, Trim$(ApellMater_h(20)))
            .Parameters.Append .CreateParameter("NOMBRE1_20", adVarChar, adParamInput, 50, Trim$(NombrePri_h(20)))
            .Parameters.Append .CreateParameter("NOMBRE2_20", adVarChar, adParamInput, 50, Trim$(NombreSeg_h(20)))
            .Parameters.Append .CreateParameter("REMU20", adVarChar, adParamInput, 50, Trim$(RemAseg_h(20)))
            .Parameters.Append .CreateParameter("FONDO20", adVarChar, adParamInput, 50, Trim$(FondoPens_h(20)))
            .Parameters.Append .CreateParameter("RETEN20", adVarChar, adParamInput, 50, Trim$(RetenRetr_h(20)))
            .Parameters.Append .CreateParameter("TOTRET20", adVarChar, adParamInput, 50, Trim$(TotalRetension_h(20)))
            .Parameters.Append .CreateParameter("TOTFOND20", adVarChar, adParamInput, 50, Trim$(TotalFonPens_h(20)))
            .Parameters.Append .CreateParameter("BARRA", adVarChar, adParamInput, 50, Trim$(strBarra_h))
            .Parameters.Append .CreateParameter("ESTADO", adVarChar, adParamInput, 50, Trim$(strEstado_h))
            .Parameters.Append .CreateParameter("NROLIQPAG", adVarChar, adParamInput, 50, Trim$(strNroliqpag_h))
            .Parameters.Append .CreateParameter("PROCESADO", adVarChar, adParamInput, 1, "0")

            .Execute
End With
            Set CmdDatos = Nothing

End Sub


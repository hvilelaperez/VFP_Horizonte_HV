VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmprincipal 
   Caption         =   "Liquidación Integra"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton Command2 
         Caption         =   "&Procesar"
         Height          =   495
         Left            =   6000
         TabIndex        =   7
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton CmbProcesar 
         Caption         =   "&Procesar"
         Height          =   495
         Left            =   6000
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   6000
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   480
         TabIndex        =   3
         Top             =   3480
         Width           =   7095
         Begin ComctlLib.ProgressBar Prbar 
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   873
            _Version        =   327682
            Appearance      =   1
         End
      End
      Begin VB.CommandButton cmdproceso 
         Caption         =   "&Procesar"
         Height          =   495
         Left            =   6000
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbProcesar_Click()
''Records para Archivo
Dim Rstcob4 As Recordset
Dim Rstcob5 As Recordset
Dim Rstcob6 As Recordset
'''Records para totales Generales
Dim RstTotal As Recordset
''''''

'''''''''''''
Dim Cantdet As Double
'

Prbar.Min = 0
'Prbar.Max = 0
 ''' Archivo Liquidación
 Open "C:\prueba" & "\Liquidacion.txt" For Output As #3
 ''''Archivo Carta
 Open "C:\prueba" & "\Carta.txt" For Output As #4

''''''''''''''''''''''''''''''''
Set Rstcob4 = Sel_Int_Cabezera11
If Rstcob4.RecordCount > 0 Then
Rstcob4.MoveLast
Rstcob4.MoveFirst
  ''Armar Encabezado
    
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
                              ; vbTab; "CAMPO133"; vbTab; "CAMPO134"; vbTab; "CAMPO135"; vbTab; "CAMPO136"; vbTab; "CAMPO137"; vbTab; "CAMPO138"; vbTab; "CAMPO139"

    
    
    Print #4, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
                              ; vbTab; "CAMPO11"; vbTab; "CAMPO12"
    

  For X = 1 To Rstcob4.RecordCount
     Prbar.Min = 0
     Prbar.Max = Rstcob4.RecordCount
     
     Prbar.Value = X + 1
     FechaLiq = Trim$(Rstcob4.Fields("cobfe5"))
     NroLiq = Trim$(Trim$(Rstcob4.Fields("cobti5")) & "-" & Trim$(Rstcob4.Fields("cobnme")))
     Razon = Trim$(IIf(IsNull(Rstcob4.Fields("cobra1")), "", Rstcob4.Fields("cobra1")))
     Ruc = Trim$(IIf(IsNull(Rstcob4.Fields("cobnr1")), "", Rstcob4.Fields("cobnr1")))
     Direct = Trim$(Rstcob4.Fields("cobdir"))
     Atencion = Trim$(IIf(IsNull(Rstcob4.Fields("cobate")), "", Rstcob4.Fields("cobate")))
     Referencia = Trim$(IIf(IsNull(Rstcob4.Fields("cobref")), "", Rstcob4.Fields("cobref")))
     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("cobadj")), "", Rstcob4.Fields("cobadj"))) & ":" & Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin")))
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
      
     ' End If
  '''Grabar El Archivo Texto de Carta
      Formatear_Archivo_Texto_Carta
      Print #4, TXTC(1); vbTab; TXTC(2); vbTab; TXTC(3); vbTab; TXTC(4); vbTab; TXTC(5); vbTab; TXTC(6); vbTab; TXTC(7); vbTab; TXTC(8); vbTab; TXTC(9); vbTab; TXTC(10) _
                              ; vbTab; TXTC(11); vbTab; TXTC(12)
  
  '''''''
  '''''Consulta Cobar6
     'Set Rstcob6 = Sel_Int_Detalle(Rstcob4.Fields("cobnme"))
     Set Rstcob6 = Sel_Int_Detalle_Liq(Rstcob4.Fields("cobnme"), Rstcob4.Fields("cobNM2"))
     
     If Rstcob6.RecordCount > 0 Then
      Rstcob6.MoveLast
      Rstcob6.MoveFirst
      '''Determinar cuantas hojas se imprimen por Razon Social
          Cantdet = 0
         If Rstcob6.RecordCount > 20 Then
            Cantdet = Int(Rstcob6.RecordCount / 20)
           If Rstcob6.RecordCount - (Cantdet * 20) > 0 Then
            Cantdet = Cantdet + 1
           Else
            Cantdet = Cantdet
           End If
         Else
            Cantdet = 1
         End If
                 
           Set RstTotal = Sel_Int_Totales_Liq(Rstcob4.Fields("cobnme"), Rstcob4.Fields("cobNM2"))
           If RstTotal.RecordCount > 0 Then
           TotalFPrr = Trim$(RstTotal.Fields("cobto1"))
           TotalInteMo = Trim$(RstTotal.Fields("cobto2"))
           TotalAd = Trim$(RstTotal.Fields("cobto3"))
           'Pagfinal = Trim(RstTotal.Fields("cobpag"))
           'Pag = Pag & Pagfinal
           End If
           
           For p = 1 To Cantdet
           Pagfinal = ""
           Set Rstcob6 = Sel_Int_Detalle_Pag_Liq(Rstcob4.Fields("cobnme"), Format(p, "000"), Rstcob4.Fields("cobNM2"))
              Rstcob6.MoveLast
              Rstcob6.MoveFirst
             If Rstcob6.RecordCount >= p * 20 Then
                For Y = 1 To 20 'Rstcob6.RecordCount
                 Cuspp(Y) = Trim$(Rstcob6.Fields("cobcus"))
                 ApellNomb(Y) = Trim$(Rstcob6.Fields("cobape") & " " & Rstcob6.Fields("cobap1") & "/" & Rstcob6.Fields("cobno6") & " " & Rstcob6.Fields("cobno7"))
                 RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                 FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                 RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                 TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                 Rstcob6.MoveNext
                Next Y
             Else
                For Y = 1 To Rstcob6.RecordCount
                 Cuspp(Y) = Trim$(Rstcob6.Fields("cobcus"))
                 ApellNomb(Y) = Trim$(Rstcob6.Fields("cobape") & " " & IIf(IsNull(Rstcob6.Fields("cobap1")), "", Rstcob6.Fields("cobap1")) & "/" & Trim$(Rstcob6.Fields("cobno6")) & " " & IIf(IsNull(Rstcob6.Fields("cobno7")), "", Rstcob6.Fields("cobno7")))
                 RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                 FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                 RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                 TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                 Rstcob6.MoveNext
                Next Y
                'Cantdet = 20 - Rstcob6.RecordCount
                'If P * 20 < Rstcob6.RecordCount Then
                    For H = Rstcob6.RecordCount + 1 To 20
                     Cuspp(H) = ""
                     ApellNomb(H) = ""
                     RemAseg(H) = ""
                     FondoPens(H) = ""
                     RetenRetr(H) = ""
                    Next H
                'End If
             End If
           ''''''''''''''''''
           'Pagfinal = Trim(Rstcob6.Fields("cobpag"))
           Pagina = Pag & Format(p, "000")
           If Format(p, "000") = Pagfinal Then
             
           End If
           ''''''CREAR ARCHIVO TEXTO LIQUIDACION PREVIA INTEGRA
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
                   ''''Formatear campos de Archivo texto
                     Formatear_Archivo_Texto
        '           Crear_Archivo_Liquidacion FechaLiq, Pag, NroLiq, Razon, Ruc, Direct, Ndepart, Urbani, Distrito, _
        '           Provincia, Departamento, TotalFPrr,,TotalInteMo,TotalAd,,Liquipre,,PeriDeve,
                   
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
                              ; vbTab; TXT(133); vbTab; TXT(134); vbTab; TXT(135); vbTab; TXT(136); vbTab; TXT(137); vbTab; TXT(138); vbTab; TXT(139)
                    
                      
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
           Next p
      ''''
     End If
  ''''
  Rstcob4.MoveNext
  Next X
End If
Close #3
Close #4

End Sub

Private Sub cmdproceso_Click()
''Records para Archivo
Dim Rstcob4 As Recordset
Dim Rstcob5 As Recordset
Dim Rstcob6 As Recordset
'''Records para totales Generales
Dim RstTotal As Recordset
''''''

'''''''''''''
Dim Cantdet As Double
'

Prbar.Min = 0
'Prbar.Max = 0
 ''' Archivo Liquidación
 Open "C:\prueba" & "\Liquidacion.txt" For Output As #3
 ''''Archivo Carta
 Open "C:\prueba" & "\Carta.txt" For Output As #4

''''''''''''''''''''''''''''''''
Set Rstcob4 = Sel_Int_Cabezera1
If Rstcob4.RecordCount > 0 Then
Rstcob4.MoveLast
Rstcob4.MoveFirst
  ''Armar Encabezado
    
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
                              ; vbTab; "CAMPO133"; vbTab; "CAMPO134"; vbTab; "CAMPO135"; vbTab; "CAMPO136"; vbTab; "CAMPO137"; vbTab; "CAMPO138"; vbTab; "CAMPO139"

    
    
    Print #4, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
                              ; vbTab; "CAMPO11"; vbTab; "CAMPO12"
    

  For X = 1 To Rstcob4.RecordCount
     Prbar.Min = 0
     Prbar.Max = Rstcob4.RecordCount
     
     Prbar.Value = X + 1
     FechaLiq = Trim$(Rstcob4.Fields("cobfe5"))
     NroLiq = Trim$(Trim$(Rstcob4.Fields("cobti5")) & "-" & Trim$(Rstcob4.Fields("cobnme")))
     Razon = Trim$(IIf(IsNull(Rstcob4.Fields("cobra1")), "", Rstcob4.Fields("cobra1")))
     Ruc = Trim$(IIf(IsNull(Rstcob4.Fields("cobnr1")), "", Rstcob4.Fields("cobnr1")))
     Direct = Trim$(Rstcob4.Fields("cobdir"))
     Atencion = Trim$(IIf(IsNull(Rstcob4.Fields("cobate")), "", Rstcob4.Fields("cobate")))
     Referencia = Trim$(IIf(IsNull(Rstcob4.Fields("cobref")), "", Rstcob4.Fields("cobref")))
     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("cobadj")), "", Rstcob4.Fields("cobadj"))) & ":" & Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin")))
  '''''Consulta Cobar 5
     Set Rstcob5 = Sel_Int_Cabezera2(Rstcob4.Fields("cobnme"))
      If Rstcob5.RecordCount > 0 Then
      Ndepart = Trim$(IIf(IsNull(Rstcob5.Fields("cobdi3")), "", Rstcob5.Fields("cobdi3")))
      Urbani = Trim$(IIf(IsNull(Rstcob5.Fields("cobdi4")), "", Rstcob5.Fields("cobdi4")))
      Distrito = Trim$(Rstcob5.Fields("cobno3"))
      Provincia = Trim$(Rstcob5.Fields("cobno4"))
      Departamento = Trim$(Rstcob5.Fields("cobno5"))
      Pag = Trim$(Rstcob5.Fields("Cobpgi"))
      
      Liquipre = Trim$(Rstcob5.Fields("cobnm2"))
      PeriDeve = Trim$(Rstcob5.Fields("cobpe2"))
      
      End If
  '''Grabar El Archivo Texto de Carta
      Formatear_Archivo_Texto_Carta
      Print #4, TXTC(1); vbTab; TXTC(2); vbTab; TXTC(3); vbTab; TXTC(4); vbTab; TXTC(5); vbTab; TXTC(6); vbTab; TXTC(7); vbTab; TXTC(8); vbTab; TXTC(9); vbTab; TXTC(10) _
                              ; vbTab; TXTC(11); vbTab; TXTC(12)
  
  '''''''
  '''''Consulta Cobar6
     Set Rstcob6 = Sel_Int_Detalle(Rstcob4.Fields("cobnme"))
     If Rstcob6.RecordCount > 0 Then
      Rstcob6.MoveLast
      Rstcob6.MoveFirst
      '''Determinar cuantas hojas se imprimen por Razon Social
          Cantdet = 0
         If Rstcob6.RecordCount > 20 Then
            Cantdet = Int(Rstcob6.RecordCount / 20)
           If Rstcob6.RecordCount - (Cantdet * 20) > 0 Then
            Cantdet = Cantdet + 1
           Else
            Cantdet = Cantdet
           End If
         Else
            Cantdet = 1
         End If
                 
           Set RstTotal = Sel_Int_Totales(Rstcob4.Fields("cobnme"))
           If RstTotal.RecordCount > 0 Then
           TotalFPrr = Trim$(RstTotal.Fields("cobto1"))
           TotalInteMo = Trim$(RstTotal.Fields("cobto2"))
           TotalAd = Trim$(RstTotal.Fields("cobto3"))
           'Pagfinal = Trim(RstTotal.Fields("cobpag"))
           'Pag = Pag & Pagfinal
           End If
           
           For p = 1 To Cantdet
           Pagfinal = ""
           Set Rstcob6 = Sel_Int_Detalle_Pag(Rstcob4.Fields("cobnme"), Format(p, "000"))
              Rstcob6.MoveLast
              Rstcob6.MoveFirst
             If Rstcob6.RecordCount >= p * 20 Then
                For Y = 1 To 20 'Rstcob6.RecordCount
                 Cuspp(Y) = Trim$(Rstcob6.Fields("cobcus"))
                 ApellNomb(Y) = Trim$(Rstcob6.Fields("cobape") & " " & Rstcob6.Fields("cobap1") & "/" & Rstcob6.Fields("cobno6") & " " & Rstcob6.Fields("cobno7"))
                 RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                 FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                 RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                 TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                 Rstcob6.MoveNext
                Next Y
             Else
                For Y = 1 To Rstcob6.RecordCount
                 Cuspp(Y) = Trim$(Rstcob6.Fields("cobcus"))
                 ApellNomb(Y) = Trim$(Rstcob6.Fields("cobape") & " " & IIf(IsNull(Rstcob6.Fields("cobap1")), "", Rstcob6.Fields("cobap1")) & "/" & Trim$(Rstcob6.Fields("cobno6")) & " " & IIf(IsNull(Rstcob6.Fields("cobno7")), "", Rstcob6.Fields("cobno7")))
                 RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                 FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                 RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                 TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                 Rstcob6.MoveNext
                Next Y
                'Cantdet = 20 - Rstcob6.RecordCount
                'If P * 20 < Rstcob6.RecordCount Then
                    For H = Rstcob6.RecordCount + 1 To 20
                     Cuspp(H) = ""
                     ApellNomb(H) = ""
                     RemAseg(H) = ""
                     FondoPens(H) = ""
                     RetenRetr(H) = ""
                    Next H
                'End If
             End If
           ''''''''''''''''''
           'Pagfinal = Trim(Rstcob6.Fields("cobpag"))
           Pagina = Pag & Format(p, "000")
           If Format(p, "000") = Pagfinal Then
             
           End If
           ''''''CREAR ARCHIVO TEXTO LIQUIDACION PREVIA INTEGRA
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
                   ''''Formatear campos de Archivo texto
                     Formatear_Archivo_Texto
        '           Crear_Archivo_Liquidacion FechaLiq, Pag, NroLiq, Razon, Ruc, Direct, Ndepart, Urbani, Distrito, _
        '           Provincia, Departamento, TotalFPrr,,TotalInteMo,TotalAd,,Liquipre,,PeriDeve,
                   
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
                              ; vbTab; TXT(133); vbTab; TXT(134); vbTab; TXT(135); vbTab; TXT(136); vbTab; TXT(137); vbTab; TXT(138); vbTab; TXT(139)
                    
                      
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
           Next p
      ''''
     End If
  ''''
  Rstcob4.MoveNext
  Next X
End If
Close #3
Close #4
End Sub

Private Sub Command1_Click()

''Records para Archivo
Dim Rstcob4 As Recordset
Dim Rstcob5 As Recordset
Dim Rstcob6 As Recordset
'''Records para totales Generales
Dim RstTotal As Recordset
''''''

'''''''''''''
Dim Cantdet As Double
'

Prbar.Min = 0
'Prbar.Max = 0
 Open "C:\prueba" & "\PRUEBAS.txt" For Output As #3
Set Rstcob4 = Sel_Int_Cabezera11
If Rstcob4.RecordCount > 0 Then
Rstcob4.MoveLast
Rstcob4.MoveFirst
  For X = 1 To Rstcob4.RecordCount
     Prbar.Min = 0
     Prbar.Max = Rstcob4.RecordCount
     
     Prbar.Value = X + 1
     FechaLiq = Rstcob4.Fields("cobar4.cobfe5")
     NroLiq = Rstcob4.Fields("cobti5") & "-" & Rstcob4.Fields("cobnme")
     Razon = IIf(IsNull(Rstcob4.Fields("cobra1")), "", Rstcob4.Fields("cobra1"))
     Ruc = IIf(IsNull(Rstcob4.Fields("cobnr1")), "", Rstcob4.Fields("cobnr1"))
     Direct = Rstcob4.Fields("cobdir")
  '''''Consulta Cobar 5
     Set Rstcob5 = Sel_Int_Cabezera2(Rstcob4.Fields("cobnme"))
      If Rstcob5.RecordCount > 0 Then
      Ndepart = IIf(IsNull(Rstcob5.Fields("cobdi3")), "", Rstcob5.Fields("cobdi3"))
      Urbani = IIf(IsNull(Rstcob5.Fields("cobdi4")), "", Rstcob5.Fields("cobdi4"))
      Distrito = Rstcob5.Fields("cobno3")
      Provincia = Rstcob5.Fields("cobno4")
      Departamento = Rstcob5.Fields("cobno5")
      Pag = Rstcob5.Fields("Cobpgi")
      
      Liquipre = Rstcob5.Fields("cobnm2")
      PeriDeve = Rstcob5.Fields("cobpe2")
      
      End If
  '''''''
  '''''Consulta Cobar6
     Set Rstcob6 = Sel_Int_Detalle(Rstcob4.Fields("cobnme"))
     If Rstcob6.RecordCount > 0 Then
      Rstcob6.MoveLast
      Rstcob6.MoveFirst
      '''Determinar cuantas hojas se imprimen por Razon Social
          Cantdet = 0
         If Rstcob6.RecordCount > 20 Then
            Cantdet = Int(Rstcob6.RecordCount / 20)
           If Rstcob6.RecordCount - (Cantdet * 20) > 0 Then
            Cantdet = Cantdet + 1
           Else
            Cantdet = Cantdet
           End If
         Else
            Cantdet = 1
         End If
                 
           Set RstTotal = Sel_Int_Totales(Rstcob4.Fields("cobnme"))
           If RstTotal.RecordCount > 0 Then
           TotalFPrr = RstTotal.Fields("cobto1")
           TotalInteMo = RstTotal.Fields("cobto2")
           TotalAd = RstTotal.Fields("cobto3")
           'Pagfinal = Trim(RstTotal.Fields("cobpag"))
           'Pag = Pag & Pagfinal
           End If
           
           For p = 1 To Cantdet
           Pagfinal = ""
           Set Rstcob6 = Sel_Int_Detalle_Pag(Rstcob4.Fields("cobnme"), Format(p, "000"))
              Rstcob6.MoveLast
              Rstcob6.MoveFirst
             If Rstcob6.RecordCount >= p * 20 Then
                For Y = 1 To 20 'Rstcob6.RecordCount
                 Cuspp(Y) = Rstcob6.Fields("cobcus")
                 ApellNomb(Y) = Rstcob6.Fields("cobape") & " " & Rstcob6.Fields("cobap1") & "/" & Rstcob6.Fields("cobno6") & " " & Rstcob6.Fields("cobno7")
                 RemAseg(Y) = Rstcob6.Fields("cobrem")
                 FondoPens(Y) = Rstcob6.Fields("cobmon")
                 RetenRetr(Y) = Rstcob6.Fields("cobmo1")
                 TotalFonPens(Y) = Rstcob6.Fields("cobtot")
                 Rstcob6.MoveNext
                Next Y
             Else
                For Y = 1 To Rstcob6.RecordCount
                 Cuspp(Y) = Rstcob6.Fields("cobcus")
                 ApellNomb(Y) = Rstcob6.Fields("cobape") & " " & Rstcob6.Fields("cobap1") & "/" & Rstcob6.Fields("cobno6") & " " & Rstcob6.Fields("cobno7")
                 RemAseg(Y) = Rstcob6.Fields("cobrem")
                 FondoPens(Y) = Rstcob6.Fields("cobmon")
                 RetenRetr(Y) = Rstcob6.Fields("cobmo1")
                 TotalFonPens(Y) = Rstcob6.Fields("cobtot")
                 Rstcob6.MoveNext
                Next Y
                'Cantdet = 20 - Rstcob6.RecordCount
                'If P * 20 < Rstcob6.RecordCount Then
                    For H = Rstcob6.RecordCount + 1 To 20
                     Cuspp(H) = ""
                     ApellNomb(H) = ""
                     RemAseg(H) = ""
                     FondoPens(H) = ""
                     RetenRetr(H) = ""
                    Next H
                'End If
             End If
           ''''''''''''''''''
           'Pagfinal = Trim(Rstcob6.Fields("cobpag"))
           Pagina = Pag & Format(p, "000")
           If Format(p, "000") = Pagfinal Then
             
           End If
           ''''''CREAR ARCHIVO TEXTO LIQUIDACION PREVIA INTEGRA
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
                   ''''Formatear campos de Archivo texto
                     Formatear_Archivo_Texto
        '           Crear_Archivo_Liquidacion FechaLiq, Pag, NroLiq, Razon, Ruc, Direct, Ndepart, Urbani, Distrito, _
        '           Provincia, Departamento, TotalFPrr,,TotalInteMo,TotalAd,,Liquipre,,PeriDeve,
                   
                    Print #3, TXT(1); vbTab; TXT(2); vbTab; TXT(3); vbTab; TXT(4); vbTab; TXT(4); vbTab; TXT(5); vbTab; TXT(6); vbTab; TXT(7); vbTab; TXT(8); vbTab; TXT(9) _
                              ; vbTab; TXT(10); vbTab; TXT(11); vbTab; TXT(12); vbTab; TXT(13); vbTab; TXT(14); vbTab; TXT(15); vbTab; TXT(16); vbTab; TXT(17); vbTab; TXT(18) _
                              ; vbTab; TXT(19); vbTab; TXT(20); vbTab; TXT(21); vbTab; TXT(22); vbTab; TXT(23); vbTab; TXT(24); vbTab; TXT(25); vbTab; TXT(26); vbTab; TXT(27); vbTab; TXT(27) _
                              ; vbTab; TXT(28); vbTab; TXT(29); vbTab; TXT(30); vbTab; TXT(31); vbTab; TXT(32); vbTab; TXT(33); vbTab; TXT(34); vbTab; TXT(35); vbTab; TXT(36); vbTab; TXT(37) _
                              ; vbTab; TXT(38); vbTab; TXT(39); vbTab; TXT(39); vbTab; TXT(40); vbTab; TXT(41); vbTab; TXT(42); vbTab; TXT(43); vbTab; TXT(44); vbTab; TXT(45); vbTab; TXT(46) _
                              ; vbTab; TXT(47); vbTab; TXT(48); vbTab; TXT(49); vbTab; TXT(50); vbTab; TXT(51); vbTab; TXT(52); vbTab; TXT(53); vbTab; TXT(54); vbTab; TXT(55); vbTab; TXT(56); vbTab; TXT(57) _
                              ; vbTab; TXT(58); vbTab; TXT(59); vbTab; TXT(60); vbTab; TXT(61); vbTab; TXT(62); vbTab; TXT(63); vbTab; TXT(64); vbTab; TXT(65); vbTab; TXT(66); vbTab; TXT(67); vbTab; TXT(68) _
                              ; vbTab; TXT(69); vbTab; TXT(70); vbTab; TXT(71); vbTab; TXT(72); vbTab; TXT(73); vbTab; TXT(74); vbTab; TXT(75); vbTab; TXT(76); vbTab; TXT(77); vbTab; TXT(78); vbTab; TXT(79) _
                              ; vbTab; TXT(80); vbTab; TXT(81); vbTab; TXT(82); vbTab; TXT(83); vbTab; TXT(84); vbTab; TXT(85); vbTab; TXT(86); vbTab; TXT(87); vbTab; TXT(88); vbTab; TXT(89); vbTab; TXT(90) _
                              ; vbTab; TXT(91); vbTab; TXT(92); vbTab; TXT(93); vbTab; TXT(94); vbTab; TXT(95); vbTab; TXT(96); vbTab; TXT(97); vbTab; TXT(98); vbTab; TXT(99); vbTab; TXT(100); vbTab; TXT(101) _
                              ; vbTab; TXT(102); vbTab; TXT(103); vbTab; TXT(104); vbTab; TXT(105); vbTab; TXT(106); vbTab; TXT(107); vbTab; TXT(108); vbTab; TXT(109); vbTab; TXT(110); vbTab; TXT(111) _
                              ; vbTab; TXT(112); vbTab; TXT(113); vbTab; TXT(114); vbTab; TXT(115); vbTab; TXT(116); vbTab; TXT(117); vbTab; TXT(118); vbTab; TXT(119); vbTab; TXT(120); vbTab; TXT(121) _
                              ; vbTab; TXT(122); vbTab; TXT(123); vbTab; TXT(124); vbTab; TXT(126); vbTab; TXT(127); vbTab; TXT(128); vbTab; TXT(129); vbTab; TXT(130); vbTab; TXT(131); vbTab; TXT(132) _
                              ; vbTab; TXT(133); vbTab; TXT(134); vbTab; TXT(135); vbTab; TXT(136); vbTab; TXT(137); vbTab; TXT(138)
                    
                      
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
           Next p
      ''''
     End If
  ''''
  Rstcob4.MoveNext
  Next X
End If
Close #3


End Sub

Private Sub Command2_Click()
''Records para Archivo
Dim Rstcob4 As Recordset
Dim Rstcob5 As Recordset
Dim Rstcob6 As Recordset
'''Records para totales Generales
Dim RstTotal As Recordset
''''''

'''''''''''''
Dim Cantdet As Double
'

Prbar.Min = 0
'Prbar.Max = 0
 ''' Archivo Liquidación
 Open "C:\prueba" & "\Liquidacion.txt" For Output As #3
 ''''Archivo Carta
 Open "C:\prueba" & "\Carta.txt" For Output As #4

''''''''''''''''''''''''''''''''
Set Rstcob4 = Sel_Int_Cabezera11
If Rstcob4.RecordCount > 0 Then
Rstcob4.MoveLast
Rstcob4.MoveFirst
  ''Armar Encabezado
    
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
                              ; vbTab; "CAMPO133"; vbTab; "CAMPO134"; vbTab; "CAMPO135"; vbTab; "CAMPO136"; vbTab; "CAMPO137"; vbTab; "CAMPO138"; vbTab; "CAMPO139"

    
    
    Print #4, "CAMPO1"; vbTab; "CAMPO2"; vbTab; "CAMPO3"; vbTab; "CAMPO4"; vbTab; "CAMPO5"; vbTab; "CAMPO6"; vbTab; "CAMPO7"; vbTab; "CAMPO8"; vbTab; "CAMPO9"; vbTab; "CAMPO10" _
                              ; vbTab; "CAMPO11"; vbTab; "CAMPO12"
    

  For X = 0 To Rstcob4.RecordCount
     Prbar.Min = 0
     Prbar.Max = Rstcob4.RecordCount
     
     Prbar.Value = X + 1
     FechaLiq = Trim$(Rstcob4.Fields("cobfe6"))
     NroLiq = Trim$(Trim$(Rstcob4.Fields("cobti7")) & "-" & Trim$(Rstcob4.Fields("cobnm1"))) & "-" & Trim$(Rstcob4.Fields("cobse2"))
     Razon = Trim$(IIf(IsNull(Rstcob4.Fields("cobra2")), "", Rstcob4.Fields("cobra2")))
     Ruc = Trim$(IIf(IsNull(Rstcob4.Fields("cobnr2")), "", Rstcob4.Fields("cobnr2")))
     Direct = Trim$(Rstcob4.Fields("cobdi2"))
     Atencion = Trim$(IIf(IsNull(Rstcob4.Fields("cobate")), "", Rstcob4.Fields("cobate")))
     Referencia = Trim$(IIf(IsNull(Rstcob4.Fields("cobref")), "", Rstcob4.Fields("cobref")))
     Adjunto = Trim$(IIf(IsNull(Rstcob4.Fields("cobadj")), "", Rstcob4.Fields("cobadj"))) & ":" & Trim$(IIf(IsNull(Rstcob4.Fields("coblin")), "", Rstcob4.Fields("coblin")))
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
      
     ' End If
  '''Grabar El Archivo Texto de Carta
      Formatear_Archivo_Texto_Carta
      Print #4, TXTC(1); vbTab; TXTC(2); vbTab; TXTC(3); vbTab; TXTC(4); vbTab; TXTC(5); vbTab; TXTC(6); vbTab; TXTC(7); vbTab; TXTC(8); vbTab; TXTC(9); vbTab; TXTC(10) _
                              ; vbTab; TXTC(11); vbTab; TXTC(12)
  
  '''''''
  '''''Consulta Cobar6
     'Set Rstcob6 = Sel_Int_Detalle(Rstcob4.Fields("cobnme"))
     Set Rstcob6 = Sel_Int_Detalle_Liq(Rstcob4.Fields("cobnm1"), Rstcob4.Fields("cobNM2"))
     
     If Rstcob6.RecordCount > 0 Then
      Rstcob6.MoveLast
      Rstcob6.MoveFirst
      '''Determinar cuantas hojas se imprimen por Razon Social
          Cantdet = 0
         If Rstcob6.RecordCount > 20 Then
            Cantdet = Int(Rstcob6.RecordCount / 20)
           If Rstcob6.RecordCount - (Cantdet * 20) > 0 Then
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
           Set Rstcob6 = Sel_Int_Detalle_Pag_Liq(Rstcob4.Fields("cobnm1"), Format(p, "000"), Rstcob4.Fields("cobNM2"))
              Rstcob6.MoveLast
              Rstcob6.MoveFirst
             If Rstcob6.RecordCount >= p * 20 Then
                For Y = 1 To 20 'Rstcob6.RecordCount
                 Cuspp(Y) = Trim$(Rstcob6.Fields("cobcus"))
                 ApellNomb(Y) = Trim$(Rstcob6.Fields("cobape") & " " & Rstcob6.Fields("cobap1") & "/" & Rstcob6.Fields("cobno6") & " " & Rstcob6.Fields("cobno7"))
                 RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                 FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                 RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                 TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                 Rstcob6.MoveNext
                Next Y
             Else
                For Y = 1 To Rstcob6.RecordCount
                 Cuspp(Y) = Trim$(Rstcob6.Fields("cobcus"))
                 ApellNomb(Y) = Trim$(Rstcob6.Fields("cobape") & " " & IIf(IsNull(Rstcob6.Fields("cobap1")), "", Rstcob6.Fields("cobap1")) & "/" & Trim$(Rstcob6.Fields("cobno6")) & " " & IIf(IsNull(Rstcob6.Fields("cobno7")), "", Rstcob6.Fields("cobno7")))
                 RemAseg(Y) = Trim$(Rstcob6.Fields("cobrem"))
                 FondoPens(Y) = Trim$(Rstcob6.Fields("cobmon"))
                 RetenRetr(Y) = Trim$(Rstcob6.Fields("cobmo1"))
                 TotalFonPens(Y) = Trim$(Rstcob6.Fields("cobtot"))
                 Rstcob6.MoveNext
                Next Y
                'Cantdet = 20 - Rstcob6.RecordCount
                'If P * 20 < Rstcob6.RecordCount Then
                    For H = Rstcob6.RecordCount + 1 To 20
                     Cuspp(H) = ""
                     ApellNomb(H) = ""
                     RemAseg(H) = ""
                     FondoPens(H) = ""
                     RetenRetr(H) = ""
                    Next H
                'End If
             End If
           ''''''''''''''''''
           'Pagfinal = Trim(Rstcob6.Fields("cobpag"))
           Pagina = Pag & Format(p, "000")
           If Format(p, "000") = Pagfinal Then
           Set RstTotal = Sel_Int_Totales_Liq(Rstcob4.Fields("cobnm1"), Rstcob4.Fields("cobNM2"))
                If RstTotal.RecordCount > 0 Then
                TotalFPrr = Trim$(RstTotal.Fields("cobto1"))
                TotalInteMo = Trim$(RstTotal.Fields("cobto2"))
                TotalAd = Trim$(RstTotal.Fields("cobto3"))
                'Pagfinal = Trim(RstTotal.Fields("cobpag"))
                'Pag = Pag & Pagfinal
                End If
           Else
           TotalFPrr = "" 'Trim$(RstTotal.Fields("cobto1"))
           TotalInteMo = "" 'Trim$(RstTotal.Fields("cobto2"))
           TotalAd = "" 'Trim$(RstTotal.Fields("cobto3"))
          
             
           End If
           ''''''CREAR ARCHIVO TEXTO LIQUIDACION PREVIA INTEGRA
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
                   ''''Formatear campos de Archivo texto
                     Formatear_Archivo_Texto
        '           Crear_Archivo_Liquidacion FechaLiq, Pag, NroLiq, Razon, Ruc, Direct, Ndepart, Urbani, Distrito, _
        '           Provincia, Departamento, TotalFPrr,,TotalInteMo,TotalAd,,Liquipre,,PeriDeve,
                   
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
                              ; vbTab; TXT(133); vbTab; TXT(134); vbTab; TXT(135); vbTab; TXT(136); vbTab; TXT(137); vbTab; TXT(138); vbTab; TXT(139)
                    
                      
           '''''''''''''''''''''''''''''''''''''''''''''''''''''
           Next p
      ''''
     End If
  ''''
  Rstcob4.MoveNext
  Next X
End If
Close #3
Close #4

End Sub

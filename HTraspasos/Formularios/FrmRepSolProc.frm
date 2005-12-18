VERSION 5.00
Begin VB.Form FrmRepSolProc 
   Caption         =   "Reporte de Solicitudes Procesadas"
   ClientHeight    =   5145
   ClientLeft      =   4740
   ClientTop       =   3960
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6330
   Begin VB.Frame Frame2 
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdRepError 
         Caption         =   "&Errores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton CmdRepo 
         Caption         =   "&Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton CmdRepoRes 
         Caption         =   "&Resumen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton CmdCruce 
         Caption         =   "&Cruce"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdProceso 
      Caption         =   "&Proceso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame FraFecSal 
      Height          =   3855
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      Begin VB.Frame Frame1 
         Caption         =   "Documento"
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
         Left            =   600
         TabIndex        =   6
         Top             =   2400
         Width           =   2055
         Begin VB.OptionButton OptCont 
            Caption         =   "Traspasos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton OptAfil 
            Caption         =   "Afiliaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.TextBox txtEnvio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtFechaB 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   0
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Envio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Ingrese Fecha de Salida o Nro de Envio :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   1995
      End
   End
End
Attribute VB_Name = "FrmRepSolProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsCruce As New Recordset       'Reporte Cruce
Public RsRepo As New Recordset        'Reporte Detalle
Public RsRepoRes As New Recordset     'Reporte Resumen
Public RsCtrlD As New Recordset      'CtrlDoc

Private Sub cmdRepError_Click()
    If RsRepo.State = 1 Then RsRepo.Close
    'SQLstr = "SELECT Solicitud.Fecsal, Solicitud.Lote, Numsol, CUSPP, ApePat, ApeMat, Nombre1, Empleadores, Cantemple, Codafpori, Solicitud.Status, Agencia , Envio FROM Solicitud, CtrlDoc WHERE Solicitud.Fecsal = '" & TxtFechaB & "' AND Solicitud.Lote = CtrlDoc.Lote AND Apepat is not null order by Solicitud.Lote, NumSol "
    'SQLstr = "SELECT Solicitud.Fecsal, Solicitud.Lote, Numsol, CUSPP, ApePat, ApeMat, Nombre1, Empleadores, Cantemple, Codafpori, Solicitud.Status, Agencia , Envio FROM Solicitud, CtrlDoc WHERE CTRLDoc.Envio = '" & txtenvio & "' AND Solicitud.Lote = CtrlDoc.Lote AND Apepat is not null order by Solicitud.Lote, NumSol "
    
    'Empleadores
    If OptAfil.Value Then
    
        SQLstr = " Select envio,agencia,c.fecsal,c.lote,numsol,cuspp,apepat,nombre1, " & _
                 " case(indlugsus) when 0 then 'N' when 1 then 'S'end 'Lugar Suscr', " & _
                 " case(indfirafi) when 0 then 'N' when 1 then 'S'end 'Firma Afilia', " & _
                 " case(indfirpro) when 0 then 'N' when 1 then 'S'end 'Firma Prom', " & _
                 " case(indhuedig) when 0 then 'N' when 1 then 'S'end 'Huella Digital', " & _
                 " case(indfirfed) when 0 then 'N' when 1 then 'S'end 'Firma Fedat', " & _
                 " error from ctrldoc c join solicitud s on c.lote=s.lote " & _
                 " where s.indDoc='C' and  ltrim(rtrim(isnull(error,''))) <> '' and secu = 0 " & IIf(TxtFechaB <> "", " and s.fecsal = '" & TxtFechaB.Text & "' ", " and c.envio = '" & Me.txtEnvio & "'")
                 
        
        
        RsRepo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRDetAfiErr.DataSource = RsRepo
        DtRDetAfiErr.Sections(2).Controls("LblFecSal").Caption = TxtFechaB
        
        DtRDetAfiErr.Show
    
    Else
    
        SQLstr = " select envio,agencia,c.fecsal,c.lote,numsol,cuspp,apepat,nombre1," & _
                 " case(indselrec) when 0 then 'N' when 1 then 'S'end 'Sello Recepcion', " & _
                 " case(indfirrep) when 0 then 'N' when 1 then 'S'end 'Firma R Afp', " & _
                 " case(indsellrep) when 0 then 'N' when 1 then 'S'end 'Sello R Afp', " & _
                 " case(indlugsus) when 0 then 'N' when 1 then 'S'end 'Lugar Suscr', " & _
                 " case(indfirafi) when 0 then 'N' when 1 then 'S'end 'Firma Afilia', " & _
                 " case(indfirpro) when 0 then 'N' when 1 then 'S'end 'Firma Prom', " & _
                 " case(indhuedig) when 0 then 'N' when 1 then 'S'end 'Huella Digital', " & _
                 " case(indfirfed) when 0 then 'N' when 1 then 'S'end 'Firma Fedat', " & _
                 " error " & _
                 " from ctrldoc c join solicitud s on c.lote=s.lote " & _
                 " where s.indDoc='T' and  ltrim(rtrim(isnull(error,''))) <> '' and secu = 0 " & IIf(TxtFechaB <> "", " and s.fecsal = '" & TxtFechaB.Text & "' ", " and c.envio = '" & Me.txtEnvio & "'")
       
               
        RsRepo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRDetSolErr.DataSource = RsRepo
        DtRDetSolErr.Sections(2).Controls("LblFecSal").Caption = TxtFechaB
        
        DtRDetSolErr.Show
            
              
    
    End If
End Sub

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "DIGITACION\DIGITACION", "BDTraspasos2"
'    DatosServidor
'***
    'Call Con.AbrirTabla("select * from TMPFechaB", RsT, adOpenStatic, adLockOptimistic)
'    CmdRepo.Enabled = False
'    CmdRepoRes.Enabled = False
'    CmdCruce.Enabled = False
'    CmdProceso.Enabled = False
 '   TxtFechaB = gDate
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub txtEnvio_LostFocus()
txtEnvio = Format(txtEnvio.Text, "000000#")
End Sub

Private Sub TxtFechaB_GotFocus()
    Marcado TxtFechaB
End Sub

Private Sub TxtFechaB_Change()
'    CmdRepo.Enabled = True
'    CmdRepoRes.Enabled = True
'    CmdCruce.Enabled = True
'    CmdProceso.Enabled = True
End Sub

Private Sub CmdRepo_Click()
    If RsRepo.State = 1 Then RsRepo.Close
    'SQLstr = "SELECT Solicitud.Fecsal, Solicitud.Lote, Numsol, CUSPP, ApePat, ApeMat, Nombre1, Empleadores, Cantemple, Codafpori, Solicitud.Status, Agencia , Envio FROM Solicitud, CtrlDoc WHERE Solicitud.Fecsal = '" & TxtFechaB & "' AND Solicitud.Lote = CtrlDoc.Lote AND Apepat is not null order by Solicitud.Lote, NumSol "
    'SQLstr = "SELECT Solicitud.Fecsal, Solicitud.Lote, Numsol, CUSPP, ApePat, ApeMat, Nombre1, Empleadores, Cantemple, Codafpori, Solicitud.Status, Agencia , Envio FROM Solicitud, CtrlDoc WHERE CTRLDoc.Envio = '" & txtenvio & "' AND Solicitud.Lote = CtrlDoc.Lote AND Apepat is not null order by Solicitud.Lote, NumSol "
    
    'Empleadores
    If Not OptAfil.Value Then
    
        SQLstr = " select c.envio, s.fecsal,agencia,c.lote,numsol, cuspp,apepat,nombre1, CantEmple = count(*), s.codafpori, " & _
                 " IndEnvMail = case when IndEnvMail = 1 then 'S' else 'N' end " & _
                 " from ctrldoc c join solicitud s on c.lote = s.lote " & _
                 " where s.indDoc='T'" & IIf(TxtFechaB <> "", " and s.fecsal = '" & TxtFechaB.Text & "' ", " and c.envio = '" & Me.txtEnvio & "'") & _
                 " group by c.envio,s.fecsal,agencia,c.lote,numsol, cuspp ,ApePat, " & _
                 " nombre1, s.codafpori,IndEnvMail "
        
        
        RsRepo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRDetSolProc.DataSource = RsRepo
        DtRDetSolProc.Sections(2).Controls("LblFecSal").Caption = TxtFechaB
        DtRDetSolProc.Show
    
    
    Else
    
        SQLstr = " select c.envio, s.fecsal,agencia,c.lote, numsol = substring(s.numsol,2,10), dni = numdoc , apepat,nombre1, CantEmple = count(*), s.codafpori, " & _
                 " IndEnvMail = case when IndEnvMail = 1 then 'S' else 'N' end " & _
                 " from ctrldoc c join solicitud s on c.lote = s.lote " & _
                 " where s.indDoc='C'" & IIf(TxtFechaB <> "", " and s.fecsal = '" & TxtFechaB.Text & "' ", " and c.envio = '" & Me.txtEnvio & "'") & _
                 " group by c.envio,s.fecsal,agencia,c.lote,s.numsol, numdoc,cuspp ,ApePat, " & _
                 " nombre1, s.codafpori,IndEnvMail "
        
        
        RsRepo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRDetAfiProc.DataSource = RsRepo
        DtRDetAfiProc.Sections(2).Controls("LblFecSal").Caption = TxtFechaB
                  
        DtRDetAfiProc.Show
    
    End If
End Sub

Private Sub CmdRepoRes_Click()
Dim TipDoc As String

 If RsRepoRes.State = 1 Then RsRepoRes.Close
 
 '   SQLstr = "SELECT c.Fecsal, c.Envio, c.Lote, c.Agencia, t.DesAgen, c.CantSol, c.CantImag, a.numS, b.numE FROM CtrlDoc c, TAgenAfp t, (SELECT s.lote, numE=count(*) FROM empleadores e, solicitud s WHERE e.numsol = s.numsol AND e.SECUENCIAL=s.SECU group by s.lote) b, (SELECT s.lote, numS=count(*) FROM solicitud s WHERE secu='0' group by s.lote) a where c.lote=a.lote and c.lote=b.lote and c.Envio = '" & txtenvio & "' and c.Agencia = t.Codigo order by c.Lote"
 If OptAfil.Value Then
    TipDoc = "C"
 Else
    TipDoc = "T"
 End If
 
    SQLstr = " select distinct c.envio,c.lote,c.agencia,desagen, codsup,s.fecsal, " & _
             " (select count(numsol) from vt_SolicitudProm where lote = " & _
             " c.lote and secu=0 and indDoc='" & TipDoc & "' and fecsal = s.fecsal and isnull(codsup,'') = isnull(p.Codsup,'')) 'solicitudes', " & _
             " (select count( numsol) from vt_SolicitudProm where lote = " & _
             " c.lote and indDoc='" & TipDoc & "' and fecsal = s.fecsal and isnull(codsup,'') = isnull(p.Codsup,'') ) 'empleadores', (select count(*) from _imagenes i " & _
             " join vt_SolicitudProm ss " & _
             " on '01'+ss.numsol=i.idimagen where ss.lote = " & _
             " c.lote and ss.indDoc='" & TipDoc & "' and fecsal = s.fecsal and isnull(codsup,'') = isnull(p.Codsup,'')) 'imagenes', " & _
             " (select count(distinct numsol) from vt_SolicitudProm where lote = " & _
             " c.lote and  ltrim(rtrim(isnull(error,''))) <> ''  and indDoc='" & TipDoc & "' and fecsal = s.fecsal and isnull(codsup,'') = isnull(p.Codsup,'') ) 'Con Error' " & _
             " from ctrldoc c " & _
             " left join solicitud s on c.lote=s.lote " & _
             " left join tagenafp a on c.agencia = a.codigo " & _
             " left join tpromotor p on p.CodPromo = s.CodPromo " & _
             " where s.indDoc= '" & TipDoc & "' " & IIf(TxtFechaB <> "", " and s.fecsal = '" & TxtFechaB.Text & "' ", " and c.envio = '" & Me.txtEnvio & "'") & _
             " group by envio,c.agencia,c.lote,numsol,desagen,codsup,p.codpromo,s.fecsal "
    
 
    RsRepoRes.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    If TipDoc = "T" Then
        DtRResSolProc.Sections("sección2").Controls("Etiqueta2").Caption = "Reporte Resumen de Solicitudes Procesadas"
    Else
        DtRResSolProc.Sections("sección2").Controls("Etiqueta2").Caption = "Reporte Resumen de Afiliaciones Procesadas"
    End If
     
    Set DtRResSolProc.DataSource = RsRepoRes
    DtRResSolProc.Show
End Sub

Private Sub CmdCruce_Click()
    If rsCruce.State = 1 Then rsCruce.Close
    'SQLstr = "SELECT CtrlDoc.Lote, Envio, Fecsal, Agencia, CantSol, CantImag, DesAgen, num=count(*) FROM CtrlDoc, TAgenAfp, _Imagenes WHERE CtrlDoc.Lote = _Imagenes.Lote AND Agencia = Codigo AND Fecsal = '" & TxtFechaB & "' Group By CtrlDoc.Lote, Envio, FecSal, Agencia, CantSol, CantImag, DesAgen Order By CtrlDoc.Lote"
    SQLstr = "SELECT CtrlDoc.Lote, Envio, Fecsal, Agencia, CantSol, CantImag, DesAgen, num=count(*) FROM CtrlDoc, TAgenAfp, _Imagenes WHERE CtrlDoc.Lote = _Imagenes.Lote AND Agencia = Codigo AND CtrlDoc.Envio = '" & txtEnvio & "' Group By CtrlDoc.Lote, Envio, FecSal, Agencia, CantSol, CantImag, DesAgen Order By CtrlDoc.Lote"
    'c.Envio = '" & txtEnvio & "'
    rsCruce.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    
    
    Set DtRCruceVeri.DataSource = rsCruce
    DtRCruceVeri.Show
End Sub

Private Sub CmdProceso_Click()
    If RsCtrlD.State = 1 Then RsCtrlD.Close
    'SQLstr = "select * from CtrlDoc WHERE Fecsal = '" & TxtFechaB & "' Order by Lote"
    SQLstr = "select * from CtrlDoc WHERE Envio = '" & txtEnvio & "' Order by Lote"
    

    RsCtrlD.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    With RsCtrlD
    If .BOF And .EOF Then   'pregunta si tabla esta vacia
        MsgBox "No existen registros ", vbCritical, "Error en Consulta por Envio"
        Exit Sub
    Else
        .MoveFirst
        TxtLotIni = !Lote
        .MoveLast
        TxtlotFin = !Lote
    End If
    End With
'****
    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = "delete tmplote"
        .Execute
        .CommandText = "insert into tmplote (lote,TotalTif,TotSol,TotDNI,TotHC,TotOtr) " & _
        "select lote,count(*),0,0,0,0 " & _
        "from _imagenes where Lote >= " & val(TxtLotIni) & " AND Lote <= " & val(TxtlotFin) & " group by lote order by lote "
        .Execute
        .CommandText = "update tmplote " & _
        "set totsol = isnull((select count(*) from _imagenes where tipimag='S' and lote=a.lote group by lote),0)," & _
        "totdni = isnull((select count(*) from _imagenes where tipimag='D' and lote=a.lote group by lote),0), " & _
        "tothc = isnull((select count(*) from _imagenes where tipimag='H' and lote=a.lote group by lote),0), " & _
        "tototr = isnull((select count(*) from _imagenes where tipimag<>'H' and tipimag<>'S' and tipimag<>'D' and lote=a.lote group by lote),0) " & _
        "from tmplote a where Lote >= " & val(TxtLotIni) & " AND Lote <= " & val(TxtlotFin) & " "
        .Execute
        .CommandText = "update ctrldoc " & _
        "set cantimag = isnull((select num = totsol + totdni from tmplote where lote=a.lote),0) " & _
        "from ctrldoc a where Lote >= " & val(TxtLotIni) & " AND Lote <= " & val(TxtlotFin) & " "
        .Execute
        MsgBox "Proceso Ejecutado", vbInformation, "Proceso de Conteo de Imagenes"
    End With

End Sub

Private Sub TxtFechaB_KeyPress(KeyAscii As Integer)
    FVFecha TxtFechaB, KeyAscii
End Sub

Private Sub TxtFechaB_LostFocus()
    ValidaFecha TxtFechaB
End Sub

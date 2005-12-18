VERSION 5.00
Begin VB.Form FrmRepDig 
   Caption         =   "Reporte por Digitador"
   ClientHeight    =   5745
   ClientLeft      =   4845
   ClientTop       =   2925
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   5640
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "&Ejecutar"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame FraRepo 
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   5295
      Begin VB.OptionButton OptDigiFecha 
         Caption         =   "Detalle Digitador por Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   3615
      End
      Begin VB.OptionButton OptProdFec 
         Caption         =   "Produccion y Pulsaciones x Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   2880
         Width           =   3855
      End
      Begin VB.OptionButton OptDigiLote 
         Caption         =   "Detalle Digitador por Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   2280
         Width           =   3735
      End
      Begin VB.OptionButton OptResDigi 
         Caption         =   "Resumen de Pulsaciones x Digitador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton OptCtaPls 
         Caption         =   "Cuenta Pulsaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   3855
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame FraFecSal 
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   3855
      Begin VB.TextBox TxtFecFin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtFecIni 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   0
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese Rango de Fechas :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmRepDig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsDetDigi As New Recordset     'Reporte Digitadores
Public RsResDigi As New Recordset     'Reporte Digitadores
Public RsFecDigi As New Recordset     'Reporte Digitadores

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "", "192.168.20.57", "BDTraspasosDig2"
'    DatosServidor
'***
    TxtFecIni = gDate
    TxtFecFin = gDate
End Sub

Private Sub CmdEjecutar_Click()
    If OptCtaPls.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "update empleadores " & _
            "Set lote = s.lote " & _
            "FROM empleadores e, solicitud s WHERE e.numsol=s.numsol and e.secuencial=s.secu and s.fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) and s.fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103) "
            .Execute
            .CommandText = "update solicitud " & _
            "set numpul=len(numsol)+len(apepat)+len(apemat)+len(nombre1)+len(nombre2)+len(tipodoc)+len(numdoc)+len(cuspp)+len(direccion)+len(numero)+len(urb)+len(ubigeo)+len(telefono)+len(codpromo)+len(bolpago)+len(pagoultapo)+len(copiasim)+len(otrodoc)+len(descotrodoc)+( case when len(fecregpro)>0 then 8 else 0 end )+len(codafpori)+( case when len(fecfirmsol)>0 then 8 else 0 end )+len(rtrim(ltrim(riaafil)))+( case when len(fecnac)>0 then 8 else 0 end )+len(sexo)+len(e_mail) + 17 " & _
            "Where secu = 0 and fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) AND fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103) "
            .Execute
            .CommandText = "update empleadores " & _
            "Set numpul = Len(numruc) + len(e.direccion)+len(e.numero)+len(e.urb)+len(e.ubigeo)+len(e.telefono)+len(rtrim(ltrim(e.riaafil))) + 6 " & _
            "from empleadores e, solicitud s " & _
            "Where e.numsol=s.numsol and s.fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) AND s.fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103) "
            .Execute
            MsgBox " Comando Ejecutado ", vbInformation, "Conteo de Pulsaciones"
        End With
    End If
    '***
    If OptResDigi.Value Then
        Llenatempo
        If RsResDigi.State = 1 Then RsResDigi.Close
        SQLstr = "SELECT codusu,nombres,apellidos,Sminxsol=SUM(minu)/SUM(numS),SnumS=SUM(numS),SnumE=SUM(numE),Shoratra=SUM(horatra),Splssol=SUM(plssol),Splsemp=SUM(plsemp),Splstot=SUM(plssol)+SUM(plsemp),veloc=(SUM(plssol)+SUM(plsemp))/SUM(horatra) from tmpdigi2 group by codusu,nombres,apellidos order by codusu"
        RsResDigi.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRResDigi.DataSource = RsResDigi
        DtRResDigi.Show
    End If
    '***
    If OptDigiFecha.Value Then
        Llenatempo
        If RsFecDigi.State = 1 Then RsFecDigi.Close
        SQLstr = "select fecsal,codusu,minu,minxsol,numS,numE,nombres,apellidos,horatra,plssol,plsemp,plstot=plssol+plsemp,veloc=(plssol+plsemp)/horatra from tmpdigi2 order by fecsal, codusu"
        RsFecDigi.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRDigiFecha.DataSource = RsFecDigi
        DtRDigiFecha.Show
    End If
    '***
    If OptDigiLote.Value Then
        If RsDetDigi.State = 1 Then RsDetDigi.Close
        SQLstr = "select a.fecsal, a.lote, a.codusu, a.numS, b.numE, c.nombres, c.apellidos from (select s.fecsal, s.lote, s.codusu, numS=count(*) from Solicitud s where secu='0' group by s.fecsal, s.lote, s.codusu) a, (select s1.fecsal, s1.lote, s1.codusu, numE=sum(cantemple) from Solicitud s1 where secu='0' group by s1.fecsal, s1.lote, s1.codusu) b, usuarios c Where a.lote = b.lote and a.codusu = c.login  and a.fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) and a.fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103) order by  a.codusu, a.fecsal, a.lote"
        RsDetDigi.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRDigiLote.DataSource = RsDetDigi
        DtRDigiLote.Show
    End If
    '***
    If OptProdFec.Value Then
        Llenatempo
        If RsResDigi.State = 1 Then RsResDigi.Close
        SQLstr = "Select fecsal,tnums=sum(numS),tnume=sum(numE),tplssol=sum(plssol),tplsemp=sum(plsemp),thoratra=sum(horatra),ppSol=sum(plssol)/sum(numS),ppEmp=sum(plsemp)/sum(numE),ppDoc=(sum(plssol)+sum(plsemp))/sum(numS) from tmpdigi2 group by fecsal order by fecsal"
        RsResDigi.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
        Set DtRProdFecha.DataSource = RsResDigi
        DtRProdFecha.Show
    End If
End Sub

Private Sub Llenatempo()
    Set comando = New Command
    With comando
        .ActiveConnection = Con.Cnx
        .CommandText = "delete tmpdigi "
        .Execute
        .CommandText = "insert into tmpdigi (a.fecsal,a.codusu,a.horaini1,a.horafin1) " & _
        "select a.fecsal, codusu, horaini1, horafin1 " & _
        "from ( select fecsal, codusu, HoraIni1=min(horfin), HoraFin1=max(horfin) from solicitud where cast(SUBSTRING(horfin,1,2) as int ) > 12 and cast(SUBSTRING(horfin,1,2) as int)< 24 and codusu is not null group by fecsal,codusu " & _
        "union " & _
        "select fecsal, codusu, HoraIni1=min(horfin), HoraFin1=max(horfin) from solicitud where cast(SUBSTRING(horfin,1,2) as int ) >= 0 and cast(SUBSTRING(horfin,1,2) as int)< 13 and codusu is not null  group by fecsal,codusu) a " & _
        "where a.fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) and a.fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103) order by a.fecsal "
        .Execute
        .CommandText = "update tmpdigi " & _
        "set minu1 = ( isnull(datediff(ss,horaini1,horafin1),0)/60 ) + 2 " & _
        "From tmpdigi "
        .Execute
        .CommandText = "delete tmpdigi2 "
        .Execute
        .CommandText = "insert into tmpdigi2 (fecsal, codusu, minu) " & _
        "select a.fecsal, a.codusu, a.minu " & _
        "from ( select fecsal, codusu, minu=sum(minu1) from tmpdigi group by fecsal, codusu ) a " & _
        "Where a.Codusu Is Not Null "
        .Execute
        .CommandText = "update tmpdigi2 " & _
        "set numS = x.numS, " & _
        "numE = x.numE, " & _
        "nombres = x.nombres, " & _
        "apellidos = x.apellidos " & _
        "from (select a.fecsal, a.codusu, a.numS, b.numE, c.nombres, c.apellidos from (select s.fecsal, s.codusu, numS=count(*) from Solicitud s where secu='0' group by s.fecsal, s.codusu) a, (select s1.fecsal, s1.codusu, numE=sum(cantemple) from Solicitud s1 where secu='0' group by s1.fecsal, s1.codusu) b, usuarios c Where a.codusu = c.login and a.codusu=b.codusu and a.fecsal=b.fecsal and a.fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) and a.fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103)) x " & _
        "where tmpdigi2.fecsal=x.fecsal and tmpdigi2.codusu=x.codusu "
        .Execute
        .CommandText = "update tmpdigi2 " & _
        "set minxsol = (isnull(minu,0))/numS " & _
        "From tmpdigi2"
        .Execute
        .CommandText = "update tmpdigi2 " & _
        "set horatra=minu/60 " & _
        "from tmpdigi2 "
        .Execute
        .ActiveConnection = Con.Cnx
        .CommandText = "update tmpdigi2 " & _
        "set plssol = totpls " & _
        "From (select fecsal, codusu, totpls=sum(numpul) from solicitud where codusu is not null and fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) AND fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103) group by fecsal, codusu) s " & _
        "Where tmpdigi2.fecsal=s.fecsal and tmpdigi2.codusu=s.codusu "
        .Execute
        .CommandText = "update tmpdigi2 " & _
        "set plsemp = totpls " & _
        "From (select s.fecsal, e.codusu, totpls=sum(e.numpul) from empleadores e, solicitud s where e.numsol=s.numsol and e.codusu is not null and s.fecsal >= CONVERT(datetime,'" & TxtFecIni & "',103) AND s.fecsal <= CONVERT(datetime,'" & TxtFecFin & "',103) group by s.fecsal, e.codusu) a  " & _
        "Where tmpdigi2.fecsal=a.fecsal and tmpdigi2.codusu=a.codusu "
        .Execute
    End With
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub TxtFecIni_GotFocus()
    Marcado TxtFecIni
End Sub

Private Sub TxtFecFin_GotFocus()
    Marcado TxtFecFin
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FVFecha TxtFecIni, KeyAscii
        TxtFecFin.SetFocus
    Else
        FVFecha TxtFecIni, KeyAscii
    End If
End Sub

Private Sub TxtFecIni_LostFocus()
    ValidaFecha TxtFecIni
End Sub

Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FVFecha TxtFecFin, KeyAscii
        CmdEjecutar.SetFocus
    Else
        FVFecha TxtFecFin, KeyAscii
    End If
End Sub

Private Sub TxtFecFin_LostFocus()
    ValidaFecha TxtFecFin
End Sub


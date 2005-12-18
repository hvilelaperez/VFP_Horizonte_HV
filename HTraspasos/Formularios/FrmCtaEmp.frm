VERSION 5.00
Begin VB.Form FrmCtaEmp 
   Caption         =   "Corrige Datos Generales"
   ClientHeight    =   4890
   ClientLeft      =   4245
   ClientTop       =   2325
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6285
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
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame FraModi 
      Caption         =   "Opciones"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   6015
      Begin VB.OptionButton OptTelGuion 
         Caption         =   "Elimina guiones de campo Telefono"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton OptRiaSolHV 
         Caption         =   "Reemplaza RIA de Solicitud x Suma Renta HV"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
      End
      Begin VB.OptionButton OptLoteEmp 
         Caption         =   "Coloca Lote en archivo Empleadores"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton OptCtaEmp 
         Caption         =   "Cuenta Empleadores"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3255
      End
   End
   Begin VB.Frame FraIngFecha 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4095
      Begin VB.TextBox TxtFecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingresa Fecha :"
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
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1470
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
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCtaEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'    Con.ConexionSQL "sa", "2326", "si-006", "BDTraspasosdig2"
'    Con.ConexionSQL "sa", "", "192.168.20.127", "BDTraspasosdig2"
'    DatosServidor
'***
    TxtFecha = gDate
End Sub

Private Sub CmdEjecutar_Click()
    If OptCtaEmp.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "Update solicitud " & _
            "set empleadores = case when isnull((select top 1 '1' from empleadores where numsol = a.numsol and empleadores.fecsal=a.fecsal),0) = 1 then '1' else '0' end, " & _
            " cantemple = isnull((select count(*) from empleadores where numsol = a.numsol and empleadores.fecsal=a.fecsal),0) from solicitud a WHERE a.fecsal= '" & TxtFecha & "' "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Conteo de Empleadores"
    End If
'*****
    If OptLoteEmp.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "update empleadores " & _
            "Set lote = b.lote " & _
            "FROM empleadores a, solicitud b WHERE a.numsol=b.numsol and a.fecsal = b.fecsal and a.secuencial=b.secu and b.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Coloca Lote a Archivo Empleador "
    End If
'*****
    If OptTelGuion.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "UPDATE solicitud " & _
            "set telefono = REPLACE(telefono,'-','') " & _
            "WHERE telefono like '%-%' " & _
            "and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
            .Execute
            .CommandText = "UPDATE empleadores " & _
            "set telefono = REPLACE(telefono,'-','') " & _
            "from empleadores e, (SELECT lote1=min(lote),lote2=max(lote) FROM CTRLDOC where fecsal=CONVERT(datetime,'" & TxtFecha & "',103)) a " & _
            "WHERE e.telefono like '%-%' " & _
            "and e.lote>=a.lote1 and e.lote<=a.lote2 "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Elimina guiones de campo Telefono "
    End If
'***** Reemplaza RIA de Solicitud x Suma Renta HV
    If OptRiaSolHV.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "update SOLICITUD " & _
            "Set riaafil=RiaTot " & _
            "from SOLICITUD s join (select a.fecsal,a.lote,a.inddoc,a.numsol,RiaTot=sum(b.renta) " & _
            "from SOLICITUD a join HOJAVIDA b on a.NUMSOL=b.NUMSOL and a.fecsal=b.fecsal and a.secu=b.sechv " & _
            "where a.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) group by a.fecsal,a.lote,a.inddoc,a.numsol) h  " & _
            "on s.NUMSOL=h.NUMSOL and s.fecsal=h.fecsal and s.lote=h.lote and s.inddoc=h.inddoc "
            .Execute
        End With
        MsgBox " Comando Ejecutado ", vbInformation, "Reemplaza Ria Emp x Ria Sol "
    End If

'*****  Reemplaza  RIA = 0  a  RIA = 1
'    If OptRia1.Value Then
'        Set comando = New Command
'        With comando
'            .ActiveConnection = Con.Cnx
'            .CommandText = "update solicitud " & _
'            "set riaafil = 1 " & _
'            "where riaafil = 0 and secu=0 and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'            .CommandText = "update empleadores " & _
'            "set riaafil = 1 " & _
'            "where riaafil = 0 and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'        End With
'        MsgBox " Comando Ejecutado ", vbInformation, "Convierte Ria = 0 a Ria = 1 "
'    End If
'***** Coloca "N" en campos vacios Doc presentados
'    If OptIndDocpres.Value Then
'        Set comando = New Command
'        With comando
'            .ActiveConnection = Con.Cnx
'            .CommandText = "update solicitud " & _
'            "set bolpago='N' " & _
'            "where bolpago<>'S' and secu=0 and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'            .CommandText = "update solicitud " & _
'            "set pagoultapo='N' " & _
'            "where pagoultapo<>'S' and secu=0 and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'            .CommandText = "update solicitud " & _
'            "set copiasim='N' " & _
'            "where copiasim<>'S' and secu=0 and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'            .CommandText = "update solicitud " & _
'            "set otrodoc='N' " & _
'            "where otrodoc<>'S' and secu=0 and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'        End With
'        MsgBox " Comando Ejecutado ", vbInformation, "Reemplaza Bol Pago,... Copia Sim,... vacios con N "
'    End If
'***** Reemplaza RIA de Solicitud en RIA Empleadores
'    If OptRiaSolEmp.Value Then
'        Set comando = New Command
'        With comando
'            .ActiveConnection = Con.Cnx
'            .CommandText = "update empleadores " & _
'            "Set riaafil = b.riaafil " & _
'            "from empleadores a, solicitud b " & _
'            "where a.numsol=b.numsol and a.secuencial = b.secu and a.riaafil <> b.riaafil and b.riaafil> 1 and b.secu=0 and b.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'        End With
'        MsgBox " Comando Ejecutado ", vbInformation, "Reemplaza Ria Emp x Ria Sol "
'    End If
'***** Reemplaza tipo doc en Imagenes  "E"  a  "S"
'    If OptEaS.Value Then
'        Set comando = New Command
'        With comando
'            .ActiveConnection = Con.Cnx
'            .CommandText = "update _IMAGENES " & _
'            "set tipimag='S' " & _
'            "WHERE tipimag='E' "
'            .Execute
'        End With
'        MsgBox " Comando Ejecutado ", vbInformation, "Coloca S al 2do empleador "
'    End If
'***** Rellena campos de Direccion vacios
'    If OptDirB.Value Then
'        Set comando = New Command
'        With comando
'            .ActiveConnection = Con.Cnx
'            .CommandText = "UPDATE solicitud " & _
'            "Set direccion = urb " & _
'            "WHERE direccion='' and urb<>'' " & _
'            "and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'            .CommandText = "UPDATE empleadores " & _
'            "Set direccion = e.urb " & _
'            "from empleadores e, solicitud s " & _
'            "where e.direccion='' and e.urb<>'' and e.numsol=s.numsol and e.secuencial=s.secu " & _
'            "and s.fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'        End With
'        MsgBox " Comando Ejecutado ", vbInformation, "Rellena campo Direccion vacios "
'    End If
'***** Modifica Telefonos de menos de 6 posiciones
'    If OptTel6.Value Then
'        Set comando = New Command
'        With comando
'            .ActiveConnection = Con.Cnx
'            .CommandText = "UPDATE solicitud " & _
'            "set telefono = isnull(replicate ('0', 7 - len(rtrim(telefono))),'') + rtrim(telefono) " & _
'            "WHERE len(telefono)<7 and telefono<>'' " & _
'            "and fecsal = CONVERT(datetime,'" & TxtFecha & "',103) "
'            .Execute
'            .CommandText = "UPDATE empleadores " & _
'            "set telefono = isnull(replicate ('0', 7 - len(rtrim(telefono))),'') + rtrim(telefono) " & _
'            "from empleadores e, (SELECT lote1=min(lote),lote2=max(lote) FROM CTRLDOC where fecsal=CONVERT(datetime,'" & TxtFecha & "',103)) a " & _
'            "WHERE len(e.telefono)<7 and e.telefono<>'' " & _
'            "and e.lote>=a.lote1 and e.lote<=a.lote2 "
'            .Execute
'        End With
'        MsgBox " Comando Ejecutado ", vbInformation, "Modifica telefonos de menos de 7 posiciones "
'    End If
End Sub
'*****

Private Sub TxtFecha_GotFocus()
    Marcado TxtFecha
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    FVFecha TxtFecha, KeyAscii
    If KeyAscii = 13 Then Me.CmdEjecutar.SetFocus
End Sub

Private Sub TxtFecha_LostFocus()
    ValidaFecha TxtFecha
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub


VERSION 5.00
Begin VB.Form frmExportacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportación"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.Frame Frame3 
         Caption         =   "Lista de Envios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3360
         TabIndex        =   15
         Top             =   2400
         Width           =   2055
         Begin VB.ListBox lstEnvio 
            Height          =   1035
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3360
         TabIndex        =   13
         Top             =   1200
         Width           =   2055
         Begin VB.CheckBox chkReingreso 
            Caption         =   "Incluir Reingresos en este Envio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdVerEnvios 
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         ToolTipText     =   "Muestra los Envio para la fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFecAcep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtEnvio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdDocId 
         Caption         =   "&Gen. Doc ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtFecSalida 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdExpoAcep 
         Caption         =   "&Exportar Aceptadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Aceptadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   420
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblFecSalida 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha "
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
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblMensajes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1080
         TabIndex        =   8
         Top             =   3120
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmExportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsExpo As New Recordset
Private RsExpoEmp As New Recordset
Private Strsqltexto As String
Private vbEnvio As String
Private RsFam As New Recordset
Private EnvioUltimo As String


Private Sub cmdDocId_Click()

If Not ValidarFechaExpo Then Exit Sub

Set comando = New Command
strPrefijoINI = LeerIni(App.Path & "\scanindexclean.ini", "seteos", "PrefijoFolio")
NumFolio = LeerIni(App.Path & "\scanindexclean.ini", "seteos", "FolioAutomatico")

With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "exec  sp_ActualizarFolio ?,?,?"
    .Parameters("@Prefijo") = strPrefijoINI
    .Parameters("@fecsal") = Trim(txtFecSalida)
    .Parameters("@numsec") = NumFolio
    
    If Con.GetFEjecutarSQL(comando) Then
        If IsNumeric(Con.vbdev) Then
           'GuardarIni App.Path & "\scanindexclean.ini", "seteos", "FolioAutomatico", Con.vbdev
        End If
    End If
    
End With

End Sub

Private Sub cmdExpoAcep_Click()
If txtFecAcep.Text = "" Or txtFecAcep.Text = "//" Then
    MsgBox "Debe ingresar la fecha de Aceptación de las solicitudes", 64, Caption
    Exit Sub
End If

ExportarAceptadas
End Sub

Private Sub cmdExportar_Click()
Dim EnvioReingreso As String
If txtEnvio.Text = "" Then MsgBox "Debe Ingresar Nro de Envio", vbCritical, Caption: Exit Sub
If txtFecSalida.Text = "" Then MsgBox "Debe Ingresar Fecha de salida o Nro de Envio", vbCritical, Caption: Exit Sub
If txtFecAcep.Text = "" Then MsgBox "Debe Ingresar Fecha de Suspención", vbCritical, Caption: Exit Sub

Call ProcesoErrores
Call SuspenderSolicitudes(txtFecSalida)
ActualizarMensaje "Iniciando Exportacion ..."
'Call Exportar(True)
'Call Exportar(False)

EnvioUltimo = fgDatosql("Select max(Envio) from CTrldoc where fecsal = '" & txtFecSalida & "' ")

If EnvioUltimo = "" Then
    MsgBox "No existen envios para la fecha de salida " & txtFecSalida, vbCritical, Caption
    Exit Sub
End If

If chkReingreso.Value = 1 Then

    EnvioReingreso = fgDatosql("Select Envio from LogExpo where FechaExpo = '" & txtFecSalida & "' ")
    
    If EnvioReingreso = "" Then
        RegistrarExpoLog txtFecSalida, txtEnvio
    Else
       If MsgBox("Los reingresos se exportaron en el envio " & EnvioReingreso & " ¿Desea Continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            Exit Sub
       End If
       
    End If
End If

Call NuevaExpo
Call Ocurrencias
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVerEnvios_Click()
MostrarEnvios txtFecSalida
End Sub

Private Sub txtEnvio_LostFocus()
txtEnvio.Text = Format(txtEnvio, "000000#")
End Sub

Private Sub txtFecAcep_KeyPress(KeyAscii As Integer)
FVFecha txtFecAcep, KeyAscii
End Sub

Private Sub txtFecSalida_KeyPress(KeyAscii As Integer)
FVFecha txtFecSalida, KeyAscii
End Sub
Sub Exportar(Normales As Boolean)
Dim sqltexto As String, Condicion As String
Dim RsExpo As Recordset
Dim rsExpo2 As Recordset
Dim rsExpo3 As Recordset
Dim antAgencia As String

Set RsExpo = New ADODB.Recordset
Set rsExpo2 = New ADODB.Recordset
Set rsExpo3 = New ADODB.Recordset

If txtEnvio.Text = "" And txtFecSalida = "" Then MsgBox "Debe Ingresar Fecha de salida o Nro de Envio", vbCritical, Caption: Exit Sub

sqltexto = "select fecsal, Envio  from Ctrldoc where "

              Condicion = " "
              
              If txtFecSalida <> "" Then
                Condicion = " convert(char, fecsal ,103) = '" & txtFecSalida.Text & "' and "
              End If

              If txtEnvio.Text <> "" Then
                Condicion = Condicion & " envio = '" & Trim(txtEnvio.Text) & "'"
              End If
              
              If Trim(Right(Condicion, 4)) = "and" Then Condicion = Left(Condicion, Len(Condicion) - 4)
              
sqltexto = sqltexto & Condicion & " group by fecsal, Envio"

With RsExpo
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
    If .RecordCount > 0 Then
    
        Do While Not .EOF
        
                'Afiliacion
        
                sqltexto = " Select a.NumSol, a.FecSal, a.indDoc, b.Envio, a.Lote, b.Agencia from Solicitud a join Ctrldoc b on a.Lote = b.Lote " & _
                           " where Envio = '" & .Fields("Envio") & "' and indtipdoc = '1' and  a.secu = 0 and (isnull(indSusp,0) = 0) order by b.Agencia, a.Lote  " 'and isnull(indSusp,0) = " & IIf(Normales, "0", "1 and indAcep is null ") & "  order by b.Agencia, a.Lote "
                                                                     
                           
                If rsExpo2.State = 1 Then rsExpo2.Close
                rsExpo2.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

                antAgencia = ""

                If rsExpo2.RecordCount > 0 Then
                        NomArch = App.Path & "\Expo\A" & Format(.Fields("fecsal"), "yymmdd") & "_" & Format(Trim(.Fields("Envio")), "00000#") & IIf(Normales, ".dat", "Rech.dat")
                        Open NomArch For Output As #1

                        If Not ExisteDirectorio(Trim(App.Path) & "\Expo") Then MsgBox "Directorio Expo no existe, debe crear el directorio Expo ", vbCritical, Caption: Exit Sub
                        Do While Not rsExpo2.EOF

                          If antAgencia <> rsExpo2.Fields("Agencia") Then
                            CabeceraDocumento rsExpo2.Fields("Agencia"), rsExpo2.Fields("Lote"), rsExpo2.Fields("Fecsal"), .Fields("Envio"), "C", Normales
                            antAgencia = rsExpo2.Fields("Agencia")
                          End If


                          ExportarAfiliacion rsExpo2.Fields("Numsol"), rsExpo2.Fields("Fecsal"), False
                          '  ExportarTraspaso .Fields("Numsol"), .Fields("Fecsal")
                          rsExpo2.MoveNext
                        Loop
                End If
                Close #1                ' A   C   T   I  V   A   R
                antAgencia = ""
                                
                sqltexto = " Select a.NumSol, a.FecSal, a.indDoc, b.Envio, a.Lote, b.Agencia from Solicitud a join Ctrldoc b on a.Lote = b.Lote " & _
                           " where Envio = '" & .Fields("Envio") & "' and indtipdoc = '0' and a.secu = 0 and isnull(indSusp,0) = 0 order by b.Agencia, a.Lote " 'and isnull(indSusp,0) = " & IIf(Normales, "0", "1 and indAcep is null ") & "    order by b.Agencia, a.Lote "
                           
                If rsExpo3.State = 1 Then rsExpo3.Close
                
                rsExpo3.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
                If rsExpo3.RecordCount > 0 Then
                        NomArch = App.Path & "\Expo\T" & Format(.Fields("fecsal"), "yymmdd") & "_" & Format(Trim(.Fields("Envio")), "00000#") & IIf(Normales, ".dat", "Rech.dat")
                        
                        Open NomArch For Output As #1
                        
                        If Not ExisteDirectorio(Trim(App.Path) & "\Expo") Then MsgBox "Directorio Expo no existe, debe crear el directorio Expo ", vbCritical, Caption: Exit Sub
                        Do While Not rsExpo3.EOF
                        
                          If antAgencia <> rsExpo3.Fields("Agencia") Then
                            CabeceraDocumento rsExpo3.Fields("Agencia"), rsExpo3.Fields("Lote"), rsExpo3.Fields("Fecsal"), .Fields("Envio"), "T", Normales
                            antAgencia = rsExpo3.Fields("Agencia")
                          End If
                        
                          ExportarTraspaso rsExpo3.Fields("Numsol"), rsExpo3.Fields("Fecsal"), False
                          rsExpo3.MoveNext
                        Loop
                        
                        Close #1
                End If
                                                
                .MoveNext
         Loop
    End If
End With
    
    Call Ocurrencias
    
    ActualizarMensaje " Exportación Concluida "
    
End Sub

Sub CabeceraDocumento(Agencia As String, Lote As String, Fecdig As String, Envio As String, TipDoc As String, vbExpoNormales As Boolean)
Dim CantDoc As Integer
CantDoc = fgDatosql("select count(*) from Solicitud a join ctrldoc b on a.lote = b.lote where agencia = '" & Agencia & "' and envio = '" & Envio & "' and secu = 0 and indDoc='" & TipDoc & "' and isnull(indSusp,0) = 0 ")  ' and isnull(indSusp,0) = " & IIf(vbExpoNormales, "0", "1") & " ")
Print #1, "0" & Right(Agencia, 3) & Format(CantDoc, "0000#") & Format(Fecdig, "yyyymmdd")
End Sub

Sub CabeceraDocumentoAcep(Agencia As String, Lote As String, Fecdig As String, Envio As String, TipDoc As String)
Dim CantDoc As Integer
CantDoc = fgDatosql("select count(*) from Solicitud a join ctrldoc b on a.lote = b.lote where agencia = '" & Agencia & "' and envio = '" & Envio & "' and secu = 0 and indDoc='" & TipDoc & "' and isnull(indSusp,0) = 1 and isnull(IndAcep,0) = 1 and convert(varchar,FecSusp,103) = '" & txtFecAcep & "'  ")
Print #1, "0" & Right(Agencia, 3) & Format(CantDoc, "0000#") & Format(Fecdig, "yyyymmdd")
End Sub


Sub ExportarTraspaso(Numsol As String, Fecsal As String, Acept As Boolean)
On Error GoTo ERRORES
'Traspaso
Dim Condicion As String

Dim strPrefijoINI As String
Dim NumFolio As String
Screen.MousePointer = vbHourglass

ActualizarMensaje "Exportando Traspasos ..."

              
              
 Strsqltexto = " pExpoTraspasos '" & Numsol & "' , '" & Fecsal & "'"
                           
If RsExpo.State = 1 Then RsExpo.Close
RsExpo.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

If Not (RsExpo.EOF And RsExpo.BOF) Then

vbEnvio = RsExpo!Envio

With RsExpo
'Open NomArch For Output As #1

        'Modificada la parte de fecha FecFirmSol 22/02/2006
    While Not .EOF
        DoEvents
        Print #1, .Fields("TipReg") & Right(Trim(.Fields("Agencia")), 3) & _
            Right(Replica("0", 10) & Trim(.Fields("PR")), 10) & Left(.Fields("CUSPP") & Replica(" ", 12), 12) & _
            Left(.Fields("Apepat") & Replica(" ", 30), 30) & Left(.Fields("ApeMat") & Replica(" ", 30), 30) & _
            Left(.Fields("Nombre1") & Replica(" ", 20), 20) & Left(.Fields("Nombre2") & Replica(" ", 20), 20) & _
            Left(.Fields("TipoDoc") & Replica("0", 2), 2) & Right(Replica("0", 8) & Trim(.Fields("NumDoc")), 8) & _
            Left(Trim(.Fields("Nacion")) & Replica("0", 3), 3) & Left(Format(Trim(.Fields("Fecnac")), "yyyymmdd") & Replica("0", 8), 8) & Left(.Fields("Sexo") & Replica(" ", 1), 1) & _
            Left(Trim(.Fields("UbigeoNacINE")) & Replica("0", 6), 6) & Left(Trim(.Fields("Direccion")) & " " & Trim(.Fields("Numero")) & Replica(" ", 100), 100) & _
            Left(Trim(.Fields("Urb")) & Replica(" ", 30), 30) & Right(Replica("0", 8) & Trim(.Fields("Telefono")), 8) & _
            Left(Trim(.Fields("Ubigeo")) & Replica("0", 6), 6) & Right(Replica("0", 5) & Trim(.Fields("CodPromo")), 5) & _
            Left(.Fields("TipTrabajador") & Replica(" ", 1), 1) & Right(" " & .Fields("BolPago"), 1) & Right(" " & .Fields("PagoUltApo"), 1) & Right(" " & .Fields("CopiaSim"), 1) & Right(" " & .Fields("OtroDoc"), 1) & _
            Right(" " & .Fields("IndFirRep"), 1) & Right(" " & .Fields("IndSellRep"), 1) & .Fields("UbiVeri") & _
            Left(Format(Trim(.Fields("FecRegPro")), "yyyymmdd") & Replica("0", 8), 8) & Left(.Fields("CodAFPOri") & Replica(" ", 2), 2) & _
            Right(" " & .Fields("TipoFonOri"), 1) & Right(" " & .Fields("TipoFonDes"), 1) & Left(Format(Trim(.Fields("FecRegPro")), "yyyymmdd") & Replica("0", 8), 8) & _
            .Fields("IndLugSus") & .Fields("IndFirAfi") & .Fields("IndFirPro") & .Fields("IndSelRec") & .Fields("IndHueDig") & .Fields("IndFirFed") & _
            Right(Replica("0", 5) & Trim(.Fields("CodCargo")), 5) & Left(.Fields("NomCorres") & Replica(" ", 50), 50) & Left(Trim(.Fields("EmailDoc")) & Replica(" ", 50), 50) & _
            Left(.Fields("EmailTra") & Replica(" ", 50), 50) & Left(Trim(.Fields("Celular")) & Replica("0", 9), 9) & Left(Trim(.Fields("DesDomCor")) & " " & Trim(.Fields("NroDomCor")) & Replica(" ", 100), 100) & _
            Left(Trim(.Fields("UrbaDomCor")) & Replica(" ", 30), 30) & Left(Trim(.Fields("CodgeoCor")) & Replica("0", 6), 6) & _
            Left(.Fields("CodEstCiv") & Replica(" ", 2), 2) & Left(Format(Trim(.Fields("FecSus")), "yyyymmdd") & Replica("0", 8), 8) & _
            Right(Replica("0", 5) & Trim(.Fields("CodInPro")), 5) & Right(Replica("0", 12) & .Fields("Renta"), 12) & _
            Left(.Fields("IndEnvMail") & Replica("N", 1), 1) & Left(Trim(.Fields("IndMail")) & Replica(" ", 1), 1) & .Fields("IndDirCorr") & _
            Left(Format(.Fields("Fecdig"), "yyyymmdd") & Replica("0", 8), 8) & IIf(Acept, "", Left(.Fields("Error") & Replica(" ", 80), 80))

            'Format(.Fields("FecNac"), "ddmmyyyy") & .Fields("Sexo") & Left(.Fields("e_mail") & Replica(" ", 50), 50) & "  1    " & IIf(.Fields("Cuenta") = "", " ", .Fields("Cuenta")) & IIf(.Fields("Boletin") = "", " ", .Fields("Boletin")) & Left(.Fields("e_mail2") & Replica(" ", 50), 50), "Ñ", "" & Chr(165) & """"
            
            .MoveNext
     Wend


   '  Close #1
End With



Strsqltexto = " select TipReg = 2, a.NumSol ,a.NumRuc, EntDig = 'PO', c.Agencia, SecEmp = Secuencial + 1 ,  Secuencial, " & _
              " Direccion = isnull(a.Direccion,''), Numero = isnull(a.Numero,''), Urb = isnull(a.Urb,'') , " & _
              " Telefono = isnull(a.Telefono,''), a.ubigeo, " & _
              " RIAafil  = convert(bigint, isnull(a.RIAafil,1)*100 ), CodSuc = isnull(CodSuc,''), " & _
              " DireEmp = isnull(d.DireEmp,''), NroEmp = isnull(d.NroEmp,''), UrbaEmp= isnull(d.UrbaEmp,''), " & _
              " Codgeo  = isnull(d.Codgeo,''), TeleOfi = isnull(d.TeleOfi,''), AnexOfi= isnull(d.AnexOfi,''), " & _
              " Refdirec = isnull(Refdirec,''),Renta= convert(bigint,isnull(Renta,'0')*100), Fecdig =isnull(b.Fecdig,''), Error,  " & _
              " RazSoc = isnull(a.RazSoc,''),  " & _
              " IndExisBol = case when isnull(a.IndExisBol,0) = 0 then '1' else '2' end " & _
              " from empleadores a " & _
              " join solicitud  b on a.Numsol = b.Numsol and a.fecsal = b.fecsal and a.Secuencial = b.Secu " & _
              " join ctrldoc c on  b.Lote = c.Lote  " & _
              " left join HojaVida d on a.Numsol = d.Numsol and a.IndDoc = d.IndDoc and " & _
              " a.Fecsal = d.Fecsal and SecHV = a.Secuencial " & _
              " left join TRuc e on a.numruc = e.numruc " & _
              " where a.numsol = '" & Numsol & "' and a.Fecsal = '" & Fecsal & "' order by a.Secuencial "
              

              
If RsExpoEmp.State = 1 Then RsExpoEmp.Close
RsExpoEmp.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

With RsExpoEmp


    While Not .EOF
        DoEvents
        Print #1, .Fields("TipReg") & Right(Trim(.Fields("Agencia")), 3) & Right(Replica("0", 10) & Trim(.Fields("NumSol")), 10) & _
            Right(Replica("0", 2) & .Fields("SecEmp"), 2) & Right(Replica("0", 11) & Trim(.Fields("Numruc")), 11) & _
            Left(.Fields("RazSoc") & Replica(" ", 50), 50) & _
            Left(Trim(.Fields("Direccion")) & " " & Trim(.Fields("Numero")) & Replica(" ", 100), 100) & _
            Left(.Fields("Urb") & Replica(" ", 30), 30) & _
            Left(Trim(.Fields("Ubigeo")) & Replica("0", 6), 6) & _
            Right(Replica("0", 8) & Trim(Trim(.Fields("Telefono"))), 8) & Right(Replica("0", 12) & .Fields("Renta"), 12) & _
            Right(Replica("0", 3) & Trim(.Fields("CodSuc")), 3) & Right(Replica("0", 8) & Trim(.Fields("TeleOfi")), 8) & _
            Right(Replica("0", 8) & Trim(.Fields("AnexOfi")), 8) & Left(Trim(.Fields("DireEmp")) & " " & Trim(.Fields("NroEmp")) & Replica(" ", 100), 100) & _
            Left(Trim(.Fields("UrbaEmp")) & Replica(" ", 30), 30) & _
            Left(Trim(Trim(.Fields("Codgeo"))) & Replica("0", 6), 6) & Left(.Fields("Refdirec") & Replica(" ", 50), 50) & _
            Left(Format(Trim(.Fields("Fecdig")), "yyyymmdd") & Replica("0", 8), 8) & Space(80) & .Fields("indExisBol")
            .MoveNext
     Wend
     
     'Call ExpoFamiliar(Numsol, Fecsal, .Fields("Secuencial"))
     Call ExpoFamiliar(Numsol, Fecsal, 0)

    ' Close #1
End With
Screen.MousePointer = vbDefault
'MsgBox "Exportación Concluida", 64, Caption

Else
    Screen.MousePointer = vbDefault
    'MsgBox "No se encontró información", 64, Caption
End If



Exit Sub
ERRORES:
Screen.MousePointer = vbDefault
MsgBox Err.Description, vbCritical, Caption
     Close #1
'RsExpoEmp

End Sub
    
Sub ExportarAfiliacion(Numsol As String, Fecsal As String, Acept As Boolean)
On Error GoTo ERRORES
'Traspaso
Dim Condicion As String

Dim strPrefijoINI As String
Dim NumFolio As String
Screen.MousePointer = vbHourglass

'If Not ValidarFechaExpo Then Exit Sub

ActualizarMensaje "Exportando Afiliaciones ..."

Strsqltexto = " pExpoAfiliacion '" & Numsol & "', '" & Fecsal & "' "
              
             
If RsExpo.State = 1 Then RsExpo.Close
RsExpo.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

If Not (RsExpo.EOF And RsExpo.BOF) Then

vbEnvio = RsExpo!Envio

With RsExpo

        'Modificada la parte de fecha FecFirmSol 22/02/2006
    While Not .EOF
    
        DoEvents
        Print #1, .Fields("TipReg") & Right(Trim(.Fields("Agencia")), 3) & _
            Right(Replica("0", 10) & Trim(Mid(.Fields("NContrato"), 2)), 10) & Left(.Fields("CUSPP") & Replica(" ", 12), 12) & _
            Left(.Fields("Apepat") & Replica(" ", 30), 30) & Left(.Fields("ApeMat") & Replica(" ", 30), 30) & _
            Left(.Fields("Nombre1") & Replica(" ", 20), 20) & Left(.Fields("Nombre2") & Replica(" ", 20), 20) & _
            Left(.Fields("Sexo") & Replica(" ", 1), 1) & Left(Trim(.Fields("UbigeoNac")) & Replica("0", 6), 6) & Left(Trim(Format(.Fields("Fecnac"), "yyyymmdd")) & Replica("0", 8), 8) & _
            Left(.Fields("TipoDoc") & Replica("0", 2), 2) & Right(Replica("0", 8) & Trim(.Fields("NumDoc")), 8) & _
            Left(.Fields("IPSS") & Replica(" ", 15), 15) & Left(Trim(.Fields("Nacion")) & Replica("0", 3), 3) & _
            Left(Trim(.Fields("Direccion")) & " " & Trim(.Fields("Numero")) & Replica(" ", 100), 100) & _
            Left(Trim(.Fields("Urb")) & Replica(" ", 30), 30) & _
            Right(Replica("0", 8) & Trim(.Fields("Telefono")), 8) & _
            Left(Trim(.Fields("Ubigeo")) & Replica("0", 6), 6) & Left(.Fields("TipTrabajador") & Replica(" ", 1), 1) & Left(.Fields("Proc19990") & Replica(" ", 1), 1) & _
            Left(Trim(.Fields("Proc20530")) & Replica(" ", 1), 1) & Left(.Fields("PrimEmpleo") & Replica(" ", 1), 1) & Left(.Fields("TipAfil") & Replica(" ", 1), 1) & Left(Trim(.Fields("CodPromo")) & Replica("0", 5), 5) & _
            .Fields("IndLugSus") & .Fields("IndFirPro") & .Fields("IndFirAfi") & .Fields("IndHueDig") & Left(Trim(Format(.Fields("FecFirmSol"), "yyyymmdd")) & Replica("0", 8), 8) & .Fields("IndFirFed") & _
            Left(Trim(.Fields("CodCargo")) & Replica("0", 5), 5) & Left(Trim(.Fields("NomCorres")) & Replica(" ", 50), 50) & Left(.Fields("EmailDoc") & Replica(" ", 50), 50) & _
            Left(.Fields("EmailTra") & Replica(" ", 50), 50) & _
            Right(Replica("0", 9) & Trim(.Fields("Celular")), 9) & Left(Trim(.Fields("DesDomCor")) & " " & Trim(.Fields("NroDomCor")) & Replica(" ", 100), 100) & _
            Left(Trim(.Fields("UrbaDomCor")) & Replica(" ", 30), 30) & Left(Trim(.Fields("CodgeoCor")) & Replica("0", 6), 6) & _
            Left(Trim(.Fields("CodEstCiv")) & Replica(" ", 2), 2) & Left(Format(Trim(.Fields("FecSus")), "yyyymmdd") & Replica("0", 8), 8) & _
            Left(Trim(.Fields("CodInPro")) & Replica("0", 5), 5) & Right(Replica("0", 12) & Trim(.Fields("Renta")), 12) & _
            Left(.Fields("IndEnvMail") & Replica("N", 1), 1) & Left(Trim(.Fields("IndMail")) & Replica(" ", 1), 1) & .Fields("IndDirCorr") & _
            Left(Format(.Fields("Fecdig"), "yyyymmdd") & Replica("0", 8), 8) & IIf(Acept, "", Left(.Fields("Error") & Replica(" ", 80), 80))
                                    
            'Format(.Fields("FecNac"), "ddmmyyyy") & .Fields("Sexo") & Left(.Fields("e_mail") & Replica(" ", 50), 50) & "  1    " & IIf(.Fields("Cuenta") = "", " ", .Fields("Cuenta")) & IIf(.Fields("Boletin") = "", " ", .Fields("Boletin")) & Left(.Fields("e_mail2") & Replica(" ", 50), 50), "Ñ", "" & Chr(165) & """"
            .MoveNext
     Wend


     'Close #1
End With



Strsqltexto = " select TipReg = 2, a.NumSol ,a.NumRuc, Nro_Cliente = isnull(Nro_Cliente,''), EntDig = 'PO', c.Agencia, Correlativo = a.numsol, a.Secuencial, SecEmp = a.Secuencial + 1, " & _
              " Direccion = isnull(a.Direccion,''), Numero = isnull(a.Numero,''), Urb = isnull(a.Urb,'') , " & _
              " Telefono = isnull(a.Telefono,''), a.ubigeo, " & _
              " RIAafil  = convert(bigint, isnull(a.RIAafil,0)*100 ), CodSuc = isnull(CodSuc,''), " & _
              " DireEmp = isnull(d.DireEmp,''), NroEmp = isnull(d.NroEmp,''), UrbaEmp= isnull(d.UrbaEmp,''), " & _
              " Codgeo  = isnull(d.Codgeo,''), TeleOfi = isnull(d.TeleOfi,''), AnexOfi= isnull(d.AnexOfi,''), " & _
              " Refdirec = isnull(Refdirec,''),Renta= convert(bigint,isnull(Renta,'0')*100), Fecdig =isnull(b.Fecdig,''), Error,  " & _
              " RazSoc = isnull(a.RazSoc,''), IndExisBol = case when isnull(a.IndExisBol,0) = 0 then '1' else '2' end " & _
              " from empleadores a " & _
              " join solicitud  b on a.Numsol = b.Numsol and a.fecsal = b.fecsal and a.Secuencial = b.Secu " & _
              " join ctrldoc c on  b.Lote = c.Lote  " & _
              " left join HojaVida d on a.Numsol = d.Numsol and a.IndDoc = d.IndDoc and " & _
              " a.Fecsal = d.Fecsal and SecHV = a.Secuencial   " & _
              " left join Truc e on a.numruc = e.numruc " & _
              " where a.numsol = '" & Numsol & "' and a.Fecsal = '" & Fecsal & "'  order by a.Secuencial "
              
              
If RsExpoEmp.State = 1 Then RsExpoEmp.Close
RsExpoEmp.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

With RsExpoEmp


    While Not .EOF
        DoEvents
        Print #1, .Fields("TipReg") & Right(Trim(.Fields("Agencia")), 3) & _
            Right(Replica("0", 10) & Trim(Mid(.Fields("Correlativo"), 2)), 10) & Right(Replica("0", 2) & .Fields("SecEmp"), 2) & Left(Trim(.Fields("Numruc")) & Replica("0", 11), 11) & _
            Left(Trim(.Fields("RazSoc")) & Replica(" ", 50), 50) & _
            Left(Trim(.Fields("Direccion")) & " " & Trim(.Fields("Numero")) & Replica(" ", 100), 100) & _
            Left(.Fields("Urb") & Replica(" ", 30), 30) & _
            Left(Trim(.Fields("ubigeo")) & Replica("0", 6), 6) & _
            Right(Replica("0", 8) & Trim(.Fields("Telefono")), 8) & _
            Right(Replica("0", 12) & .Fields("Renta"), 12) & _
            Left(Trim(.Fields("CodSuc")) & Replica("0", 3), 3) & _
            Right(Replica("0", 8) & Trim(.Fields("TeleOfi")), 8) & _
            Right(Replica("0", 8) & Trim(.Fields("AnexOfi")), 8) & _
            Left(Trim(.Fields("DireEmp")) & " " & Trim(.Fields("NroEmp")) & Replica(" ", 100), 100) & _
            Left(Trim(.Fields("UrbaEmp")) & Replica(" ", 30), 30) & _
            Left(Trim(.Fields("Codgeo")) & Replica("0", 6), 6) & Left(Trim(.Fields("Refdirec")) & Replica(" ", 50), 50) & _
            Left(Format(.Fields("Fecdig"), "yyyymmdd") & Replica("0", 8), 8) & _
            Space(80) '& .Fields("IndExisBol")
            
            
            .MoveNext
     Wend
     
     'Call ExpoFamiliar(Numsol, Fecsal, .Fields("Secuencial"))
     Call ExpoFamiliar(Numsol, Fecsal, 0)

     'Close #1
End With
Screen.MousePointer = vbDefault
'MsgBox "Exportación Concluida", 64, Caption

Else
    Screen.MousePointer = vbDefault
  '  MsgBox "No se encontró información", 64, Caption
End If



Exit Sub
ERRORES:
Close #1
Screen.MousePointer = vbDefault
MsgBox Err.Description, vbCritical, Caption
     
'RsExpoEmp

End Sub

Private Sub txtFecSalida_LostFocus()
    ValidaFecha txtFecSalida
End Sub
Function ValidarFechaExpo() As Boolean

ValidarFechaExpo = True
If Trim(txtFecSalida) = "" Then
    ValidarFechaExpo = False
    MsgBox "Debe ingresar Fecha de Salida", vbCritical, Caption
    txtFecSalida.SetFocus
End If

End Function

Private Sub Form_Load()
'Dim Resp As Long
'Resp = SetWindowLong(Me.hwnd, -20, &H20&)
'frmExportacion.Refresh
End Sub

Sub ExpoFamiliar(Numsol As String, Fecsal As String, NumEmpleador As Integer)
Dim inicio As Integer
Strsqltexto = " Select distinct TipReg = 3, a.Numsol , a.indDoc, a.Fecsal, a.Secu, " & _
              " a.ApePat, a.ApeMat, a.NomUno, a.Secuencia, " & _
              " a.NomDos, CodPare = isnull(a.CodPare,''), Sexo = isnull(a.Sexo,'') , " & _
              " Fecnac =  isnull(convert(char,a.Fecnac,103),''), d.Agencia, c.Fecdig, Errores ='', Correlativo= c.Numsol from Familiar a " & _
              " join HojaVida b on a.Numsol = b.Numsol and a.indDoc = b.indDoc " & _
              " and a.Fecsal = b.Fecsal and a.SecHV = b.SecHV " & _
              " join Solicitud c on  b.Numsol = c.Numsol and b.indDoc = c.indDoc " & _
              " and  b.Fecsal = c.Fecsal and b.Secu = c.Secu " & _
              " join CtrlDoc d on c.Lote = d.Lote " & _
              " Where c.Secu = 0 and a.numsol = '" & Numsol & "' and a.Fecsal = '" & Fecsal & "'  order by a.Secuencia "
              '" Where c.Secu = 0 And b.SecHV = " & NumEmpleador & " and a.numsol = '" & Numsol & "' and a.Fecsal = '" & Fecsal & "'  order by a.Secuencia "
                            
If RsFam.State = 1 Then RsFam.Close
RsFam.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

With RsFam

While Not .EOF
        DoEvents
        inicio = "1"
        If .Fields("indDoc") = "C" Then inicio = "2"
        
        Print #1, .Fields("TipReg") & Right(Trim(.Fields("Agencia")), 3) & _
            Right(Replica("0", 10) & Trim(Mid(.Fields("Correlativo"), inicio)), 10) & Right(Replica("0", 2) & .Fields("secuencia"), 2) & _
            Left(.Fields("ApePat") & Replica(" ", 30), 30) & _
            Left(.Fields("ApeMat") & Replica(" ", 30), 30) & Left(.Fields("NomUno") & Replica(" ", 20), 20) & _
            Left(.Fields("NomDos") & Replica(" ", 20), 20) & Left(.Fields("CodPare") & Replica(" ", 1), 1) & Left(.Fields("Sexo") & Replica(" ", 1), 1) & _
            Left(Format(Trim(.Fields("Fecnac")), "yyyymmdd") & Replica("0", 8), 8) & _
            Left(Format(.Fields("Fecdig"), "yyyymmdd") & Replica("0", 8), 8)
            'Left(.Fields("Errores") & Replica("0", 80), 80)
        .MoveNext
     Wend


End With
    

End Sub

Sub ProcesoErrores()
Dim sqltexto As String
Dim Condicion As String
Dim cmdErro As ADODB.Command
Dim sqlDetalleErrores As String
Set cmdErro = New ADODB.Command

ActualizarMensaje "Procesando errores ..."
sqlDetalleErrores = " DetErr26 = case when indIAApel   = 1 then 'Apellidos, ' else '' end + case when indIANomb = 1 then 'Nombres, ' else ''end + " & _
                               " case when indIAFecNac = 1 then 'Fec.Nac, ' else '' end + case when indIATipDoc = 1 then 'Tipo Doc, ' else '' end  + " & _
                               " case when indIANroDoc = 1 then 'Nro Documento, ' else '' end + case when indIACuspp = 1 then 'Cuspp, ' else '' end +" & _
                               " case when indIASexo   = 1 then 'Sexo, ' else '' end + case when indIAIpss = 1 then 'Carnet Ipss, ' else '' end + " & _
                               " case when indIANacio  = 1 then 'Nacionalidad, ' else '' end + case when indIADirAfi = 1 then 'Dirección, ' else '' end + " & _
                               " case when indIAFirRev = 1 then 'Firma Reverso, ' else '' end + case when indIAUbig = 1 then 'Ubigeo, ' else '' end + " & _
                               " case when indIAFalDat = 1 then 'Faltan Datos en formatos copias, ' else '' end , " & _
                               " DetErr33 = case when indIERuc = 1 then 'Ruc, ' else ''end + case when indIERazSoc = 1 then 'Razón Social, ' else '' end + " & _
                               " case when indIENroEmp = 1 then 'Nro. Empleadores, ' else '' end + case when indIEDir = 1 then 'Dirección, ' else '' end + " & _
                               " case when indIEUbigeo = 1 then 'Ubigeo, ' else '' end, " & _
                               " DetErr36 = case when indIVFecSus = 1 then 'Fec. Suscrip, ' else '' end + case when indIVTipFon = 1 then 'Elección Tip Fondo, ' else '' end + " & _
                               " case when indIVDecTra = 1 then 'Declaración Traspaso, ' else ''end + case when indIVVerSol = 1 then 'Verificación Solicitud, '  else '' end ," & _
                               " DetErr24 = case when indEEDatAfi = 1 then 'Datos Afiliado, ' else '' end + case when indEEDatEmp = 1 then 'Datos Empleador, '  else '' end + " & _
                               " case when indEEDatPro = 1 then 'Datos Proveedor, ' else '' end + case when indEEAntPrev = 1 then 'Antecedentes Previsionales, ' else '' end + " & _
                               " case when indEEDecSol = 1 then 'Verificación o Declaración de solicitud, ' else '' end, " & _
                               " DetErr25 = case when indDDApeAfi = 1 then 'Apellidos del Afiliado, ' else '' end + case when indDDNomAfi = 1 then 'Nombre Afiliado, ' else '' end + " & _
                               " case when indDDFecNac=1 then 'Fecha de Nacimiento, ' else '' end + case when indDDTipDoc = 1 then 'Tipo de Documento, ' else '' end  + " & _
                               " case when indDDNumDoc=1 then 'Nro Documento, ' else '' end"


sqltexto = " Update Solicitud set Error = case when IndFalDni = 1 then '01' else '' end + " & _
           " case when IndDniLeg = 1 then '02' else '' end + case when IndDniCad = 1 then '03' else '' end + " & _
           " case when IndDNIFal = 1 then '04' else '' end + case when IndSellONP= 1 then '05' else '' end + " & _
           " case when IndHueDig= 0 then  '90' else '' end + case when IndHueIle= 1 then  '06' else '' end + case when IndHueDif= 1 then  '07' else '' end + " & _
           " case when IndFirAfi= 0 then '95' else ''end + case when IndFirPro= 0 then '96' else ''end + " & _
           " case when IndFirFed= 0 then '97' else ''end + case when IndFirAfiDif= 1 then '08' else '' end + " & _
           " case when IndSinAfpO =1 and indDoc = 'T' then '12' else '' end + case when IndTipFon= 1 and indDoc = 'T' then '13' else '' end + " & _
           " case when IndCusBlan= 1 and indDoc = 'T' then '14' else '' end + case when IndCusErr= 1 and indDoc = 'T' then '15' else '' end + " & _
           " case when IndCodGeoDoc= 1 then '16' else '' end + case when IndCodGeoEmp = 1 then '17' else '' end + " & _
           " case when IndCodGeoNac= 1 then '18' else '' end + case when (Select top 1 IndBolFal from Solicitud where numsol = a.numsol and fecsal = a.fecsal and IndBolFal = '1' ) = '1' then '19' else '' end + " & _
           " case when IndForAdi= 1 then '20' else '' end + case when IndForInc= 1 then '21' else '' end + " & _
           " case when IndHojVid= 1 then '22' else '' end + case when IndFalDec= 1 then '23' else '' end + " & _
           " case when IndForEmn= 1 then '24' else '' end + case when IndDGenDif= 1 then '25'else '' end + " & _
           " case when IndLlenInc= 1 then '26'else '' end + case when IndAfiOtraAfp= 1 and indDoc = 'T' then '27'else '' end + " & _
           " case when IndAfiSpp= 1 and indDoc = 'C' then '28' else '' end + case when IndDocdup= 1 then '29' else '' end + " & _
           " case when IndMultiAfi= 1 then '30' else '' end + case when IndHomi= 1 then '31' else '' end + " & _
           " case when IndLlenIncEmp= 1 then '33' else '' end  + " & _
           " case when IndLlenIncPro= 1 then '34' else '' end  + " & _
           " case when IndLlenIncAnt= 1 and indDoc = 'C' then '35' else '' end  + " & _
           " case when IndLlenIncVrS= 1 then '36' else '' end + " & _
           " case when IndCSBSErr= 1 then '38' else '' end  + " & _
           " case when IndCIntErr= 1 then '39' else '' end  + " & _
           " case when IndErrPro = 1 then '43' else '' end + case when IndRAM = 1 then '40' else '' end, " & _
           sqlDetalleErrores & _
           " from solicitud a join Ctrldoc b on a.Lote = b.Lote where secu = 0 "
                                
           '" case when IndFecDec= 1 then '10' else '' end + case when IndFecVer= 1 then '11' else '' end + " & _
           '+ case when IndFecSus= 1 then '09' else '' end
           '" case when IndRucN= 1 then '32' else '' end  + " & _

      If txtFecSalida <> "" And txtFecSalida <> "//" Then
        Condicion = " and convert(char,b.fecsal ,103) = '" & txtFecSalida.Text & "' "
      End If

      If txtEnvio.Text <> "" Then
        Condicion = Condicion & " and envio = '" & Trim(txtEnvio.Text) & "'"
      End If
      
      
      If Trim(Right(Condicion, 4)) = "and" Then Condicion = Left(Condicion, Len(Condicion) - 4)
      
      sqltexto = sqltexto & Condicion
      
      With cmdErro
      .ActiveConnection = Con.Cnx
      .CommandText = sqltexto
      .Execute
      End With
    
'
'End With
'


'Suspender Solicitudes con errores 06, 07, 08




End Sub
Sub ActualizarMensaje(Ms As String)
'Iniciando Exportacion ...

lblMensajes.Caption = Ms
lblMensajes.Refresh

End Sub
Sub Ocurrencias()
Dim rsOcurrencia As ADODB.Recordset
Set rsOcurrencia = New ADODB.Recordset


Dim sqltexto As String
sqltexto = " select envio, a.Lote , a.numsol, Afiliado = ltrim(rtrim(isnull(ApePat,''))) +' ' + ltrim(rtrim(isnull(ApeMat,''))) + ' ' + ltrim(rtrim(isnull(Nombre1,''))) + ' ' + ltrim(rtrim(isnull(Nombre2,''))), " & _
           " a.FecSal, Agencia, CodPromo = isnull(CodPromo,''), CodInPro = isnull(CodInPro,''), a.IndDoc , " & _
           " Ocurrencia = case when IndFecSus = 1 and a.indDoc = 'C' then '09' else '' end + " & _
           "            case when IndFecDec = 1 and a.indDoc = 'T' then '10' else '' end + " & _
           "            case when IndFecVer = 1 and a.indDoc = 'T' then '11' else '' end + " & _
           "            case when IndFecRevr = 1 and a.indDoc = 'C' then '37' else '' end + " & _
           "            case when IndRucN= 1 then '32' else '' end + " & _
           "            case when IndSelRec = 0 and a.indDoc = 'T' then '91' else '' end + " & _
           "            case when IndFirRep = 0 and a.indDoc = 'T' then '92' else '' end + " & _
           "            case when IndSellRep = 0 and a.indDoc = 'T' then '93' else '' end + " & _
           "            case when IndLugSus = 0 then '94' else '' end  " & _
           " from Solicitud a join ctrldoc b " & _
           " on a.lote = b.lote and a.fecsal = b.fecsal " & _
           " left join HojaVida c on a.numsol=c.numsol and a.fecsal=c.fecsal and a.secu=c.secu where "

'and indDoc = 'T'

'"            case when IndRAM = 1 then '40' else '' end + " & _

           If txtFecSalida <> "" And txtFecSalida <> "//" Then
                Condicion = " convert(char, a.fecsal ,103) = '" & txtFecSalida.Text & "' and "
           End If

           If txtEnvio.Text <> "" Then
                Condicion = Condicion & " envio = '" & Trim(txtEnvio.Text) & "' and"
           End If
           'IndCamFec = 1 or
                         
           If Trim(Right(Condicion, 4)) = "and" Then Condicion = Left(Condicion, Len(Condicion) - 4)
           Condicion = Condicion & " and (IndFecDec = 1 or IndFecVer = 1 or IndFecRevr = 1 or IndFirRep = 0 or IndSellRep = 0 or IndSelRec = 0 or IndLugSus = 0 or IndRAM = 1 or  IndRucN = 1) and a.secu = 0 "
              
sqltexto = sqltexto & Condicion
With rsOcurrencia
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
    If .RecordCount > 0 Then
    
    
        NomArch = App.Path & "\Expo\Ocurrencia" & Format(.Fields("fecsal"), "yymmdd") & "_" & Format(Trim(.Fields("Envio")), "00000#") & ".xls"
    
       Open NomArch For Output As #1
            'Cabecera
            Linea = "Envio" & vbTab & "Lote" & vbTab & "Solicitud" & vbTab & " Tipo Doc. " & vbTab & "Afiliado" & vbTab & "Fecha de Salida" & vbTab & "Agencia" & vbTab & "Cod. Prom SBS" & vbTab & " Cod. Prom. Int." & vbTab & " Ocurrencia "
            Print #1, Linea
    
        Do While Not .EOF
        
            If .Fields("Ocurrencia") <> "" Then
                Linea = "'" & .Fields("Envio") & vbTab & "'" & .Fields("Lote") & vbTab & "'" & Right(Replica("0", 10) & Trim(.Fields("numsol")), 10) & vbTab & .Fields("Inddoc") & _
                        vbTab & Left(.Fields("Afiliado") & Replica(" ", 50), 50) & _
                        vbTab & Left(Format(Trim(.Fields("Fecsal")), "dd/mm/yyyy") & Replica("0", 10), 10) & vbTab & "'" & Trim(Right(Trim(.Fields("Agencia")), 3)) & _
                        vbTab & "'" & Left(.Fields("Codpromo") & Replica(" ", 5), 5) & vbTab & "'" & Left(Trim(.Fields("CodInPro")) & Replica(" ", 5), 5) & _
                        vbTab & .Fields("Ocurrencia")
                        
                Print #1, Linea
            End If
            .MoveNext
        Loop
    
        Close #1
    
    
    End If
End With
  

End Sub

Function ObtenerNumEnvio(Envio As String, Fecsal As String) As String
Dim RsEnvio As Recordset
Set RsEnvio = New Recordset
With RsEnvio
    .Open "Select * from Envio where Fecsal = '" & Fecsal & "' order by Envio ", Con.Cnx, adOpenStatic, adLockReadOnly
    
    
If .RecordCount > 0 Then
    .Find "Envio = '" & Envio & "'"
    If Not .EOF Then
        ObtenerNumEnvio = Format(.AbsolutePosition, "0#")
    End If
    
Else
    ObtenerNumEnvio = "00"
End If
    
End With
End Function

Sub NuevaExpo()
    'On Error GoTo errores
'Determinar el nombre de los archivos

' TRASPASOS
Dim NomArch  As String
NumEnvio = ObtenerNumEnvio(txtEnvio, txtFecSalida)

If NumEnvio = "00" Then MsgBox "Fecha y Envio no corresponden", vbCritical, Caption: Exit Sub

'ExpoArchivoSBS App.Path & "\Expo\IAI1" & Format(Date, "mmdd") & ".PRO", txtEnvio

NomArch = App.Path & "\Expo\TRSN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNTraspaso txtEnvio.Text, FechaTransm(txtFecSalida), NomArch

NomArch = App.Path & "\Expo\TRTN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoTelefonos txtEnvio.Text, FechaTransm(txtFecSalida), NomArch, "T"

NomArch = App.Path & "\Expo\TRDN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNDirecciones txtEnvio.Text, FechaTransm(txtFecSalida), NomArch, "T"
    
NomArch = App.Path & "\Expo\TREN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNEmpleadoresTraspasos txtEnvio.Text, FechaTransm(txtFecSalida), NomArch

NomArch = App.Path & "\Expo\TRHN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNHojaVida txtEnvio.Text, FechaTransm(txtFecSalida), NomArch, "T"

' AFILIACIONES


NomArch = App.Path & "\Expo\CASN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNAfiliacion txtEnvio.Text, FechaTransm(txtFecSalida), NomArch

NomArch = App.Path & "\Expo\CATN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoTelefonos txtEnvio.Text, FechaTransm(txtFecSalida), NomArch, "C"

NomArch = App.Path & "\Expo\CADN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNDirecciones txtEnvio.Text, FechaTransm(txtFecSalida), NomArch, "C"

NomArch = App.Path & "\Expo\CAEN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNEmpleadoresAfiliacion txtEnvio.Text, FechaTransm(txtFecSalida), NomArch

NomArch = App.Path & "\Expo\CAHN" & EnvioForFecha(txtFecSalida) & NumEnvio & ".PRO"
ExpoNHojaVida txtEnvio.Text, FechaTransm(txtFecSalida), NomArch, "C"





MsgBox "Proceso Terminado", 64, Caption

'Exit Sub
'errores:
'MsgBox Err.Description, vbCritical, Caption

End Sub

Sub CabeTraspaso(FechaTransmision As String)
Print #1, "E1914" & FechaTransmision & Space(1901)
End Sub
Sub CabeAfiliacion(FechaTransmision As String)
Print #1, "E1833" & FechaTransmision & Space(1820)
End Sub
Sub SumarioTraspaso(nregistro As Integer)
Print #1, "S" & Right(Replica("0", 8) & nregistro, 8) & Space(1905)
End Sub
Sub SumarioAfiliacion(nregistro As Integer)
Print #1, "S" & Right(Replica("0", 8) & nregistro, 8) & Space(1824)
End Sub
Sub ExpoNTraspaso(Envio As String, FecTran As String, ARCHIVO As String)
Dim rsTraspaso As Recordset
Set rsTraspaso = New Recordset

Open ARCHIVO For Output As #1


With rsTraspaso
    'If Envio = EnvioUltimo Then
    If chkReingreso.Value = 1 Then
        .Open " exec pObtenerDatosDetalle3 '" & Envio & "', 'T', '" & txtFecAcep & "'  ", Con.Cnx, adOpenStatic, adLockReadOnly
    Else
        .Open " exec pObtenerDatosDetalle2 '" & Envio & "', 'T'  ", Con.Cnx, adOpenStatic, adLockReadOnly
    End If
    
 Call CabeTraspaso(FecTran)
    
 If .RecordCount > 0 Then
  
    If RevisionNulos(rsTraspaso, "Archivo Traspasos") Then
        Close #1
        Exit Sub
    End If
    Do While Not .EOF
    
    Print #1, !TipReg & FrmtCampo("PR" & Format(!Numsol, "0000000#"), "A", 15) & !Fecsus & FrmtCampo(!Nomuno, "A", 20) & FrmtCampo(!Nomdos, "A", 20) & _
        FrmtCampo(!ApePat, "A", 30) & FrmtCampo(!ApeMat, "A", 30) & !Fecnac & !Sexo & FrmtCampo(!estciv, "A", 1) & FrmtCampo(!TipDoc, "N", 2) & FrmtCampo(!numdoc, "A", 15) & _
        FrmtCampo(!codnac, "N", 3) & FrmtCampo(!ubinac, "N", 6) & !codpai & FrmtCampo(!ubidir, "N", 6) & FrmtCampo(!prefijodir, "N", 2) & FrmtCampo(!DetDir, "A", 100) & FrmtCampo(!prefijoloc, "N", 2) & _
        FrmtCampo(!detdirloc, "A", 100) & FrmtCampo(!refdirtra, "A", 100) & FrmtCampo(!codddnTra, "N", 3) & FrmtCampo(!telefono1, "N", 9) & FrmtCampo(!EmailTra, "A", 80) & FrmtCampo(!rucemp, "N", 11) & _
        FrmtCampo(!razsocemp, "A", 120) & FrmtCampo(!ubiemp, "N", 6) & FrmtCampo(!prefijodiremp, "N", 2) & FrmtCampo(!diremp, "A", 100) & FrmtCampo(!prefijolocemp, "N", 2) & FrmtCampo(!detdirlocemp, "A", 100) & _
        FrmtCampo(!RefDirEmp, "A", 100) & FrmtCampo(!CodDDNEmp, "N", 3) & FrmtCampo(!teleEmp, "N", 9) & !TipoProceso & FrmtCampo(!tiptra, "A", 1) & FrmtCampo(!firtra, "A", 1) & FrmtCampo(!huella, "A", 1) & FrmtCampo(!firpro, "A", 1) & FrmtCampo(!copdni, "A", 1) & _
        FrmtCampo(!copbolpag, "A", 1) & FrmtCampo(!PagoUltApo, "A", 1) & FrmtCampo(!Cuspp, "A", 12) & FrmtCampo(!codoriafp, "N", 2) & !anibodas & FrmtCampo(!codcargo, "N", 3) & FrmtCampo(!codpaisnac, "N", 3) & FrmtCampo(!codocupac, "N", 4) & _
        FrmtCampo(!codgironec, "N", 5) & FrmtCampo(!autoenvio, "A", 1) & FrmtCampo(!numhijos, "N", 2) & FrmtCampo(!Ram, "N", 11) & FrmtCampo(!derbono, "A", 1) & FrmtCampo(!SitBono, "N", 1) & FrmtCampo(!periodicidad, "A", 1) & _
        FrmtCampo(!codadm, "A", 10) & FrmtCampo(!CodPromo, "A", 10) & FrmtCampo(!OtrosDoc, "A", 1) & FrmtCampo(!DesOtroDoc, "A", 50) & FrmtCampo(!TipoFonOri, "A", 5) & FrmtCampo(!TipoFonDes, "A", 5) & FrmtCampo(!IndSelRec, "A", 1) & FrmtCampo(!IndSelVer, "A", 1) & _
        FrmtCampo(!NombVer, "A", 40) & FrmtCampo(!TipoDocVer, "N", 2) & FrmtCampo(!NumDocVer, "N", 15) & !FecVeri & FrmtCampo(!codDDNfaxemp, "N", 3) & FrmtCampo(!numfaxemp, "N", 8) & FrmtCampo(!AnxFaxEmp, "N", 5) & FrmtCampo(!Dptotrab, "A", 80) & FrmtCampo(!Provtrab, "A", 80) & _
        FrmtCampo(!Disttrab, "A", 80) & FrmtCampo(!DptoEmp, "A", 80) & FrmtCampo(!ProvEmp, "A", 80) & FrmtCampo(!DistEmp, "A", 80) & FrmtCampo(!DesCargo, "A", 80) & FrmtCampo(!DesOcup, "A", 80) & FrmtCampo(!DesGiro, "A", 50) & FrmtCampo(!AnxTel, "N", 5)
    
        .MoveNext
    Loop
 
 End If
 
 SumarioTraspaso .RecordCount + 2
 Close #1
        
End With

End Sub
Sub ExpoNAfiliacion(Envio As String, FecTran As String, ARCHIVO As String)
Dim rsAfiliacion As Recordset
Set rsAfiliacion = New Recordset

Open ARCHIVO For Output As #1

Call CabeAfiliacion(FecTran)

'Detalles
With rsAfiliacion

    'If Envio = EnvioUltimo Then
    If chkReingreso.Value = 1 Then
    .Open " exec pObtenerDatosDetalle3 '" & Envio & "', 'C', '" & txtFecAcep & "'  ", Con.Cnx, adOpenStatic, adLockReadOnly
    Else
    .Open " exec pObtenerDatosDetalle2 '" & Envio & "', 'C'  ", Con.Cnx, adOpenStatic, adLockReadOnly
    End If
    
 If .RecordCount > 0 Then
    
  If RevisionNulos(rsAfiliacion, "Archivo Contratos") Then
    Close #1
    Exit Sub
  End If
  
  Do While Not .EOF
    Print #1, !TipReg & FrmtCampo(!Numsol, "A", 15) & !Fecsus & FrmtCampo(!Nomuno, "A", 20) & FrmtCampo(!Nomdos, "A", 20) & _
        FrmtCampo(!ApePat, "A", 30) & FrmtCampo(!ApeMat, "A", 30) & !Fecnac & !Sexo & FrmtCampo(!estciv, "A", 1) & FrmtCampo(!TipDoc, "N", 2) & FrmtCampo(!numdoc, "A", 15) & _
        FrmtCampo(!codnac, "N", 3) & FrmtCampo(!ubinac, "N", 6) & !codpai & FrmtCampo(!ubidir, "N", 6) & FrmtCampo(!prefijodir, "N", 2) & FrmtCampo(!DetDir, "A", 100) & FrmtCampo(!prefijoloc, "N", 2) & _
        FrmtCampo(!detdirloc, "A", 100) & FrmtCampo(!refdirtra, "A", 100) & FrmtCampo(!codddnTra, "N", 3) & FrmtCampo(!telefono1, "N", 9) & FrmtCampo(!EmailTra, "A", 80) & FrmtCampo(!rucemp, "N", 11) & _
        FrmtCampo(!razsocemp, "A", 120) & FrmtCampo(!ubiemp, "N", 6) & FrmtCampo(!prefijodiremp, "N", 2) & FrmtCampo(!diremp, "A", 100) & FrmtCampo(!prefijolocemp, "N", 2) & FrmtCampo(!detdirlocemp, "A", 100) & _
        FrmtCampo(!RefDirEmp, "A", 100) & FrmtCampo(!CodDDNEmp, "N", 3) & FrmtCampo(!teleEmp, "N", 9) & FrmtCampo(!tiptra, "A", 1) & FrmtCampo(!firtra, "A", 1) & FrmtCampo(!firpro, "A", 1) & FrmtCampo(!huella, "A", 1) & FrmtCampo(!Cuspp, "A", 12) & FrmtCampo(!copdni, "A", 1) & _
        FrmtCampo(!ley19990, "A", 1) & FrmtCampo(!ley20530, "A", 1) & FrmtCampo(!periodicidad, "A", 1) & FrmtCampo(!priempleo, "A", 1) & FrmtCampo(!essalud, "A", 15) & !anibodas & FrmtCampo(!codcargo, "A", 3) & FrmtCampo(!codpaisnac, "N", 3) & FrmtCampo(!codocupac, "A", 4) & _
        FrmtCampo(!compor, "N", 5) & FrmtCampo(!comser, "N", 5) & !fecinilab & FrmtCampo(!firfed, "A", 1) & FrmtCampo(!firfedrev, "A", 1) & FrmtCampo(!codgironec, "N", 5) & FrmtCampo(!autoenvio, "A", 1) & FrmtCampo(!derbono, "A", 1) & FrmtCampo(!numhijos, "N", 2) & FrmtCampo(!Ram, "N", 11) & _
        FrmtCampo(!decjur50, "N", 1) & FrmtCampo(!CodPromo, "A", 10) & FrmtCampo(!codDDNfaxemp, "N", 3) & FrmtCampo(!numfaxemp, "N", 8) & FrmtCampo(!AnxFaxEmp, "N", 5) & FrmtCampo(!tipafil, "A", 1) & FrmtCampo(!firafirev, "A", 1) & FrmtCampo(!firafprev, "A", 1) & FrmtCampo(!codadm, "A", 10) & _
        FrmtCampo(!codoriafp, "A", 2) & FrmtCampo(!copbolpag, "A", 1) & FrmtCampo(!Dptotrab, "A", 80) & FrmtCampo(!Provtrab, "A", 80) & FrmtCampo(!Disttrab, "A", 80) & FrmtCampo(!DptoEmp, "A", 80) & FrmtCampo(!ProvEmp, "A", 80) & FrmtCampo(!DistEmp, "A", 80) & FrmtCampo(!DesCargo, "A", 80) & FrmtCampo(!DesOcup, "A", 80) & _
        FrmtCampo(!DesGiro, "A", 50) & FrmtCampo(!AnxTel, "N", 5) & !FecsusRev
                
        .MoveNext
    Loop
    ' + 2 Encabezado Sumario
 End If
 
 Call SumarioAfiliacion(.RecordCount + 2)
        
 Close #1
        

        
            
End With


End Sub

Sub ExpoTelefonos(Envio As String, FecTran As String, ARCHIVO As String, TipoDoc As String)

Dim RsTelefono As Recordset
Set RsTelefono = New ADODB.Recordset


With RsTelefono

'If Envio = EnvioUltimo Then
If chkReingreso.Value = 1 Then
.Open "pObtenerTelefonos3 '" & Envio & "','" & TipoDoc & "','" & txtFecAcep & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
Else
.Open "pObtenerTelefonos2 '" & Envio & "','" & TipoDoc & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
End If



Open ARCHIVO For Output As #1

 CabeTelefonos FecTran, "E", "0032", 19

 If .RecordCount > 0 Then
     
    If RevisionNulos(RsTelefono, "Archivo Teléfonos") Then
      Close #1
      Exit Sub
    End If
    
    
    Do While Not .EOF
    
      
      Print #1, !TipReg & FrmtCampo(IIf(TipoDoc = "T", "PR0", "") & !Numsol, "A", 15) & FrmtCampo(!TipTel, "A", 1) & FrmtCampo(!DDN, "N", 3) & FrmtCampo(!numtel, "N", 9) & FrmtCampo(!Anx, "N", 4)
      .MoveNext
    Loop

 End If
 
 Call Sumario("S", Format(CStr(.RecordCount + 2), "0000000#"), 23)

Close #1
End With

End Sub
Sub CabeTelefonos(FechaTransmision As String, TipReg As String, Tamano As String, EspaciosB As String)
Print #1, TipReg & Tamano & FechaTransmision & Space(EspaciosB)
End Sub
Sub Sumario(TipReg As String, NroReg As String, EspaciosB As String)
Print #1, TipReg & Format(NroReg, "0000000#") & Space(EspaciosB)
End Sub
Sub ExpoNDirecciones(Envio As String, FecTran As String, ARCHIVO As String, TipoDoc As String)
Dim RsDirecciones As Recordset
Set RsDirecciones = New ADODB.Recordset


With RsDirecciones
'If Envio = EnvioUltimo Then
If chkReingreso.Value = 1 Then
.Open "pObtenerDirecciones3 '" & Envio & "','" & TipoDoc & "','" & txtFecAcep & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
Else
.Open "pObtenerDirecciones2 '" & Envio & "','" & TipoDoc & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
End If

Open ARCHIVO For Output As #1

CabeDire FecTran, "E", "0567", 554

 If .RecordCount > 0 Then

  If RevisionNulos(RsDirecciones, "Archivo Direcciones") Then Exit Sub

  Do While Not .EOF
    Print #1, !TipReg & FrmtCampo(IIf(TipoDoc = "T", "PR0", "") & !Numsol, "A", 15) & FrmtCampo(!Tipdir, "A", 1) & FrmtCampo(!Ubigeo, "N", 6) & FrmtCampo(!PrefDir, "N", 2) & FrmtCampo(!DetDir, "A", 100) & FrmtCampo(!PrefLoc, "N", 2) & _
              FrmtCampo(!DetLoc, "A", 100) & FrmtCampo(!RefDir, "A", 100) & FrmtCampo(!Departamento, "A", 80) & FrmtCampo(!Provincia, "A", 80) & FrmtCampo(!Distrito, "A", 80)
    .MoveNext
  Loop
    
 End If
 
 Call Sumario("S", CStr(.RecordCount + 2), 558)

 Close #1

End With

End Sub
Sub CabeDire(FechaTransmision As String, TipReg As String, Tamano As String, EspaciosB As String)
Print #1, TipReg & Tamano & FechaTransmision & Space(EspaciosB)
End Sub

Sub ExpoNEmpleadoresTraspasos(Envio As String, FecTran As String, ARCHIVO As String)
Dim RsEmpleadores As Recordset
Set RsEmpleadores = New ADODB.Recordset


With RsEmpleadores
'If Envio = EnvioUltimo Then
If chkReingreso.Value = 1 Then
.Open " pObtenerEmpleadores3 '" & Envio & "','T', '" & txtFecAcep & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
Else
.Open " pObtenerEmpleadores2 '" & Envio & "','T' ", Con.Cnx, adOpenStatic, adLockReadOnly
End If

Open ARCHIVO For Output As #1

  CabeDire FecTran, "E", "0795", 782

 If .RecordCount > 0 Then
 
  If RevisionNulos(RsEmpleadores, "Archivo Empleadores Traspasos") Then
    Exit Sub
    Close #1
  End If
    
  Do While Not .EOF
    Print #1, !TipReg & FrmtCampo("PR0" & !Numsol, "A", 15) & FrmtCampo(!numruc, "N", 11) & FrmtCampo(!RazSoc, "A", 120) & FrmtCampo(!CodPais, "N", 3) & FrmtCampo(!Ubigeo, "N", 6) & FrmtCampo(!PrefDir, "N", 2) & _
              FrmtCampo(!DetDir, "A", 100) & FrmtCampo(!PreFijLoc, "N", 2) & FrmtCampo(!DetLoc, "A", 100) & FrmtCampo(!RefDir, "A", 100) & FrmtCampo(!CodDDN, "N", 3) & FrmtCampo(!TelEmp, "N", 9) & _
              FrmtCampo(!CodGirNec, "N", 5) & FrmtCampo(!Ram, "N", 11) & FrmtCampo(!CodDDNEmp, "N", 3) & FrmtCampo(!NroEmpFax, "N", 9) & FrmtCampo(!AnxFaxEmp, "N", 5) & FrmtCampo(!Departamento, "A", 80) & _
              FrmtCampo(!Provincia, "A", 80) & FrmtCampo(!Distrito, "A", 80) & FrmtCampo(!DesGiro, "A", 50) & FrmtCampo(!AnxTelEmp, "N", 5)
              
    .MoveNext
  Loop



 End If

Call Sumario("S", CStr(.RecordCount + 2), 786)
Close #1
End With

End Sub
Sub ExpoNEmpleadoresAfiliacion(Envio As String, FecTran As String, ARCHIVO As String)
Dim RsEmpleadores As Recordset
Set RsEmpleadores = New ADODB.Recordset

With RsEmpleadores
 'If Envio = EnvioUltimo Then
 If chkReingreso.Value = 1 Then
  .Open "pObtenerEmpleadores3 '" & Envio & "','C', '" & txtFecAcep & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
 Else
  .Open "pObtenerEmpleadores2 '" & Envio & "','C' ", Con.Cnx, adOpenStatic, adLockReadOnly
 End If

Open ARCHIVO For Output As #1

 CabeDire FecTran, "E", "0795", 782

 If .RecordCount > 0 Then
  
  If RevisionNulos(RsEmpleadores, "Archivo Empleadores Contratos") Then
    Exit Sub
    Close #1
  End If
    
    
  Do While Not .EOF
    Print #1, !TipReg & FrmtCampo(!Numsol, "A", 15) & FrmtCampo(!numruc, "N", 11) & FrmtCampo(!RazSoc, "A", 120) & FrmtCampo(!CodPais, "N", 3) & FrmtCampo(!Ubigeo, "N", 6) & FrmtCampo(!PrefDir, "N", 2) & _
              FrmtCampo(!DetDir, "A", 100) & FrmtCampo(!PreFijLoc, "N", 2) & FrmtCampo(!DetLoc, "A", 100) & FrmtCampo(!RefDir, "A", 100) & FrmtCampo(!CodDDN, "N", 3) & FrmtCampo(!TelEmp, "N", 9) & _
              FrmtCampo(!CodGirNec, "N", 5) & FrmtCampo(!CodDDNEmp, "N", 3) & FrmtCampo(!NroEmpFax, "N", 9) & FrmtCampo(!AnxFaxEmp, "N", 5) & FrmtCampo(!Departamento, "A", 80) & _
              FrmtCampo(!Provincia, "A", 80) & FrmtCampo(!Distrito, "A", 80) & FrmtCampo(!DesGiro, "A", 50) & FrmtCampo(!AnxTelEmp, "N", 5) & FrmtCampo(!Ram, "N", 11)
              
    .MoveNext
  Loop

  
 End If
 
 Call Sumario("S", CStr(.RecordCount + 2), 786)

 Close #1

End With

End Sub


Sub ExpoNHojaVida(Envio As String, FecTran As String, ARCHIVO As String, TipDoc As String)
Dim RsHojaVida As Recordset
Set RsHojaVida = New ADODB.Recordset

With RsHojaVida

'If Envio = EnvioUltimo Then
If chkReingreso.Value = 1 Then
.Open " pObtenerHojaVida3 '" & Envio & "','" & TipDoc & "', '" & txtFecAcep & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
Else
.Open " pObtenerHojaVida2 '" & Envio & "','" & TipDoc & "' ", Con.Cnx, adOpenStatic, adLockReadOnly
End If
    
Open ARCHIVO For Output As #1


 If TipDoc = "T" Then
  CabeDire FecTran, "E", "1914", 1901
 Else
  CabeDire FecTran, "E", "1833", 1820
 End If


If .RecordCount > 0 Then

  If RevisionNulos(RsHojaVida, "Archivo Hoja de Vida") Then
     Close #1
     Exit Sub
  End If
   
  Do While Not .EOF
    Print #1, !TipReg & FrmtCampo(IIf(TipDoc = "T", "PR0", "") & !Numsol, "A", 15) & FrmtCampo(!Nomuno, "A", 20) & FrmtCampo(!Nomdos, "A", 20) & FrmtCampo(!ApePat, "A", 30) & FrmtCampo(!ApeMat, "A", 30) & FrmtCampo(!TipDoc, "N", 2) & _
              FrmtCampo(!numdoc, "A", 15) & FrmtCampo(!DesCargo, "A", 50) & FrmtCampo(!CodDDNCasa, "N", 3) & FrmtCampo(!TelefCasa, "N", 9) & FrmtCampo(!CodDDNTel, "N", 3) & FrmtCampo(!Celular, "N", 9) & _
              FrmtCampo(!EmailCasa, "A", 80) & FrmtCampo(!EmailTra, "A", 80) & FrmtCampo(!RefDir, "A", 100) & FrmtCampo(!CodDDNEmp, "N", 3) & FrmtCampo(!TelEmp, "N", 9) & _
              FrmtCampo(!AnxEmp, "N", 5) & FrmtCampo(!Tipdir, "A", 1) & FrmtCampo(!UbigeoDir, "N", 6) & FrmtCampo(!PrefDir, "N", 2) & FrmtCampo(!DetDirEmp, "A", 100) & FrmtCampo(!PrefLocEmp, "N", 2) & _
              FrmtCampo(!DetLocEmp, "A", 100) & FrmtCampo(!RefDirEmp, "A", 100) & FrmtCampo(!Departamento, "A", 80) & FrmtCampo(!Provincia, "A", 80) & FrmtCampo(!Distrito, "A", 80) & FrmtCampo(!CodSup, "A", 10) & _
              FrmtCampo(!CodInPro, "A", 10) & FrmtCampo(!Renta, "N", 9) & FrmtCampo(!Agencia, "A", 3)
              
    .MoveNext
  Loop

  Call Sumario("S", CStr(.RecordCount + 2), 786)
  
End If

Close #1

End With

End Sub

Sub SuspenderSolicitudes(Fecha As String)

Dim cmdSusp As New ADODB.Command


With cmdSusp
    .ActiveConnection = Con.Cnx
    .CommandText = " pProcSuspender ?"
    .Parameters("@fecsal") = Fecha
    .Execute , , adExecuteNoRecords
End With

Set cmdSusp = Nothing

End Sub


Sub ExportarAceptadas()
Dim sqltexto As String, Condicion As String
Dim RsExpo As Recordset
Dim rsExpo2 As Recordset
Dim rsExpo3 As Recordset
Dim antAgencia As String

Set RsExpo = New ADODB.Recordset
Set rsExpo2 = New ADODB.Recordset
Set rsExpo3 = New ADODB.Recordset

'If txtEnvio.Text = "" And txtFecSalida = "" Then MsgBox "Debe Ingresar Fecha de salida o Nro de Envio", vbCritical, Caption: Exit Sub

sqltexto = "select a.fecsal, Envio  from Ctrldoc a join Solicitud b on a.lote = b.lote  where "

              Condicion = " "
              
              If txtFecAcep <> "" Then
                Condicion = " convert(char, b.FecSusp ,103) = '" & txtFecAcep.Text & "' "
              End If

'              If txtEnvio.Text <> "" Then
'                Condicion = Condicion & " envio = '" & Trim(txtEnvio.Text) & "'"
'              End If
              
              If Trim(Right(Condicion, 4)) = "and" Then Condicion = Left(Condicion, Len(Condicion) - 4)
              
sqltexto = sqltexto & Condicion & " group by a.fecsal, Envio"

With RsExpo
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
    If .RecordCount > 0 Then
    
        Do While Not .EOF
        
                'Afiliacion
        
                sqltexto = " Select a.NumSol, a.FecSal, a.indDoc, b.Envio, a.Lote, b.Agencia from Solicitud a join Ctrldoc b on a.Lote = b.Lote " & _
                           " where Envio = '" & .Fields("Envio") & "' and indtipdoc = '1' and  a.secu = 0  and isnull(indSusp,0) = 1 and  isnull(IndAcep,0) = 1 and convert(varchar,FecSusp,103) = '" & Me.txtFecAcep & "' order by b.Agencia, a.Lote "
                                                                     
                           
                If rsExpo2.State = 1 Then rsExpo2.Close
                rsExpo2.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

                antAgencia = ""

                If rsExpo2.RecordCount > 0 Then
                        NomArch = App.Path & "\Expo\A" & Format(.Fields("fecsal"), "yymmdd") & "_" & Format(Trim(.Fields("Envio")), "00000#") & "Acep.dat"
                        Open NomArch For Output As #1


                        If Not ExisteDirectorio(Trim(App.Path) & "\Expo") Then MsgBox "Directorio Expo no existe, debe crear el directorio Expo ", vbCritical, Caption: Exit Sub
                        
                        Do While Not rsExpo2.EOF

                          If antAgencia <> rsExpo2.Fields("Agencia") Then
                            CabeceraDocumentoAcep rsExpo2.Fields("Agencia"), rsExpo2.Fields("Lote"), rsExpo2.Fields("Fecsal"), .Fields("Envio"), "C"
                            antAgencia = rsExpo2.Fields("Agencia")
                          End If


                          ExportarAfiliacion rsExpo2.Fields("Numsol"), rsExpo2.Fields("Fecsal"), True
                          '  ExportarTraspaso .Fields("Numsol"), .Fields("Fecsal")
                          rsExpo2.MoveNext
                        Loop
                End If
                Close #1                ' A   C   T   I  V   A   R
                antAgencia = ""
                                
                sqltexto = " Select a.NumSol, a.FecSal, a.indDoc, b.Envio, a.Lote, b.Agencia from Solicitud a join Ctrldoc b on a.Lote = b.Lote " & _
                           " where Envio = '" & .Fields("Envio") & "' and indtipdoc = '0' and a.secu = 0 and isnull(indSusp,0) = 1 and  isnull(IndAcep,0) = 1 and convert(varchar,FecSusp,103) = '" & Me.txtFecAcep & "' order by b.Agencia, a.Lote "
                           
                If rsExpo3.State = 1 Then rsExpo3.Close
                rsExpo3.Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
                If rsExpo3.RecordCount > 0 Then
                        NomArch = App.Path & "\Expo\T" & Format(.Fields("fecsal"), "yymmdd") & "_" & Format(Trim(.Fields("Envio")), "00000#") & "Acep.dat"
                        
                        Open NomArch For Output As #1
                        
                        If Not ExisteDirectorio(Trim(App.Path) & "\Expo") Then MsgBox "Directorio Expo no existe, debe crear el directorio Expo ", vbCritical, Caption: Exit Sub
                        Do While Not rsExpo3.EOF
                        
                          If antAgencia <> rsExpo3.Fields("Agencia") Then
                            CabeceraDocumentoAcep rsExpo3.Fields("Agencia"), rsExpo3.Fields("Lote"), rsExpo3.Fields("Fecsal"), .Fields("Envio"), "T"
                            antAgencia = rsExpo3.Fields("Agencia")
                          End If
                        
                          ExportarTraspaso rsExpo3.Fields("Numsol"), rsExpo3.Fields("Fecsal"), True
                          rsExpo3.MoveNext
                        Loop
                        
                        Close #1
                End If
                
                
                
                .MoveNext
         Loop
    End If
End With
    
        
ActualizarMensaje " Exportación Rechazados Concluida "
    
End Sub

Sub ExpoArchivoSBS(ARCHIVO, Envio As String)
On Error GoTo ERRORES
Dim RsSbs As Recordset
Set RsSbs = New ADODB.Recordset

With RsSbs

.Open "select * from vt_IA1 where fecsal = '" & txtFecSalida & "' order by  Apepat, Apemat, Nombre1, Nombre2 ", Con.Cnx, adOpenStatic, adLockReadOnly

    Open ARCHIVO For Output As #1
    
    
     If .RecordCount > 0 Then
      
     If RevisionNulos(RsSbs, "Archivo para SBS") Then
        Close #1
        Exit Sub
     End If
     
     Do While Not .EOF
        
                  
        Print #1, "PR" & FrmtCampo(!Numsol, "A", 7) & FrmtCampo(!ApePat, "A", 20) & FrmtCampo(!ApeMat, "A", 20) & FrmtCampo(!Nombre1, "A", 10) & FrmtCampo(!Nombre2, "A", 10) & FrmtCampo(!Fecnac, "A", 8) & _
                  FrmtCampo(!Fecsus, "A", 8) & FrmtCampo(!Nacionalidad, "A", 3) & FrmtCampo(!TipNumdoc, "A", 10) & FrmtCampo(!TipTrabajador, "A", 1) & FrmtCampo(!Origen, "A", 1) & FrmtCampo(!IPSS, "A", 15) & _
                  FrmtCampo(!Sexo, "A", 1) & FrmtCampo(!CodPromo, "N", 5) & FrmtCampo(!AFP, "A", 2) & "00"
                  
        .MoveNext
      Loop
    
      
     End If
     
    
    Close #1

End With


Exit Sub
ERRORES:
MsgBox Err.Description, vbCritical, "Generación archivo IAI"


End Sub

Function RevisionNulos(ByRef registro As Recordset, Titulo As String) As Boolean
RevisionNulos = False

With registro

Do While Not .EOF
    For i = 0 To registro.Fields.Count - 1
        If IsNull(registro.Fields(i)) Then
            
            MsgBox "campo nulo " & .Fields(i).Name & " nro solicitud " & IIf(IsNull(.Fields(1)), "", .Fields(1)), vbCritical, Titulo
            RevisionNulos = True
            Exit Function
        End If
    Next
    .MoveNext
Loop

.MoveFirst
End With
End Function

Sub MostrarEnvios(Fecha As String)
Dim sqltexto As String
sqltexto = "select distinct envio, c1 = envio from ctrldoc where fecsal = '" & Fecha & "' order by envio "
'ccombosql lstEnvio, sqltexto
CargarListaSql lstEnvio, sqltexto

End Sub

Sub RegistrarExpoLog(Fecha As String, Envio As String)

Dim sentencia As String
Dim cmd As New ADODB.Command

sentencia = " Insert into LogExpo (FechaExpo, Envio) values ('" & Fecha & "', '" & Envio & "'  ) "

With cmd
    .ActiveConnection = Con.Cnx
    .CommandText = sentencia
    .Execute , , adExecuteNoRecords
End With

Set cmd = Nothing





End Sub

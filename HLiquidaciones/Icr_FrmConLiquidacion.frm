VERSION 5.00
Begin VB.Form Icr_FrmConLiquidacion 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consulta Liquidación Procesada"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   Icon            =   "Icr_FrmConLiquidacion.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5220
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txtimagen 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busqueda Id Imagen"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busqueda Standar"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   5895
         Begin VB.ComboBox txtitem 
            Height          =   315
            ItemData        =   "Icr_FrmConLiquidacion.frx":030A
            Left            =   4800
            List            =   "Icr_FrmConLiquidacion.frx":034A
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   360
            Width           =   750
         End
         Begin VB.ComboBox cbodato 
            Height          =   315
            ItemData        =   "Icr_FrmConLiquidacion.frx":0395
            Left            =   1800
            List            =   "Icr_FrmConLiquidacion.frx":03C3
            TabIndex        =   15
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Campo a Modificar"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblitem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nro Item"
            Height          =   255
            Left            =   3960
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Criterios de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   5895
         Begin VB.TextBox Txtliquidacion 
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtpag 
            Height          =   285
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Liquidacion"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nro Pagina"
            Height          =   255
            Left            =   3960
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   6120
         TabIndex        =   8
         Top             =   4080
         Width           =   1695
      End
      Begin VB.ComboBox cmbBatch 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdactualizar 
         Caption         =   "Actualizar"
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Frame Frdatos 
         BackColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   3000
         Width           =   5895
         Begin VB.TextBox txtdato 
            BackColor       =   &H0080FFFF&
            Height          =   405
            Left            =   2280
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label lblcampo 
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
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdconsulta 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   6120
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblimagen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Imagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblEnvio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nùmero de batch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   765
         Width           =   1410
      End
   End
End
Attribute VB_Name = "Icr_FrmConLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xcampo As String
Dim StrPag As String
Dim Strid As Double
Dim StrNombreCampo  As String

Private Sub Command1_Click()

End Sub

Private Sub cbodato_Change()
    If cbodato.ListIndex = 0 Or cbodato.ListIndex = 1 Then
    lblitem.Visible = False
    txtitem.Visible = False
    Else
    lblitem.Visible = True
    txtitem.Visible = True
    End If
End Sub

Private Sub cbodato_Click()
If cbodato.ListIndex = 0 Or cbodato.ListIndex = 1 Or cbodato.ListIndex = 5 Or cbodato.ListIndex = 6 Or cbodato.ListIndex = 7 Or cbodato.ListIndex = 8 Or cbodato.ListIndex = 9 Or cbodato.ListIndex = 11 Or cbodato.ListIndex = 12 Then
    lblitem.Visible = False
    txtitem.Visible = False
    Limpiar1
    Else
    lblitem.Visible = True
    txtitem.Visible = True
    Limpiar1
    End If
End Sub

Private Sub cmdactualizar_Click()
If Opt(0).Value = True Then
strEjecutar = "Update tb_liquidacion set  " & xcampo & "= '" & txtdato.Text & "' where ID_liq=" & Strid & ""
Else
strEjecutar = "Update tb_liquidacion set  " & xcampo & "= '" & txtdato.Text & "' where csid='" & txtimagen.Text & ".tif" & "'"
End If
cnServer.Execute strEjecutar
MsgBox "Proceso concluido Satisfactoriamente", vbInformation, "Horizonte"
Limpiar
End Sub

Private Sub cmdconsulta_Click()

        If rsSQL.State = 1 Then rsSQL.Close
            StrPag = Format(txtpag.Text, "00000")
            'rsSQL.Source = "Select id_hojapedido from rechazos where id_tiporechazo=1 or id_tiporechazo=2 or id_tiporechazo=3"
            If Opt(0).Value = True Then
            rsSQL.Source = "Select * from tb_liquidacion where  DS_BARRA='" & Mid(Txtliquidacion.Text, 3, 11) & Trim$((StrPag)) & "' and ds_batch='" & cmbBatch & "'"
            Else
            rsSQL.Source = "Select * from tb_liquidacion where csid='" & txtimagen.Text & ".tif" & "'"
            End If
            
            rsSQL.Open
            If rsSQL.RecordCount > 0 Then
              If rsSQL.RecordCount < 2 Then
                   
                   If rsSQL.RecordCount = 1 Then
                        If cbodato.ListIndex = 0 Or cbodato.ListIndex = 1 Or cbodato.ListIndex = 5 Or cbodato.ListIndex = 6 Or cbodato.ListIndex = 7 Or cbodato.ListIndex = 8 Or cbodato.ListIndex = 9 Then
                        xcampo = campos(cbodato.Text)
                        Else
                          xcampo = campos(cbodato.Text) & txtitem
                        End If
                        Strid = IIf(IsNull(rsSQL.Fields("id_liq")), 0, rsSQL.Fields("id_liq"))
                        StrNombreCampo = Nombre_campo(cbodato.Text)
                        lblcampo.Caption = StrNombreCampo
                        txtdato.Text = "" & rsSQL.Fields(xcampo)
                        txtdato.Visible = True
                        lblcampo.Visible = True
                        Frdatos.Caption = "Actualizar Dato"
                   Else
                        MsgBox "Existe mas de un registro con estos datos", vbInformation, "Horizonte"
                        Limpiar
                   End If
               Else
               MsgBox "El registro se encuentra Duplicado", vbInformation, "Horizonte"
               End If
                'lstMensajes.AddItem "Liquidaciones Previas rechazados : " & rsSQL.RecordCount
            Else
                MsgBox "El registro no se encuentra", vbInformation, "Horizonte"
                Limpiar
            End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
        rsSQL.CursorType = adOpenDynamic
        rsSQL.LockType = adLockBatchOptimistic
        rsSQL.ActiveConnection = cnServer
        ListarBatch
        If cmbBatch.ListCount > 0 Then
        cmbBatch.ListIndex = 0
        End If
        cbodato.ListIndex = 0
        txtitem.ListIndex = 0
End Sub

Sub ListarBatch()
    cmbBatch.Clear
    If rsSQL.State = 1 Then rsSQL.Close
    'rsSQL.Source = "Select distinct ds_batch from afiliacion order by 1"
    rsSQL.Source = "Select distinct ds_batch from TB_liquidacion order by 1"
    rsSQL.Open
    If rsSQL.RecordCount > 0 Then
        rsSQL.MoveFirst
        Do While Not rsSQL.EOF
            cmbBatch.AddItem rsSQL!DS_batch
            rsSQL.MoveNext
        Loop
    End If
    If rsSQL.State = 1 Then rsSQL.Close
End Sub

Function campos(xcampo As String) As String
    Select Case xcampo
     Case "Nro_Liquidacion"
     campos = "ds_nroliq"
     Case "Pag"
     campos = "ds_Pag"
     Case "Cod"
     campos = "ds_cod"
     Case "Fecha Incio"
     campos = "ds_fec_ini"
     Case "Fecha Final"
     campos = "ds_fec_ce"
     Case "Nro Tramite"
     campos = "ds_numerotra"
     Case "Email"
     campos = "ds_email"
     Case "Telefono"
     campos = "ds_telefono"
     Case "Barra"
     campos = "ds_barra"
     Case "Dni"
     campos = "ds_dni"
     Case "Remuneracion"
     campos = "ds_rem_as"
     Case "Ruc R.Social"
     campos = "ds_razon_ruc"
     Case "Firma"
     campos = "ds_nombre"
     Case "Planilla"
     campos = "ds_pla"
    End Select
End Function

Function Nombre_campo(xcampo As String) As String
    Select Case xcampo
     Case "Nro_Liquidacion"
     Nombre_campo = "Nro Liquidacion:"
     Case "Pag"
     Nombre_campo = "Pagina:"
     Case "Cod"
     Nombre_campo = "Codigo"
     Case "Fecha Incio"
     Nombre_campo = "Fecha Inicio:"
     Case "Fecha Final"
     Nombre_campo = "Fecha Final:"
     Case "Nro Tramite"
     Nombre_campo = "Nro Tramite:"
     Case "Email"
     Nombre_campo = "Email:"
     Case "Telefono"
     Nombre_campo = "Telefono:"
     Case "Barra"
     Nombre_campo = "Barra:"
     Case "Dni"
     Nombre_campo = "Dni:"
     Case "Remuneracion"
     Nombre_campo = "Remuneracion:"
     Case "Ruc R.Social"
     Nombre_campo = "Ruc_Razon_Social:"
     Case "Firma"
     Nombre_campo = "Nombre_Firma:"
     Case "Planilla"
     Nombre_campo = "Planilla:"
    End Select
End Function

Public Sub Limpiar()
txtdato.Text = ""
lblcampo.Caption = ""
Txtliquidacion.Text = ""
txtpag.Text = ""
txtitem.Text = ""

txtdato.Visible = False
lblcampo.Visible = False
End Sub

Public Sub Limpiar1()
txtdato.Text = ""
lblcampo.Caption = ""
'Txtliquidacion.Text = ""
'txtpag.Text = ""
'txtitem.Text = ""

txtdato.Visible = False
lblcampo.Visible = False
End Sub

Private Sub Opt_Click(Index As Integer)
  If Opt(1).Value = True Then
  cmbBatch.Enabled = False
  Frame2.Enabled = False
  lblimagen.Visible = True
  txtimagen.Visible = True
  Else
  lblimagen.Visible = False
  txtimagen.Visible = False
  cmbBatch.Enabled = True
  Frame2.Enabled = True
  End If
End Sub

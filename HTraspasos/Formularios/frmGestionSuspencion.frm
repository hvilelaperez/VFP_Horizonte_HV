VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGestionSuspencion 
   Caption         =   "Aceptar o Rechazar Solicitudes"
   ClientHeight    =   8655
   ClientLeft      =   300
   ClientTop       =   1680
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSuspender 
      Caption         =   "&Suspender"
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
      Left            =   5400
      TabIndex        =   11
      Top             =   8040
      Width           =   1695
   End
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
      Height          =   495
      Left            =   13080
      TabIndex        =   10
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar "
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
      Left            =   9360
      TabIndex        =   9
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptarS 
      Caption         =   "&Aceptar "
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
      Left            =   7440
      TabIndex        =   8
      Top             =   8040
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   11033
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Envio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha Salida"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Agencia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nro Solicitud"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nombre "
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Estado"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fec. Susp."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Fec. Acept."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.ComboBox cboEstado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmGestionSuspencion.frx":0000
         Left            =   3840
         List            =   "frmGestionSuspencion.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFecSalFin 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtFecSalIni 
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
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Estado"
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
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Salida "
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
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGestionSuspencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsSol As Recordset

Sub BuscarSolicitudes()
Dim texto As String
Dim X As ListItem
'Screen.MousePointer = MousePointerConstants.vbHourglass

texto = " select b.Envio, a.Fecsal,  b.Agencia,  a.Numsol,  Nombrec = rtrim(isnull(Apepat,'')) + ' ' + rtrim(isnull(Apemat,'')) + ' ' + rtrim(isnull(Nombre1,'')) + ' ' +  rtrim(isnull(Nombre2,'')) , FecSusp = convert(varchar,FecSusp,103)  , FecAprob = convert(varchar,FecAprob,103) , " & _
        " EstadoSol = case when indSusp = 1 and indAcep is null  then 'Suspendida' " & _
                         " when indSusp = 1 and indAcep = 0 then 'Rechazada' " & _
                         " when indSusp = 1 and indAcep = 1 then 'Aceptada' end " & _
                         " from Solicitud a join Ctrldoc b on a.Lote = b.Lote " & _
                         " where indSusp = 1 "
        
If txtFecSalIni <> "" And txtFecSalFin <> "" Then
    Condicion = " and a.fecsal >= '" & txtFecSalIni & "' and a.fecsal <= '" & txtFecSalFin & "'"
End If
    
If cboEstado.ListIndex <> -1 Then

    Select Case cboEstado.Text
    
    Case "Todos"
    Case "Suspendidas"
        Condicion = Condicion & " and IndAcep is null "
    Case "Aceptadas"
        Condicion = Condicion & " and IndAcep = 1 "
    Case "Rechazadas"
        Condicion = Condicion & " and IndAcep = 0 "
    End Select
    
End If
    
        
texto = texto & Condicion & " order by a.fecsal, a.numsol "
        
If RsSol.State = 1 Then RsSol.Close
RsSol.Open texto, Con.Cnx, adOpenStatic
RsSol.Supports adResync
'********
ListView1.ListItems.Clear
ListView1.Refresh
ListView1.Checkboxes = True
If Not (RsSol.EOF And RsSol.BOF) Then
RsSol.MoveFirst
While Not RsSol.EOF()
    
    Set X = ListView1.ListItems.Add(, , IIf(IsNull(Trim(RsSol.Fields("Envio"))), "", Trim(RsSol.Fields("Envio"))))
    'If BuscarBenefGrabado(RsxBene.Fields("Sec_Ben")) Then X.Checked = True
    X.SubItems(1) = IIf(IsNull(Trim(RsSol.Fields("Fecsal"))), "", Trim(RsSol.Fields("Fecsal")))
    X.SubItems(2) = IIf(IsNull(Trim(RsSol.Fields("Agencia"))), "", Trim(RsSol.Fields("Agencia")))
    X.SubItems(3) = IIf(IsNull(Trim(RsSol.Fields("Numsol"))), "", Trim(RsSol.Fields("Numsol")))
    X.SubItems(4) = CStr(IIf(IsNull(Trim(RsSol.Fields("Nombrec"))), "", Trim(RsSol.Fields("Nombrec"))))
    X.SubItems(5) = CStr(IIf(IsNull(Trim(RsSol.Fields("EstadoSol"))), "", Trim(RsSol.Fields("EstadoSol"))))
    X.SubItems(6) = IIf(IsNull(Trim(RsSol.Fields("FecSusp"))), "", Trim(RsSol.Fields("FecSusp")))
    X.SubItems(7) = IIf(IsNull(Trim(RsSol.Fields("FecAprob"))), "", Trim(RsSol.Fields("FecAprob")))
      
    RsSol.MoveNext
Wend

RsSol.MoveFirst

Else
      
End If
Screen.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub cmdAceptarS_Click()
ProcesarAceptador
BuscarSolicitudes
End Sub

Private Sub cmdBuscar_Click()
BuscarSolicitudes
End Sub

Private Sub cmdRechazar_Click()
ProcesarRechazos
BuscarSolicitudes
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub


Private Sub cmdSuspender_Click()
ProcesarSuspendidos
BuscarSolicitudes
End Sub

Private Sub Form_Load()
Set RsSol = New ADODB.Recordset
txtFecSalIni = Date
txtFecSalFin = Date
End Sub
Sub Rechazar(Fecsal As String, Numsol As String)

Dim cmdRechazar As Command
Set cmdRechazar = New ADODB.Command

With cmdRechazar
    .ActiveConnection = Con.Cnx
    .CommandText = " Update solicitud set indAcep = 0, FecAprob = getdate()  where fecsal = '" & Fecsal & "' and numsol = '" & Numsol & "' "
    .Execute , , adExecuteNoRecords
End With

End Sub

Sub Aprobar(Fecsal As String, Numsol As String)

Dim cmdRechazar As ADODB.Command
Set cmdRechazar = New ADODB.Command

With cmdRechazar
    .ActiveConnection = Con.Cnx
    .CommandText = " Update solicitud set indAcep = 1, FecAprob = getdate()  where fecsal = '" & Fecsal & "' and numsol = '" & Numsol & "' "
    .Execute , , adExecuteNoRecords
End With

End Sub

Sub Suspender(Fecsal As String, Numsol As String)

Dim cmdRechazar As ADODB.Command
Set cmdRechazar = New ADODB.Command

With cmdRechazar
    .ActiveConnection = Con.Cnx
    .CommandText = " Update solicitud set indAcep = isnull, FecAprob = Null, FecSusp = getdate()  where fecsal = '" & Fecsal & "' and numsol = '" & Numsol & "' "
    .Execute , , adExecuteNoRecords
End With

End Sub

Sub ProcesarSuspendidos()
        For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    'Fecha - Numsol
                    Call Suspender(ListView1.ListItems(i).SubItems(1), ListView1.ListItems(i).SubItems(3))
                End If
        Next
        MsgBox "Proceso Terminado", 64, Caption
End Sub

Sub ProcesarRechazos()
        For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    'Fecha - Numsol
                    Call Rechazar(ListView1.ListItems(i).SubItems(1), ListView1.ListItems(i).SubItems(3))
                End If
        Next
        MsgBox "Proceso Terminado", 64, Caption
End Sub


Sub ProcesarAceptador()
        For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    'Fecha - Numsol
                    Call Aprobar(ListView1.ListItems(i).SubItems(1), ListView1.ListItems(i).SubItems(3))
                End If
        Next
                
        MsgBox "Proceso Terminado", 64, Caption
End Sub

Private Sub txtFecSalFin_KeyPress(KeyAscii As Integer)
FVFecha txtFecSalFin, KeyAscii
End Sub

Private Sub txtFecSalFin_LostFocus()
ValidaFecha txtFecSalFin
End Sub

Private Sub txtFecSalIni_KeyPress(KeyAscii As Integer)
FVFecha txtFecSalIni, KeyAscii
End Sub

Private Sub txtFecSalIni_LostFocus()
ValidaFecha txtFecSalIni
End Sub

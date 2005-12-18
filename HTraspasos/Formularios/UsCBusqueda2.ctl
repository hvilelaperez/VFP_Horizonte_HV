VERSION 5.00
Begin VB.UserControl UsCBusqueda2 
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   ScaleHeight     =   2535
   ScaleWidth      =   4230
   Begin VB.ListBox lstCons 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   4215
   End
   Begin VB.TextBox txtCons 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "UsCBusqueda2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strConsulta As String
Private RsBusqueda As New ADODB.Recordset
Private campo As String
Private Cnx As New Connection
Dim strsql As String
Public Event Pulsacion(KeyAscii As Integer)
Public Event KeyUP(KeyAscii As Integer)
Private vbBusqinicial As Boolean
Private nroCarateresB As Integer
Private vbBusRecord As Boolean 'si la busqueda se hace x Recordset

Private Sub lstCons_DblClick()
   txtCons.Text = Trim(lstCons.Text)
   lstCons.Visible = False
   UserControl.Height = txtCons.Height
End Sub

Private Sub lstCons_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtCons.Text = Trim(lstCons.Text)
   lstCons.Visible = False
   UserControl.Height = txtCons.Height
   'TabSig KeyAscii
   RaiseEvent Pulsacion(KeyAscii)
End If
End Sub

'Public Event Keypress()
Private Sub lstCons_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
   txtCons.Text = lstCons.Text
End If
End Sub

Private Sub txtCons_GotFocus()
  lstCons.Visible = True
  UserControl.Height = txtCons.Height + lstCons.Height + 30
End Sub

Private Sub txtCons_KeyPress(KeyAscii As Integer)
RaiseEvent Pulsacion(KeyAscii)
End Sub

Private Sub txtCons_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then
    If lstCons.Visible Then
      lstCons.SetFocus
      If lstCons.ListCount > 0 Then
         lstCons.ListIndex = 0
         txtCons = lstCons.Text
      End If
    Else
      lstCons.Visible = True
      lstCons.SetFocus
    End If
  Exit Sub
End If

If Trim(txtCons.Text) <> "" Then
    If vbBusqinicial Then
        Call Busqueda
     Else
        If Len(txtCons.Text) >= nroCarateresB Then
           Call Busqueda
        End If
    End If
End If
RaiseEvent KeyUP(KeyCode)
End Sub

Private Sub UserControl_ExitFocus()
lstCons.Visible = False
UserControl.Height = txtCons.Height + 30
End Sub

Private Sub UserControl_GotFocus()
  lstCons.Visible = True
  UserControl.Height = txtCons.Height + lstCons.Height + 20
End Sub

Private Sub UserControl_LostFocus()
lstCons.Visible = False
UserControl.Height = txtCons.Height
End Sub

Private Sub UserControl_Resize()
txtCons.Width = Width
lstCons.Width = Width
End Sub
Public Property Let Consulta(texto As String)
     strConsulta = texto
End Property

Public Property Get Consulta() As String
    Consulta = strConsulta
End Property

Public Property Let Campocondicion(texto As String)
    campo = texto
End Property

Private Function Busqueda()

If Not vbBusRecord Then
    With RsBusqueda
     If .State = 1 Then RsBusqueda.Close
     strsql = strConsulta & " where " & campo & " like '%" & Trim(txtCons.Text) & "%' order by " & campo & ""
     RsBusqueda.Open strsql, Cnx, adOpenStatic
     
     If Not (.EOF And .BOF) Then
        Call LlenarLista(RsBusqueda)
     End If
    End With
Else
    RsBusqueda.Filter = 0
    RsBusqueda.Filter = campo & " like '%" & Trim(txtCons.Text) & "*'"
    If RsBusqueda.RecordCount > 0 Then
        Call LlenarLista(RsBusqueda)
    End If

End If

'Leo
'93075762

End Function

Private Function LlenarLista(Record As Recordset)
With lstCons
    .Clear
     Record.MoveFirst
     While Not Record.EOF
        .AddItem Record(0).Value
        Record.MoveNext
     Wend
End With
End Function

Public Property Get cTexto() As String
'Propiedad q me devuelve el contenido del texto
     cTexto = Trim(txtCons.Text)
End Property
Public Property Let cTexto(valor As String)
'Propiedad q me devuelve el contenido del texto
     txtCons.Text = valor
End Property

Public Property Set cConexion(conex As Connection)
     Set Cnx = conex
End Property

Public Property Let Enabledf(valor As Boolean)
UserControl.Enabled = valor
txtCons.Enabled = valor
lstCons.Enabled = valor
End Property

Public Property Let Busqinicial(valor As Boolean)
'indica si se realiza la busqueda despues de 3 caracteres
'o en el 1ero
 vbBusqinicial = valor
End Property
Public Property Get RsResultado() As Recordset
 Set RsResultado = RsBusqueda
End Property

Public Property Get nroCaracteres() As Integer
'Propiedad q me devuelve el contenido del texto
     nroCaracteres = nroCarateresB
End Property
Public Property Let nroCaracteres(valor As Integer)
'Propiedad q me devuelve el contenido del texto
     nroCarateresB = valor
End Property

Public Property Let BusqRecord(valor As Boolean)
'indica si la busqueda se realiza por Recordset
 vbBusRecord = valor
End Property
Public Property Set RecordBusqueda(Rsx As Recordset)
     Set RsBusqueda = Rsx
End Property

Public Property Get RecordBusqueda() As Recordset
'Propiedad q me devuelve el contenido del texto
     Set RecordBusqueda = RsBusqueda
End Property


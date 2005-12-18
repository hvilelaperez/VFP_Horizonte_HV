VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmSolicitudesD2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitudes : Segunda Digitacion"
   ClientHeight    =   6135
   ClientLeft      =   4155
   ClientTop       =   2265
   ClientWidth     =   12135
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid dtgLotes 
      Height          =   4215
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483644
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "Fecsal"
         Caption         =   "Fec. Salida"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Lote"
         Caption         =   "Lote"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "envio"
         Caption         =   "Envio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Agencia"
         Caption         =   "Agencia"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "CantSol"
         Caption         =   "Cant. Sol."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "CantImag"
         Caption         =   "Cant. Imag"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Status"
         Caption         =   "Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "indTipDoc"
         Caption         =   "Doc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "codmod"
         Caption         =   "Usuario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "CodmodD2"
         Caption         =   "Usuario D2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1110.047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtenvio 
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
      Left            =   6000
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtnrolote 
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
      Left            =   3600
      TabIndex        =   6
      Top             =   360
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
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Buscar"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDigitar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Digitar"
      Enabled         =   0   'False
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
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Envio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Lote "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblFecSalida 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmSolicitudesD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsLote As New Recordset
Dim horaini As String * 8
Dim lcCadena As String
Dim Condicion As String
Dim vbrefrecar As Boolean

Private Sub cmdBuscar_Click()
CargarDatos
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub
Sub CargarDatos()
On Error GoTo TRATA_ERROR
Dim Strsqltexto As String
'horaini = ObtenerHora


Screen.MousePointer = MousePointerConstants.vbHourglass
DoEvents

    Strsqltexto = " select a.CodAfp, a.Fecsal, a.Lote, a.Envio, a.Agencia, a.CantSol , a.CantImag, a.Codusu , Codmod , CodmodD2, " & _
                  " indTipDoc = case when indTipDoc = 1 then 'C' else 'T' end  from ctrldoc a join _scan b on a.lote = b.lotefisico "
                     
    lcCadena = Trim(Strsqltexto)
    Condicion = " Where  (b.Estado = '3' or  b.Estado = '4' )  "
    
    If Trim(txtFecSalida.Text) <> "" Then Condicion = Condicion & " and convert(char,fecsal,103) = '" & Trim(txtFecSalida.Text) & "' "
    If Trim(txtnrolote.Text) <> "" Then Condicion = Condicion & " and a.Lote = '" & Trim(txtnrolote.Text) & "' "
    If Trim(txtEnvio.Text) <> "" Then Condicion = Condicion & " and a.Envio = '" & Trim(txtEnvio.Text) & "'"
    
'    If Condicion = " Where " Then
'       Condicion = ""
'    Else
'       Condicion = Left(Condicion, Len(Condicion) - 4)
'    End If
    lcCadena = lcCadena & Condicion & " order by a.Lote"
    
    If rsLote.State = 1 Then rsLote.Close
    With rsLote
         .CursorType = adOpenStatic
         .CursorLocation = adUseClient
         .LockType = adLockReadOnly
    End With
        
    rsLote.Open lcCadena, Con.Cnx
    
    If Not (rsLote.EOF And rsLote.BOF) Then
        cmdDigitar.Enabled = True
        dtgLotes.Enabled = True
        Set dtgLotes.DataSource = rsLote
        cmdDigitar.Enabled = True
     Else
        dtgLotes.Enabled = False
        cmdDigitar.Enabled = False
        MsgBox "No se encontró información", 64, "Lotes"
    End If
     dtgLotes.Refresh
     Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
TRATA_ERROR:
    MsgBox Err.Description

End Sub

Private Sub cmdIngAfiliados_Click()
frmConsSelLote2.Show 1
End Sub

Private Sub cmdDigitar_Click()
vbrefrecar = True
vbGLte = rsLote.Fields("Lote")
IngresarDatos

End Sub

Private Sub Form_Activate()
If vbrefrescar Then
  CargarDatos
End If
vbrefrescar = False
End Sub

Private Sub Form_Load()
 vbrefrescar = True
 CenterForm Me
End Sub

Private Sub Form_Resize()
'With frmSolicitudesD2
'.Cls
'.AutoRedraw = True
'.DrawStyle = 6
'.DrawMode = 13
'.DrawWidth = 2
'.ScaleMode = 3
'.ScaleHeight = (256 * 2)
'For i = 0 To 255
'Line (0, Y)-(Width, Y + 2), RGB(5, 0, i), BF
' Y = Y + 2
'Next i
'End With
End Sub

Private Sub txtEnvio_LostFocus()
  txtEnvio.Text = Format(txtEnvio.Text, "000000#")
   
End Sub

Private Sub txtFecSalida_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
FVFecha txtFecSalida, KeyAscii
End Sub

Private Sub txtFecSalida_LostFocus()
ValidaFecha txtFecSalida
End Sub

Sub IngresarDatos()
   
'Valida que el Lote no es digitado
'Estado Pendiente
If IsNull(rsLote.Fields("codmodd2")) Then

        If Trim(UCase(rsLote.Fields("codmod"))) = Trim(UCase(gNomUsu)) Then
            MsgBox "No puede verificar un lote que usted mismo digito", vbCritical, "Verificar Lotes"
            Exit Sub
        End If

        If MsgBox("¿Desea Verificar los datos del Lote " & rsLote.Fields("Lote") & "?", vbQuestion + vbYesNo, "Verificar Lotes") = vbYes Then
        
        Set comando = New Command
        With comando
        
          .ActiveConnection = Con.Cnx
          .CommandText = "Update Ctrldoc set codmodd2 = '" & gNomUsu & "',Status = 5  where Lote = '" & rsLote.Fields("Lote") & "' "
          .Execute
          
          CambiarEstadoLote rsLote.Fields("Lote"), "4"
                                
          ' rsLote.Bookmark = marca
               
           MsgBox "El Lote " & rsLote.Fields("Lote") & " ha sido asignado a " & gNomUsu & "  ", 64, "Ingreso de Lotes"
             If rsLote.Fields("indTipDoc") = "T" Then
                frmTraspasos2Dig.IndDoc = "T"
                frmTraspasos2Dig.Show
             Else
                frmTraspasos2Dig.IndDoc = "C"
                frmTraspasos2Dig.Show
             End If
            
        End With
       End If
Else

    If Trim(rsLote.Fields("codmodd2")) = gNomUsu Then
            If rsLote.Fields("indTipDoc") = "T" Then
                frmTraspasos2Dig.IndDoc = "T"
                frmTraspasos2Dig.Show
             Else
                frmTraspasos2Dig.IndDoc = "C"
                frmTraspasos2Dig.Show
             End If
        
    Else
        MsgBox "Lote " & rsLote.Fields("Lote") & " ya está Asignado", 64, "Mensaje"
    End If
 
End If



End Sub




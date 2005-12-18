VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGenLst 
   Caption         =   "Generador LST"
   ClientHeight    =   7440
   ClientLeft      =   2310
   ClientTop       =   2130
   ClientWidth     =   10215
   Icon            =   "frmGenLST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10215
   Begin VB.TextBox txtplantillaAfi 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6480
      TabIndex        =   28
      Text            =   "AFILIACIONES TEMPORALES"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtplantillaTras 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6480
      TabIndex        =   27
      Text            =   "TRASPASOS TEMPORALES"
      Top             =   120
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Modo 1"
      TabPicture(0)   =   "frmGenLST.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Modo 2"
      TabPicture(1)   =   "frmGenLST.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frmDocumentos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Documentos Aprobados"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2535
         Left            =   5160
         TabIndex        =   22
         Top             =   480
         Width           =   4575
         Begin VB.CommandButton cmdExpAprob 
            Caption         =   "Exportar Documentos Aprobados"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1080
            TabIndex        =   25
            Top             =   960
            Width           =   2175
         End
      End
      Begin VB.Frame frmDocumentos 
         Caption         =   "Envios"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2655
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton cmdActualizarDoc 
            Caption         =   "&Actualizar"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            TabIndex        =   24
            ToolTipText     =   "Actualizar"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ListBox lstEnvios 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2040
            Left            =   1080
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fecha Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2535
         Left            =   -68160
         TabIndex        =   16
         Top             =   360
         Width           =   2895
         Begin VB.TextBox fechaHas 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1080
            TabIndex        =   18
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox fechaDes 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1080
            TabIndex        =   17
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lotes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2535
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   6495
         Begin VB.ListBox LstLotes 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1770
            Left            =   3840
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "&Actualizar"
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
            Left            =   4080
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   11
            ToolTipText     =   "Actualizar"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txthas 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   10
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtdes 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   9
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Rango:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.ComboBox cboTipo 
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
      ItemData        =   "frmGenLST.frx":047A
      Left            =   2520
      List            =   "frmGenLST.frx":0484
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtmsj 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4800
      Width           =   9975
   End
   Begin VB.CommandButton cmdSalir 
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
      Left            =   8640
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
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
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Generar LST"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Plantilla Afiliación:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Plantilla Traspaso:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo de Documento:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblproceso 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   9975
   End
End
Attribute VB_Name = "frmGenLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RutaIni As String
Private Cnx As Connection
Private intFileHnd As Integer

Private Sub Form_Load()

   ' gUsuario = "polydoc"
    'gPassword = "master"
    'gServidor = "dg-73"
    'gBaseDatos = "TraspasosProfuturo"
    'Call Conectar
   
  
    Set Cnx = Con.Cnx
    cboTipo.ListIndex = 0
    
    Call cmdActualizar_Click
    Call cmdActualizarDoc_Click
    
    txtmsj.Text = ""
    RutaIni = App.Path & "\Config.ini"
   
End Sub

Private Sub cboTipo_Click()
 Call cmdActualizar_Click
End Sub

Public Sub CargarListaSql(lst As ListBox, sql As String)
On Error GoTo TRATA_ERROR
Dim wrs As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
Set wrs = New ADODB.Recordset
With cmd
      .ActiveConnection = Cnx
      .CommandText = sql
End With
Set wrs = cmd.Execute
Set cmd.ActiveConnection = Nothing
Do Until wrs.EOF
   DoEvents

   lst.AddItem wrs.Fields(0)
   wrs.MoveNext
Loop
wrs.Close
Exit Sub

TRATA_ERROR:
If wrs.State = 1 Then
   wrs.Close
End If
   MsgBox Err.Description
End Sub


Private Sub cmdActualizar_Click()
    'Actualiza lista de lotes
    LstLotes.Clear
    
    If (cboTipo.ListIndex <> -1) Then
        CargarListaSql LstLotes, "select distinct lote from ctrldoc " & _
        "where indTipDoc='" & cboTipo.ListIndex & "'"
    End If
End Sub

Private Sub cmdActualizarDoc_Click()
'Actualiza lista de documentos
    lstEnvios.Clear
    
    If (cboTipo.ListIndex <> -1) Then
        CargarListaSql lstEnvios, "select distinct envio from ctrldoc order by envio"
    End If
End Sub

Private Sub cmdExpAprob_Click()
    Call lstVolumen("", "", True, "")
End Sub

'Generar LST
Private Sub cmdGenerar_Click()
    If validarLotes And txtplantillaAfi.Text <> "" And txtplantillaTras.Text <> "" Then
      Call lstVolumen("", "", False, "")
    Else
      MsgBox "Debe seleccionar al menos un lote e ingresar una plantilla", vbOKOnly, "Error"
    End If
End Sub

Function validarLotes() As Boolean
    validarLotes = True
    If LstLotes.SelCount = 0 Then
        If txtdes = "" And txthas = "" Then
            If lstEnvios.SelCount = 0 Then
                validarLotes = False
            End If
        End If
    End If
End Function

' --------------------------------1ER NIVEL ---------------------------------------
Sub lstVolumen(list As String, listDocumento As String, val As Boolean, strRuta As String)
    On Error GoTo errores
    
Dim strFile As String
Dim Volumenant As Integer
Dim RsLotesRg As ADODB.Recordset
Set RsLotesRg = New ADODB.Recordset
Dim strtexto As String
Dim rsVolumen As ADODB.Recordset
Set rsVolumen = New ADODB.Recordset
Dim ListaLotes As String
Dim vpCaja As String
Dim vpBco As String
Dim Conta As Integer
Conta = 0
Dim directorio As String
Dim listaDoc As String
listaDoc = ""
Dim CmdAcc As ADODB.Command
Dim listaT As String
Dim listaA As String
listaT = ""
listaA = ""

       intRespuesta = MsgBox(" ¿Desea Iniciar el proceso de Generación ?", vbQuestion + vbYesNo, Caption)
       
       If strRuta = "" Then
            directorio = BrowseFolder("Escoja una carpeta para ubicación de imágenes " & cboTipo.Text)
        Else
           directorio = strRuta
       End If
          
      If directorio = "" Then
        MsgBox "debe seleccionar un directorio"
        Exit Sub
      End If
      
       'selecciona Si
       txtmsj.Text = ""
       
       'selecciona NO
      If (intRespuesta = vbNo) Then
            Exit Sub
      End If
      
      Set CmdAcc = New ADODB.Command

     'With CmdAcc
     '   .ActiveConnection = Con.Cnx
     '   .CommandText = "update solicitud set indExpo1=0"
     '   .Execute
     'End With
      
         
       ListaLotes = ""
       
If (val = True) Then
     
     cboTipo.ListIndex = 0
         strtexto = "select * from ctrldoc c join solicitud s " & _
                    "on c.lote=s.lote where indCruce=1 and indExpo1=1 and indExpo2=0 and indTipDoc='" & cboTipo.ListIndex & "'"
        
         RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
         
         If RsLotesRg.PageCount = 0 Then
            EscribeMensaje ("Error, no hay documentos cruzados para exportar en Traspasos")
            
         Else
         
         Do While Not RsLotesRg.EOF
            listaT = listaT & "'" & Trim(RsLotesRg.Fields("lote")) & "',"
            listaDoc = listaDoc & "'" & Trim(RsLotesRg.Fields("numsol")) & "',"
            RsLotesRg.MoveNext
         Loop
         
         listaDoc = Left(listaDoc, Len(Trim(listaDoc)) - 1)
         Call lstVolumen(listaT, listaDoc, False, directorio)
         
         End If
         RsLotesRg.Close
         
          cboTipo.ListIndex = 1
         strtexto = "select * from ctrldoc c join solicitud s " & _
                    "on c.lote=s.lote where indCruce=1 and indExpo1=1 and indExpo2=0 and indTipDoc='" & cboTipo.ListIndex & "'"
        
         RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
         
         If RsLotesRg.PageCount = 0 Then
            EscribeMensaje ("Error, no hay documentos cruzados para exportar en Afiliaciones")
            
         Else
         
         Do While Not RsLotesRg.EOF
            listaA = listaA & "'" & Trim(RsLotesRg.Fields("lote")) & "',"
            listaDoc = listaDoc & "'" & Trim(RsLotesRg.Fields("numsol")) & "',"
            RsLotesRg.MoveNext
         Loop
         'RsLotesRg.Close
         listaDoc = Left(listaDoc, Len(Trim(listaDoc)) - 1)
         Call lstVolumen(listaA, listaDoc, False, directorio)
         
         End If
         RsLotesRg.Close
         
         Exit Sub
ElseIf (list <> "") Then
    ListaLotes = list
    listaDoc = listDocumento
    If Len(listaDoc) > 0 Then val = True
    
Else
       
    If LstLotes.SelCount <> 0 Then
        'Genera la lista de Lotes seleccionados
        For i = 0 To LstLotes.ListCount - 1
            If LstLotes.Selected(i) Then
              ListaLotes = ListaLotes & "'" & Trim(Left(LstLotes.list(i), 10)) & "',"
            End If
        Next
        
    ElseIf (txtdes.Text <> "" And txthas.Text <> "") Then
       
            'verifica que los campos hayan sido llenados correctamente
            If txtdes.Text > txthas.Text Then
                 MsgBox "El valor inicial debe ser menor o igual al valor final", vbOKOnly, "Error"
                Exit Sub
            End If
       
        'genera la lista desde el menor hasta el mayor
         strtexto = "select distinct lote from ctrldoc " & _
        "where indTipDoc='" & cboTipo.ListIndex & "'" & _
         " and lote>='" & Trim(txtdes.Text) & "' and lote<='" & Trim(txthas.Text) & "'"
         
         RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
         
         If RsLotesRg.PageCount = 0 Then
            MsgBox "Error, lotes invalidos"
            Exit Sub
         End If
         
         Do While Not RsLotesRg.EOF
            ListaLotes = ListaLotes & "'" & RsLotesRg.Fields("lote") & "',"
            RsLotesRg.MoveNext
         Loop
         RsLotesRg.Close
         
    ElseIf (fechaDes.Text <> "" And fechaHas.Text <> "") Then
        
        'genera la lista por la fecha de salida
         strtexto = "select distinct lote from ctrldoc " & _
        "where indTipDoc='" & cboTipo.ListIndex & "'" & _
        " and fecsal between '" & fechaDes.Text & "' and '" & fechaHas.Text & "' order by lote"
         
         RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
         
         If RsLotesRg.PageCount = 0 Then
            MsgBox "Error, lotes invalidos"
            Exit Sub
         End If
         
         Do While Not RsLotesRg.EOF
            ListaLotes = ListaLotes & "'" & RsLotesRg.Fields("lote") & "',"
            RsLotesRg.MoveNext
         Loop
         RsLotesRg.Close
         
    ElseIf lstEnvios.SelCount <> 0 Then
         ListaLotes = ""
          For i = 0 To lstEnvios.ListCount - 1
            If lstEnvios.Selected(i) Then
              lista = lista & "'" & Trim(Left(lstEnvios.list(i), 10)) & "',"
            End If
          Next
        
          lista = Left(lista, Len(Trim(lista)) - 1)
          cboTipo.ListIndex = 0
          strtexto = "select distinct lote from ctrldoc " & _
        "where indTipDoc='" & cboTipo.ListIndex & "'" & _
         " and envio in (" & lista & ")"
         
         RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
         
         If RsLotesRg.PageCount = 0 Then
            MsgBox "Error, el envio no tiene lotes"
            'Exit Sub
         
         Else
         ListaLotes = ""
         Do While Not RsLotesRg.EOF
            listaT = listaT & "'" & RsLotesRg.Fields("lote") & "',"
            RsLotesRg.MoveNext
         Loop
         
        Call lstVolumen(listaT, "", False, directorio)
        End If
         RsLotesRg.Close
         cboTipo.ListIndex = 1
          strtexto = "select distinct lote from ctrldoc " & _
        "where indTipDoc='" & cboTipo.ListIndex & "'" & _
         " and envio in (" & lista & ")"
         
         RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
         
         If RsLotesRg.PageCount = 0 Then
            MsgBox "Error, el envio no tiene lotes"
            Exit Sub
         End If
         
         ListaLotes = ""
         Do While Not RsLotesRg.EOF
            listaA = listaA & "'" & RsLotesRg.Fields("lote") & "',"
            RsLotesRg.MoveNext
         Loop
         RsLotesRg.Close
         Call lstVolumen(listaA, "", False, directorio)
        Exit Sub
    End If
End If
       
       'Quita la ultima coma
       ListaLotes = Left(ListaLotes, Len(Trim(ListaLotes)) - 1)
  
       strtexto = "select distinct v.idvolumen,v.nombre from _scan s " & _
       "join _volumenes v on s.idvolumen=v.idvolumen " & _
       "where s.lotefisico in (" & ListaLotes & ")"
  
       RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
           
       If RsLotesRg.RecordCount = 0 Then MsgBox "No existe información para el intervalo", 64, Caption: Exit Sub
  
       Do While Not RsLotesRg.EOF
     
            DoEvents
                If (cboTipo.ListIndex = 0) Then
                      strFile = directorio & "\TRASPASOS_PROFUTURO.lst"
                ElseIf (cboTipo.ListIndex = 1) Then
                     strFile = directorio & "\AFILIACIONES_PROFUTURO.lst"
                Else
                    strFile = directorio & "\_PROFUTURO.lst"
                End If

                             
                If Conta = 0 Then
                EscribeMensaje "Iniciando Generación de Lst : " & Now
                  
                 If Dir(strFile, vbArchive) <> "" Then
                    Kill strFile
                 End If
                 intFileHnd = OpenFile(strFile, "W")
                 Print #intFileHnd, "LASERFICHE IMPORT LIST"
                          
                     Print #intFileHnd, "FOLDERS (VOLUMENES)"
                      If (cboTipo.ListIndex = 0) Then
                        Print #intFileHnd, "FOLDERS (SOLICITUD DE TRASPASO)"
                    Else
                        Print #intFileHnd, "FOLDERS (CONTRATOS DE AFILIACION)"
                     End If

                     
                 End If
                    Print #intFileHnd, "FOLDERS (" & RsLotesRg.Fields("nombre") & ")"
                    
                    If val Then
                         Call LstFecha(RsLotesRg.Fields("idvolumen"), ListaLotes, Trim(listaDoc), "1")
                    Else
                      Call LstFecha(RsLotesRg.Fields("idvolumen"), ListaLotes, "", "0")
                    End If
                     
                     Print #intFileHnd, "FOLDERS ()"
                     Conta = Conta + 1
            RsLotesRg.MoveNext
            
       Loop
       
       EscribeMensaje "Se genero el archivo " & UCase(strFile) & ""
       
       Print #intFileHnd, "FOLDERS ()"
       Print #intFileHnd, "FOLDERS ()"
       Close intFileHnd
       
 RsLotesRg.Close
    Set RsLotesRg = Nothing
    
    Set CmdAcc = New ADODB.Command

    If val Then
    
    With CmdAcc
   
        .ActiveConnection = Con.Cnx
        .CommandText = "update solicitud set indExpo2=1 where indExpo1=1 and indCruce=1"
        .Execute
   End With
    End If
    lblproceso.Caption = "Proceso Concluido"
    
    MsgBox "Proceso concluido", 64, Caption
    
    Exit Sub
    
errores:
    MsgBox Err.Description, vbCritical, Caption
    Close intFileHnd
End Sub
'

' --------------------------------2DO NIVEL ---------------------------------------

Sub LstFecha(ByVal Volumen As String, ByVal ListaLotes As String, _
                                        Optional ByVal listaDoc As String, Optional ByVal exc As String)
  On Error GoTo errores

Dim RsLotesRg As ADODB.Recordset
Set RsLotesRg = New ADODB.Recordset
Dim RsDocRg As ADODB.Recordset
Dim strtexto As String
Dim vpCaja As String
Dim vpBco As String
                  
       strtexto = "select distinct fecsal from solicitud s " & _
       "join _scan sc on s.lote = sc.lotefisico " & _
       "join _volumenes v on sc.idvolumen=v.idvolumen " & _
       "where sc.lotefisico in (" + ListaLotes + ") And v.idvolumen = '" & Volumen & "'"
         
       RsLotesRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
        
        If RsLotesRg.RecordCount = 0 Then Exit Sub
        
       Do While Not RsLotesRg.EOF
     
            DoEvents
                      
                Print #intFileHnd, "FOLDERS (" & RsLotesRg.Fields("fecsal") & ")"
                
                If Len(listaDoc) = 0 Then
                    strtexto = "select distinct numsol,lotefisico from solicitud s " & _
                    "join _scan sc on s.lote=sc.lotefisico" & _
                    " where sc.lotefisico in (" & ListaLotes & _
                    ") and s.fecsal='" & RsLotesRg.Fields("fecsal") & "'"
                
                Else
                    strtexto = "select distinct numsol,lotefisico from solicitud s " & _
                    "join _scan sc on s.lote=sc.lotefisico" & _
                    " where lotefisico in (" & ListaLotes & _
                    ") and s.fecsal='" & RsLotesRg.Fields("fecsal") & "' " & _
                    "and s.numsol in (" & listaDoc & ")"
                End If
                
                
                Set RsDocRg = New ADODB.Recordset
                RsDocRg.Open strtexto, Cnx, adOpenStatic, adLockReadOnly
                
                Do While Not RsDocRg.EOF
                DoEvents
                
                   lblproceso.Caption = "Solicitud en proceso: " + RsDocRg.Fields("numsol")
                    Call LstDocumento(RsDocRg.Fields("lotefisico"), RsDocRg.Fields("numsol"), exc)
                RsDocRg.MoveNext
            
                Loop
                
                 RsDocRg.Close
                Set RsDocRg = Nothing
                Print #intFileHnd, "FOLDERS ()"
                     
            RsLotesRg.MoveNext
            
       Loop

RsLotesRg.Close
    Set RsLotesRg = Nothing

 Exit Sub
    
errores:
    MsgBox Err.Description, vbCritical, Caption
    Close intFileHnd
End Sub

' --------------------------------3ER NIVEL ---------------------------------------

Sub LstDocumento(lotefisico As String, idDoc As String, exc As String)
  On Error GoTo errores

Dim DocData As Recordset
Set DocData = New Recordset

Dim DocImage As Recordset
Set DocImage = New Recordset

Dim CmdAcc As ADODB.Command

Strsqltexto = "select nro_cliente,numsol,s.fecsal,isnull(rtrim(ltrim(apepat)),'')+' '+ " & _
"isnull(rtrim(ltrim(apemat)),'')+' '+isnull(rtrim(ltrim(nombre1)),'')+' '+ " & _
"isnull(rtrim(ltrim(nombre2)),'') 'nombre', s.cuspp " & _
"from solicitud s " & _
"join _scan sc on s.lote=sc.lotefisico where secu=0 " & _
"and rtrim(ltrim(numsol))='" & Trim(idDoc) & "' " & _
"and sc.lotefisico='" & Trim(lotefisico) & "'"

DocData.Open Strsqltexto, Cnx, adOpenStatic, adLockReadOnly

    Print #intFileHnd, _
IIf(cboTipo.ListIndex = 0, _
"DOCUMENT(CAT02-" & Trim(DocData.Fields("numsol")) & ")", _
"DOCUMENT(CAT01-" & IIf(exc = 1, _
    Trim(DocData.Fields("cuspp")), _
    Mid(Trim(DocData.Fields("numsol")), 2, Len(Trim(DocData.Fields("numsol")))) - 1) & _
    ")")
    Print #intFileHnd, "STARTFIELDS"
    Print #intFileHnd, "" & _
                IIf(cboTipo.ListIndex = 0, UCase(Trim(txtplantillaTras.Text)), _
                                            UCase(Trim(txtplantillaAfi.Text)))
    Print #intFileHnd, "" & (IIf(cboTipo.ListIndex = 0, "Traspaso de entrada", "Contrato de afiliación"))
    Print #intFileHnd, ""
    Print #intFileHnd, "" & Trim(IIf(IsNull(DocData.Fields("cuspp")), "", DocData.Fields("cuspp")))
    Print #intFileHnd, "" & Trim(DocData.Fields("nombre"))
    Print #intFileHnd, "" & Trim(IIf(exc = 1, DocData.Fields("nro_cliente"), DocData.Fields("numsol")))
    Print #intFileHnd, "" & Trim(DocData.Fields("fecsal"))
    Print #intFileHnd, "ENDFIELDS"
                                        
    Print #intFileHnd, "STARTLIST"
    
With DocImage
    
    Strsqltexto = _
    "Select ruta+'\'+ i.lotefisico + '\' + imagen + '.tif' 'ruta' from solicitud s left join _imagenes i " & _
    "on '01'+s.numsol=i.idimagen and i.lotefisico=s.lote " & _
    " join _volumenes v " & _
    "on i.idvolumen = v.idvolumen where s.lote='" & Trim(lotefisico) & "'" & _
    "and s.numsol='" & idDoc & "' and s.secu = 0 and tipimag <> '" & IIf(exc = 1, "S", "A") & "' ORDER BY i.posicion "
    
  .Open Strsqltexto, Cnx, adOpenStatic, adLockReadOnly
                                        
    If .RecordCount > 0 Then
        Do While Not .EOF
            Print #intFileHnd, "" & Trim(.Fields("Ruta"))
            .MoveNext
        Loop
    Else
        EscribeMensaje "No se encontraron imágenes para el Documento " & UCase(Trim(DocData.Fields("numsol")))
    End If
        
    DocImage.Close
    Set DocImage = Nothing
End With
     Print #intFileHnd, "ENDLIST"
     
     Set CmdAcc = New ADODB.Command
     
     Dim cadena As String
     cadena = "update solicitud set indExpo1=1 " & _
        "from solicitud s left join _imagenes i " & _
        "on '01'+s.numsol=i.idimagen join _volumenes v " & _
        "on i.idvolumen = v.idvolumen where lotefisico='" & Trim(lotefisico) & "'" & _
        "and s.numsol='" & Trim(idDoc) & "'"

    With CmdAcc
        .ActiveConnection = Con.Cnx
        .CommandText = cadena
        .Execute
   End With

DocData.Close
Set DocData = Nothing
 Exit Sub
    
errores:
    MsgBox Err.Description, vbCritical, Caption

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Sub EscribeMensaje(texto As String)
If txtmsj.Text = "" Then
    txtmsj.Text = texto
Else
    txtmsj.Text = txtmsj.Text + vbCrLf + texto
End If
End Sub

Public Function OpenFile(ByVal sFile As String, Modo As String) As Integer
    On Error GoTo EHOPENFILE
    Dim nFileNumber As Integer
    nFileNumber = FreeFile
    If Modo = "L" Then
        Open sFile For Input As #nFileNumber
    Else
        Open sFile For Append As #nFileNumber
    End If
    OpenFile = nFileNumber
    Exit Function
EHOPENFILE:
    MsgBox Err.Description & "(" & Err.Number & ")."
    
    OpenFile = 0
End Function

Private Sub fechaHas_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
    FVFecha fechaHas, KeyAscii
End Sub

Private Sub fechaHas_LostFocus()
    If (IsDate(fechaHas.Text) = False) Then
        MsgBox "La fecha ingresada es invalida", vbOKOnly, "Error"
        fechaHas.SetFocus
    End If
End Sub

Private Sub fechaDes_KeyPress(KeyAscii As Integer)
TabSig KeyAscii
    FVFecha fechaDes, KeyAscii
End Sub

Private Sub fechaDes_LostFocus()
    If (IsDate(fechaDes.Text) = False) Then
        MsgBox "La fecha ingresada es invalida", vbOKOnly, "Error"
        fechaDes.SetFocus
    End If
End Sub

'Private Sub txtFechaDoc_KeyPress(KeyAscii As Integer)
'TabSig KeyAscii
'FVFecha txtFechaDoc, KeyAscii
'End Sub

Private Sub txtFechaDoc_LostFocus()
    If (IsDate(txtFechaDoc.Text) = False) Then
        MsgBox "La fecha ingresada es invalida", vbOKOnly, "Error"
        txtFechaDoc.SetFocus
    End If
End Sub


Private Sub txtdes_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 44 Or KeyAscii > 44)) Then
         If (KeyAscii <> 8) Then KeyAscii = 0
     End If
End Sub

Private Sub txthas_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 44 Or KeyAscii > 44)) Then
         If (KeyAscii <> 8) Then KeyAscii = 0
     End If
End Sub

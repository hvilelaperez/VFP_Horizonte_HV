VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmOpciones 
   Caption         =   "Opciones"
   ClientHeight    =   6540
   ClientLeft      =   870
   ClientTop       =   1680
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   13845
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   11640
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtloteu1 
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtloteu2 
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   5760
      TabIndex        =   8
      Top             =   840
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Eliminación  de Lotes Digitación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   5175
      Begin VB.CommandButton cmdEliminar2 
         Caption         =   "&Eliminar"
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
         Left            =   2760
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtloted2 
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtloted1 
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Eliminación  de Lotes Digitalización"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      Begin VB.CommandButton cmdEliminar1 
         Caption         =   "Eliminar"
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
         Left            =   2760
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtLote2 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtLote1 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsimagenes As New Recordset

Private Sub cmdActualizar_Click()
On Error Resume Next
Set comando = New Command
With comando

    .ActiveConnection = Con.Cnx
    .CommandText = "Update _imagenes set lote = lotefisico where (lote >= " & txtloteu1 & "  and  lote <= " & txtloteu2 & ") and  lote <> lotefisico  "
    .Execute

End With
End Sub

Private Sub CmdBuscar_Click()
If Trim(txtloteu1.Text) = "" Or txtloteu2.Text = "" Then MsgBox "Debe ingresar lote de busqueda", vbCritical, Caption: Exit Sub

With rsimagenes
  If .State = 1 Then .Close
    .Open "select * from _imagenes where (lote >= " & txtloteu1 & "  and  lote <= " & txtloteu2 & ") and  lote <> lotefisico  ", Con.Cnx, adOpenStatic, adLockReadOnly
  
If .EOF And .BOF Then
    MsgBox "No se encontraron registros", 64, Caption
End If

  Set DataGrid1.DataSource = rsimagenes

End With

End Sub

Private Sub cmdEliminar1_Click()
On Error GoTo errores
Dim Vol As Integer

If Trim(TxtLote1.Text) = "" Or TxtLote2.Text = "" Then MsgBox "Debe ingresar lote para eliminacion", vbCritical, Caption: Exit Sub

Vol = Val(fgDatosql("select idvolumen  from _Imagenes where  lote =" & TxtLote1 & ""))
If Vol = 0 Then MsgBox "Lote no " & Me.TxtLote1.Text & " no existe ", 64, Caption: Exit Sub
Con.Cnx.BeginTrans

Set comando = New Command
With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete _Imagenes where lote >= " & TxtLote1 & "  and  lote <= " & TxtLote2 & ""
    .Execute
End With

With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete _PreScan where lote >= " & TxtLote1 & "  and  lote <= " & TxtLote2 & ""
    .Execute
End With

With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete _Scan where lote >= " & TxtLote1 & "  and  lote <= " & TxtLote2 & ""
    .Execute
End With


'----------Digitacion----------


With comando
    .ActiveConnection = Con.Cnx
    .CommandText = " delete  empleadores where numsol + convert(char,secuencial) " & _
                   " in (select numsol + convert(char,secu) from  solicitud where lote >= " & TxtLote1 & "  and  lote <= " & TxtLote2 & "  ) "

    .Execute
End With

With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete solicitud where lote >= " & TxtLote1 & "  and  lote <= " & TxtLote2 & ""
    .Execute
End With


With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete ctrldoc where lote >= " & TxtLote1 & "  and  lote <= " & TxtLote2 & ""
    .Execute
End With



With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete ctrldoc where lote >= " & TxtLote1 & "  and  lote <= " & TxtLote2 & ""
    .Execute
End With


With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "exec sp_ActualizarVolumen ?"
    .Parameters(0) = Vol
    .Execute
End With
 


Con.Cnx.CommitTrans
MsgBox "Eliminacion Satisfactoria Digitalizacion", 64, Caption
Exit Sub
errores:
Con.Cnx.RollbackTrans
MsgBox Err.Description, vbCritical, Caption

End Sub

Private Sub cmdEliminar2_Click()

'----------Digitacion----------
On Error GoTo errores


If Trim(txtloted1.Text) = "" Or txtloted2.Text = "" Then MsgBox "Debe ingresar lote para eliminacion", vbCritical, Caption: Exit Sub

Con.Cnx.BeginTrans
Set comando = New Command


With comando
    .ActiveConnection = Con.Cnx
    .CommandText = " delete  empleadores where numsol + convert(char,secuencial) " & _
                   " in (select numsol + convert(char,secu) from  solicitud where lote >= " & txtloted1 & "  and  lote <= " & txtloted2 & "  ) "

    .Execute
End With

With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete solicitud where lote >= " & txtloted1 & "  and  lote <= " & txtloted2 & ""
    .Execute
End With


With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "delete ctrldoc where lote >= " & txtloted1 & "  and  lote <= " & txtloted2 & ""
    .Execute
End With


With comando
    .ActiveConnection = Con.Cnx
    .CommandText = "update _imagenes set status = 1 where lote >= " & txtloted1 & "  and  lote <= " & txtloted2 & ""
    .Execute
End With


Con.Cnx.CommitTrans
MsgBox "Eliminacion Satisfactoria Digitación", 64, Caption
Exit Sub
errores:
Con.Cnx.RollbackTrans
MsgBox Err.Description, vbCritical, Caption

End Sub

Function validarLotes() As Boolean
    validarLotes = True
    If Trim(txtloteu1.Text) = "" Or Trim(txtloteu2.Text) = "" Then
        validarLotes = False
        MsgBox "Debe ingresar los lotes", vbCritical, Caption
        txtloted1.SetFocus
        Exit Function
    End If
End Function



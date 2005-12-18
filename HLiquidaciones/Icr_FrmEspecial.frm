VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Icr_FrmEspecial 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ingreso de Datos Especiales"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "Icr_FrmEspecial.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7065
   ScaleWidth      =   10425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.CommandButton cmdagregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3720
         Width           =   9735
      End
      Begin VB.Frame Fracantidad 
         BackColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   5040
         TabIndex        =   24
         Top             =   240
         Width           =   3255
         Begin VB.TextBox TXTCANTIDAD 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   960
            MaxLength       =   2
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cantidad"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Salir"
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton Cmdgrabar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Frame Fradetalle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos Detalle"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   9735
         Begin VB.TextBox txtnro 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   120
            MaxLength       =   2
            TabIndex        =   7
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtpla 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   8040
            TabIndex        =   13
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtremu 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   6600
            TabIndex        =   12
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtfechafin 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   5160
            MaxLength       =   6
            TabIndex        =   11
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtfechaini 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   3720
            MaxLength       =   6
            TabIndex        =   10
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtcodigo 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   9
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TXTcuspp 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   8
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cuspp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   615
         End
         Begin VB.Line Line7 
            X1              =   1080
            X2              =   1080
            Y1              =   120
            Y2              =   2640
         End
         Begin VB.Line Line6 
            X1              =   7920
            X2              =   7920
            Y1              =   120
            Y2              =   2640
         End
         Begin VB.Line Line5 
            X1              =   6480
            X2              =   6480
            Y1              =   120
            Y2              =   2640
         End
         Begin VB.Line Line4 
            X1              =   5040
            X2              =   5040
            Y1              =   120
            Y2              =   2640
         End
         Begin VB.Line Line3 
            X1              =   3600
            X2              =   3600
            Y1              =   120
            Y2              =   2640
         End
         Begin VB.Line Line2 
            X1              =   2760
            X2              =   2760
            Y1              =   120
            Y2              =   2640
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   0
            X2              =   9720
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nro Planilla y Banco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7920
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Remuneracion Asegurable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6600
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Cese (dd-mm-aa)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5160
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio (dd-mm-aa)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3720
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Codigo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cuspp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Fracabecera 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Datos Cabecera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         Begin VB.TextBox txtpagina 
            Height          =   375
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   5
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox TXTliquidacion 
            Height          =   375
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pagina"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nro Liquidacion"
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid Griding 
         Height          =   1935
         Left            =   240
         TabIndex        =   26
         Top             =   4200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "ds_reg"
            Caption         =   "Nro"
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
            DataField       =   "ds_cuspp"
            Caption         =   "Cuspp"
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
            DataField       =   "ds_codigo"
            Caption         =   "Codigo"
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
            DataField       =   "ds_fechaini"
            Caption         =   "Fecha Ini"
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
            DataField       =   "ds_fechafin"
            Caption         =   "Fecha Cese"
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
            DataField       =   "ds_remu"
            Caption         =   "Remuneracion"
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
            DataField       =   "ds_pla"
            Caption         =   "Planilla"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Icr_FrmEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdagregar_Click()
 Dim IntReg As Integer
 Dim StrPag As String
 Dim StrCupss As String
 IntReg = Val(TXTCANTIDAD.Text)
    
 If Len(txtnro.Text) > 0 Then
   If Len(Txtcuspp.Text) <> 0 Then
       StrPag = "PAG.:" & Format(txtpagina, "000")
       StrCupss = "CUSPP" & txtnro
       If CupssProcesada(Txtliquidacion, StrPag, StrCupss) Then
       
                If rsSQL.State = 1 Then rsSQL.Close
                rsSQL.Source = "Select * from tb_tempo"
                rsSQL.Open
                If rsSQL.RecordCount > 0 Then
                 Set Griding.DataSource = rsSQL
                End If
                        
                If rsSQL.RecordCount + 1 <= IntReg Then
                        If Len(Txtcuspp.Text) <> 0 Then
                            If Len(txtnro.Text) <> 0 Or Len(txtcodigo.Text) <> 0 Or Len(txtfechaini.Text) <> 0 Or Len(txtfechafin.Text) <> 0 Or Len(txtremu.Text) <> 0 Or Len(txtpla.Text) <> 0 Then
                              
                              strCadenaSQL = "EXEC sp_InsertarTempo '" & ("" & txtnro.Text) & "','" & ("" & UCase(Txtcuspp.Text)) & "','" & ("" & txtcodigo.Text) & "','" & ("" & txtfechaini.Text) & "','" & txtfechafin.Text & "','" & txtremu.Text & "','" & txtpla.Text & "'"
                              cnServer.Execute strCadenaSQL
                              bolError = True
                            Else
                            MsgBox "Se Debe ingresar los datos deL Cuspp ", vbInformation, "Liquidacion Previa"
                            End If
                        Else
                          MsgBox "Debe ingresar el Cuspp ", vbInformation, "Liquidacion Previa"
                        End If
                            If rsSQL.State = 1 Then rsSQL.Close
                            rsSQL.Source = "Select * from tb_tempo"
                            rsSQL.Open
                            If rsSQL.RecordCount > 0 Then
                               Set Griding.DataSource = rsSQL
                               Cmdgrabar.Enabled = True
                               
                            Else
                               Set Griding.DataSource = Nothing
                            End If
                 Else
                 MsgBox "El limite de registros es de: " & TXTCANTIDAD, vbInformation, "Liquidacion Previa"
                 End If
       Else
       MsgBox "El numero de registro ya tiene un Cuspp con dato ", vbInformation, "Liquidacion Previa"
       End If
     Else
        MsgBox "Se Debe ingresar los datos deL Cuspp ", vbInformation, "Liquidacion Previa"
     End If
Else
     MsgBox "Se debe Ingresar el numero de registros del cupss", vbInformation, "Liquidacion Previa"
End If
End Sub

Private Sub Cmdgrabar_Click()
Dim Strliqui As String
Dim IntReg As Integer
Dim IntReg1 As Integer
Dim Xnro As String
Dim Xcussp As String
Dim Xcodigo As String
Dim XFechIni As String
Dim XFechCe As String
Dim XRemu As String
Dim XPla As String

Dim ds_cod As String
Dim ds_fec_ini As String
Dim ds_fec_ce As String
Dim ds_rem_as As String
Dim ds_pla As String


Dim cuspp As String
Strliqui = Txtliquidacion & Format(txtpagina, "000")
If MsgBox("Esta seguro de Insertar estos registros a la base de datos", vbYesNo) = vbYes Then

   'If Griding.Row.Count > 0 Then
                If LiquidacionFormulario(Strliqui) Then
                  If LiquidacionProcesada(Trim$(Txtliquidacion), "PAG.:" & Format(txtpagina, "000")) Then
                              If Griding.ApproxCount = TXTCANTIDAD.Text Then
                                  If Griding.ApproxCount - 1 >= 0 Then
                                   For IntReg = 0 To Griding.ApproxCount - 1
                                               Xnro = Griding.Columns(0).Text
                                               Xcussp = Griding.Columns(1).Text
                                               Xcodigo = Griding.Columns(2).Text
                                               XFechIni = Griding.Columns(3).Text
                                               XFechCe = Griding.Columns(4).Text
                                               XRemu = Griding.Columns(5).Text
                                               XPla = Griding.Columns(6).Text
                                               
                                               cuspp = "CUSPP" & Val(Xnro)
                                               ds_cod = "ds_cod" & Val(Xnro)
                                               ds_fec_ini = "ds_fec_ini" & Val(Xnro)
                                               ds_fec_ce = "ds_fec_ce" & Val(Xnro)
                                               ds_rem_as = "ds_rem_as" & Val(Xnro)
                                               ds_pla = "ds_pla" & Val(Xnro)
                                               
                                               
                                              'If rsSQL.State = 1 Then rsSQL.Close
                                              strEjecutar = "Update tb_formularios set " & cuspp & "='" & Xcussp & "' where nroliqpag = '" & Strliqui & "'"
                                              cnServer.Execute strEjecutar
                                               
                                              'If rsSQL.State = 1 Then rsSQL.Close
                                              strEjecutar = "Update tb_liquidacion set " & ds_cod & "='" & Xcodigo & "'," & ds_fec_ini & "='" & XFechIni & "'," & ds_fec_ce & "='" & XFechCe & "'," & ds_rem_as & "='" & XRemu & "'," & ds_pla & "='" & XPla & "' where ds_nroliq= '" & Trim$(Txtliquidacion) & "' and ds_pag ='PAG.:" & Format(txtpagina, "000") & "'"
                                              cnServer.Execute strEjecutar
                                          If Griding.Row < Griding.ApproxCount - 1 Then
                                          Griding.Row = Griding.Row + 1
                                          Else
                                          End If
                                              
                                              
                                   Next
                                   
                                   MsgBox "Se Agregaron Los Registros Satisfactoriamente ", vbInformation, "Liquidacion Previa"
                                   Limpiar
                                  End If
                                Else
                                   MsgBox "El nro de Registros no coincide con la cantidad ingresada ", vbInformation, "Liquidacion Previa"
                                End If
                  Else
                  MsgBox "El nro de Liquidacion No esta Procesada ", vbInformation, "Liquidacion Previa"
                  End If
                Else
                MsgBox "El nro de Liquidacion No existe ", vbInformation, "Liquidacion Previa"
                End If
                
    'Else
    '  MsgBox "No existen Registros Agregados", vbInformation, "Liquidacion Previa"
    'End If
End If

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    rsSQL.CursorType = adOpenDynamic
    rsSQL.LockType = adLockBatchOptimistic
    rsSQL.ActiveConnection = cnServer
    
    If rsSQL.State = 1 Then rsSQL.Close
    rsSQL.Source = "delete  from tb_tempo"
    rsSQL.Open

End Sub
Public Sub Limpiar()
Fracabecera.Enabled = True
Griding.ClearFields
Griding.ClearSelCols

Fracantidad.Enabled = True
Fradetalle.Enabled = False
txtnro.Text = ""
Txtcuspp.Text = ""
txtcodigo.Text = ""
txtfechaini.Text = ""
txtfechafin.Text = ""
txtremu.Text = ""
txtpla.Text = ""
Txtliquidacion.Text = ""
txtpagina.Text = ""
TXTCANTIDAD.Text = ""
Txtliquidacion.SetFocus
 End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not (rsSQL Is Nothing) Then
        If rsSQL.State = 1 Then rsSQL.Close
        Set rsSQL = Nothing
    End If

End Sub

Private Sub TXTCANTIDAD_Change()
If Len(TXTCANTIDAD.Text) > 0 Then
  If IsNumeric(Val(TXTCANTIDAD.Text)) Then
   If Val(TXTCANTIDAD.Text) > 0 Then
        'TXTCANTIDAD.Text
          If Len(Txtliquidacion.Text) > 0 And Len(txtpagina) > 0 And Len(TXTCANTIDAD) > 0 Then
          Fracabecera.Enabled = False
          'Fracantidad.Enabled = False
          Fradetalle.Enabled = True
          txtnro.SetFocus
          End If
        
    Else
        MsgBox "La cantidad debe ser mayor 0 ", vbInformation, "Liquidacion Previa"
    End If
  Else
        MsgBox "Verifique la Cantidad debe ser numerico ", vbInformation, "Liquidacion Previa"
        TXTCANTIDAD.Text = ""
        TXTCANTIDAD.SetFocus
  
  End If
End If
End Sub

Private Sub TXTCANTIDAD_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If Len(TXTCANTIDAD.Text) > 0 Then
          If IsNumeric(Val(TXTCANTIDAD.Text)) Then
          'TXTCANTIDAD.Text
            If Val(TXTCANTIDAD.Text) > 0 Then
                  If Len(Txtliquidacion.Text) > 0 And Len(txtpagina) > 0 And Len(TXTCANTIDAD) > 0 Then
                  Fracabecera.Enabled = False
                  Fracantidad.Enabled = False
                  Fradetalle.Enabled = True
                  Txtcuspp.SetFocus
                  End If
                
                End If
            Else
            MsgBox "La cantidad debe ser mayor 0 ", vbInformation, "Liquidacion Previa"
            End If
        Else
                MsgBox "Verifique la Cantidad debe ser numerico ", vbInformation, "Liquidacion Previa"
                TXTCANTIDAD.Text = ""
                TXTCANTIDAD.SetFocus
        End If
End If
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(txtcodigo.Text) > 0 Then
        txtfechaini.SetFocus
        End If
End If

End Sub

Private Sub txtfechafin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(txtfechafin.Text) > 0 Then
        txtremu.SetFocus
        End If
End If

End Sub

Private Sub TXTliquidacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Len(Txtliquidacion.Text) = 10 Then
      If ExisteNroliq(Txtliquidacion.Text) = True Then
        intIntento = intIntento + 1
            If intIntento >= 3 Then
              'Fradetalle.Enabled = True
            txtpagina.SetFocus
            Else
            MsgBox "El nro de Liquidacion Es Correcto Intento= " & intIntento, vbInformation, "Liquidacion Previa"
            Txtliquidacion.Text = ""
            Txtliquidacion.SetFocus
            End If
      Else
       MsgBox "El nro de Liquidacion No existe ", vbInformation, "Liquidacion Previa"
      End If
    Else
      MsgBox "Ingrese Correctamente nro de Liquidacion ", vbInformation, "Liquidacion Previa"
    End If
  End If
End Sub

Private Sub txtnro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(txtnro.Text) > 0 Then
        Txtcuspp.SetFocus
        End If
End If
End Sub

Private Sub txtpagina_Change()
If Len(txtpagina.Text) > 0 Then
  If IsNumeric(Val(txtpagina.Text)) Then
  'TXTCANTIDAD.Text
    If Len(Txtliquidacion.Text) > 0 And Len(txtpagina) > 0 And Len(TXTCANTIDAD) > 0 Then
    Fracabecera.Enabled = False
    Fracantidad.Enabled = False
    Fradetalle.Enabled = True
    Txtcuspp.SetFocus
    End If
  Else
  MsgBox "Verifique el Nro de Pagina debe ser numerico ", vbInformation, "Liquidacion Previa"
  txtpagina.Text = ""
  txtpagina.SetFocus
  End If
End If
End Sub

Private Sub txtpagina_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Len(txtpagina.Text) > 0 Then
  If IsNumeric(Val(txtpagina.Text)) Then
  'TXTCANTIDAD.Text
    If Len(Txtliquidacion.Text) > 0 And Len(txtpagina) > 0 And Len(TXTCANTIDAD) > 0 Then
    Fracabecera.Enabled = False
    Fracantidad.Enabled = False
    Fradetalle.Enabled = True
    Txtcuspp.SetFocus
    Else
    TXTCANTIDAD.SetFocus
    
    End If
  Else
  MsgBox "Verifique el Nro de Pagina debe ser numerico ", vbInformation, "Liquidacion Previa"
  txtpagina.Text = ""
  txtpagina.SetFocus
  End If
End If

End If
End Sub

Private Sub txtremu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(txtpla.Text) > 0 Then
        cmdagregar.SetFocus
        End If
End If
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHabilitado 
   Caption         =   "Habilitado"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   5415
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "&Importar"
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
      Left            =   1665
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   3225
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Importar archivo IAI de la SBB"
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
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmHabilitado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnExcel As ADODB.Connection
Private rsLEXCEL As ADODB.Recordset

Private Sub cmdBuscar_Click()
     CommonDialog1.Filter = "Archivo de Texto *.dat | *.dat"
     CommonDialog1.FileName = ""
     CommonDialog1.ShowOpen
     
     If CommonDialog1.FileName <> "" Then
           txtArchivo.Text = CommonDialog1.FileName
     End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdImportar_Click()
Importar
End Sub

Private Sub Form_Load()
Set cnExcel = New ADODB.Connection
Set rsLEXCEL = New ADODB.Recordset

End Sub

Sub Importar()
On Error GoTo ERRORES
Dim cmdProm As ADODB.Command
Set cmdProm = New ADODB.Command
Dim Insertxt As String
Dim nreg As Double

nreg = 0

'Conexion

    'Eliminar tabla de Promotores
    With cmdProm
        .ActiveConnection = Con.Cnx
        .CommandText = "Delete IAI1"
        .Execute
    End With

vbcont = 0


NomArchcmp = txtArchivo

Open NomArchcmp For Input As #1
          While Not EOF(1)
                Line Input #1, CadTemp
                
                    If Len(Trim(CadTemp)) > 20 Then
                                      
                                      
                        Id_Form = Trim(Left(CadTemp, 9))
                        Cuspp = Replace(Trim(Mid(CadTemp, 14, 12)), "¥", "Ñ")
                        
                        If Cuspp <> "" Then
                        
                            vApepat = Replace(Trim(Mid(CadTemp, 26, 20)), "¥", "Ñ")
                            vApemat = Replace(Trim(Mid(CadTemp, 46, 20)), "¥", "Ñ")
                            vNomb1 = Replace(Trim(Mid(CadTemp, 66, 10)), "¥", "Ñ")
                            vNomb2 = Replace(Trim(Mid(CadTemp, 76, 10)), "¥", "Ñ")
                            Nac = Trim(Mid(CadTemp, 86, 3))
                            NumIden = Trim(Mid(CadTemp, 89, 10))
                            FecIngreso = Trim(Mid(CadTemp, 99, 8))
                            FecIngreso = Right(FecIngreso, 2) & "/" & Mid(FecIngreso, 5, 2) & "/" & Left(FecIngreso, 4)
                            Estado = Trim(Mid(CadTemp, 107, 3))
                                                                                        
                               
                            'Crear Campo en la base de datos
                            Insertxt = "Insert into IAI1 (Id_Form,Cuspp,Apepat,Apemat,Nomuno, " & _
                            " Nomdos,Nacionalidad, NumIden,FecIngreso,Estado )  " & _
                            " values ('" & Id_Form & "','" & Cuspp & "', '" & vApepat & "', " & _
                            " '" & vApemat & "', '" & vNomb1 & "', '" & vNomb2 & "', '" & Nac & "', '" & NumIden & "','" & FecIngreso & "','" & Estado & "')"
                                         
                            cmdProm.ActiveConnection = Con.Cnx
                            cmdProm.CommandText = Insertxt
                            cmdProm.Execute
                                   
                            nreg = nreg + 1
                        
                        End If
                    
                    End If

             
        Wend
        
        Close #1

    MsgBox "Importación concluida :" & nreg & " registros importados ", 64, Caption
    
    
    GenerarReportes App.Path & "\Expo\AfiliadosLima" & Format(Date, "ddmmyyyy") & ".xls", 1
    GenerarReportes App.Path & "\Expo\AfiliadosProvincia" & Format(Date, "ddmmyyyy") & ".xls", 2
    GenerarReportes App.Path & "\Expo\EmpleadoresLima" & Format(Date, "ddmmyyyy") & ".xls", 3
    GenerarReportes App.Path & "\Expo\EmpleadoresProvincia" & Format(Date, "ddmmyyyy") & ".xls", 4

MsgBox "Archivos Generados", 64, Caption

Exit Sub

ERRORES:
MsgBox Err.Description, vbCritical, Caption
Close #1
'Or .Fields("CodPromo") Is Null
End Sub
Sub GenerarReportes(NombreArchivo As String, indice As String)

On Error GoTo ERRORES
Dim cabecera As String
Dim detalle As String
Dim RsReport As New ADODB.Recordset
Dim comando As String
With RsReport

'    If Dir(vbGRutaStand, vbDirectory) = "" Then
'        MsgBox "Ruta no existe " & vbGRutaStand & ", debe crear la ruta para generar archivos xls", vbCritical, Caption
'        Exit Sub
'    End If
    
    Dim CusRucant As String
    Dim numReg As Integer
    
    numReg = 1
    CusRucant = ""
    
    Select Case indice
    Case 1
        comando = "exec pRepAfiIAILima"
    Case 2
        comando = "exec pRepAfiIAIProvincias"
    Case 3
        comando = "exec pRepEmpIAILima"
    Case 4
        comando = "exec pRepEmpIAIProvincias"
    End Select
    
    .Open comando, Con.Cnx, adOpenStatic, adLockReadOnly
    
    'Expo Cabecera
    cabecera = "Num" & vbTab
    For i = 0 To .Fields.Count - 1
                cabecera = cabecera & .Fields(i).Name & IIf(i < .Fields.Count - 1, vbTab, "")
    Next i
    
    If .RecordCount > 0 Then
            .MoveFirst
                    
                       
           Open NombreArchivo For Output As #1
                   Print #1, cabecera
             
                    Do While Not .EOF
                        detalle = ""
                        detalle = numReg & vbTab
                        'If CusRucant = .Fields("Cuspp/Ruc")  Then
                        
                            For i = 0 To .Fields.Count - 1
                                detalle = detalle & IIf(.Fields(i).Name = "Cuspp/Ruc", "'", "") & .Fields(i).Value & IIf(i < .Fields.Count, vbTab, "")
                            Next i
                            
                            numReg = numReg + 1
                            Print #1, detalle
                        'End If
                        
                        'CusRucant = .Fields("Cuspp/Ruc")
                        .MoveNext
                        
                        DoEvents
                    Loop
                    
                Close #1
            
    End If
    
End With


Exit Sub
ERRORES:
MsgBox Err.Description, vbCritical, Caption

End Sub







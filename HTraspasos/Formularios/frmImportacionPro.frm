VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportacionPro 
   Caption         =   "Importación - Promotor"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3225
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "&Importar"
      Height          =   375
      Left            =   1665
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "frmImportacionPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnExcel As ADODB.Connection
Private rsLEXCEL As ADODB.Recordset

Private Sub cmdBuscar_Click()
     CommonDialog1.Filter = "Archivo de Texto *.txt | *.txt"
     CommonDialog1.FileName = ""
     CommonDialog1.ShowOpen
     
     If CommonDialog1.FileName <> "" Then
           txtArchivo.Text = CommonDialog1.FileName
     End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Command1_Click()

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
        .CommandText = "Delete TPromotor"
        .Execute
    End With

vbcont = 0


NomArchcmp = txtArchivo

Open NomArchcmp For Input As #1
          While Not EOF(1)
                Line Input #1, CadTemp
                If vbcont > 0 Then
                                      
                    vcodpro = Format(Trim(Left(CadTemp, 12)), "0000#")
                    vcodadm = Format(Trim(Mid(CadTemp, 13, 8)), "0000#")
                    vbIden = Left(Trim(Mid(CadTemp, 34, 11)), 10)
                    vApepat = Left(Trim(Mid(CadTemp, 48, 23)), 25)
                    vApemat = Left(Trim(Mid(CadTemp, 71, 23)), 25)
                    vNomb1 = Left(Trim(Mid(CadTemp, 94, 16)), 20)
                    vNomb2 = Left(Trim(Mid(CadTemp, 110, 16)), 20)
                    vbNombres = Left(vNomb1 & " " & vNomb2, 30)
                    vTipSan = Left(Trim(Mid(CadTemp, 140, 1)), 1)
                    vCodsup = Left(Trim(Mid(CadTemp, 154, 11)), 11)
                    vCodage = Left(Trim(Mid(CadTemp, 172, 3)), 3)
                    vbSubApepat = Trim(Left(Trim(Mid(CadTemp, 186, 23)), 23))
                    vbSubApemat = Trim(Left(Trim(Mid(CadTemp, 209, 23)), 23))
                    vbSubNomb1 = Trim(Left(Trim(Mid(CadTemp, 232, 17)), 17))
                    vbSubNomb2 = Trim(Left(Trim(Mid(CadTemp, 249, 17)), 17))
                    
                    vbSubApepat = IIf(UCase(vbSubApepat) = "NULL", "", vbSubApepat)
                    vbSubApemat = IIf(UCase(vbSubApemat) = "NULL", "", vbSubApemat)
                    vbSubNomb1 = IIf(UCase(vbSubNomb1) = "NULL", "", vbSubNomb1)
                    vbSubNomb2 = IIf(UCase(vbSubNomb2) = "NULL", "", vbSubNomb2)
                    
                    vNomSup = vbSubApepat & " " & vbSubApemat & " " & vbSubNomb1 & " " & vbSubNomb2
                    vbDesPromo = vApepat & " " & vApemat & " " & vbNombres
                                                            
                    vbcont = vbcont + 1
                                                                                
                    If Not (vcodpro = "00000") Then
                        
                       If fgDatosql("Select codpromo from Tpromotor where codpromo = '" & vcodpro & "'  ") = "" Then
                       
                           If vTipSan = "R" Then vTipSan = ""
                           
                           'Crear Campo en la base de datos
                           Insertxt = "Insert into TPromotor (CodPromo, Ape_Pat, Ape_Mat, " & _
                           " Nombres, Agencia, DesPromo, DocIden, CodProAd, Tipsanc, Codsup, Codage, Nomsup  )  " & _
                           " values ('" & vcodpro & "','" & vApepat & "', '" & vApemat & "', " & _
                           " '" & vbNombres & "', '" & vCodage & "', '" & vbDesPromo & "', '" & vbIden & "', '" & vcodadm & "','" & vTipSan & "','" & vCodsup & "', '" & vCodage & "', '" & vNomSup & "'   )"
                                 
                           cmdProm.ActiveConnection = Con.Cnx
                           cmdProm.CommandText = Insertxt
                           cmdProm.Execute
                           
                           nreg = nreg + 1
                       End If
                    
                    End If
                   
                
                End If
                
                 vbcont = vbcont + 1
             
             
        Wend
        
        Close #1

    MsgBox "Importación concluida :" & nreg & " registros importados ", 64, Caption
    



Exit Sub

ERRORES:
MsgBox Err.Description, vbCritical, Caption
Close #1
'Or .Fields("CodPromo") Is Null
End Sub

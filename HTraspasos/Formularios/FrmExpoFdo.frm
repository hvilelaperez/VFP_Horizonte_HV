VERSION 5.00
Begin VB.Form FrmExpoFdo 
   Caption         =   "Exporta Archivos de Tipos Fondo"
   ClientHeight    =   5910
   ClientLeft      =   3960
   ClientTop       =   2610
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7200
   Begin VB.CommandButton CmdSalir 
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
      Height          =   615
      Left            =   4680
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton CmdExpo 
      Caption         =   "&Exportar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame FraModi 
      Caption         =   "Exporta Archivos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.OptionButton OptErrores2 
         Caption         =   "Solic Ingresadas <> BD Carga   (Ing vs. Orig) - AFP,CUSPP,Nombres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   6495
      End
      Begin VB.OptionButton OptErrNSol 
         Caption         =   "Posibles Num Solicitudes Erradas  (FisNoBD vs. BDNoFis)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   5535
      End
      Begin VB.OptionButton OptIngre 
         Caption         =   "Archivo Registros Ingresados (Ing)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   4080
         Width           =   4215
      End
      Begin VB.OptionButton OptOrig 
         Caption         =   "Archivo Carga Original (Orig)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   4215
      End
      Begin VB.OptionButton OptFdoDif22 
         Caption         =   "Solicitudes con Fondos <> 2 - 2  (Ing)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   4815
      End
      Begin VB.OptionButton OptFisNoBD 
         Caption         =   "Solicitudes Fisicos sin Registro en BD Carga (Ing)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   4815
      End
      Begin VB.OptionButton OptBDNoFis 
         Caption         =   "Registros de Solicitudes en BD sin Fisico (Orig)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   4815
      End
      Begin VB.OptionButton OptErrores 
         Caption         =   "Solic Ingresadas <> BD Carga   (Ing vs. Orig) - Solo Fondo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Width           =   5535
      End
   End
End
Attribute VB_Name = "FrmExpoFdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsExpoFdo As New Recordset
Dim Strsqltexto As String
Dim xarchi As String
Dim xindexpo As String

Private Sub CmdExpo_Click()

    xarchi = ""
    xindexpo = "0"
    
    If OptFdoDif22.Value Then
        Strsqltexto = "SELECT * FROM FdoDif22 order by Anno,Mes,Lote,NumSol "
        xarchi = "FdoDif22.txt"
    End If
    If OptFisNoBD.Value Then
        Strsqltexto = "SELECT * FROM FisNoBD order by Anno,Mes,Lote,NumSol "
        xarchi = "FisNoBD.txt"
    End If
    If OptBDNoFis.Value Then
        Strsqltexto = "SELECT * FROM BDNoFis order by Anno,Mes,Lote,NumSol "
        xarchi = "BDNoFis.txt"
    End If
    If OptErrores.Value Then
        Strsqltexto = "SELECT * FROM FdoERROR order by Anno,Mes,Lote,NumSol "
        xarchi = "FdoERROR.txt"
    End If
    If OptErrores2.Value Then
        Strsqltexto = "SELECT * FROM FdoERROR2 order by Anno,Mes,Lote,NumSol "
        xarchi = "FdoERROR2.txt"
    End If
    If OptErrNSol.Value Then
        Strsqltexto = "SELECT * FROM FdoNSol order by Anno,Mes,Lote,NumSol "
        xarchi = "FdoNSol.txt"
    End If
    If OptOrig.Value Then
        Strsqltexto = "SELECT Anno,Mes,Lote,NumSol=N_Solici,Cuspp,Ape_Pat,Ape_Mat,Nombre1,Nombre2,CodAfp=AFP,TipFdoOri=Ori,TipFdoDes=Des,cruzo,sec FROM RevFdoCarga order by Anno,Mes,Lote,NumSol "
        xarchi = "RevFdoCarga.txt"
        xindexpo = "1"
    End If
    If OptIngre.Value Then
        Strsqltexto = "SELECT Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,Ape_Mat,Nombre1,Nombre2,CodAfp,TipFdoOri,TipFdoDes,cruzo,sec,usudig,fecdig FROM RevFdodig order by Anno,Mes,Lote,NumSol "
        xarchi = "RevFdodig.txt"
        xindexpo = "2"
    End If

    Call Exportar

End Sub

Sub Exportar()

    If RsExpoFdo.State = 1 Then RsExpoFdo.Close
    RsExpoFdo.Open Strsqltexto, Con.Cnx, adOpenStatic, adLockReadOnly

    If Not (RsExpoFdo.EOF And RsExpoFdo.BOF) Then

        With RsExpoFdo
        Open App.Path & "\expo\" & xarchi For Output As #1
            'Exporta Archivos resultados
            If xindexpo = "0" Then
                Print #1, "Año" & "," & "Mes" & "," & "Lote" & "," & "NumSol" & "," & _
                    "CUSPP" & "," & "Ape_Pat" & "," & "Ape_Mat" & "," & "Nombre1" & "," & "Nombre2" & "," & _
                    "codafp" & "," & "TipFdoOri" & "," & "TipFdoDes" & "," & _
                    "Nom_campo" & "," & "Val_Dif" & "  "
        
                While Not .EOF
                    Print #1, .Fields("Anno") & "," & .Fields("Mes") & "," & .Fields("Lote") & "," & Trim(.Fields("NumSol")) & "," & _
                        .Fields("CUSPP") & "," & .Fields("Ape_Pat") & "," & .Fields("Ape_Mat") & "," & .Fields("Nombre1") & "," & .Fields("Nombre2") & "," & _
                        .Fields("codafp") & "," & .Fields("TipFdoOri") & "," & .Fields("TipFdoDes") & "," & _
                        .Fields("Nom_campo") & "," & .Fields("Val_Dif") & "  "
                        .MoveNext
                Wend
            Else
                'Exporta Original - Archivo Cargado
                If xindexpo = "1" Then
                    Print #1, "Año" & "," & "Mes" & "," & "Lote" & "," & "NumSol" & "," & _
                        "CUSPP" & "," & "Ape_Pat" & "," & "Ape_Mat" & "," & "Nombre1" & "," & "Nombre2" & "," & _
                        "codafp" & "," & "TipFdoOri" & "," & "TipFdoDes" & "," & _
                        "Cruzo" & "," & "Sec" & "  "
        
                    While Not .EOF
                        Print #1, .Fields("Anno") & "," & .Fields("Mes") & "," & .Fields("Lote") & "," & Trim(.Fields("NumSol")) & "," & _
                        .Fields("CUSPP") & "," & .Fields("Ape_Pat") & "," & .Fields("Ape_Mat") & "," & .Fields("Nombre1") & "," & .Fields("Nombre2") & "," & _
                        .Fields("codafp") & "," & .Fields("TipFdoOri") & "," & .Fields("TipFdoDes") & "," & _
                        .Fields("Cruzo") & "," & .Fields("Sec") & "  "
                        .MoveNext
                    Wend
                Else
                    'Exporta Archivo Ingresado
                    Print #1, "Año" & "," & "Mes" & "," & "Lote" & "," & "NumSol" & "," & _
                        "CUSPP" & "," & "Ape_Pat" & "," & "Ape_Mat" & "," & "Nombre1" & "," & "Nombre2" & "," & _
                        "codafp" & "," & "TipFdoOri" & "," & "TipFdoDes" & "," & _
                        "Cruzo" & "," & "Sec" & "," & "UsuDig" & "," & "FecDig" & "  "
        
                    While Not .EOF
                        Print #1, .Fields("Anno") & "," & .Fields("Mes") & "," & .Fields("Lote") & "," & Trim(.Fields("NumSol")) & "," & _
                        .Fields("CUSPP") & "," & .Fields("Ape_Pat") & "," & .Fields("Ape_Mat") & "," & .Fields("Nombre1") & "," & .Fields("Nombre2") & "," & _
                        .Fields("codafp") & "," & .Fields("TipFdoOri") & "," & .Fields("TipFdoDes") & "," & _
                        .Fields("Cruzo") & "," & .Fields("Sec") & "," & .Fields("UsuDig") & "," & Format(.Fields("FecDig"), "ddmmyyyy") & "  "
                        .MoveNext
                    Wend
                End If
            End If
    
    
            Close #1
        End With
        MsgBox "Exportación Realizada", 64, Caption
    Else
        MsgBox "No se encontró información", 64, Caption
    End If

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub



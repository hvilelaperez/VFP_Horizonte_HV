VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmProcFdo 
   Caption         =   "Procesos y Reportes"
   ClientHeight    =   8925
   ClientLeft      =   525
   ClientTop       =   1470
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   13830
   Begin VB.Frame Frame1 
      Caption         =   "  Cruza Archivos (Orig vs Ing)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9120
      TabIndex        =   6
      Top             =   240
      Width           =   3735
      Begin VB.CommandButton CmdCruza 
         Caption         =   "&Procesar"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame FraModi 
      Caption         =   "Genera archivos y Reportes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   8535
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   2880
         Width           =   5535
      End
      Begin VB.CommandButton CmdRepo 
         Caption         =   "&Reporte Vista Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6840
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdEjecutar 
         Caption         =   "&Ejecutar"
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
         Left            =   6840
         TabIndex        =   11
         Top             =   600
         Width           =   1335
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
         TabIndex        =   5
         Top             =   1920
         Width           =   5535
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
         TabIndex        =   4
         Top             =   1440
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
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   4815
      End
   End
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
      Left            =   11760
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DtGFondos 
      Height          =   5055
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   8916
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "anno"
         Caption         =   "Año"
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
         DataField       =   "Mes"
         Caption         =   "Mes"
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
      BeginProperty Column03 
         DataField       =   "numsol"
         Caption         =   "N Solicitud"
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
         DataField       =   "cuspp"
         Caption         =   "CUSPP"
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
         DataField       =   "codafp"
         Caption         =   "AFP"
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
         DataField       =   "tipfdoori"
         Caption         =   "Fdo Origen"
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
         DataField       =   "tipfdodes"
         Caption         =   "Fdo Destino"
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
         DataField       =   "nom_campo"
         Caption         =   "Campo"
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
         DataField       =   "Val_Dif"
         Caption         =   "Valor Dif"
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
      BeginProperty Column10 
         DataField       =   "ape_pat"
         Caption         =   "Ape Paterno"
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
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   5999.812
         EndProperty
      EndProperty
   End
   Begin VB.Label LblTotReg 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   9240
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Registros :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9240
      TabIndex        =   8
      Top             =   2280
      Width           =   1425
   End
End
Attribute VB_Name = "FrmProcFdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsConsFdo As New Recordset
Public RsRepoFdo As New Recordset

Private Sub DtGFondos_HeadClick(ByVal ColIndex As Integer)
    RsConsFdo.Sort = DtGFondos.Columns(ColIndex).DataField
End Sub

Private Sub CmdCruza_Click()
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "update RevFdoCarga " & _
            "set cruzo = 'N' "
            .Execute
            .CommandText = "update RevFdoDig " & _
            "set cruzo = 'N' "
            .Execute
            .CommandText = "update RevFdoCarga " & _
            "Set RevFdoCarga.lote = RevFdoDig.lote, " & _
            "RevFdoCarga.cruzo = 'S' " & _
            "FROM RevFdoCarga, RevFdoDig WHERE RevFdoCarga.n_solici=RevFdoDig.numsol "
            .Execute
            .CommandText = "update RevFdoDig " & _
            "Set RevFdoDig.sec = RevFdoCarga.sec, " & _
            "RevFdoDig.cruzo = 'S' " & _
            "FROM RevFdoCarga, RevFdodig WHERE RevFdoDig.numsol=RevFdoCarga.n_solici "
            .Execute
            .CommandText = "update RevFdoCarga " & _
            "set cruzo = 'F' " & _
            "where cruzo<>'S' "
            .Execute
            .CommandText = "update RevFdoDig " & _
            "set cruzo = 'F' " & _
            "where cruzo<>'S' "
            .Execute
        End With
        MsgBox " Proceso Ejecutado ", vbInformation, "Cruzo RevFdoCarga vs. RevFdoDig"
End Sub

Private Sub CmdEjecutar_Click()
    If RsConsFdo.State = 1 Then RsConsFdo.Close
'*** 1
    If OptFdoDif22.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "delete FdoDif22"
            .Execute
            .CommandText = "insert into FdoDif22 (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,Codafp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,'FdoDif22','*' " & _
            "from RevFdoDig where TipFdoOri <> '2' or TipFdoDes <> '2' order by Anno,Mes,Lote,NumSol "
            .Execute
        End With
        SQLstr = "SELECT * FROM FdoDif22 ORDER BY Lote, NumSol "
    End If
'*** 2
    If OptFisNoBD.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "delete FisNoBD"
            .Execute
            .CommandText = "insert into FisNoBD (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,'FisNoBD','*' " & _
            "from RevFdoDig where cruzo = 'F' order by Anno,Mes,Lote,NumSol "
            .Execute
        End With
        SQLstr = "SELECT * FROM FisNoBD ORDER BY Lote, NumSol "
    End If
'*** 3
    If OptBDNoFis.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "delete BDNoFis"
            .Execute
            .CommandText = "insert into BDNoFis (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select Anno,Mes,Lote,N_solici,Cuspp,Ape_Pat,Afp,Ori,Des,'BDNoFis','*' " & _
            "from RevFdoCarga where cruzo = 'F' order by Anno,Mes,Lote,N_solici "
            .Execute
        End With
        SQLstr = "SELECT * FROM BDNoFis ORDER BY Lote, NumSol "
    End If
'*** 4
    If OptErrores.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "delete FdoERROR"
            .Execute
            .CommandText = "insert into FdoERROR (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select b.Anno,b.Mes,b.Lote,b.NumSol,b.Cuspp,b.Ape_Pat,b.CodAfp,b.TipFdoOri,b.TipFdoDes,'FDO ORIGEN',a.ORI " & _
            "from RevFdoCarga a, RevFdoDig b where a.n_solici=b.numsol and a.ORI<>b.TipFdoOri order by b.Anno,b.Mes,b.Lote,b.NumSol "
            .Execute
            .CommandText = "insert into FdoERROR (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select b.Anno,b.Mes,b.Lote,b.NumSol,b.Cuspp,b.Ape_Pat,b.CodAfp,b.TipFdoOri,b.TipFdoDes,'FDO DESTINO',a.DES " & _
            "from RevFdoCarga a, RevFdoDig b where a.n_solici=b.numsol and a.DES<>b.TipFdoDes order by b.Anno,b.Mes,b.Lote,b.NumSol "
            .Execute
        End With
        SQLstr = "SELECT * FROM FdoERROR ORDER BY Lote, NumSol "
    End If
'*** 5
    If OptErrores2.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "delete FdoERROR2"
            .Execute
            .CommandText = "insert into FdoERROR2 (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select b.Anno,b.Mes,b.Lote,b.NumSol,b.Cuspp,b.Ape_Pat,b.CodAfp,b.TipFdoOri,b.TipFdoDes,'CUSPP',a.cuspp " & _
            "from RevFdoCarga a, RevFdoDig b where a.n_solici=b.numsol and a.cuspp<>b.cuspp order by b.Anno,b.Mes,b.Lote,b.NumSol "
            .Execute
            .CommandText = "insert into FdoERROR2 (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select b.Anno,b.Mes,b.Lote,b.NumSol,b.Cuspp,b.Ape_Pat,b.CodAfp,b.TipFdoOri,b.TipFdoDes,'AFP',a.AFP " & _
            "from RevFdoCarga a, RevFdoDig b where a.n_solici=b.numsol and a.afp<>b.codafp order by b.Anno,b.Mes,b.Lote,b.NumSol "
            .Execute
            .CommandText = "insert into FdoERROR2 (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select b.Anno,b.Mes,b.Lote,b.NumSol,b.Cuspp,b.Ape_Pat,b.CodAfp,b.TipFdoOri,b.TipFdoDes,'APE_PAT',a.Ape_Pat " & _
            "from RevFdoCarga a, RevFdoDig b where a.n_solici=b.numsol and a.Ape_Pat<>b.Ape_Pat order by b.Anno,b.Mes,b.Lote,b.NumSol "
            .Execute
            
            
        End With
        SQLstr = "SELECT * FROM FdoERROR2 ORDER BY Lote, NumSol "
    End If
'*** 6
    If OptErrNSol.Value Then
        Set comando = New Command
        With comando
            .ActiveConnection = Con.Cnx
            .CommandText = "delete FdoNSol"
            .Execute
            .CommandText = "insert into FdoNSol (Anno,Mes,Lote,NumSol,Cuspp,Ape_Pat,CodAfp,TipFdoOri,TipFdoDes,Nom_Campo,Val_Dif) " & _
            "select b.Anno,b.Mes,b.Lote,b.NumSol,b.Cuspp,b.Ape_Pat,b.CodAfp,b.TipFdoOri,b.TipFdoDes,'FisNoBD / BDNoFis',a.numsol " & _
            "from BDNoFis a, FisNoBD b where a.cuspp=b.cuspp order by b.Anno,b.Mes,b.Lote,b.NumSol "
            .Execute
        End With
        SQLstr = "SELECT * FROM FdoNSol ORDER BY Lote, NumSol "
    End If
    
    RsConsFdo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    LlenaControles
End Sub

Private Sub LlenaControles()
    LblTotReg = 0
    With RsConsFdo
        If .BOF And .EOF Then   'pregunta si tabla esta vacia
            MsgBox "No existen registros ", vbCritical, "Error en Consulta de Fondos"
            DtGFondos.Refresh
            Exit Sub
        Else
            Set DtGFondos.DataSource = RsConsFdo
            DtGFondos.Refresh
            If DtGFondos.ApproxCount <> 0 Then
                .MoveFirst
                For i = 1 To DtGFondos.ApproxCount
                    DtGFondos.Col = 0
                    LblTotReg = Val(LblTotReg) + 1
                    .MoveNext
                Next i
            End If
            .MoveFirst
        End If
    End With
End Sub

Private Sub CmdRepo_Click()
    If RsRepoFdo.State = 1 Then RsRepoFdo.Close
    
    If OptFdoDif22.Value Then
        SQLstr = "SELECT * FROM FdoDif22 order by Anno,Mes,Lote,NumSol "
    End If
    If OptFisNoBD.Value Then
        SQLstr = "SELECT * FROM FisNoBD order by Anno,Mes,Lote,NumSol "
    End If
    If OptBDNoFis.Value Then
        SQLstr = "SELECT * FROM BDNoFis order by Anno,Mes,Lote,NumSol "
    End If
    If OptErrores.Value Then
        SQLstr = "SELECT * FROM FdoERROR order by Anno,Mes,Lote,NumSol "
    End If
    If OptErrores2.Value Then
        SQLstr = "SELECT * FROM FdoERROR2 order by Anno,Mes,Lote,NumSol "
    End If
    If OptErrNSol.Value Then
        SQLstr = "SELECT * FROM FdoNSol order by Anno,Mes,Lote,NumSol "
    End If

    RsRepoFdo.Open SQLstr, Con.Cnx, adOpenStatic, adLockReadOnly
    Set DtRFdo.DataSource = RsRepoFdo
    DtRFdo.Show

End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub


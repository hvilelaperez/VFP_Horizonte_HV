VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCruce 
   Caption         =   "Cruce Información Re-Afiliación"
   ClientHeight    =   6645
   ClientLeft      =   1950
   ClientTop       =   2610
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   11700
   Begin VB.CommandButton cmdCruzar 
      Caption         =   "&Cruce"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
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
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   4
         Top             =   320
         Width           =   1215
      End
      Begin VB.TextBox txtFecsal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dtgAfiliacion 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7858
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "numsol"
         Caption         =   "Documento"
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
         DataField       =   "numdoc"
         Caption         =   "Doc Iden"
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
         DataField       =   "Fecsal"
         Caption         =   "Fec Sal"
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
         DataField       =   "Apepat"
         Caption         =   "Ape Pat"
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
         DataField       =   "Apemat"
         Caption         =   "Ape Mat"
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
         DataField       =   "Nombre1"
         Caption         =   "Nombre1"
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
         DataField       =   "Nombre2"
         Caption         =   "Nombre2"
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
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCruce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCruce As Recordset
Private Sub cmdBuscar_Click()
    Buscar
End Sub
Sub Buscar()
Dim sqltexto As String

sqltexto = " Select a.NumSol, a.indDoc, a.FecSal, a.Secu, a.PosIma, a.Lote, a.Nro_Cliente, b.NumCli, b.cuspp, b.[NroDoc], " & _
           " a.ApePat , a.ApeMat, a.Nombre1, a.Nombre2, a.TipoDoc, a.NumDoc, a.NumSec, b.Clavesp " & _
           " from solicitud a join PolydocAfiliaciones..TBL001011 b on " & _
           " a.NumDoc = b.[NroDoc] " & _
           " where indcruce = 0 and inddoc = 'C' and b.SecAfi Is Null "
           
           '" a.ApePat = b.Paterno and a.ApeMat = b.Materno and  a.Nombre1 = b.Nombre1 and " & _
           " a.Nombre2 = b.Nombre2 " & _


           
           If txtFecsal.Text <> "" Then
                sqltexto = sqltexto & " and fecsal = '" & txtFecsal.Text & "'"
           End If

With rsCruce
    If .State = 1 Then .Close
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
    
    
    Set dtgAfiliacion.DataSource = rsCruce
    
    If .RecordCount = 0 Then MsgBox "No hay información para Cruzar", 64, Caption
    
End With

End Sub

Private Sub cmdCruzar_Click()
Call Cruzar
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rsCruce = New ADODB.Recordset
End Sub
Sub CruzarInfo(Numsol As String, Secafi As Double, NroDoc As String, Clavesp As String, Numcli As String, Cuspp As String)
'ApePat As String, ApeMat As String, Nombre1 As String, Nombre2 As String
On Error GoTo errores
Dim sqltexto As String
Dim rsImagPolyDoc As Recordset
Dim rsImagTraspas As Recordset

Set rsImagPolyDoc = New ADODB.Recordset
Set rsImagTraspas = New ADODB.Recordset

sqltexto = " select Ruta = Ruta + '\' +  right('00000000' + a.Lote,8) + '\' + imagen + '.tif', imagen, Lote = right('00000000' + a.Lote,8)   from PolydocAfiliaciones..documento a join PolydocAfiliaciones..imagenes b " & _
 " on a.idcliente = b.idcliente and a.idproyecto = b.idproyecto and a.idvolumen = b.idvolumen " & _
 " and a.idlote = b.idlote and a.nrodoc1 = b.nrodoc1 and a.nrodoc2 = b.nrodoc2 " & _
 " join PolydocAfiliaciones..volumen c on a.idcliente = c.idcliente and a.idproyecto = c.idproyecto " & _
 " and a.idvolumen = c.idvolumen " & _
 " where a.NC1 = '" & Clavesp & "' and b.barcode = 0 "
 
  'paterno, materno, nombre1, nombre2
'" Where paterno = '" & ApePat & "' and  materno = '" & ApeMat & "' and nombre1 = '" & Nombre1 & "' and nombre2 = '" & Nombre2 & "' and " & _

 'NroDoc
 
 With rsImagPolyDoc
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
 End With
 
 With rsImagTraspas
    sqltexto = " select Ruta = Ruta +'\'+ LoteFisico +'\',a.idVolumen, Lote, idImagen, Imagen," & _
               " Tipo, BarCode, Posicion,LoteFisico, Status, TipImag  " & _
               " from _imagenes a join _volumenes b " & _
               " on a.idvolumen = b.idvolumen " & _
               " Where idimagen = '01" & Numsol & " ' and TipImag = 'S' order by posicion "
    .Open sqltexto, Con.Cnx, adOpenStatic, adLockReadOnly
    
 End With
    
    If rsImagTraspas.RecordCount > 0 Then
        'Comparar
        For i = 1 To rsImagTraspas.RecordCount
        
            rsImagTraspas.AbsolutePosition = i
            
            If rsImagPolyDoc.RecordCount < i Then Exit For
            rsImagPolyDoc.AbsolutePosition = i

            
            
            'idVolumen, Lote, idImagen, Imagen,
            'Tipo , BarCode, Posicion, lotefisico, Status, TipImag
            
            If Err.Number = 0 Then
            'Copiar Imagen
                If CopiarFile(rsImagPolyDoc.Fields("Ruta"), rsImagTraspas.Fields("Ruta") & rsImagPolyDoc.Fields("imagen") & ".tif") Then
                    With rsImagTraspas
                        If Not InsertarImagen(.Fields("idVolumen"), .Fields("Lote"), .Fields("idImagen"), rsImagPolyDoc.Fields("Imagen"), 0, _
                                          .Fields("posicion"), .Fields("Lotefisico"), .Fields("Status"), "A") Then
                               'Eliminar Imagen
                               Kill rsImagTraspas.Fields("Ruta") & rsImagPolyDoc.Fields("imagen") & ".tif"
                               Exit Sub
                        Else
                            'Actualizar NumCli en Afiliacion
                            
                            ActualizarAfiliacion Secafi, Numcli, Cuspp
                            
                            'Actualizar num. secuencial en tbl
                            ActualizarTbl Secafi, Clavesp
                            
                        End If
                        
                        
                    End With
                End If

            End If
             
        
        Next
        
    End If
    
    



Exit Sub

errores:
'Select Case Err.Number

    MsgBox Err.Description, vbCritical, Caption
'End Select

End Sub

Function InsertarImagen(idvolumen As String, Lote As String, idImagen As String, _
        Imagen As String, Barcode As String, Posicion As String, lotefisico As String, _
        Status As String, TipImag As String) As Boolean
InsertarImagen = True
Dim cmdRegistro As ADODB.Command
Set cmdRegistro = New ADODB.Command

On Error GoTo errores

With cmdRegistro
    .ActiveConnection = Con.Cnx
    .CommandText = "pInserImagReAfi ?,?,?,?,?,?,?,?,? "
    .Parameters("@idVolumen") = idvolumen
    .Parameters("@Lote") = Lote
    .Parameters("@idImagen") = idImagen
    .Parameters("@Imagen") = Imagen
    .Parameters("@BarCode") = Barcode
    .Parameters("@Posicion") = Posicion
    .Parameters("@LoteFisico") = lotefisico
    .Parameters("@Status") = Status
    .Parameters("@TipImag") = TipImag
    .Execute
End With

Exit Function
errores:
InsertarImagen = False
MsgBox Err.Description, vbCritical, "Insercion de Imagen"

End Function

Function ActualizarAfiliacion(Numsec As Double, Numcli As String, Cuspp As String) As Boolean

Dim cmdAfil As ADODB.Command
Set cmdAfil = New ADODB.Command

With cmdAfil

    .ActiveConnection = Con.Cnx
    .CommandText = " Update Solicitud set Nro_Cliente = '" & Numcli & "', CUSPP = '" & Cuspp & "', " & _
               " indcruce = 1 where numsec = " & Numsec & " "
    .Execute
                     
End With


End Function


Function ActualizarTbl(Numsec As Double, Clavesp) As Boolean

Dim cmdAfil As ADODB.Command
Set cmdAfil = New ADODB.Command

With cmdAfil

    .ActiveConnection = Con.Cnx
    .CommandText = " Update PolydocAfiliaciones..TBL001011  set SecAfi = " & Numsec & _
                   " where clavesp = '" & Clavesp & "'"
    .Execute
                     
End With


End Function

Sub Cruzar()
On Error GoTo errores
With rsCruce
    If .RecordCount > 0 Then
        .MoveFirst
                
        Do While Not .EOF
            CruzarInfo .Fields("numsol"), .Fields("numsec"), .Fields("NumDoc"), .Fields("Clavesp"), .Fields("NumCli"), .Fields("cuspp")
            .MoveNext
        Loop
        MsgBox "Proceso de Cruce Terminado", 64, Caption
    End If
End With
Exit Sub
errores:
MsgBox Err.Description, vbCritical, Caption
End Sub


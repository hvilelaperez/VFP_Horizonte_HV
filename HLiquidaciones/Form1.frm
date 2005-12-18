VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cobar7"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cobar6"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cobar5"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cobar4"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Vartexto As String
Dim Lproceso As String
Dim NroCargo As String
Dim Modcarta As String
Dim Tipcargo As String
Dim Localidad  As String
Dim FechTexto As String
Dim Quien As String
Dim Razon As String
Dim direc1 As String
Dim direc2 As String
Dim TipoIdent As String
Dim NroIdent As String
Dim NomDistr As String
Dim Nomprov As String
Dim Atencion As String
Dim Referencia As String
Dim Adjunto As String
Dim Linea1 As String
Dim Linea2 As String
Dim Linea3 As String
Dim Linea4 As String
Dim Linea5 As String
Open "C:\prueba" & "\COBAR4.txt" For Input As #1
'''923 Caracteres longitud de Texto
Do Until EOF(1)
   Line Input #1, Vartexto
        Lproceso = Mid(Vartexto, 1, 11)
        NroCargo = Mid(Vartexto, 12, 9)
        Modcarta = Mid(Vartexto, 21, 1)
        Tipcargo = Mid(Vartexto, 22, 4)
        Localidad = Mid(Vartexto, 26, 10)
        FechTexto = Mid(Vartexto, 36, 20)
        Quien = Mid(Vartexto, 56, 15)
        Razon = Mid(Vartexto, 71, 100)
        direc1 = Mid(Vartexto, 171, 75)
        direc2 = Mid(Vartexto, 246, 75)
        TipoIdent = Mid(Vartexto, 321, 3)
        NroIdent = Mid(Vartexto, 324, 16)
        NomDistr = Mid(Vartexto, 340, 30)
        Nomprov = Mid(Vartexto, 370, 63)
        Atencion = Mid(Vartexto, 433, 45)
        Referencia = Mid(Vartexto, 478, 75)
        Adjunto = Mid(Vartexto, 553, 21)
        Linea1 = Mid(Vartexto, 574, 70)
        Linea2 = Mid(Vartexto, 644, 70)
        Linea3 = Mid(Vartexto, 714, 70)
        Linea4 = Mid(Vartexto, 784, 70)
        Linea5 = Mid(Vartexto, 854, 70)
   
   ''Grabar en dbf
   
   Grabar_Dbf Lproceso, NroCargo, Modcarta, Tipcargo, Localidad, _
              FechTexto, Quien, Razon, direc1, direc2, TipoIdent, NroIdent, _
              NomDistr, Nomprov, Atencion, Referencia, Adjunto, Linea1, Linea2, _
              Linea3, Linea4, Linea5
   
   Loop
Close #1
End Sub

Private Sub Command2_Click()
Dim Vartexto As String
Dim Lproceso As String
Dim NroCargo As String
Dim NroLiq As String
Dim Tipcargo As String
Dim Secuen  As String
Dim TipImpr As String
Dim Periodo As String
Dim FechTexto As String
Dim Pag As String
Dim Razon As String
Dim TipoIdent As String
Dim NroIdent As String
'Dim TipoIdent As String
Dim direcav As String
Dim direcnu As String
Dim direcurb As String
Dim NomDistr As String
Dim Nomprov As String
Dim NomDepar As String

Open "C:\prueba" & "\COBAR5.txt" For Input As #1
'''368 Caracteres longitud de Texto
Do Until EOF(1)
   Line Input #1, Vartexto
        Lproceso = Mid(Vartexto, 1, 11)
        NroCargo = Mid(Vartexto, 12, 9)
        NroLiq = Mid(Vartexto, 21, 10)
        Tipcargo = Mid(Vartexto, 31, 4)
        Secuen = Mid(Vartexto, 35, 2)
        TipImpr = Mid(Vartexto, 37, 1)
        Periodo = Mid(Vartexto, 38, 7)
        FechTexto = Mid(Vartexto, 45, 20)
        Pag = Mid(Vartexto, 65, 5)
        Razon = Mid(Vartexto, 70, 100)
        TipoIdent = Mid(Vartexto, 170, 3)
        NroIdent = Mid(Vartexto, 173, 16)
        direcav = Mid(Vartexto, 189, 50)
        direcnu = Mid(Vartexto, 239, 15)
        direcurb = Mid(Vartexto, 254, 25)
        NomDistr = Mid(Vartexto, 279, 30)
        Nomprov = Mid(Vartexto, 309, 30)
        NomDepar = Mid(Vartexto, 339, 30)
        
   
   ''Grabar en dbf
   Grabar_Dbf1 Lproceso, NroCargo, NroLiq, Tipcargo, Secuen, _
              TipImpr, Periodo, FechTexto, Pag, Razon, TipoIdent, NroIdent, _
              direcav, direcnu, direcurb, NomDistr, Nomprov, NomDepar
   Loop
Close #1
End Sub

Private Sub Command3_Click()
Dim Vartexto As String
Dim Lproceso As String
Dim NroCargo As String
Dim NroLiq As String
Dim Tipdet As String
Dim Appat  As String
Dim Apmat As String
Dim Nom1 As String
Dim Nom2 As String
Dim Cusp As String
Dim Remun As String
Dim Montf As String
Dim Montr As String
'Dim TipoIdent As String
Dim Totalaf As String
Dim Totalfdo As String
Dim Totalint As String
Dim Totalgen As String
Dim Pag As String


Open "C:\prueba" & "\COBAR6.txt" For Input As #1
'''199 Caracteres longitud de Texto
Do Until EOF(1)
   Line Input #1, Vartexto
        Lproceso = Mid(Vartexto, 1, 11)
        NroCargo = Mid(Vartexto, 12, 9)
        NroLiq = Mid(Vartexto, 21, 10)
        Tipdet = Mid(Vartexto, 31, 1)
        Appat = Mid(Vartexto, 32, 20)
        Apmat = Mid(Vartexto, 52, 20)
        Nom1 = Mid(Vartexto, 72, 15)
        Nom2 = Mid(Vartexto, 87, 15)
        Cusp = Mid(Vartexto, 102, 12)
        Remun = Mid(Vartexto, 114, 10)
        Montf = Mid(Vartexto, 124, 10)
        Montr = Mid(Vartexto, 134, 10)
        Totalaf = Mid(Vartexto, 144, 10)
        Totalfdo = Mid(Vartexto, 154, 12)
        Totalint = Mid(Vartexto, 166, 12)
        Totalgen = Mid(Vartexto, 178, 14)
        Pag = Mid(Vartexto, 192, 8)
   
   ''Grabar en dbf
   
   Grabar_Dbf2 Lproceso, NroCargo, NroLiq, Tipdet, Appat, _
              Apmat, Nom1, Nom2, Cusp, Remun, Montf, Montr, _
              Totalaf, Totalfdo, Totalint, Totalgen, Pag
   
   Loop
Close #1

End Sub

Private Sub Command4_Click()
Dim Tippro As String
Dim Codmen  As String
Dim Sec As String
Dim Mens As String


Open "C:\prueba" & "\COBAR7.txt" For Input As #1
'''159 Caracteres longitud de Texto
Do Until EOF(1)
   Line Input #1, Vartexto
        Tippro = Mid(Vartexto, 1, 3)
        Codmen = Mid(Vartexto, 4, 3)
        Sec = Mid(Vartexto, 7, 3)
        Mens = Mid(Vartexto, 10, 150)
                
   
   ''Grabar en dbf
   
   Grabar_Dbf3 Tippro, Codmen, Sec, Mens
   
   Loop
Close #1

End Sub


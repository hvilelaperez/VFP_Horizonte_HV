Attribute VB_Name = "Mod_separador_nomapell"
Global gApePat As String
Global gApeMat As String
Global gNombres As String
Global Palabras(1 To 20) As String
Global gNumPalabras As Integer

Public Sub Separar_Apellidos_Nombres(pTexto As String, pOrden As Integer)
gApePat = ""
gApeMat = ""
gNombres = ""
If HayComaEnTexto(pTexto) = 1 Then Exit Sub
Separar_Palabras pTexto
If pOrden = 0 Then
  Obtener_Apellidos_Nombres
Else
  Obtener_Nombres_Apellidos
End If
End Sub

Public Function HayComaEnTexto(pTexto As String) As Integer
Dim l_strApellidos As String
On Error GoTo SalirFor
HayComaEnTexto = 0
For i = 1 To Len(pTexto)
  If Mid(pTexto, i, 1) = "," Then
    l_strApellidos = UCase(Trim(Mid(pTexto, 1, i - 1)))
    gNombres = UCase(Trim(Mid(pTexto, i + 1, Len(pTexto) - i)))
    Separar_Palabras l_strApellidos
    For j = 1 To gNumPalabras
      Select Case Palabras(j)
      Case "VDA.", "VIUDA"
        If gApeMat <> "" Then
          gApeMat = gApeMat & " " & Palabras(j)
          If j < gNumPalabras Then
            gApeMat = gApeMat & " " & Palabras(j + 1)
            j = j + 1
            If Palabras(j) = "DE" And j < gNumPalabras Then
              gApeMat = gApeMat & " " & Palabras(j + 1)
              j = j + 1
            End If
          End If
        Else
          gApePat = Trim(gApePat & " " & Palabras(j))
          If j < gNumPalabras Then
            gApePat = gApePat & " " & Palabras(j + 1)
            j = j + 1
            If Palabras(j) = "DE" And j < gNumPalabras Then
              gApePat = gApePat & " " & Palabras(j + 1)
              j = j + 1
            End If
          End If
        End If
      Case "Y"
        If gApeMat <> "" Then
          gApeMat = gApeMat & " " & Palabras(j)
          If j < gNumPalabras Then
            gApeMat = gApeMat & " " & Palabras(j + 1)
            j = j + 1
          End If
        Else
          gApePat = Trim(gApePat & " " & Palabras(j))
          If j < gNumPalabras Then
            gApePat = gApePat & " " & Palabras(j + 1)
            j = j + 1
          End If
        End If
      Case Else
        If gApePat = "" Then
          gApePat = Palabras(j)
        Else
          If gApePat = "DE" Or gApePat = "LA" Or _
             gApePat = "DE LA" Or gApePat = "VON" Or gApePat = "DEL" Then
            gApePat = gApePat & " " & Palabras(j)
          Else
            If gApeMat = "" Then
              gApeMat = Palabras(j)
            Else
              gApeMat = gApeMat & " " & Palabras(j)
            End If
          End If
        End If
      End Select
    Next j
    HayComaEnTexto = 1
    Exit For
  End If
Next i
Exit Function
SalirFor:
  HayComaEnTexto = 0
End Function

Public Sub Separar_Palabras(pTexto As String)
Dim l_StartWord As Long
On Error GoTo HayError
Erase Palabras
l_StartWord = 0
gNumPalabras = 0
For i = 1 To Len(pTexto)
  If Mid(pTexto, i, 1) = " " Then
    If l_StartWord > 0 Then
      gNumPalabras = gNumPalabras + 1
      Palabras(gNumPalabras) = UCase(Mid(pTexto, l_StartWord, i - l_StartWord))
      l_StartWord = 0
    End If
  Else
    If l_StartWord = 0 Then
      l_StartWord = i
    End If
  End If
Next i
If l_StartWord > 0 Then
  gNumPalabras = gNumPalabras + 1
  Palabras(gNumPalabras) = UCase(Mid(pTexto, l_StartWord, (Len(pTexto) + 1) - l_StartWord))
End If
Exit Sub
HayError:
  gNumPalabras = 0
  Erase Palabras
End Sub

Public Sub Obtener_Apellidos_Nombres()
On Error GoTo HayError
For i = 1 To gNumPalabras
  Select Case Palabras(i)
  Case "VDA.", "VIUDA"
    If gNombres = "" Then
      If gApeMat <> "" Then
        gApeMat = gApeMat & " " & Palabras(i)
        If i < gNumPalabras Then
          gApeMat = gApeMat & " " & Palabras(i + 1)
          i = i + 1
          If Palabras(i) = "DE" And i < gNumPalabras Then
            gApeMat = gApeMat & " " & Palabras(i + 1)
            i = i + 1
          End If
        End If
      Else
        gApePat = Trim(gApePat & " " & Palabras(i))
        If i < gNumPalabras Then
          gApePat = gApePat & " " & Palabras(i + 1)
          i = i + 1
          If Palabras(i) = "DE" And i < gNumPalabras Then
            gApePat = gApePat & " " & Palabras(i + 1)
            i = i + 1
          End If
        End If
      End If
    End If
  Case "Y"
    If gNombres <> "" Then
      gNombres = gNombres & " " & Palabras(i)
      If i < gNumPalabras Then
        gNombres = gNombres & " " & Palabras(i + 1)
        i = i + 1
      End If
    Else
      If gApeMat <> "" Then
        gApeMat = gApeMat & " " & Palabras(i)
        If i < gNumPalabras Then
          gApeMat = gApeMat & " " & Palabras(i + 1)
          i = i + 1
        End If
      Else
        If gApePat <> "" Then
          gApePat = gApePat & " " & Palabras(i)
        Else
          gApePat = Palabras(i)
        End If
        If i < gNumPalabras Then
          gApePat = gApePat & " " & Palabras(i + 1)
          i = i + 1
        End If
      End If
    End If
  Case Else
    If gApePat = "" Then
      gApePat = Palabras(i)
    Else
      If gApePat = "DE" Or gApePat = "LA" Or _
         gApePat = "DE LA" Or gApePat = "VON" Or gApePat = "DEL" Then
        gApePat = gApePat & " " & Palabras(i)
      Else
        If gApeMat = "" Then
          gApeMat = Palabras(i)
        Else
          If gApeMat = "DE" Or gApeMat = "LA" Or _
             gApeMat = "DE LA" Or gApeMat = "VON" Or gApeMat = "DEL" Then
            gApeMat = gApeMat & " " & Palabras(i)
          Else
            If gNombres = "" Then
              gNombres = Palabras(i)
            Else
              gNombres = gNombres & " " & Palabras(i)
            End If
          End If
        End If
      End If
    End If
  End Select
Next i
Exit Sub
HayError:
  gApePat = ""
  gApeMat = ""
  gNombres = ""
End Sub

Public Sub Obtener_Nombres_Apellidos()
On Error GoTo HayError
For i = gNumPalabras To 1 Step -1
  If gApeMat = "" Then
    gApeMat = Palabras(i)
  Else
    Select Case Palabras(i)
      Case "VDA.", "VIUDA"
        If gApePat <> "" Then
          gApePat = Palabras(i) & " " & gApePat
        Else
          gApeMat = Trim(Palabras(i) & " " & gApeMat)
        End If
      Case "DE", "LA", "LOS", "VON", "Y", "LAS", "DEL"
        If gNombres <> "" Then
           gNombres = Palabras(i) & " " & gNombres
           If Palabras(i) = "Y" And i > 1 Then
             gNombres = Palabras(i - 1) & " " & gNombres
             i = i - 1
           End If
        Else
           If gApePat <> "" Then
             gApePat = Palabras(i) & " " & gApePat
             If Palabras(i) = "Y" And i > 1 Then
               gApePat = Palabras(i - 1) & " " & gApePat
               i = i - 1
             End If
           Else
             gApeMat = Palabras(i) & " " & gApeMat
             If Palabras(i) = "Y" And i > 1 Then
               gApeMat = Palabras(i - 1) & " " & gApeMat
               i = i - 1
             End If
           End If
        End If
    Case Else
      If gApePat = "" Then
        gApePat = Palabras(i)
      Else
        If gNombres = "" Then
          gNombres = Palabras(i)
        Else
          gNombres = Palabras(i) & " " & gNombres
        End If
      End If
    End Select
  End If
Next i
Exit Sub
HayError:
  gApePat = ""
  gApeMat = ""
  gNombres = ""
End Sub





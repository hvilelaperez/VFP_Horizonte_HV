Attribute VB_Name = "FuncionFecha"
Option Explicit
'***************************************************************************
'*  La funcion VALIDAFECHA debe ir en el evento LOST_focus del TEXTBOX,
'*  junto con la funcion FVFECHA que ira en el evento KEYPRESS del TEXTBOX.
'*  CAJA es la caja de texto que contiene el dato de tipo fecha.
'***************************************************************************
Public Function ValidaFecha(Caja As Object)
    If Trim(Caja.Text) <> "//" And Trim(Caja.Text) <> "" Then
        If IsDate(Caja.Text) And Len(Caja.Text) = 10 Then
            If Len(Caja) < 10 Then
                If Mid(Caja, 2, 1) = "/" Then
                    Caja.Text = "0" + Caja.Text
                End If
                If Mid(Caja, 5, 1) = "/" Then
                    Caja.Text = Mid(Caja, 1, 3) + "0" + Mid(Caja, 4, 6)
                End If
            End If
        Else
            MsgBox "Fecha no Valida"
            Caja.SetFocus
        End If
    End If
End Function
'********************************************************************
'*  Esta funcion debe ir en el evento KEYPRESS del TEXTBOX.
'*  Caja = TEXTBOX
'*  KEY = KEYCODE
'*  EJEMPLO: FVFecha (TEXT1, KEYASCII)
'******************************************************************
Public Function FVFecha(Caja As TextBox, key As Integer)
    If (key < Asc("0") Or key > Asc("9")) And (key <> 8) Then
       key = 0
       Exit Function
    End If
    If Caja.SelLength = Len(Caja) Then Caja.Text = "": Caja.SelStart = 1
    If Len(Trim(Caja)) = 0 Or Trim(Caja.Text) = "/" Then
       Caja.Text = "//"
    End If
    If Caja.SelStart = 2 And key <> 8 Then
       Caja.SelStart = Caja.SelStart + 1
       Caja.MaxLength = 10
    End If
    If Caja.SelStart = 5 And key <> 8 Then
       Caja.SelStart = Caja.SelStart + 1
       Caja.MaxLength = 10
    End If
    '******************
    BackSpace Caja, key
    '******************
End Function

Public Sub Suprimir(Caja As TextBox, key As Integer)
    If key = 46 Then
        If Mid(Caja, Caja.SelStart + 1, 1) = "/" Then key = 0: Caja.SelStart = Caja.SelStart + 1
        Recorrer Caja, key
    End If
End Sub

Private Sub BackSpace(Caja As TextBox, key As Integer)
    If key = 8 Then
        If Caja.SelStart >= 1 Then
            If Mid(Caja, Caja.SelStart, 1) = "/" Then key = 0: SendKeys "{LEFT}"
        End If
        If Caja.Text = "//" Then key = 0
        Recorrer Caja, key
    End If
End Sub

Private Function Recorrer(Caja As TextBox, key As Integer)
    Dim i As Integer
    If Caja.SelLength > 0 Then
        For i = 1 To Caja.SelLength
            If Mid(Trim(Caja.SelText), i, 1) = "/" Then key = 0: Exit For
        Next
    End If
End Function



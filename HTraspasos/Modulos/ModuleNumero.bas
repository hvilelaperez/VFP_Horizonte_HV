Attribute VB_Name = "ModuleNumero"
Option Explicit
'****************************************************************************************************************************************
'* Este procedimiento permite ingresar un número en un campo texto, teniendo
'* en cuenta la cantidad de enteros y de decimales que tiene
'* Este procedimiento debe ser llamado en el evento keypress del texto.
'* Caja : nombre del texbox; key : "keyascii"
'* Ent : La cantidad de enteros que tiene este texto
'* Dec   : La cantidad de decimales que tiene este tecto
'****************************************************************************************************************************************
Public Function FEntDec(Caja As Control, key As Integer, ent As Integer, dec As Integer)
    If (key < 48 Or key > 57) And (key <> 8 And key <> 46) Then key = 0: Exit Function 'Restringe teclas
    If (key = 46) Then If (InStr(Caja.Text, ".") > 0) Or dec = 0 Then key = 0: Exit Function 'Valida un solo Punto
    Call BackSpacePunto(Caja, key, ent, dec)
    'No permite colocar el punto cuando supera el número de decimales
    If key = 46 And Len(Right(Caja, Len(Caja) - Caja.SelStart)) > dec Then key = 0: Exit Function
    'Verifica si el número ingresado es parte entera o decimal
    If (InStr(Caja.Text, ".") = 0 Or Caja.SelStart < InStr(Caja.Text, ".")) Then
            If key <> 8 And key <> 46 And FCEntero(Caja.Text, dec) = ent And InStr(Caja, ".") = 0 And dec <> 0 Then
                SendKeys "{End}": SendKeys (".")
            Else
                If key <> 8 And key <> 46 And FCEntero(Caja.Text, dec) >= ent And InStr(Caja, ".") = 0 And dec <> 0 And Caja.SelStart = ent Then
                    SendKeys "{End}": SendKeys ("."): SendKeys (Chr(key)): key = 0: Exit Function
                Else: If key <> 8 And key <> 46 And FCEntero(Caja.Text, dec) >= ent Then key = 0
                End If
            End If
    Else: If key <> 8 And Len(Right(Caja, Len(Caja) - InStr(Caja, "."))) >= dec Then key = 0
    End If 'Len(Right(Caja, Len(Caja) - InStr(Caja, "."))) Devuelve los decimales de un número
End Function
'Cuenta la parte entera de un numero
Private Function FCEntero(Caja As String, dec As Integer) As Integer
    If InStr(Caja, ".") > 0 Then
        FCEntero = Len(Mid(Caja, 1, InStr(Caja, ".") - 1))
        Else: FCEntero = IIf(dec > 0, Len(Trim(Caja)) + 1, Len(Trim(Caja)))
    End If
End Function
'Sirve para Suprimir con la tecla SUPR; debe ser llamado en el evento KeyDown del texto
Public Sub SuprimirPunto(Caja As Control, key As Integer, ent As Integer, dec As Integer)
If key = 46 And Mid(Caja, Caja.SelStart + 1, 1) = "." Then
    Caja.SelStart = IIf(Not ((FCEntero(Caja.Text, dec) + Len(Right(Caja, Len(Caja) - InStr(Caja, ".")))) <= ent), Caja.SelStart + 1, Caja.SelStart)
End If
End Sub
'Sirve para Borrar con la tecla BACKSPACE
Private Sub BackSpacePunto(Caja As Control, key As Integer, ent As Integer, dec As Integer)
    If key = 8 And Caja.SelStart > 0 Then
        If Mid(Caja, Caja.SelStart, 1) = "." Then
            Caja.SelStart = IIf(Not ((FCEntero(Caja.Text, dec) + Len(Right(Caja, Len(Caja) - InStr(Caja, ".")))) <= ent), Caja.SelStart - 1, Caja.SelStart)
        End If
    End If
End Sub

# Función Número en Letras

Convierte un número en su representación textual en inglés (dólares y centavos).

## Código VBA

```vba
' Autor: Henry Vilela - hvilelaperez@gmail.com
Function NumeroEnLetras(ByVal MiNumero)
    Dim Dolares, Centavos, Temp
    Dim PosicionDecimal, Contador
    ReDim Lugar(9) As String
    Lugar(2) = " Thousand "
    Lugar(3) = " Million "
    Lugar(4) = " Billion "
    Lugar(5) = " Trillion "

    ' Representación en cadena del monto
    MiNumero = Trim(Str(MiNumero))

    ' Posición del punto decimal, 0 si no existe
    PosicionDecimal = InStr(MiNumero, ".")

    ' Convertir centavos y ajustar MiNumero al monto en dólares
    If PosicionDecimal > 0 Then
        Centavos = ObtenerDecenas(Left(Mid(MiNumero, PosicionDecimal + 1) & _
                  "00", 2))
        MiNumero = Trim(Left(MiNumero, PosicionDecimal - 1))
    End If

    Contador = 1
    Do While MiNumero <> ""
        Temp = ObtenerCentenas(Right(MiNumero, 3))
        If Temp <> "" Then Dolares = Temp & Lugar(Contador) & Dolares
        If Len(MiNumero) > 3 Then
            MiNumero = Left(MiNumero, Len(MiNumero) - 3)
        Else
            MiNumero = ""
        End If
        Contador = Contador + 1
    Loop

    Select Case Dolares
        Case ""
            Dolares = "No Dollars"
        Case "One"
            Dolares = "One Dollar"
        Case Else
            Dolares = Dolares & " Dollars"
    End Select

    Select Case Centavos
        Case ""
            Centavos = " and No Cents"
        Case "One"
            Centavos = " and One Cent"
        Case Else
            Centavos = " and " & Centavos & " Cents"
    End Select

    NumeroEnLetras = Dolares & Centavos
End Function

' Convierte un número de 100-999 a texto
Function ObtenerCentenas(ByVal MiNumero)
    Dim Resultado As String
    If Val(MiNumero) = 0 Then Exit Function
    MiNumero = Right("000" & MiNumero, 3)

    ' Convertir las centenas
    If Mid(MiNumero, 1, 1) <> "0" Then
        Resultado = ObtenerDigito(Mid(MiNumero, 1, 1)) & " Hundred "
    End If

    ' Convertir decenas y unidades
    If Mid(MiNumero, 2, 1) <> "0" Then
        Resultado = Resultado & ObtenerDecenas(Mid(MiNumero, 2))
    Else
        Resultado = Resultado & ObtenerDigito(Mid(MiNumero, 3))
    End If

    ObtenerCentenas = Resultado
End Function

' Convierte un número de 10 a 99 a texto
Function ObtenerDecenas(TensText)
    Dim Resultado As String
    Resultado = ""

    ' Si el valor está entre 10-19...
    If Val(Left(TensText, 1)) = 1 Then
        Select Case Val(TensText)
            Case 10: Resultado = "Ten"
            Case 11: Resultado = "Eleven"
            Case 12: Resultado = "Twelve"
            Case 13: Resultado = "Thirteen"
            Case 14: Resultado = "Fourteen"
            Case 15: Resultado = "Fifteen"
            Case 16: Resultado = "Sixteen"
            Case 17: Resultado = "Seventeen"
            Case 18: Resultado = "Eighteen"
            Case 19: Resultado = "Nineteen"
            Case Else
        End Select
    Else
        ' Si el valor está entre 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Resultado = "Twenty "
            Case 3: Resultado = "Thirty "
            Case 4: Resultado = "Forty "
            Case 5: Resultado = "Fifty "
            Case 6: Resultado = "Sixty "
            Case 7: Resultado = "Seventy "
            Case 8: Resultado = "Eighty "
            Case 9: Resultado = "Ninety "
            Case Else
        End Select
        Resultado = Resultado & ObtenerDigito _
            (Right(TensText, 1))
    End If

    ObtenerDecenas = Resultado
End Function

' Convierte un número de 1 a 9 a texto
Function ObtenerDigito(Digit)
    Select Case Val(Digit)
        Case 1: ObtenerDigito = "One"
        Case 2: ObtenerDigito = "Two"
        Case 3: ObtenerDigito = "Three"
        Case 4: ObtenerDigito = "Four"
        Case 5: ObtenerDigito = "Five"
        Case 6: ObtenerDigito = "Six"
        Case 7: ObtenerDigito = "Seven"
        Case 8: ObtenerDigito = "Eight"
        Case 9: ObtenerDigito = "Nine"
        Case Else: ObtenerDigito = ""
    End Select
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función principal** | `NumeroEnLetras` - Convierte número a texto en inglés (dólares) |
| **Función auxiliar** | `ObtenerCentenas` - Convierte números de 100-999 |
| **Función auxiliar** | `ObtenerDecenas` - Convierte números de 10-99 |
| **Función auxiliar** | `ObtenerDigito` - Convierte dígitos de 1-9 |
| **Escalas** | Soporta hasta trillones ( Thousand, Million, Billion, Trillion) |
| **Formato** | Incluye "Dollars" y "Cents" en el resultado |
| **Uso** | `=NumeroEnLetras(A1)` para convertir un número a texto legible |

---

*Módulo: 31_CodigosGenerales | Archivo: 43 - Ham doc so thanh chu tieng anh.txt*

# Uso de FirstIndex con Expresiones Regulares

Función que utiliza la propiedad `FirstIndex` de las expresiones regulares para encontrar la posición de la primera coincidencia en un texto.

## Código VBA

```vba
Function MYMATCH(strValue As String, strPattern As String, Optional blnCase As Boolean = True, Optional blnBoolean = True) As String
    Dim objRegEx As Object
    Dim strPosition As Integer
    Dim RegMC

    ' Crear expresión regular.
    Set objRegEx = CreateObject("VBScript.RegExp")
    With objRegEx
        .Pattern = strPattern
        .IgnoreCase = blnCase
        If .test(strValue) Then
            Set RegMC = .Execute(strValue)
            MYMATCH = RegMC(0).firstindex + 1
        Else
            MYMATCH = "sin coincidencia"
        End If
    End With
End Function

Sub TestMe()
    MsgBox MYMATCH("test 1", "\d+")
End Sub
```

## Descripción

| Elemento | Descripción |
|---|---|
| `MYMATCH` | Función que retorna la posición (1-based) de la primera coincidencia del patrón en el texto de entrada |
| `strValue` | Cadena de texto donde se busca el patrón |
| `strPattern` | Expresión regular a buscar (ej: `\d+` para dígitos) |
| `blnCase` | Opcional. Indica si diferencia mayúsculas/minúsculas (True = sensible) |
| `FirstIndex` | Propiedad del objeto Match que indica la posición (0-based) de la coincidencia; se suma 1 para convertir a posición 1-based |
| `TestMe` | Sub de prueba que busca el primer número en "test 1" (retorna posición 6) |

---

*Módulo: 25 - VBA Regular expression | Archivo: 01 - Ung dung FirstIndex cua Regular Expression.txt*

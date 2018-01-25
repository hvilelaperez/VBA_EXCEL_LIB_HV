# Separar Números de Cadena

Extrae números o texto de una cadena usando expresiones regulares.

## Código VBA

```vba
Function ObtenerNumeroCadena(txt As String, obtenerTexto As Boolean) As String
'Si obtenerTexto = true o 1 --> Obtiene solo texto
'Si obtenerTexto = false o 0 --> Obtiene solo números

    With CreateObject("VBScript.RegExp")
    .Pattern = IIf(obtenerTexto = True, "\d+", "\D+")
    .Global = True
    ObtenerNumeroCadena = .Replace(txt, "")
    End With
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `VBScript.RegExp` | Objeto de expresión regular |
| `\d+` | Patrón para dígitos (números) |
| `\D+` | Patrón para no-dígitos (letras) |
| `Global` | Reemplaza todas las ocurrencias |
| `obtenerTexto` | True = extrae texto, False = extrae números |

### Ejemplo de uso

```vba
' Para obtener solo números:
resultado = ObtenerNumeroCadena("ABC123XYZ", False)  ' Retorna "123"

' Para obtener solo texto:
resultado = ObtenerNumeroCadena("ABC123XYZ", True)   ' Retorna "ABCXYZ"
```

---

*Módulo: 31_CodigosGenerales | Archivo: 24_TachNumerosCadena.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

# Separar Números y Letras con RegExp

Función que utiliza expresiones regulares para extraer solo números o solo letras de una cadena de texto.

## Código VBA

```vba
Function GetStringNumber(txt As String, getStr As Boolean) As String
'Si getStr = True o 1 --> Obtener solo letras
'Si getStr = False o 0 --> Obtener solo números

    With CreateObject("VBScript.RegExp")
    .Pattern = IIf(getStr = True, "\d+", "\D+")
    .Global = True
    GetStringNumber = .Replace(txt, "")
    End With
End Function
```

## Descripción

| Elemento | Descripción |
|---|---|
| `GetStringNumber` | Función que separa números de letras en una cadena |
| `txt` | Cadena de entrada a procesar |
| `getStr` | Parámetro booleano que determina qué extraer |
| `\d+` | Patrón que coincide con uno o más dígitos (0-9) |
| `\D+` | Patrón que coincide con uno o más caracteres no numéricos |
| `True` | Si es True, elimina dígitos y retorna solo letras |
| `False` | Si es False, elimina letras y retorna solo números |

---

*Módulo: 25 - VBA Regular expression | Archivo: 03 - Ung dung RegExp de tach chu a so.txt*

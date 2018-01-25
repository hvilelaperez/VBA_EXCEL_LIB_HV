# Obtener Color de Rango

Obtiene el color de fondo de celdas en diferentes formatos (HEX, RGB, IDX).

## Código VBA

```vba
Public Function iColor(rng As Range, Optional formatType As String) As Variant
'formatType: Hex para #RRGGBB, RGB para (R, G, B) e IDX para ColorIndex de VBA
    Dim colorVal As Variant
    colorVal = rng.DisplayFormat.Interior.Color
    Select Case UCase(formatType)
        Case "HEX"
            iColor = "#" & Format(Hex(colorVal Mod 256),"00") & _
                           Format(Hex((colorVal \ 256) Mod 256),"00") & _
                           Format(Hex((colorVal \ 65536)),"00")
        Case "RGB"
            iColor = Format((colorVal Mod 256),"00") & ", " & _
                     Format(((colorVal \ 256) Mod 256),"00") & ", " & _
                     Format((colorVal \ 65536),"00")
        Case "IDX"
            iColor = rng.Interior.ColorIndex
        Case Else
            iColor = colorVal
    End Select
End Function

'Ejemplo de uso de la función iColor
Sub Obtener_Formatos_Color()
    Dim rng As Range

    For Each rng In Selection.Cells
        rng.Offset(0, 1).Value = iColor(rng, "HEX")
        rng.Offset(0, 2).Value = iColor(rng, "RGB")
    Next
End Sub
```

## Descripción

| Parámetro | Valores | Formato Resultado |
|-----------|---------|-------------------|
| `"HEX"` | #RRGGBB | `#FF0000` (rojo) |
| `"RGB"` | R, G, B | `255, 0, 0` (rojo) |
| `"IDX"` | Número | `3` (ColorIndex de VBA) |
| Otro | Valor numérico | `16711680` (valor Long) |

### Funciones utilizadas

| Función | Descripción |
|---------|-------------|
| `DisplayFormat.Interior.Color` | Obtiene el color visible (incluye formato condicional) |
| `Interior.Color` | Obtiene el color de la celda |
| `Hex()` | Convierte número a hexadecimal |
| `Format()` | Formatea el resultado con ceros a la izquierda |

---

*Módulo: 31_CodigosGenerales | Archivo: 27_ObtenerColorRango.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

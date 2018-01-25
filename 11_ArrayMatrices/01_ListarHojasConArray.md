# Listar Nombres de Hojas con Array

Código VBA para obtener los nombres de todas las hojas de un libro y listarlos en un rango usando un array.

## Código VBA

```vba
Sub QUERY_NAME_SHEET()

Dim sh As Worksheet
Dim arr As Variant, i As Integer

ReDim arr(1 To ThisWorkbook.Worksheets.Count)
i = 1
For Each sh In Worksheets
    arr(i) = sh.Name
    i = i + 1
Next sh

If i = 1 Then Exit Sub
Sheet2.Range("H2").Resize(UBound(arr), 1).Value = Application.Transpose(arr)
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `arr` | Array que almacena los nombres de las hojas |
| `ReDim` | Redimensiona el array según el número de hojas |
| `Application.Transpose` | Convierte el array horizontal a vertical para volcarlo |

### Proceso

1. Cuenta el número total de hojas en el libro
2. Crea un array dimensionado exactamente
3. Recorre cada hoja y almacena su nombre
4. Usa `Transpose` para volcar los datos verticalmente en la celda H2 de Sheet2

### Ejemplo de salida

| Celda | Valor |
|-------|-------|
| H2 | Nombre de Hoja 1 |
| H3 | Nombre de Hoja 2 |
| H4 | Nombre de Hoja 3 |

---

*Módulo: 05 - VBA Array | Archivo: 00 - Su dung mang de liet ke ten sheets*
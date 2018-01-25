# Función Verificar Existencia de Hoja

Función que verifica si una hoja específica existe en el libro de trabajo.

## Código VBA

```vba
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `sheetExists` | Función que retorna `True` si la hoja fue encontrada |
| `sheetToFind` | Nombre de la hoja que se desea buscar |
| Iteración | Recorre todas las hojas (`Worksheets`) del libro activo |
| `Exit Function` | Sale inmediatamente al encontrar la hoja, optimizando rendimiento |

---

*Módulo: 06 - VBA Worksheets | Archivo: 08 - Chevk if SheetExists.txt*

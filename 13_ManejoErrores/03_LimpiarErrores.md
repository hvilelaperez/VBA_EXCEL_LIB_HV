# Limpiar Errores con Err.Clear

Uso de `Err.Clear` para limpiar el objeto de error después de manejarlo.

## Código VBA

```vba
Private Function FirstEmptyRow() As Long

    Dim Clm As Long

    With ActiveSheet
        On Error GoTo ErrHandler
        Clm = WorksheetFunction.Match("Style", .Rows(3), 0)
        FirstEmptyRow = .Cells(.Rows.Count, Clm).End(xlUp).Row + 1
    End With
ErrHandler:
    Err.Clear
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `FirstEmptyRow` | Función que retorna la primera fila vacía en una columna |
| `WorksheetFunction.Match` | Busca el valor "Style" en la fila 3 de la hoja activa |
| `On Error GoTo ErrHandler` | Redirige a la etiqueta de manejo si `Match` falla |
| `Err.Clear` | Limpia las propiedades del objeto `Err` eliminando el error registrado |
| `End(xlUp).Row + 1` | Calcula la siguiente fila vacía después de la última con datos |

---

*Módulo: 07 - VBA Error Handler | Archivo: 02 - Loai bo error bang bien phap clear.txt*

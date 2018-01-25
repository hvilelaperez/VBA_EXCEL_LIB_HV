# Colorear Celdas Activas

Evento de hoja que resalta la fila y columna de la celda activa con un color de fondo amarillo.

## Código VBA

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$D$2" Or Target.Address = "$F$2" Then Call Report2
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Resaltar fila y columna de la celda activa

    'Exit Sub   '<---------Quitar esta línea para activar el código

    Dim iRow As Long, iCol As Long, i As Long, j As Long
    Application.ScreenUpdating = False

    Cells.Interior.ColorIndex = 0   'Eliminar color de fondo anterior
    iRow = ActiveCell.Row           'Obtener fila de la celda actual'
    iCol = ActiveCell.Column        'Obtener columna de la celda actual'

    For i = 1 To iRow
        Cells(i, iCol).Interior.ColorIndex = 6  'Colorear celdas de la misma columna'
    Next i

    For j = 1 To iCol
        Cells(iRow, j).Interior.ColorIndex = 6  'Colorear celdas de la misma fila'
    Next j

    Application.ScreenUpdating = True
End Sub
```

## Descripción

| Evento | Descripción |
|--------|-------------|
| `Worksheet_Change` | Se ejecuta cuando cambia el valor de una celda; llama a Report2 si se modifican D2 o F2 |
| `Worksheet_SelectionChange` | Se ejecuta cuando se cambia la selección de celda; resalta fila y columna |

| Elemento técnico | Descripción |
|------------------|-------------|
| `ColorIndex = 0` | Elimina el color de fondo de todas las celdas |
| `ColorIndex = 6` | Aplica color amarillo (índice 6 en la paleta de Excel) |
| `ActiveCell.Row/Column` | Obtiene las coordenadas de la celda actualmente seleccionada |

> **Nota:** Para activar el resaltado, eliminar o comentar la línea `Exit Sub`. Este evento se ejecuta en cada cambio de selección, lo que puede afectar el rendimiento en hojas grandes.

---

*Módulo: 31_CodigosGenerales | Archivo: 10_ColorearCeldasActivas.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

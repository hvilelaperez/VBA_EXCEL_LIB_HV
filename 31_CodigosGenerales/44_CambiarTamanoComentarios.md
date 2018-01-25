# Cambiar Tamaño de Comentarios

Macro para ajustar automáticamente el tamaño de todos los comentarios en una hoja de cálculo.

## Código VBA

```vba
Sub AjustarComentarios1()
    Application.ScreenUpdating = False
    Dim x As Range, y As Long
    For Each x In Cells.SpecialCells(xlCellTypeComments)
        Select Case True
        Case Len(x.NoteText) <> 0
            With x.Comment
                .Shape.TextFrame.AutoSize = True
                If .Shape.Width > 250 Then
                    y = .Shape.Width * .Shape.Height
                    .Shape.Width = 150
                    .Shape.Height = (y / 200) * 1.3
                End If
            End With
        End Select
    Next x
    Application.ScreenUpdating = True
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `ScreenUpdating = False` | Desactiva la actualización de pantalla para mayor velocidad |
| `SpecialCells(xlCellTypeComments)` | Selecciona todas las celdas que tienen comentarios |
| `x.NoteText` | Obtiene el texto del comentario |
| `.Shape.TextFrame.AutoSize = True` | Ajusta automáticamente el tamaño del cuadro de texto |
| `.Shape.Width` | Ancho actual del comentario |
| `.Shape.Height` | Alto actual del comentario |

### Lógica de ajuste

1. Recorre todas las celdas con comentarios en la hoja
2. Verifica que el comentario no esté vacío
3. Si el ancho del comentario supera 250 puntos:
   - Calcula el área actual (ancho × alto)
   - Establece el ancho a 150 puntos
   - Recalcula el alto manteniendo proporción con un factor de 1.3

### Fórmula de cálculo

```
Área original = Ancho × Alto
Nuevo Alto = (Área original / 200) × 1.3
```

### Parámetros de ajuste

| Valor | Descripción |
|-------|-------------|
| 250 | Umbral máximo de ancho antes del ajuste |
| 150 | Nuevo ancho fijo después del ajuste |
| 200 | Divisor para recalcular el alto |
| 1.3 | Factor de incremento para el alto |

---

*Módulo: 96 - VBA Codes general | Archivo: 40 - Thay doi kich thuoc cua comments.txt*

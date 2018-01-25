# Cambiar Tamaño y Alinear Contenido de Comentarios

Macro para ajustar el tamaño de los comentarios y alinear su contenido desde la esquina superior izquierda.

## Código VBA

```vba
Sub AjustarComentarios2()
    Application.ScreenUpdating = False
    Dim x As Range, y As Long
    For Each x In Cells.SpecialCells(xlCellTypeComments)
        Select Case True
        Case Len(x.NoteText) <> 0
            With x.Comment
                .Shape.TextFrame.AutoSize = True
                If .Shape.Width > 250 Then
                    y = .Shape.Width * .Shape.Height
                    .Shape.ScaleHeight 0.9, msoFalse, msoScaleFromTopLeft
                    .Shape.ScaleWidth 1#, msoFalse, msoScaleFromTopLeft
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
| `ScreenUpdating = False` | Desactiva la actualización de pantalla para mayor rendimiento |
| `SpecialCells(xlCellTypeComments)` | Obtiene todas las celdas con comentarios |
| `x.NoteText` | Obtiene el texto del comentario |
| `.Shape.TextFrame.AutoSize = True` | Ajusta automáticamente el tamaño del cuadro |
| `.Shape.ScaleHeight` | Escala la altura de la forma |
| `.Shape.ScaleWidth` | Escala el ancho de la forma |
| `msoScaleFromTopLeft` | Escala desde la esquina superior izquierda |

### Diferencia con AjustarComentarios1

| Característica | AjustarComentarios1 | AjustarComentarios2 |
|----------------|---------------------|---------------------|
| Método de ajuste | Cambia dimensiones directamente | Escala proporcionalmente |
| Ancho resultante | Fijo a 150 puntos | 100% del original |
| Alto resultante | Recalculado con fórmula | 90% del original |
| Alineación | Predeterminada | Esquina superior izquierda |

### Parámetros de escalado

| Valor | Descripción |
|-------|-------------|
| 0.9 | Factor de escala para altura (90% del original) |
| 1# | Factor de escala para ancho (100% del original, sin cambio) |
| `msoFalse` | No efecto de transformación |
| `msoScaleFromTopLeft` | Punto de origen de la escala |

### Flujo de ejecución

1. Se desactiva la actualización de pantalla
2. Se recorren todos los comentarios de la hoja
3. Se ajusta automáticamente el tamaño del cuadro de texto
4. Si el ancho supera 250 puntos:
   - Se guarda el área actual
   - Se reduce la altura al 90%
   - Se mantiene el ancho al 100%
   - La escala se aplica desde la esquina superior izquierda
5. Se reactiva la actualización de pantalla

---

*Módulo: 96 - VBA Codes general | Archivo: 41 - Thay doi kich thuoc va canh giua noi dung comment.txt*

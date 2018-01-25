# Redimensionar Imágenes para Ajustar a Celdas

Ajusta el tamaño de las filas y columnas para que coincidan con las dimensiones de las imágenes insertadas en la hoja.

## Código VBA

```vba
Sub RedimensionarImagenesCeldas()
    'Dim Imagen As Variant
    For Each Imagen In ActiveSheet.DrawingObjects
        PictureTop = Imagen.Top
        PictureLeft = Imagen.Left
        PictureHeight = Imagen.Height
        PictureWidth = Imagen.Width
        For N = 2 To 256
            If Columns(N).Left > PictureLeft Then
                PictureColumn = N - 1
                Exit For
            End If
        Next N
        For N = 2 To 65536
            If Rows(N).Top > PictureTop Then
                PictureRow = N - 1
                Exit For
            End If
        Next N
        Rows(PictureRow).RowHeight = PictureHeight
        Columns(PictureColumn).ColumnWidth = PictureWidth * (54.29 / 288)
        Imagen.Top = Cells(PictureRow, PictureColumn).Top
        Imagen.Left = Cells(PictureRow, PictureColumn).Left
    Next Imagen
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Recorrido** | Itera sobre todos los objetos de dibujo (imágenes) de la hoja activa |
| **Ubicación** | Determina qué columna y fila contiene cada imagen |
| **Ajuste de fila** | Modifica la altura de la fila para igualar la altura de la imagen |
| **Ajuste de columna** | Modifica el ancho de la columna con factor de conversión (54.29/288) |
| **Reubicación** | Mueve la imagen para que se ajuste perfectamente a la celda |

| Cálculo | Fórmula |
|---------|---------|
| **Ancho de columna** | `PictureWidth * (54.29 / 288)` - factor de conversión de píxeles a puntos |
| **Fila objetivo** | Primera fila cuya posición vertical supere la imagen |
| **Columna objetivo** | Primera columna cuya posición horizontal supere la imagen |

---

*Módulo: 12 - VBA Images | Archivo: 01 - ResizePictureCells*

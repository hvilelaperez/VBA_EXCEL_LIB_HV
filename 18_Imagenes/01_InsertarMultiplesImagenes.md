# Insertar Múltiples Imágenes en Celdas

Permite seleccionar múltiples imágenes desde el explorador de archivos e insertarlas en celdas consecutivas de una columna.

## Código VBA

```vba
Sub InsertarMultiplesImagenes()
    Dim ListaImagenes() As Variant, lLoop As Variant
    Dim FormatoImagen As String, xIndiceColumna As Integer, xIndiceFila As Integer
    Dim Rng As Range
    Dim sShape As Shape
    On Error Resume Next
    ListaImagenes = Application.GetOpenFilename(FormatoImagen, MultiSelect:=True)
    xIndiceColumna = Application.ActiveCell.Column
    If IsArray(ListaImagenes) Then
        xIndiceFila = Application.ActiveCell.Row
        For lLoop = LBound(ListaImagenes) To UBound(ListaImagenes)
            Set Rng = Cells(xIndiceFila, xIndiceColumna)
            Set sShape = ActiveSheet.Shapes.AddPicture(ListaImagenes(lLoop), msoFalse, msoCTrue, Rng.Left, Rng.Top, Rng.Width, Rng.Height)
            xIndiceFila = xIndiceFila + 1
        Next
    End If
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Selección** | Utiliza `GetOpenFilename` con `MultiSelect:=True` para elegir varias imágenes |
| **Posición inicial** | Comienza desde la celda activa (fila y columna actual) |
| **Insertado** | Cada imagen se coloca en la celda correspondiente usando `Shapes.AddPicture` |
| **Distribución** | Las imágenes se apilan verticalmente, una por fila |

| Parámetro de AddPicture | Valor |
|------------------------|-------|
| Filename | Ruta completa de la imagen |
| LinkToFile | msoFalse (incrusta la imagen) |
| SaveWithDocument | msoCTrue (guarda con el documento) |
| Left/Top | Posición de la celda destino |
| Width/Height | Dimensiones de la celda destino |

---

*Módulo: 12 - VBA Images | Archivo: 00 - InsertMultiImage*

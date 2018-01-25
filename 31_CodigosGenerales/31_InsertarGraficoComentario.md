# Insertar Gráfico en Comentario

Macro que exporta un gráfico existente como imagen temporal y lo inserta dentro de un comentario de celda.

## Código VBA

```vba
' Autor: Henry Vilela - hvilelaperez@gmail.com
Sub InsertarGrafico()
    Dim x As String, z As Range
    Application.ScreenUpdating = False

    ' Ruta temporal para almacenar la imagen del gráfico
    ' Modificar según sea necesario
    x = "C:\GraficoTemporal.gif"

    ' Celda donde se insertará el comentario con el gráfico
    Set z = Worksheets("ChartInComment").Range("A3")

    ' Eliminar comentario existente si lo hay
    On Error Resume Next
    z.Comment.Delete
    On Error GoTo 0

    ' Seleccionar y exportar el gráfico
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Export x

    ' Crear comentario nuevo, ajustar tamaño e insertar gráfico como imagen
    With z.AddComment
        With .Shape
            .Height = 322
            .Width = 465
            .Fill.UserPicture x
        End With
    End With

    ' Eliminar archivo de imagen temporal
    Kill x
    Range("A1").Activate
    Application.ScreenUpdating = True
    Set z = Nothing
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Procedimiento** | `InsertarGrafico` - Inserta un gráfico de Excel dentro de un comentario |
| **Proceso** | Exporta gráfico como GIF → crea comentario → inserta imagen → elimina archivo temporal |
| **Celda objetivo** | Modificar `Worksheets("ChartInComment").Range("A3")` según necesidad |
| **Gráfico** | Modificar `"Chart 1"` por el nombre de tu gráfico |
| **Dimensiones** | Alto: 322 píxeles, Ancho: 465 píxeles (ajustable) |
| **Uso** | Visualizar gráficos dentro de comentarios sin ocupar espacio en la hoja |

---

*Módulo: 31_CodigosGenerales | Archivo: 42 - Dua chart vao comment.txt*

# Colores de Celdas a Gráfico (Múltiples Series)

Macro avanzada para aplicar colores de celdas a múltiples series de un gráfico, incluyendo tanto el relleno como el borde de los puntos.

## Código VBA

```vba
Sub CellColorsToChart()

    Dim xChart As Chart
    Dim I As Long, J As Long
    Dim xRowsOrCols As Long, xSCount As Long
    Dim xRg As Range, xCell As Range
    On Error Resume Next
    Set xChart = ActiveSheet.ChartObjects("Graphique 6").Chart
    If xChart Is Nothing Then Exit Sub
    xSCount = xChart.SeriesCollection.count
    For I = 1 To xSCount
        J = 1
        With xChart.SeriesCollection(I)
            Set xRg = ActiveSheet.Range(Split(Split(.Formula, ",")(2), "!")(1))
            If xSCount > 4 Then
                xRowsOrCols = xRg.Columns.count
            Else
                xRowsOrCols = xRg.Rows.count
            End If
            For Each xCell In xRg
                .Points(J).Format.Fill.ForeColor.RGB = ThisWorkbook.Colors(xCell.Interior.ColorIndex)
                .Points(J).Format.Line.ForeColor.RGB = ThisWorkbook.Colors(xCell.Interior.ColorIndex)
                J = J + 1
            Next
        End With
    Next

End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función** | Colorea múltiples series de un gráfico según los colores de las celdas de origen |
| **Objetivo** | Aplicar colores tanto al relleno como al borde de cada punto en todas las series |
| **Requisitos** | Un gráfico llamado "Graphique 6" en la hoja activa |
| **Método** | Recorre todas las series y aplica el color de cada celda al punto correspondiente |

### Pasos del proceso

1. Obtiene referencia al gráfico "Graphique 6" en la hoja activa
2. Cuenta el número de series en el gráfico
3. Para cada serie, extrae el rango de datos de la fórmula
4. Determina si recorrer por columnas (más de 4 series) o por filas
5. Para cada celda, aplica el color de fondo al relleno y al borde del punto

### Notas

- Si hay más de 4 series, se recorre por columnas; de lo contrario, por filas
- Se colorea tanto el relleno (Fill) como el borde (Line) de cada punto

---

*Módulo: 13 - VBA Charts | Archivo: 01 - CellColorstoChart.txt*

# Colores de Celdas a Gráfico

Macro para aplicar los colores de las celdas de origen a los puntos de un gráfico en Excel.

## Código VBA

```vba
Sub ColorChartColumnsbyCellColor()

    Dim xChart As Chart
    Dim I As Long, xRows As Long
    Dim xRg As Range, xCell As Range
    On Error Resume Next
    Set xChart = ActiveSheet.ChartObjects("Chart 1").Chart
    If xChart Is Nothing Then Exit Sub
    With xChart.SeriesCollection(1)
        Set xRg = ActiveSheet.Range(Split(Split(.Formula, ",")(1), "!")(1))
        xRows = xRg.Rows.Count
        Set xRg = xRg(1)
        For I = 1 To xRows
            .Points(I).Format.Fill.ForeColor.RGB = ThisWorkbook.Colors(xRg.Offset(I - 1, 0).Interior.ColorIndex)
        Next
    End With
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función** | Colorea los puntos de la serie del gráfico según el color de fondo de las celdas de origen |
| **Objetivo** | Sincronizar visualmente los colores de celdas con los elementos del gráfico |
| **Requisitos** | Un gráfico llamado "Chart 1" en la hoja activa |
| **Método** | Extrae la fórmula de la serie para obtener el rango de datos y aplica el color de cada celda al punto correspondiente |

### Pasos del proceso

1. Obtiene referencia al gráfico "Chart 1" en la hoja activa
2. Extrae el rango de datos de la primera serie usando la fórmula
3. Recorre cada celda del rango de origen
4. Aplica el color de fondo de cada celda al punto del gráfico correspondiente

---

*Módulo: 13 - VBA Charts | Archivo: 00 - CellsColortoChartsimply.txt*

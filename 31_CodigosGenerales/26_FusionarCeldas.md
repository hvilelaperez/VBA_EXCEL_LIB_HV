# Fusionar Celdas

Obtiene y deshace celdas fusionadas en un rango.

## Código VBA

```vba
    'Deshacer todas las fusiones de celdas
    Dim celda As Range, celdasFusionadas As Range
    For Each celda In rangoUsado
        If celda.MergeCells Then
            Set celdasFusionadas = celda.MergeArea
            celda.MergeCells = False
            celdasFusionadas.Value = celda.Value
        End If
    Next celda


'Esta función retorna todas las celdas fusionadas en un rango.
Function ObtenerCeldasFusionadas(rngFila As Range) As Range

        Dim rngCelda     As Range
        Dim rngFusion    As Range

        For Each rngCelda In rngFila.Cells
            If rngCelda.MergeCells Then

              If rngFusion Is Nothing Then
                Set rngFusion = rngCelda
              Else
                Set rngFusion = Union(rngFusion, rngCelda)
              End If

            End If
        Next

        Set ObtenerCeldasFusionadas = rngFusion

    End Function
Sub test()

    Dim rngTest As Range
    Set rngTest = ObtenerCeldasFusionadas(Selection)

    If Not rngTest Is Nothing Then
        MsgBox "Celdas fusionadas:" & rngTest.Address
    Else
        MsgBox "No hay celdas fusionadas"
    End If

    End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `MergeCells` | Propiedad que indica si la celda está fusionada |
| `MergeArea` | Retorna el rango completo de la fusión |
| `Union` | Combina múltiples rangos en uno solo |

### Código para deshacer fusiones

```vba
' Deshace todas las fusiones preservando valores
For Each celda In rangoUsado
    If celda.MergeCells Then
        Set celdasFusionadas = celda.MergeArea
        celda.MergeCells = False
        celdasFusionadas.Value = celda.Value
    End If
Next celda
```

---

*Módulo: 31_CodigosGenerales | Archivo: 26_FusionarCeldas.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

# Función Celda Máxima

Función que retorna la última celda con datos en un rango o en la hoja activa, o la celda A1 si la hoja está vacía.

## Código VBA

```vba
Public Function GetMaxCell(Optional ByRef rng As Range = Nothing) As Range

    'Retorna la última celda que contiene un valor, o A1 si la hoja está vacía

    Const NONEMPTY As String = "*"
    Dim lRow As Range, lCol As Range

    If rng Is Nothing Then Set rng = Application.ActiveWorkbook.ActiveSheet.UsedRange
    If WorksheetFunction.CountA(rng) = 0 Then
        Set GetMaxCell = rng.Parent.Cells(1, 1)
    Else
        With rng
            Set lRow = .Cells.Find(What:=NONEMPTY, LookIn:=xlFormulas, _
                                        After:=.Cells(1, 1), _
                                        SearchDirection:=xlPrevious, _
                                        SearchOrder:=xlByRows)
            If Not lRow Is Nothing Then
                Set lCol = .Cells.Find(What:=NONEMPTY, LookIn:=xlFormulas, _
                                            After:=.Cells(1, 1), _
                                            SearchDirection:=xlPrevious, _
                                            SearchOrder:=xlByColumns)

                Set GetMaxCell = .Parent.Cells(lRow.Row, lCol.Column)
            End If
        End With
    End If
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `GetMaxCell` | Función que retorna un objeto Range de la última celda con datos |
| `rng` (opcional) | Rango a analizar; si se omite, usa el UsedRange de la hoja activa |
| `NONEMPTY` | Constante "*" que busca cualquier celda con contenido |

| Flujo de ejecución | Descripción |
|--------------------|-------------|
| Si `rng` es Nothing | Usa el rango utilizado de la hoja activa |
| Si `CountA = 0` | La hoja está vacía; retorna celda A1 |
| Si hay datos | Busca la última fila y última columna con datos |
| Resultado | Retorna la intersección de la última fila y última columna |

**Ejemplos de uso:**
```vba
' Obtener la última celda de la hoja activa
Set ultimaCelda = GetMaxCell()

' Obtener la última celda de un rango específico
Set ultimaCelda = GetMaxCell(Range("A1:Z100"))
```

---

*Módulo: 31_CodigosGenerales | Archivo: 11_FuncionCeldaMaxima.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

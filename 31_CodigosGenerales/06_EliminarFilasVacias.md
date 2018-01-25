# Eliminar Filas Vacías

Macro para eliminar filas completamente vacías en un rango seleccionado, y función auxiliar para convertir colecciones a arreglos.

## Código VBA

```vba
Option Explicit

Sub REMOVE_EMPTY_ROW()
    Dim rRange As Range
    Dim firstRow As Long, lastRow As Long, usedRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim firstColTxt As String, lastColTxt As String, usedCol As Long
    Dim i As Long, Target As Range, j As Long
    Set rRange = Application.InputBox(Prompt:="Seleccionar filas a eliminar", Title:="Eliminar filas vacías", Type:=8)

    firstRow = rRange.Row
    usedRow = rRange.Rows.Count
    lastRow = usedRow + firstRow - 1
    firstColTxt = Split(rRange.Address, "$")(1)
    lastColTxt = Split(rRange.Address, "$")(3)
    firstCol = rRange.Column
    usedCol = rRange.Columns.Count
    lastCol = usedCol + firstCol - 1
    Application.ScreenUpdating = False
    For i = lastRow To firstRow Step -1
        For j = firstCol To lastCol
            Set Target = ActiveSheet.Cells(i, j)
            If Application.CountA(Target) = 0 Then
                Target.Delete Shift:=xlUp
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub

Sub COLLECTION_TO_ARRAY()
    Dim coll As New Collection
    Dim arr As Variant
    Dim sh As Worksheet
    Dim rng As Range, rngi As Range, i As Integer

    Set sh = ActiveSheet
    Set rng = sh.Range("A6:A17")
    For Each rngi In rng
        coll.Add rngi.Value
    Next rngi

    'Caso arreglo unidimensional: 1 fila, muchas columnas
    'ReDim arr(1 To coll.Count)
    'For i = 1 To coll.Count
    '    arr(i) = coll.Item(i)
    'Next i
    'sh.Range("B1").Resize(UBound(arr, 1)) = Application.Transpose(arr)

    'Caso arreglo bidimensional: n filas, 1 columna
    ReDim arr(1 To coll.Count, 1 To 1)
    For i = 1 To coll.Count
        arr(i, 1) = coll.Item(i)
    Next i
    sh.Range("B1").Resize(UBound(arr, 1), 1) = arr
End Sub
```

## Descripción

| Macro | Descripción |
|-------|-------------|
| `REMOVE_EMPTY_ROW` | Solicita un rango al usuario y elimina filas vacías desde abajo hacia arriba |
| `COLLECTION_TO_ARRAY` | Convierte una colección de celdas en un arreglo bidimensional y lo escribe en la hoja |

| Elemento técnico | Descripción |
|------------------|-------------|
| `Application.InputBox(..., Type:=8)` | Cuadro de diálogo que permite seleccionar un rango en la hoja |
| `CountA(Target)` | Cuenta celdas con contenido; si retorna 0, la celda está vacía |
| `Target.Delete Shift:=xlUp` | Elimina la celda y desplaza las celdas inferiores hacia arriba |
| `Step -1` | Recorre las filas de abajo hacia arriba para evitar errores de índice |

---

*Módulo: 31_CodigosGenerales | Archivo: 06_EliminarFilasVacias.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

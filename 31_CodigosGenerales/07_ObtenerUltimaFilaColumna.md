# Obtener Última Fila y Columna

Función versátil para obtener la última fila, última columna o última celda con datos en una hoja de cálculo, con múltiples métodos de búsqueda.

## Código VBA

```vba
Function getLastRowCol(choice As Long, sh As Worksheet, Optional InputCol As Variant, Optional InputRw As Long)
' 1 = última fila
' 2 = última columna
' 3 = última celda
' 4 = última fila de una columna dada
' 5 = última columna de una fila dada
    Dim lRw&, lCol&, i&
    Dim s As String, s1 As String

    With sh
        Select Case choice
        Case 1:
            On Error Resume Next
            getLastRowCol = .Cells.Find(What:="*", _
                            After:=.Cells(1, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
            On Error GoTo 0

        Case 2:
            On Error Resume Next
            getLastRowCol = .Cells.Find(What:="*", _
                            After:=.Cells(1, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0

        Case 3:
            On Error Resume Next
            lRw = .Cells.Find(What:="*", _
                           After:=.Cells(1, 1), _
                           LookAt:=xlPart, _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlPrevious, _
                           MatchCase:=False).Row
            lCol = .Cells.Find(What:="*", _
                            After:=.Cells(1, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            getLastRowCol = Cells(lRw, lCol).Address(False, False)
            If Err.Number > 0 Then
                getLastRowCol = .Cells(1, 1).Address(False, False)
                Err.Clear
            End If
            On Error GoTo 0

        Case 4:
            On Error Resume Next
            If Not IsNumeric(InputCol) Then
                InputCol = sh.Columns(InputCol).Column
            End If
            getLastRowCol = .Columns(InputCol).Find(What:="*", _
                            After:=.Cells(1, InputCol), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
            On Error GoTo 0

        Case 5
            On Error Resume Next
            getLastRowCol = .Rows(InputRw).Find(What:="*", _
                            After:=.Cells(InputRw, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0

        End Select
    End With
End Function
```

## Descripción

| Parámetro | Descripción |
|-----------|-------------|
| `choice = 1` | Retorna el número de la última fila con datos |
| `choice = 2` | Retorna el número de la última columna con datos |
| `choice = 3` | Retorna la dirección de la última celda con datos (formato A1) |
| `choice = 4` | Retorna la última fila con datos en una columna específica |
| `choice = 5` | Retorna la última columna con datos en una fila específica |

| Elemento técnico | Descripción |
|------------------|-------------|
| `Find(..., SearchDirection:=xlPrevious)` | Busca desde el final hacia atrás para encontrar el último valor |
| `LookIn:=xlFormulas` | Busca en fórmulas (incluye celdas calculadas) |
| `On Error Resume Next` | Maneja errores cuando la hoja está vacía |
| `InputCol` | Puede recibirse como letra de columna ("A") o número (1) |

---

*Módulo: 31_CodigosGenerales | Archivo: 07_ObtenerUltimaFilaColumna.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

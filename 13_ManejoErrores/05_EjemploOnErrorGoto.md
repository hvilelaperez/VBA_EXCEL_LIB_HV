# Ejemplo de On Error GoTo

Ejemplo práctico de uso de `On Error GoTo` con bucle y manejo de errores.

## Código VBA

```vba
Sub test()

    f = 5

    On Error GoTo message

check:
    Do Until Cells(f, 1).Value = ""

        Cells.Find(what:=refnumber, After:=ActiveCell, LookIn:=xlFormulas, _
              lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
              MatchCase:=False, SearchFormat:=False).Activate
    Loop

    Exit Sub

message:
    MsgBox "Se produjo un error"
    f = f + 1
    GoTo check

End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `test` | Procedimiento que busca un valor en las celdas |
| `On Error GoTo message` | Redirige a la etiqueta `message` si ocurre un error |
| `check` | Etiqueta de inicio del bucle de búsqueda |
| `Cells.Find` | Busca el valor `refnumber` en las celdas |
| `Do Until Cells(f, 1).Value = ""` | Continúa hasta encontrar una celda vacía |
| `message` | Etiqueta de manejo de error que incrementa `f` y reintenta |
| `GoTo check` | Retoma la búsqueda desde la etiqueta `check` |

---

*Módulo: 07 - VBA Error Handler | Archivo: 05 - VD hay su dung on error goto.txt*

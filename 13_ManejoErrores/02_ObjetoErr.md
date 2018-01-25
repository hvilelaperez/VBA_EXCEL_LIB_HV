# Objeto Err

Uso del objeto `Err` para manejar errores al calcular raíces cuadradas.

## Código VBA

```vba
Option Explicit
 
Public Const COL_A = "A"
Public Const COL_B = "B"
 
Sub ErrorHanlding_Demo()
    Dim i As Integer
     
    On Error GoTo InvalidValue
     
    For i = 1 To 5
        Cells(i, COL_B).Value = Sqr(Cells(i, COL_A).Value)
    Next i
     
    Exit Sub
     
InvalidValue:
    MsgBox "Error: " & Err.Number & " : " & Err.Description
    Resume Next
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `Err.Number` | Propiedad que contiene el número del error occurred |
| `Err.Description` | Propiedad que contiene la descripción del error |
| `On Error GoTo InvalidValue` | Redirige el flujo a la etiqueta `InvalidValue` si ocurre un error |
| `Exit Sub` | Sale del procedimiento si no hay errores |
| `Resume Next` | Continúa ejecutando la siguiente línea después del manejo del error |
| `Const COL_A / COL_B` | Constantes públicas que definen las columnas de entrada y salida |

---

*Módulo: 07 - VBA Error Handler | Archivo: 01 - Doi tuong Err.txt*

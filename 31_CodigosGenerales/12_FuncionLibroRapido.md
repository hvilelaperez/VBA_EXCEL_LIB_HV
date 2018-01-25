# Función Libro Rápido

Macros para optimizar el rendimiento de Excel deshabilitando múltiples características de la aplicación y hojas de cálculo durante la ejecución de código.

## Código VBA

```vba
Public Sub FastWB(Optional ByVal opt As Boolean = True)
    With Application
        .Calculation = IIf(opt, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not opt
        .DisplayStatusBar = Not opt
        .EnableAnimations = Not opt
        .EnableEvents = Not opt
        .ScreenUpdating = Not opt
    End With
    FastWS , opt
End Sub

Public Sub FastWS(Optional ByVal ws As Worksheet = Nothing, _
                  Optional ByVal opt As Boolean = True)
    If ws Is Nothing Then
        For Each ws In Application.ActiveWorkbook.Sheets
            EnableWS ws, opt
        Next
    Else
        EnableWS ws, opt
    End If
End Sub

Private Sub EnableWS(ByVal ws As Worksheet, ByVal opt As Boolean)
    With ws
        .DisplayPageBreaks = False
        .EnableCalculation = Not opt
        .EnableFormatConditionsCalculation = Not opt
        .EnablePivotTable = Not opt
    End With
End Sub
```

## Descripción

| Macro | Descripción |
|-------|-------------|
| `FastWB` | Optimiza toda la aplicación Excel (modo rápido o restaurar) |
| `FastWS` | Optimiza todas las hojas o una hoja específica |
| `EnableWS` | Configura las propiedades de optimización de una hoja individual |

| Propiedad | Optimizado (True) | Restaurado (False) |
|-----------|-------------------|---------------------|
| `Calculation` | Manual | Automático |
| `DisplayAlerts` | Desactivado | Activado |
| `DisplayStatusBar` | Desactivado | Activado |
| `EnableAnimations` | Desactivado | Activado |
| `EnableEvents` | Desactivado | Activado |
| `ScreenUpdating` | Desactivado | Activado |
| `DisplayPageBreaks` | Desactivado | - |
| `EnableCalculation` | Desactivado | Activado |
| `EnableFormatConditionsCalculation` | Desactivado | Activado |
| `EnablePivotTable` | Desactivado | Activado |

**Uso recomendado:**
```vba
Sub MiMacroRapida()
    FastWB True      ' Activar modo rápido

    ' ... código de la macro ...

    FastWB False     ' Restaurar configuración normal
End Sub
```

---

*Módulo: 31_CodigosGenerales | Archivo: 12_FuncionLibroRapido.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

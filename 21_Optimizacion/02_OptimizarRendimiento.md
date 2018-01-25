# Optimizar Rendimiento (Stack Overflow)

Solución completa de optimización de rendimiento para VBA en Excel, incluyendo optimización a nivel de aplicación y de hojas de trabajo individuales.

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

### Funciones

| Función | Alcance | Descripción |
|---------|---------|-------------|
| `FastWB` | Público | Optimiza toda la aplicación Excel |
| `FastWS` | Público | Optimiza hojas de trabajo específicas o todas |
| `EnableWS` | Privado | Configura propiedades de una hoja individual |

### Parámetros

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `opt` | Boolean | `True` para optimizar, `False` para restaurar (default: True) |
| `ws` | Worksheet | Hoja a optimizar (opcional, null = todas las hojas) |

### Propiedades optimizadas a nivel de aplicación

| Propiedad | Valor optimizado | Efecto |
|-----------|------------------|--------|
| `Calculation` | Manual | Evita recálculos automáticos |
| `DisplayAlerts` | False | Oculta alertas molestas |
| `DisplayStatusBar` | False | Oculta barra de estado |
| `EnableAnimations` | False | Desactiva animaciones |
| `EnableEvents` | False | Impide ejecución de eventos |
| `ScreenUpdating` | False | Congela la pantalla |

### Propiedades optimizadas a nivel de hoja

| Propiedad | Valor optimizado | Efecto |
|-----------|------------------|--------|
| `DisplayPageBreaks` | False | Oculta saltos de página |
| `EnableCalculation` | False | Desactiva cálculo de la hoja |
| `EnableFormatConditionsCalculation` | False | Desactiva cálculo de formatos condicionales |
| `EnablePivotTable` | False | Desactiva tablas dinámicas |

### Uso

```vba
Sub MiMacroLenta()
    FastWB True          ' Activar optimización
    
    ' ... código que tarda mucho ...
    
    FastWB False         ' Restaurar configuración
End Sub
```

---

*Módulo: 15 - VBA Optimize | Archivo: 01 - Optimize - StackoverFlow.txt*

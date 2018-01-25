# Optimizar Código VBA

Macros para optimizar el rendimiento del código VBA en Excel, desactivando temporalmente cálculos, eventos y actualizaciones de pantalla.

## Código VBA

```vba
Option Explicit

Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimzeCode_Begin()
'=============================================================='
'Habilitar modo Optimizar
'=============================================================='

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

'=============================================================='

End Sub


Sub OptimizeCode_End()

'=============================================================='
'Restaurar valores predeterminados después del modo Optimizar
'=============================================================='

    ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    
End Sub
```

## Descripción

### Funciones

| Función | Descripción |
|---------|-------------|
| `OptimzeCode_Begin` | Desactiva elementos que ralentizan la ejecución |
| `OptimizeCode_End` | Restaura la configuración predeterminada de Excel |

### Variables globales

| Variable | Tipo | Descripción |
|----------|------|-------------|
| `CalcState` | Long | Almacena el estado del cálculo automático |
| `EventState` | Boolean | Almacena el estado de los eventos |
| `PageBreakState` | Boolean | Almacena el estado de los saltos de página |

### Qué se desactiva

| Propiedad | Efecto |
|-----------|--------|
| `ScreenUpdating` | Evita actualizaciones de pantalla durante la ejecución |
| `EnableEvents` | Impide que se ejecuten eventos de Excel |
| `Calculation` | Cambia a cálculo manual para evitar recálculos |
| `DisplayPageBreaks` | Oculta los saltos de página para mayor velocidad |

### Uso recomendado

```vba
Sub MiMacro()
    OptimzeCode_Begin
    
    ' ... código optimizado aquí ...
    
    OptimizeCode_End
End Sub
```

---

*Módulo: 15 - VBA Optimize | Archivo: 00 - OptimizeCode - HEO.txt*

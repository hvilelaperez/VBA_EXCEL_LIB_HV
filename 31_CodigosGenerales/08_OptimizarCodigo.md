# Optimizar Código VBA

Macros para optimizar el rendimiento de macros VBA deshabilitando actualizaciones de pantalla, eventos y cálculos durante la ejecución.

## Código VBA

```vba
Option Explicit

Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimzeCode_Begin()
    '=============================================================='
    'Activar modo Optimizar
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
    'Restaurar configuración predeterminada después del modo Optimizar
    '=============================================================='

    ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
```

## Descripción

| Macro | Descripción |
|-------|-------------|
| `OptimzeCode_Begin` | Desactiva funcionalidades pesadas antes de ejecutar código |
| `OptimizeCode_End` | Restaura todas las configuraciones al estado original |

| Variable/Propósito | Descripción |
|--------------------|-------------|
| `ScreenUpdating` | Desactiva la actualización visual de la pantalla |
| `EnableEvents` | Desactiva la ejecución de eventos de Excel |
| `Calculation` | Cambia a cálculo manual (evita recálculos automáticos) |
| `DisplayPageBreaks` | Oculta los saltos de página para mayor velocidad |

**Uso recomendado:**
```vba
Sub MiMacroOptimizada()
    OptimzeCode_Begin

    ' ... código de la macro ...

    OptimizeCode_End
End Sub
```

---

*Módulo: 31_CodigosGenerales | Archivo: 08_OptimizarCodigo.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*

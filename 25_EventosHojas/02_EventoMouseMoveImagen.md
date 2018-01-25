# Evento MouseMove para Efecto Visual en Imágenes

Macros VBA de eventos `MouseMove` en un UserForm que aplican efectos visuales de elevación a imágenes cuando el cursor del mouse pasa sobre ellas, restaurándose al estado normal al salir.

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
Option Explicit

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image1
        .SpecialEffect = fmSpecialEffectRaised
        .Left = 103
        .Top = 19
    End With
End Sub

Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image2
        .SpecialEffect = fmSpecialEffectRaised
        .Left = 103
        .Top = 85
    End With
End Sub

Private Sub Image3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image3
        .SpecialEffect = fmSpecialEffectRaised
        .Left = 103
        .Top = 151
    End With
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image1
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 102
        .Top = 18
    End With
    
    With Image2
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 102
        .Top = 84
    End With
    
    With Image3
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 102
        .Top = 150
    End With
End Sub


Private Sub cbClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 1 Then
        'No hacer nada
    Else
        Cancel = True
        MsgBox "¡No se puede cerrar de esta manera!", vbOKOnly + vbCritical
    End If
End Sub
```

## Descripción

| Evento | Descripción |
|---|---|
| `Image1_MouseMove` | Eleva visualmente la imagen 1 al pasar el mouse |
| `Image2_MouseMove` | Eleva visualmente la imagen 2 al pasar el mouse |
| `Image3_MouseMove` | Eleva visualmente la imagen 3 al pasar el mouse |
| `UserForm_MouseMove` | Restaura todas las imágenes al estado plano cuando el mouse sale de ellas |
| `cbClose_Click` | Cierra el formulario |
| `UserForm_QueryClose` | Controla el cierre, solo permite cerrar con el botón |

### Efecto Visual Implementado

| Estado | `SpecialEffect` | Posición | Sensación |
|---|---|---|---|
| **Normal** (fuera) | `fmSpecialEffectFlat` | Posición base | Plano |
| **Hover** (encima) | `fmSpecialEffectRaised` | Desplazado (+1 px) | Elevado/botón presionado |

### Puntos Clave

- `fmSpecialEffectRaised` crea un efecto 3D de elevación en el control
- El desplazamiento de 1 pixel (`Left + 1`, `Top + 1`) refuerza la ilusión de elevación
- El evento `MouseMove` del UserForm actúa como "reset" para todas las imágenes
- `CloseMode = 1` corresponde a `vbFormControlMenu` (cierre desde la X)

---

*Módulo: 20 - VBA Worksheets Events | Archivo: 01 - Event mouse move effect image*

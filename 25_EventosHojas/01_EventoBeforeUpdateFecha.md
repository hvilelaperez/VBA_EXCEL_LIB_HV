# Evento BeforeUpdate para Validación de Fechas y Horas

Macros VBA para validar y formatear entradas de fecha y hora en cuadros de texto de formularios, asegurando el formato correcto antes de actualizar el valor.

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
Private Sub tbxStart_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Me.tbxStart Like "??:??" Then
        MsgBox "Por favor use el formato 'hh:mm'"
        Cancel = True
        Exit Sub
    End If
    
    myVar = Application.WorksheetFunction.Text(Me.tbxStart, "hh:mm am")
    Me.tbxStart = myVar
End Sub


Private Sub tbxEndDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    tbxEndDate.Value = CDate(tbxEndDate.Value)
End Sub


Private Sub tbxStartDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    tbxStartDate.Value = CDate(tbxStartDate.Value)
End Sub
```

## Descripción

| Sub/Evento | Control | Descripción |
|---|---|---|
| `tbxStart_BeforeUpdate` | Cuadro de texto de hora | Valida formato `hh:mm` y formatea con AM/PM |
| `tbxEndDate_BeforeUpdate` | Cuadro de texto de fecha fin | Convierte el valor a tipo Fecha |
| `tbxStartDate_BeforeUpdate` | Cuadro de texto de fecha inicio | Convierte el valor a tipo Fecha |

### Flujo de Validación de Hora (`tbxStart`)

1. Verifica si el valor coincide con el patrón `"??:??"` (exactamente 2 dígitos, dos puntos, 2 dígitos)
2. Si no coincide: muestra mensaje de error y cancela la actualización
3. Si coincide: convierte al formato `hh:mm AM/PM` usando `WorksheetFunction.Text`

### Flujo de Validación de Fechas

1. Intenta convertir el valor del cuadro de texto a tipo `Date` usando `CDate()`
2. Si la conversión falla, el error se ignora (`On Error Resume Next`) y el valor permanece sin cambios

### Puntos Clave

- El evento `BeforeUpdate` se ejecuta **antes** de que el valor se guarde en el control
- `Cancel = True` cancela la actualización y mantiene el valor anterior
- `Like "??:??"` usa comodines para validar la estructura de la hora
- `CDate()` es flexible: acepta múltiples formatos de fecha según la configuración regional

---

*Módulo: 20 - VBA Worksheets Events | Archivo: 00 - Event Before Update cho ngay thang*

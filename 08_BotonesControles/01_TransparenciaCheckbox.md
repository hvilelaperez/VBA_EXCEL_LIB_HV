# Transparencia de Checkbox

Código VBA para controlar la transparencia y apariencia de controles CheckBox en hojas de cálculo Excel.

## Código VBA

```vba
Option Explicit

Sub checkBox_Transapency()
Dim chk As CheckBox
'Llamar OptimzeCode_Begin
    For Each chk In ActiveSheet.CheckBoxes
        With chk.Fill
            If chk.Value = True Then
                .Visible = msoTrue
                '.ForeColor.RGB = RGB(255, 0, 0)
                .BackColor = RGB(0, 0, 0)
                
            Else
                .Visible = msoTrue
                '.ForeColor.RGB = RGB(128, 128, 128)
                .BackColor = RGB(128, 128, 128)
            End If
        End With
    Debug.Print chk.Name
    Next chk

'Llamar OptimizeCode_End
End Sub


Sub checkBox_Transapency1()
Dim chk As Object
Dim nameChk As String
Dim ws As Worksheet
Set ws = ActiveSheet
For Each chk In ws.CheckBoxes
        Debug.Print "---------------------------------------------"
        Debug.Print "Control de formulario en hoja: " & ws.Name
        Debug.Print "Ubicación en hoja: " & chk.TopLeftCell.Address
        Debug.Print "Nombre del componente: " & chk.Name
        Debug.Print "Tipo de objeto: CheckBox"
Next chk
End Sub

Sub CheckboxLoop()
'PROPÓSITO: Recorrer cada CheckBox de Control de Formulario en la hoja activa
'FUENTE: www.TheSpreadsheetGuru.com/the-code-vault

Dim cb As Shape

'Recorrer los CheckBox
  For Each cb In ActiveSheet.Shapes
    If cb.Type = msoFormControl Then
      If cb.FormControlType = xlCheckBox Then
        If cb.ControlFormat.Value = 1 Then
          'Hacer algo si está marcado...
          cb.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ElseIf cb.ControlFormat.Value = -4146 Then
          'Hacer algo si no está marcado...
          cb.Fill.ForeColor.RGB = RGB(128, 128, 128)
        ElseIf cb.ControlFormat.Value = 2
          'Hacer algo si está en estado mixto...
          cb.Fill.ForeColor.RGB = RGB(128, 128, 128)
        End If
      End If
    End If
  Next cb

End Sub
```

## Descripción

| Subrutina | Función |
|-----------|---------|
| `checkBox_Transapency` | Cambia el color de fondo de CheckBox según su estado (marcado/desmarcado) |
| `checkBox_Transapency1` | Lista información detallada de cada CheckBox en la hoja activa |
| `CheckboxLoop` | Recorre todos los CheckBox del formulario y aplica colores según estado |

### Estados del CheckBox

| Valor | Significado | Color |
|-------|-------------|-------|
| `1` | Marcado | Negro `RGB(0, 0, 0)` |
| `-4146` | Desmarcado | Gris `RGB(128, 128, 128)` |
| `2` | Mixto | Gris `RGB(128, 128, 128)` |

---

*Módulo: 02 - VBA Button, checkbox, canvas | Archivo: 00 - Checkbox tranparency*
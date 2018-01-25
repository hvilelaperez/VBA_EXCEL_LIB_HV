# Vista Previa de Impresión

Macro para mostrar la vista previa de impresión de una hoja específica del libro.

## Código VBA

```vba
Private Sub cbVistaPrevia_Click()
    Dim shtMenu As Worksheet

    On Error GoTo hojaInvalida
    Set shtModelo = ThisWorkbook.Sheets("02-MODEL")
    With shtModelo
        If Not IsVisibleSheet(.Range("A1")) Then
        .Visible = True
        End If
        Me.Hide
        Application.Goto Reference:="Modelo_Factura"
        ActiveWindow.SelectedSheets.PrintPreview
        Me.Show
'        .PrintPreview
'        .Visible = xlHidden
    End With
hojaInvalida:
    If Err.Number = 9 Then
        Set shtModelo = shModelo
        Resume Next
    End If
    
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `cbVistaPrevia_Click` | Evento de clic del botón de vista previa |
| `shtModelo` | Variable para la hoja modelo |
| `IsVisibleSheet` | Verifica si la hoja es visible |
| `Application.Goto` | Navega a la referencia nombrada "Modelo_Factura" |
| `PrintPreview` | Muestra la vista previa de impresión |
| `Me.Hide / Me.Show` | Oculta/muestra el UserForm durante la vista previa |

### Flujo de ejecución

1. Se intenta acceder a la hoja "02-MODEL"
2. Si la hoja no es visible, se hace visible
3. Se oculta el formulario actual
4. Se navega a la referencia "Modelo_Factura"
5. Se muestra la vista previa de impresión
6. Se muestra el formulario nuevamente

### Manejo de errores

| Error | Código | Descripción |
|-------|--------|-------------|
| Hoja no encontrada | 9 | Si la hoja "02-MODEL" no existe, se usa la variable `shModelo` como respaldo |

### Nota

> La función `IsVisibleSheet` debe estar definida en el proyecto para verificar la visibilidad de la hoja.

---

*Módulo: 96 - VBA Codes general | Archivo: 36 - PintPreview.txt*

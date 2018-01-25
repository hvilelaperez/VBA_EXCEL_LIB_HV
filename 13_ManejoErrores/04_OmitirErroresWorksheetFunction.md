# Omitir Errores con WorksheetFunction

Ejemplo de manejo de errores al usar `WorksheetFunction` con `On Error Resume Next`.

## Código VBA

```vba
Private Sub tbName_Change()
	On Error Resume Next
	Me.lblDesc = Application.WorksheetFunction.VLookup(Me.tbName, ThisWorkbook.Sheets(...))

	If Err.Number <> 0 Then
		Me.lblDesc = ""
	End If
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `tbName_Change` | Procedimiento que se ejecuta al cambiar el cuadro de texto |
| `On Error Resume Next` | Omite errores y continúa con la siguiente línea |
| `WorksheetFunction.VLookup` | Busca un valor en la primera columna de un rango |
| `Err.Number <> 0` | Verifica si ocurrió un error después de la operación |
| `lblDesc = ""` | Limpia la etiqueta si el `VLookup` no encontró resultados |

---

*Módulo: 07 - VBA Error Handler | Archivo: 03 - Bo qua loi khi su dung worksheetfunction trong vba.txt*

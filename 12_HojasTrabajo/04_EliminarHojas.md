# Eliminar Hojas de Trabajo

Procedimiento para eliminar todas las hojas excepto una específica (`Feuil1`).

## Código VBA

```vba
Sub DELETE_SHEETS()
    Dim sht As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each sht In Worksheets
        If sht.Name <> "Feuil1" Then sht.Delete
    Next
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `DELETE_SHEETS` | Procedimiento que elimina todas las hojas excepto `Feuil1` |
| `Application.ScreenUpdating = False` | Desactiva la actualización de pantalla para mayor eficiencia |
| `Application.DisplayAlerts = False` | Desactiva las alertas para evitar confirmaciones de eliminación |
| `sht.Name <> "Feuil1"` | Condición que protege la hoja `Feuil1` de ser eliminada |
| Restauración | Se reactivan `ScreenUpdating` y `DisplayAlerts` al finalizar |

---

*Módulo: 06 - VBA Worksheets | Archivo: 05 - Delete Sheets.txt*

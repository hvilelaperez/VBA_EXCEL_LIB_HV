# Prohibir Eliminar Columna o Celdas

Macros para prevenir la eliminación de contenido en rangos específicos de la hoja de cálculo.

## Código VBA

```vba
Private Sub Worksheet_Change(ByVal target As Range)
    If Application.EnableEvents = False Then Application.EnableEvents = True
       If Not Intersect(target, Range("A1:A65356")) Is Nothing And target.Rows.Count <> ActiveSheet.Rows.Count Then
        Dim rngi As Range
        For Each rngi In target
            Select Case LCase(rngi.Value)
                    Case "w"
                        rngi.Resize(, 6).Interior.ColorIndex = 3
                    Case "a"
                        rngi.Resize(, 6).Interior.ColorIndex = 4
                    Case Else
                        rngi.Resize(, 6).Interior.ColorIndex = xlColorIndexNone
            End Select
        Next rngi
    
    
    ElseIf Not Intersect(target, Range("A1:A65356")) Is Nothing And target.Rows.Count = ActiveSheet.Rows.Count Then
        On Error GoTo Salida
        Application.EnableEvents = False
        If Application.WorksheetFunction.CountA(target) = 0 Then
            Application.Undo
            MsgBox " No puede eliminar todo el contenido de las celdas de este rango " _
            , vbCritical, "Kutools for Excel"
        End If
    End If
Salida:
        Application.EnableEvents = True
End Sub


'No permite eliminar celdas especiales
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("A1:E7")) Is Nothing Then Exit Sub
    On Error GoTo Salida
    Application.EnableEvents = False
    If Not IsDate(Target(1)) Then
        Application.Undo
        MsgBox " No puede eliminar el contenido de las celdas de este rango " _
        , vbCritical, "Kutools for Excel"
    End If
Salida:
    Application.EnableEvents = True
End Sub
```

## Descripción

### Primer procedimiento: Cambio de color y protección

| Elemento | Descripción |
|----------|-------------|
| `Worksheet_Change` | Evento que se ejecuta al modificar cualquier celda |
| `Intersect` | Verifica si el cambio ocurrió en el rango A1:A65356 |
| `target.Rows.Count` | Cuenta las filas modificadas |
| `LCase(rngi.Value)` | Convierte el valor a minúsculas para la comparación |

| Valor | Color aplicado | ColorIndex |
|-------|----------------|------------|
| "w" | Rojo | 3 |
| "a" | Verde | 4 |
| Otro | Sin color | xlColorIndexNone |

### Segundo procedimiento: Protección contra eliminación

| Elemento | Descripción |
|----------|-------------|
| `Intersect` | Verifica si el cambio ocurrió en el rango A1:E7 |
| `IsDate` | Verifica si la celda contiene una fecha |
| `Application.Undo` | Revierte el cambio si no es válido |
| `MsgBox` | Muestra un mensaje de error al usuario |

### Lógica de protección

1. Se detecta un cambio en la hoja
2. Se verifica si el cambio está dentro del rango protegido
3. Si se intenta eliminar contenido, se revierte la acción con `Application.Undo`
4. Se muestra un mensaje de error al usuario

### Nota importante

> `Application.EnableEvents = False/True` se usa para evitar que el evento `Worksheet_Change` se ejecute recursivamente al modificar la hoja programáticamente.

---

*Módulo: 96 - VBA Codes general | Archivo: 31 - Khong cho phep xoa ca mot cot.txt*
